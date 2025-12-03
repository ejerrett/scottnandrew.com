<?php
namespace Grav\Plugin;

use Grav\Common\Page\PageInterface;
use Grav\Common\Plugin;
use PhpOffice\PhpWord\IOFactory;

class CvdocxPlugin extends Plugin
{
    public static function getSubscribedEvents(): array
    {
        return [
            'onTwigInitialized' => ['onTwigInitialized', 0],
        ];
    }

    public function onTwigInitialized(): void
    {
        // Ensure composer deps (PhpWord) are available
        $autoload = __DIR__ . '/vendor/autoload.php';
        if (is_file($autoload)) {
            require_once $autoload;
        } else {
            $this->grav['log']->warning('cvdocx: vendor/autoload.php missing; run composer install in user/plugins/cvdocx');
        }

        $twig = $this->grav['twig']->twig();

        // {{ docx_html(page)|raw }} or {{ docx_html(page, 'cv.docx')|raw }}
        $twig->addFunction(new \Twig\TwigFunction('docx_html', function ($page, ?string $filename = null) {
            return $this->renderDocxHtml($page, $filename);
        }, ['is_safe' => ['html']]));

        // {{ cvdocx()|raw }} or {{ cvdocx('cv.docx')|raw }} (uses current page)
        $twig->addFunction(new \Twig\TwigFunction('cvdocx', function (?string $filename = null) {
            $page = $this->grav['page'];
            return $this->renderDocxHtml($page, $filename);
        }, ['is_safe' => ['html']]));
    }

    /**
     * Render DOCX to sanitized HTML with caching.
     *
     * @param PageInterface|\Grav\Common\Page\Page $page
     * @param string|null $filename Specific file or null to auto-pick newest .docx
     */
    // keep this one as your main entry point (unchanged signature)
    private function renderDocxHtml($page, ?string $filename): string
    {
        // 1) resolve a path (supports explicit filename or newest .docx)
        $path = $this->resolveDocxPath($page, $filename);
        if (!$path) {
            return '<div class="cvdocx-msg">No DOCX found for this page.</div>';
        }

        // 2) cache: invalidate when file mtime changes
        $key = $this->cacheKey($path);
        $cache = $this->grav['cache'];
        if ($cached = $cache->get($key)) {
            return $cached;
        }

        // 3) render the file to HTML (helper below) and cache it
        $html = $this->renderDocxToHtml($path);
        $cache->save($key, $html);

        return $html;
    }

    // helper: turn a resolved filesystem path into clean, scoped HTML
    private function renderDocxToHtml(string $path): string
    {
        $phpWord = \PhpOffice\PhpWord\IOFactory::load($path, 'Word2007');
        $writer = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');

        ob_start();
        $writer->save('php://output');
        $html = ob_get_clean();

        // keep only <body>…</body>
        $html = preg_replace('/^.*?<body[^>]*>/is', '', $html);
        $html = preg_replace('/<\/body>.*$/is', '', $html);

        // remove inlined CSS/links that PhpWord emits
        $html = preg_replace('#<style[^>]*>.*?</style>#is', '', $html);
        $html = preg_replace('#<link[^>]*>#is', '', $html);

        // tidy spacing
        $html = preg_replace('#(?:<br\s*/?>\s*){3,}#i', '<br><br>', $html);
        $html = preg_replace('#<p[^>]*>\s*(?:&nbsp;|\s)*</p>#i', '', $html);

        return '<div class="cvdocx-scope">' . $html . '</div>';
    }

    // cache key that changes when the DOCX does
    private function cacheKey(string $path): string
    {
        $mt = @filemtime($path) ?: 0;
        return 'cvdocx:' . md5($path . ':' . $mt);
    }


    /**
     * If $filename provided, resolve that; otherwise pick the newest *.docx in the page folder.
     */
    private function resolveDocxPath($page, ?string $filename): ?string
    {
        // Prefer page->path(), fall back to dirname(filePath()).
        $pageDir = '';
        try {
            if (method_exists($page, 'path')) {
                $pageDir = rtrim($page->path(), '/'); // folder containing the .md
            }
            if (!$pageDir && method_exists($page, 'filePath') && $page->filePath()) {
                $pageDir = rtrim(dirname($page->filePath()), '/');
            }
        } catch (\Throwable $e) {
            $pageDir = '';
        }
        if (!$pageDir || !is_dir($pageDir)) {
            throw new \RuntimeException('cvdocx: could not resolve page directory');
        }

        if ($filename) {
            $path = $pageDir . '/' . ltrim($filename, '/');
            return is_file($path) ? $path : null;
        }

        // Auto-pick newest .docx (ignore hidden/temp/lock files)
        $candidates = glob($pageDir . '/*.docx') ?: [];
        $candidates = array_values(array_filter($candidates, function ($p) {
            $bn = basename($p);
            if ($bn === '' || $bn[0] === '.')
                return false;                   // hidden
            if (preg_match('/(^~\$|~$|\#|\$)/', $bn))
                return false;           // temp/lock
            return is_file($p);
        }));

        if (!$candidates) {
            return null;
        }

        usort($candidates, function ($a, $b) {
            return (@filemtime($b) <=> @filemtime($a));
        });
        return $candidates[0] ?? null;
    }

    private function sanitizeCvHtml(string $html): string
    {
        // Allow basic structure + tables + links
        $allowed_tags = [
            'h1',
            'h2',
            'h3',
            'h4',
            'h5',
            'h6',
            'p',
            'div',
            'span',
            'br',
            'hr',
            'strong',
            'b',
            'em',
            'i',
            'u',
            'sup',
            'sub',
            'ul',
            'ol',
            'li',
            'table',
            'thead',
            'tbody',
            'tr',
            'th',
            'td',
            'a'
        ];

        // Strip all tags except the above
        $html = strip_tags($html, '<' . implode('><', $allowed_tags) . '>');

        // Whitelist a few safe attributes (href, colspan/rowspan, align)
        // Remove everything else (especially inline styles from Word)
        $html = preg_replace_callback(
            '#<([a-z0-9]+)\b([^>]*)>#i',
            function ($m) {
                $tag = strtolower($m[1]);
                $attr = $m[2];

                // Keep only permitted attributes per tag
                $keep = [];
                if ($tag === 'a') {
                    if (preg_match('#\bhref=("|\')(.*?)\1#i', $attr, $mm)) {
                        $href = htmlspecialchars($mm[2], ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
                        $keep[] = 'href="' . $href . '"';
                        $keep[] = 'target="_blank"';
                        $keep[] = 'rel="noopener"';
                    }
                } elseif (in_array($tag, ['td', 'th'], true)) {
                    foreach (['colspan', 'rowspan'] as $k) {
                        if (preg_match('#\b' . $k . '=("|\')(\d+)\1#i', $attr, $mm)) {
                            $keep[] = strtolower($k) . '="' . $mm[2] . '"';
                        }
                    }
                } elseif (in_array($tag, ['p', 'div', 'td', 'th'], true)) {
                    if (preg_match('#\balign=("|\')(left|right|center)\1#i', $attr, $mm)) {
                        $keep[] = 'data-align="' . $mm[2] . '"';
                    }
                }
                return '<' . $tag . ($keep ? ' ' . implode(' ', $keep) : '') . '>';
            },
            $html
        );

        return $html;
    }

}
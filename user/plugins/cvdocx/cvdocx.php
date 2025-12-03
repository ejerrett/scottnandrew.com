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
    private function renderDocxHtml($page, ?string $filename): string
    {
        try {
            $path = $this->resolveDocxPath($page, $filename);
            if (!$path) {
                return '<div class="cvdocx-error">No DOCX found for this page.</div>';
            }

            // Build a cache key that changes when the chosen file changes
            $key = $this->cacheKey($page, $path);

            $cache = $this->grav['cache'];
            $cached = $cache->fetch($key);
            if (is_string($cached) && $cached !== '') {
                return $cached;
            }

            // Load and convert to HTML
            $phpword = IOFactory::load($path);
            $writer = IOFactory::createWriter($phpword, 'HTML');
            ob_start();
            $writer->save('php://output');
            $html = ob_get_clean() ?: '';

            // Sanitize Wordy HTML so it doesn't fight your dark theme
            $html = $this->sanitizeDocxHtml($html);

            // Wrap for scoping
            $htmlWrapped = '<div class="cvdocx">' . $html . '</div>';
            $cache->save($key, $htmlWrapped);

            return $htmlWrapped;
        } catch (\Throwable $e) {
            $this->grav['log']->error('cvdocx render error: ' . $e->getMessage());
            return '<div class="cvdocx-error">Unable to render CV.</div>';
        }
    }

    /**
     * If $filename provided, resolve that; otherwise pick the newest *.docx in the page folder.
     */
    private function resolveDocxPath($page, ?string $filename): ?string
    {
        $pageDir = dirname($page->filePath());

        if ($filename) {
            $path = $pageDir . '/' . ltrim($filename, '/');
            return is_file($path) ? $path : null;
        }

        // Auto-pick newest .docx (ignore temp/hidden files)
        $candidates = glob($pageDir . '/*.docx') ?: [];
        $candidates = array_filter($candidates, function ($p) {
            $bn = basename($p);
            if ($bn[0] === '.')
                return false;              // hidden
            if (preg_match('/(~|\#|\$|^~\$)/', $bn))
                return false; // temp/lock files
            return is_file($p);
        });

        if (!$candidates) {
            return null;
        }

        usort($candidates, function ($a, $b) {
            return (@filemtime($b) <=> @filemtime($a)); // newest first
        });

        return $candidates[0] ?? null;
    }

    private function cacheKey($page, string $path): string
    {
        $route = method_exists($page, 'route') ? $page->route() : (string) $page;
        $mtime = @filemtime($path) ?: 0;
        return 'cvdocx:' . md5($route . '|' . $path . '|' . $mtime);
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
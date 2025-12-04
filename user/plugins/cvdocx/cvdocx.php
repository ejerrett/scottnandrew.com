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

        // Grav returns false on cache miss
        $cached = $cache->fetch($key);
        if ($cached !== false) {
            return $cached;
        }

        // 3) render and cache
        $html = $this->renderDocxToHtml($path);
        $html = $this->tidyCvHtml($html);
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

    /**
     * Normalize PhpWord HTML: remove inline colors/fonts, unwrap spans,
     * and add an indent class to paragraphs that DON'T start with a year.
     */
    private function tidyCvHtml(string $html): string
    {
        // Drop any embedded <style> blocks from PhpWord
        $html = preg_replace('~<style\b[^>]*>.*?</style>~is', '', $html);

        // Remove inline CSS that forces black text / fonts from Word
        $html = preg_replace('~\s*color\s*:\s*#[0-9a-fA-F]{3,6}\s*;?~i', '', $html);
        $html = preg_replace('~\s*font-family\s*:\s*[^;"]+;?~i', '', $html);
        $html = preg_replace('~\s*font-size\s*:\s*[^;"]+;?~i', '', $html);

        // Clean empty style=""
        $html = preg_replace('~\sstyle="(\s*;?\s*)*"~i', '', $html);

        // Unwrap spans so we’re not fighting span soup
        $html = preg_replace('~</?span\b[^>]*>~i', '', $html);

        // Add indent class to <p> that do NOT begin with a year (e.g., "2024", "2019–")
        // Leave true year lines alone.
        $html = preg_replace_callback('~<p([^>]*)>(.*?)</p>~is', function ($m) {
            $text = trim(strip_tags($m[2]));
            // match 4-digit year at start, optionally followed by dash/en-dash/em-dash/space
            if ($text === '' || preg_match('~^(19|20)\d{2}(\s|–|-|—)~u', $text)) {
                return "<p{$m[1]}>{$m[2]}</p>";
            }
            return "<p class=\"cv-cont\"{$m[1]}>{$m[2]}</p>";
        }, $html);

        return $html;
    }


}
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
            $writer  = IOFactory::createWriter($phpword, 'HTML');
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
            if ($bn[0] === '.') return false;              // hidden
            if (preg_match('/(~|\#|\$|^~\$)/', $bn)) return false; // temp/lock files
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
        $route = method_exists($page, 'route') ? $page->route() : (string)$page;
        $mtime = @filemtime($path) ?: 0;
        return 'cvdocx:' . md5($route . '|' . $path . '|' . $mtime);
    }

    private function sanitizeDocxHtml(string $html): string
    {
        // Remove <style>…</style>
        $html = preg_replace('#<style\b[^>]*>.*?</style>#is', '', $html);

        // Strip inline color/background styles so your dark theme rules win
        $html = preg_replace_callback(
            '#\sstyle="([^"]*)"#i',
            function ($m) {
                $style = $m[1];

                // remove color & background / background-color
                $style = preg_replace('/(^|;)\s*(color|background(?:-color)?)\s*:[^;"]*/i', '', $style);

                // Tidy
                $style = trim(preg_replace('/;{2,}/', ';', $style), " ;");

                return $style ? ' style="' . $style . '"' : '';
            },
            $html
        );

        // Remove <body> wrapper if present
        $html = preg_replace('#</?body[^>]*>#i', '', $html);

        return $html;
    }
}

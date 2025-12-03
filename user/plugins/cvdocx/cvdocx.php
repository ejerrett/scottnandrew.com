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

        // Your existing call style: {{ docx_html(page, 'cv.docx')|raw }}
        $twig->addFunction(new \Twig\TwigFunction('docx_html', function ($page, string $filename) {
            return $this->renderDocxHtml($page, $filename);
        }, ['is_safe' => ['html']]));

        // Convenience: {{ cvdocx('cv.docx')|raw }} using current page
        $twig->addFunction(new \Twig\TwigFunction('cvdocx', function (string $filename) {
            $page = $this->grav['page'];
            return $this->renderDocxHtml($page, $filename);
        }, ['is_safe' => ['html']]));
    }

    /**
     * Render DOCX to sanitized HTML with caching.
     *
     * @param PageInterface|\Grav\Common\Page\Page $page
     * @param string $filename
     * @return string HTML
     */
    private function renderDocxHtml($page, string $filename): string
    {
        try {
            // Resolve the file in the current page folder
            $pageDir = dirname($page->filePath());
            $path = $pageDir . '/' . ltrim($filename, '/');

            if (!is_file($path)) {
                return '<div class="cvdocx-error">CV file not found: ' . htmlspecialchars($filename) . '</div>';
            }

            // Build a cache key that changes when the file changes
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

            // Sanitize (strip style blocks + inline color/background + body wrapper)
            $html = $this->sanitizeDocxHtml($html);

            // Wrap and store
            $htmlWrapped = '<div class="cvdocx">' . $html . '</div>';
            $cache->save($key, $htmlWrapped);

            return $htmlWrapped;
        } catch (\Throwable $e) {
            $this->grav['log']->error('cvdocx render error: ' . $e->getMessage());
            return '<div class="cvdocx-error">Unable to render CV.</div>';
        }
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

        // Strip inline color/background styles
        $html = preg_replace_callback(
            '#\sstyle="([^"]*)"#i',
            function ($m) {
                $style = $m[1];

                // remove color & background / background-color declarations
                $style = preg_replace('/(^|;)\s*(color|background(?:-color)?)\s*:[^;"]*/i', '', $style);

                // optionally strip font-family/size as well (uncomment if needed)
                $style = preg_replace('/(^|;)\s*font-family\s*:[^;"]*/i', '', $style);
                $style = preg_replace('/(^|;)\s*font-size\s*:[^;"]*/i', '', $style);

                // clean up
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

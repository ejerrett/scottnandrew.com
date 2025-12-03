<?php
namespace Grav\Plugin;

$autoload = __DIR__ . '/vendor/autoload.php';
if (is_file($autoload)) {
    require_once $autoload;
} else {
    // Optional: Fail gracefully in the page instead of a fatal error
    if (isset($this->grav)) {
        $this->grav['log']->warning('cvdocx: vendor/autoload.php missing; run composer install in user/plugins/cvdocx');
    }
}

use Grav\Common\Grav;
use Grav\Common\Plugin;
use Grav\Common\Page\Page;
use PhpOffice\PhpWord\IOFactory;

class CvdocxPlugin extends Plugin
{
    public static function getSubscribedEvents(): array
    {
        return [
            'onTwigExtensions' => ['onTwigExtensions', 0],
        ];
    }

    public function onTwigExtensions(): void
    {
        $this->grav['twig']->twig()->addFunction(
            new \Twig\TwigFunction('docx_html', function ($page, $filename = null) {
                return $this->renderDocx($page, $filename);
            })
        );
    }

    private function renderDocx($page, $filename)
    {
        // Resolve page
        if (!($page instanceof Page)) {
            $page = Grav::instance()['page'];
        }
        if (!$page) { return ''; }

        // Pick first .docx if not specified
        $media = $page->media()->all();
        $docxPath = null;
        if ($filename && isset($media[$filename])) {
            $docxPath = $media[$filename]->path();
        } else {
            foreach ($media as $item) {
                if (str_ends_with(strtolower($item->path()), '.docx')) {
                    $docxPath = $item->path();
                    break;
                }
            }
        }
        if (!$docxPath || !is_file($docxPath)) {
            return ''; // nothing to render
        }

        // Cache location keyed by file mtime
        $mtime = filemtime($docxPath) ?: time();
        $cacheDir = Grav::instance()['locator']->findResource('user-data://cvdocx', true, true);
        if (!is_dir($cacheDir)) {
            mkdir($cacheDir, 0775, true);
        }
        $cacheFile = $cacheDir . '/' . $page->id() . '-' . $mtime . '.html';

        if (!is_file($cacheFile)) {
            // Convert with PhpWord
            $phpWord = IOFactory::load($docxPath);
            $writer = IOFactory::createWriter($phpWord, 'HTML');

            // Capture HTML
            ob_start();
            $writer->save('php://output');
            $html = ob_get_clean();

            // Lightweight cleanups (optional)
            // Normalize headings/lists a bit:
            $style = <<<CSS
<style>
.cvdocx * { color: inherit; }
.cvdocx h1,.cvdocx h2,.cvdocx h3 { margin: 0.6em 0 0.3em; font-weight: 600; }
.cvdocx p { margin: 0.4em 0; }
.cvdocx ul, .cvdocx ol { margin: 0.4em 0 0.6em 1.2em; }
.cvdocx table { border-collapse: collapse; width: 100%; }
.cvdocx td, .cvdocx th { border: 0; padding: 2px 4px; vertical-align: top; }
</style>
CSS;

            $html = '<div class="cvdocx">'.$style.$html.'</div>';

            file_put_contents($cacheFile, $html);
            // Garbage collect older cache files for this page
            foreach (glob($cacheDir . '/' . $page->id() . '-*.html') as $old) {
                if ($old !== $cacheFile) @unlink($old);
            }
        }

        return file_get_contents($cacheFile) ?: '';
    }
}

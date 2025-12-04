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

    // Add this helper inside the class
    private function postProcessHtml(string $html): string
    {
        // Scope wrapper so styles don't leak
        $html = '<div class="cvdocx-scope"><div class="cv-body">' . $html . '</div></div>';

        libxml_use_internal_errors(true);
        $dom = new \DOMDocument('1.0', 'UTF-8');
        $dom->loadHTML('<?xml encoding="UTF-8">' . $html, LIBXML_HTML_NOIMPLIED | LIBXML_HTML_NODEFDTD);
        libxml_clear_errors();

        $xp = new \DOMXPath($dom);

        // YEAR forms we accept:
        //  2026 …
        //  2024–2025 …
        //  2014 - Present …
        //  2014–Present …
        $yearPattern = '/^\s*(\d{4})(?:\s*[–—\-]\s*(Present|\d{4}))?\b(.*)$/u';

        // Which tags should break the "year block"
        $breakTags = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'];

        // Work on a live list: gather <p> under our scope in document order
        $paras = iterator_to_array($xp->query('//div[contains(@class,"cv-body")]//p'));
        $inYearBlock = false;

        foreach ($paras as $p) {
            if (!$p->parentNode)
                continue;
            $text = trim($p->textContent ?? '');

            // Blank lines just end a block
            if ($text === '') {
                $inYearBlock = false;
                continue;
            }

            // If this <p> is inside or preceded by a heading, end block
            $tag = strtolower($p->nodeName);
            if (in_array($tag, $breakTags, true)) {
                $inYearBlock = false;
                continue;
            }

            if (preg_match($yearPattern, $text, $m)) {
                // New year row
                $inYearBlock = true;

                $year = $m[1] . (isset($m[2]) && $m[2] !== '' ? '–' . $m[2] : '');

                // reconstruct original inner HTML then strip the leading year/range once
                $inner = '';
                foreach (iterator_to_array($p->childNodes) as $cn) {
                    $inner .= $dom->saveHTML($cn);
                }
                $innerClean = preg_replace(
                    '/^\s*' . preg_quote($m[1], '/') . '(?:\s*[–—\-]\s*(?:Present|\d{4}))?\s*/u',
                    '',
                    $inner,
                    1
                );
                if ($innerClean === null || $innerClean === '') {
                    $innerClean = htmlspecialchars(trim($m[3] ?? ''), ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
                }

                $row = $dom->createElement('div');
                $row->setAttribute('class', 'cv-row');

                $yearEl = $dom->createElement('div');
                $yearEl->setAttribute('class', 'cv-year');
                $yearEl->appendChild($dom->createTextNode($year));

                $detailEl = $dom->createElement('div');
                $detailEl->setAttribute('class', 'cv-detail');
                $frag = $dom->createDocumentFragment();
                $frag->appendXML($innerClean);
                $detailEl->appendChild($frag);

                $row->appendChild($yearEl);
                $row->appendChild($detailEl);

                $p->parentNode->replaceChild($row, $p);
                continue;
            }

            // Not a year line:
            if ($inYearBlock) {
                // Convert this paragraph into a right-column-only row
                $row = $dom->createElement('div');
                $row->setAttribute('class', 'cv-row');

                $yearEl = $dom->createElement('div');
                $yearEl->setAttribute('class', 'cv-year');
                // keep empty cell to align grid
                $yearEl->appendChild($dom->createTextNode(''));

                $detailEl = $dom->createElement('div');
                $detailEl->setAttribute('class', 'cv-detail');

                // Preserve inner HTML
                $inner = '';
                foreach (iterator_to_array($p->childNodes) as $cn) {
                    $inner .= $dom->saveHTML($cn);
                }
                $frag = $dom->createDocumentFragment();
                // guard in case inner has stray top-level text that needs escaping
                if (@$frag->appendXML($inner) === false) {
                    $detailEl->appendChild($dom->createTextNode($text));
                } else {
                    $detailEl->appendChild($frag);
                }

                $row->appendChild($yearEl);
                $row->appendChild($detailEl);

                $p->parentNode->replaceChild($row, $p);
            } else {
                // Outside a year block: leave paragraph as-is
            }
        }

        $out = $dom->saveHTML();

        // Trim any dumped inline Word CSS blob that sometimes appears at the top
        $out = preg_replace('/^.*?body\s*\{[^}]+\}.*?\}/s', '', $out, 1) ?? $out;

        return $out;
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

        // 1) Allow table markup
        $allowedTags = [
            'p',
            'h1',
            'h2',
            'h3',
            'h4',
            'h5',
            'h6',
            'strong',
            'em',
            'u',
            'a',
            'ul',
            'ol',
            'li',
            'br',
            'span',
            'sup',
            'sub',
            'table',
            'tbody',
            'thead',
            'tr',
            'td',
            'th'   // <-- add these
        ];

        // keep scrubbing inline styles, but don't nuke alignment completely
        $allowedCss = [
            'text-align',       // to let left/right headers survive
            'font-weight',
            'font-style',
            'text-decoration',
            // (intentionally no color / font-family here)
        ];

        // after you build $html (and after purifying), run these normalizers:
        $html = $this->cvdocxNormalizeTables($html);
        $html = $this->cvdocxYearRows($html);

        $html = $this->tidyCvHtml($html);
        $cache->save($key, $html);

        return $html;
    }

    private function cvdocxNormalizeTables(string $html): string
    {
        // Mark the first table as header (name, title, contact)
        $html = preg_replace('/<table\b(?![^>]*class=)/i', '<table class="cv-header"', $html, 1);

        // Any remaining plain tables become year-entry grids
        // (don’t touch tables already carrying a class)
        $html = preg_replace_callback(
            '/<table\b((?:(?!class=)[^>])*)>/i',
            function ($m) {
                $attrs = trim($m[1] ?? '');
                return '<table class="cv-yeargrid" ' . $attrs . '>';
            },
            $html,
            -1,
            $count
        );

        return $html;
    }

    private function cvdocxYearRows(string $html): string
    {
        // Heuristic: for paragraphs that begin with a year or year range and
        // are NOT already inside a table, wrap them as "cv-row"
        // We only touch simple <p>…</p> lines.
        $pattern = '#<p([^>]*)>\s*(\d{4}(?:\s*[–-]\s*(?:\d{4}|Present))?)\s+(.*?)</p>#u';
        $replace = '<div class="cv-row"><span class="cv-year">$2</span><span class="cv-detail">$3</span></div>';

        // Avoid double-processing: skip if this paragraph already sits within a table
        // by splitting on tables and only transforming outside segments
        $parts = preg_split('#(</?table[^>]*>)#i', $html, -1, PREG_SPLIT_DELIM_CAPTURE);
        if ($parts === false)
            return $html;

        for ($i = 0; $i < count($parts); $i++) {
            // Outside table segments are odd indices? We used DELIM_CAPTURE, so:
            // parts like [text, <table>, text, </table>, text ...]
            if (!preg_match('#^</?table#i', $parts[$i] ?? '')) {
                $parts[$i] = preg_replace($pattern, $replace, $parts[$i]);
            }
        }
        return implode('', $parts);
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

        $html = $this->postProcessHtml($html);

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
<?php

/**
 * DOCX HTML Fixers
 * - Overrides PHPWord's ListItemRun writer to output <li> tags.
 * - Provides ListHtmlFixer::rebuildLists() to wrap <li> into <ul>/<ol>.
 */

// 1) Override the PHPWord HTML writer for list items
namespace PhpOffice\PhpWord\Writer\HTML\Element;

use PhpOffice\PhpWord\Writer\HTML\Element\TextRun;

/**
 * Custom HTML writer for ListItemRun elements
 */
class ListItemRun extends TextRun {
    public function write() {
        // Ensure this element is a ListItemRun
        if (! $this->element instanceof \PhpOffice\PhpWord\Element\ListItemRun) {
            return '';
        }

        $depth = $this->element->getDepth();
        $style = $this->element->getStyle();
        $numFmt = method_exists($style, 'getNumStyle') ? $style->getNumStyle() : 'bullet';
        $numId  = $style->getNumId();

        $html = sprintf(
            '<li data-depth="%d" data-liststyle="%s" data-numId="%s">',
            $depth,
            $numFmt,
            $numId
        );

        // Write child elements inside the <li>
        foreach ($this->element->getElements() as $child) {
            $childClass  = get_class($child);
            $writerClass = str_replace(
                'PhpOffice\\PhpWord\\Element',
                __NAMESPACE__,
                $childClass
            );
            if (class_exists($writerClass)) {
                // Pass true to avoid extra <p> wrappers
                $writer = new $writerClass($this->parentWriter, $child, true);
                $html .= $writer->write();
            }
        }

        $html .= '</li>' . PHP_EOL;
        return $html;
    }
}

// 2) Provide a fixer to wrap <li> into <ul>/<ol>
namespace TSM_WP_SPP\DocxHtml;

class ListHtmlFixer {
    /**
     * Wrap <li> tags (with data-depth, data-liststyle) into proper <ul>/<ol> lists.
     *
     * @param string $html Raw HTML from PHPWord
     * @return string Fixed HTML with valid lists
     */
    public static function rebuildLists($html) {
        $lines = explode("\n", $html);
        $out = '';
        $stack = [];
        $currentDepth = 0;
        $currentNumId = null;

        foreach ($lines as $line) {
            if (preg_match(
                '/^<li data-depth="(\d+)" data-liststyle="([^"]+)" data-numId="([^"]+)">/',
                $line,
                $m
            )) {
                list(, $depth, $style, $numId) = $m;
                $depth = (int) $depth;
                $tag   = ($style === 'bullet') ? 'ul' : 'ol';

                // New list sequence?
                if ($currentNumId !== null && $numId !== $currentNumId) {
                    // Close all open lists
                    while ($currentDepth > 0) {
                        $out .= '</' . array_pop($stack) . '>' . "\n";
                        $currentDepth--;
                    }
                }

                // Open lists for increased depth
                while ($currentDepth < $depth) {
                    $out .= "<$tag>\n";
                    $stack[] = $tag;
                    $currentDepth++;
                }
                // Close lists for decreased depth
                while ($currentDepth > $depth) {
                    $out .= '</' . array_pop($stack) . '>' . "\n";
                    $currentDepth--;
                }

                // Extract the <li> content without data attributes
                $inner = preg_replace('/^<li[^>]*>/', '', $line);
                $inner = preg_replace('/<\/li>$/', '', $inner);
                $out  .= '  <li>' . $inner . '</li>' . "\n";

                $currentNumId = $numId;
            } else {
                // Not a list item: close lists if open
                if ($currentDepth > 0) {
                    while ($currentDepth > 0) {
                        $out .= '</' . array_pop($stack) . '>' . "\n";
                        $currentDepth--;
                    }
                    $currentNumId = null;
                }
                $out .= $line . "\n";
            }
        }

        // Close any remaining lists
        while ($currentDepth > 0) {
            $out .= '</' . array_pop($stack) . '>' . "\n";
            $currentDepth--;
        }

        return $out;
    }
}

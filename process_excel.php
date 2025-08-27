<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

// Check if it's an AJAX POST request with a file
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $filePath = $_FILES['file']['tmp_name'];

    try {
        // --- 1. Load Spreadsheet ---
        $reader = new Xlsx(); 
        $reader->setReadEmptyCells(false); 
        $spreadsheet = $reader->load($filePath); 
        
        $sheetNames = $spreadsheet->getSheetNames();
        $sheetToLoad = $_POST['sheetName'] ?? $spreadsheet->getActiveSheet()->getTitle();
        $sheet = $spreadsheet->getSheetByName($sheetToLoad);

        // Convert sheet to array
        $rows = $sheet->toArray(null, true, true, true);
        $rowCount = count($rows);
        $colCount = count($rows[1] ?? []);

        // --- 2. Determine Column Widths Based on Content ---
        $colWidths = [];
        $minWidth = 60;   // minimum px
        $maxWidth = 400;  // optional max width for wrapping
        $charWidth = 8;   // approx width per character

        for ($c = 1; $c <= $colCount; $c++) {
            $colLetter = Coordinate::stringFromColumnIndex($c);
            $maxLen = 0;

            for ($r = 1; $r <= $rowCount; $r++) {
                $val = (string)($rows[$r][$colLetter] ?? '');
                $len = mb_strlen($val);
                if ($len > $maxLen) $maxLen = $len;
            }

            $width = max($minWidth, $maxLen * $charWidth);
            $width = min($width, $maxWidth); // cap width
            $colWidths[$c] = $width;
        }

        // --- 3. Determine Row Heights ---
        $rowHeights = [];
        for ($r = 1; $r <= $rowCount; $r++) {
            $h = $sheet->getRowDimension($r)->getRowHeight();
            $rowHeights[$r] = round(($h ?: 18) * 1.3);
        }

        // --- 4. Column Labels ---
        $colLabels = '<div class="col-labels-inner">';
        for ($c = 1; $c <= $colCount; $c++) {
            $label = Coordinate::stringFromColumnIndex($c);
            $colLabels .= '<div class="col-label" style="width:' . $colWidths[$c] . 'px;">' . $label . '</div>';
        }
        $colLabels .= '</div>';

        // --- 5. Row Labels ---
        $rowLabels = '<div class="row-labels-inner">';
        for ($r = 1; $r <= $rowCount; $r++) {
            $rowLabels .= '<div class="row-label" data-row="' . $r . '" style="height:' . $rowHeights[$r] . 'px;">' . $r . '</div>';
        }
        $rowLabels .= '</div>';

        // --- 6. Handle Merged Cells ---
        $mergedMap = [];
        $mergeSpans = [];
        foreach ($sheet->getMergeCells() as $range) {
            [$start, $end] = explode(':', $range);
            $startC = Coordinate::columnIndexFromString(preg_replace('/\d+/', '', $start));
            $startR = (int)preg_replace('/[A-Z]+/', '', $start);
            $endC   = Coordinate::columnIndexFromString(preg_replace('/\d+/', '', $end));
            $endR   = (int)preg_replace('/[A-Z]+/', '', $end);

            $mergedMap[$startR][$startC] = "root";
            $mergeSpans[$startR][$startC] = ['cols' => $endC - $startC + 1, 'rows' => $endR - $startR + 1];
            for ($rr = $startR; $rr <= $endR; $rr++) {
                for ($cc = $startC; $cc <= $endC; $cc++) {
                    if ($rr == $startR && $cc == $startC) continue;
                    $mergedMap[$rr][$cc] = "skip";
                }
            }
        }

        // --- 7. Build Grid Style ---
        $colTemplates = [];
        for ($c = 1; $c <= $colCount; $c++) {
            $colTemplates[] = $colWidths[$c] . 'px';
        }
        $rowTemplates = [];
        for ($r = 1; $r <= $rowCount; $r++) {
            $rowTemplates[] = $rowHeights[$r] . 'px';
        }

        $gridStyle = 'grid-template-columns: ' . implode(' ', $colTemplates) . ';';
        $gridStyle .= ' grid-auto-rows: min-content;'; 
        $gridStyle .= ' grid-template-rows: ' . implode(' ', $rowTemplates) . ';';

        // --- 8. Build Grid HTML with Word Wrap ---
        $sheetHtml = '<div class="grid-inner" style="' . $gridStyle . '">';

        for ($r = 1; $r <= $rowCount; $r++) {
            for ($c = 1; $c <= $colCount; $c++) {
                if (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "skip") continue;

                $colLetter = Coordinate::stringFromColumnIndex($c);
                $value = htmlspecialchars((string)($rows[$r][$colLetter] ?? ''));

                $colspan = 1;
                $rowspan = 1;
                if (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "root") {
                    $span = $mergeSpans[$r][$c];
                    $colspan = $span['cols'];
                    $rowspan = $span['rows'];
                }

                $css = [];
                $css[] = "grid-column: " . $c . " / span " . $colspan;
                $css[] = "grid-row: " . $r . " / span " . $rowspan;
                $css[] = "display:flex";
                $css[] = "align-items:flex-start";
                $css[] = "justify-content:flex-start";
                $css[] = "word-break:break-word";
                $css[] = "max-width:" . ($colWidths[$c]) . "px"; // wrap at max column width

                $sheetHtml .= '<div class="cell' . 
                    (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "root" ? ' merged-root' : '') . 
                    '" data-row="' . $r . '" data-col="' . $c . '" style="' . implode(";", $css) . '">' . 
                    '<span class="cell-content" style="white-space:normal;">' . $value . '</span></div>';
            }
        }
        $sheetHtml .= '</div>';

        // --- 9. Send JSON Response ---
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode([
            'html' => $sheetHtml,
            'colLabels' => $colLabels,
            'rowLabels' => $rowLabels,
            'info' => "$rowCount rows × $colCount cols",
            'sheets' => $sheetNames,
            'currentSheet' => $sheetToLoad
        ]);
        exit;

    } catch (Exception $e) {
        http_response_code(500);
        echo json_encode(['error' => $e->getMessage()]);
        exit;
    }
}

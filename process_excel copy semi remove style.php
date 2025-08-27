<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// This file contains ONLY the PHP backend logic.

// Check if it's an AJAX POST request with a file
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $filePath = $_FILES['file']['tmp_name'];

    try {
        // --- 1. Load Spreadsheet (HYBRID: Fast for XLSX, Auto for others) ---
        $ext = strtolower(pathinfo($_FILES['file']['name'], PATHINFO_EXTENSION));

        if ($ext === 'xlsx') {
            // Optimized path for XLSX
            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
            $reader->setReadEmptyCells(false);
        } else {
            // Auto-detect for other formats (.xls, .csv, .ods, etc.)
            $type = IOFactory::identify($filePath);
            $reader = IOFactory::createReader($type);
            $reader->setReadDataOnly(false);
        }

        $spreadsheet = $reader->load($filePath);

        $sheetNames = $spreadsheet->getSheetNames();
        $sheetToLoad = $_POST['sheetName'] ?? $spreadsheet->getActiveSheet()->getTitle();
        $sheet = $spreadsheet->getSheetByName($sheetToLoad);

        // Convert sheet to array
        $rows = $sheet->toArray(null, true, true, true);
        $rowCount = count($rows);
        $colCount = count($rows[1] ?? []);

        // --- 2. Auto-size Baseline (Optimized) ---
        for ($c = 1; $c <= $colCount; $c++) {
            $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($c))->setAutoSize(true);
        }
        $sheet->calculateColumnWidths();

        // --- 3. Column Widths ---
        $colWidths = [];
        $minWidth = 60; 
        $widthFactor = 8; 
        for ($c = 1; $c <= $colCount; $c++) {
            $colLetter = Coordinate::stringFromColumnIndex($c);
            $baseWidth = round($sheet->getColumnDimension($colLetter)->getWidth() * $widthFactor);
            $colWidths[$c] = max($minWidth, $baseWidth);
        }

        // --- 4. Row Heights ---
        $rowHeights = [];
        for ($r = 1; $r <= $rowCount; $r++) {
            $h = $sheet->getRowDimension($r)->getRowHeight();
            $rowHeights[$r] = round(($h ?: 18) * 1.3);
        }

        // --- 5. Column Labels ---
        $colLabels = '<div class="col-labels-inner">';
        for ($c = 1; $c <= $colCount; $c++) {
            $label = Coordinate::stringFromColumnIndex($c);
            $colLabels .= '<div class="col-label" style="width:' . $colWidths[$c] . 'px;">' . $label . '</div>';
        }
        $colLabels .= '</div>';

        // --- 6. Row Labels ---
        $rowLabels = '<div class="row-labels-inner">';
        for ($r = 1; $r <= $rowCount; $r++) {
            $rowLabels .= '<div class="row-label" data-row="' . $r . '" style="height:' . $rowHeights[$r] . 'px;">' . $r . '</div>';
        }
        $rowLabels .= '</div>';

        // --- 7. Handle Merged Cells ---
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

        // --- 8. Build Grid Style ---
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

        // --- 9. Build Grid HTML ---
        $sheetHtml = '<div class="grid-inner" style="' . $gridStyle . '">';

        for ($r = 1; $r <= $rowCount; $r++) {
            for ($c = 1; $c <= $colCount; $c++) {
                if (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "skip") {
                    continue;
                }
                $colLetter = Coordinate::stringFromColumnIndex($c);
                $value = htmlspecialchars((string)($rows[$r][$colLetter] ?? ''));
                $style = $sheet->getStyleByColumnAndRow($c, $r); 
                $colspan = 1;
                $rowspan = 1;
                if (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "root") {
                    $span = $mergeSpans[$r][$c];
                    $colspan = $span['cols'];
                    $rowspan = $span['rows'];
                }
                $css = [];
                $classes = ['cell'];
                $font = $style->getFont();
                if ($font->getBold()) $css[] = "font-weight:bold";
                if ($font->getItalic()) $css[] = "font-style:italic";
                if ($font->getColor()->getRGB() !== "000000") $css[] = "color:#" . $font->getColor()->getRGB();
                $align = $style->getAlignment();
                switch ($align->getHorizontal()) {
                    case Alignment::HORIZONTAL_CENTER: $css[] = "text-align:center;justify-content:center"; break;
                    case Alignment::HORIZONTAL_RIGHT:  $css[] = "text-align:right;justify-content:flex-end"; break;
                    default: $css[] = "text-align:left;justify-content:flex-start"; break;
                }
                switch ($align->getVertical()) {
                    case Alignment::VERTICAL_TOP:    $css[] = "align-items:flex-start;vertical-align:top"; break;
                    case Alignment::VERTICAL_BOTTOM: $css[] = "align-items:flex-end;vertical-align:bottom"; break;
                    default: $css[] = "align-items:center;vertical-align:middle"; break;
                }
                if ($align->getWrapText()) $classes[] = 'word-wrap';
                $fill = $style->getFill();
                if ($fill->getFillType() !== Fill::FILL_NONE) {
                    $color = $fill->getStartColor()->getRGB();
                    if ($color !== "FFFFFF") $css[] = "background-color:#" . $color;
                }
                $css[] = "grid-column: " . $c . " / span " . $colspan;
                $css[] = "grid-row: " . $r . " / span " . $rowspan;
                $sheetHtml .= '<div class="' . implode(' ', $classes) . (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "root" ? ' merged-root' : '') . '" data-row="' . $r . '" data-col="' . $c . '" style="' . implode(";", $css) . '"><span class="cell-content">' . $value . '</span></div>';
            }
        }
        $sheetHtml .= '</div>';

        // Send JSON response to the client
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

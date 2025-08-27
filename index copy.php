<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

/**
 * Measure text width in pixels using GD + TTF font
 */
function measureText($text, $fontSize = 11, $fontFile = null) {
    if (!$fontFile) {
        // fallback to built-in Arial substitute (Linux: DejaVuSans)
        $fontFile = __DIR__ . "/arial.ttf";
        if (!file_exists($fontFile)) {
            return strlen($text) * ($fontSize * 0.6) + 12; // fallback approx
        }
    }
    $box = imagettfbbox($fontSize, 0, $fontFile, $text);
    return abs($box[2] - $box[0]) + 12; // + padding
}

// ===== AJAX handler for file load =====
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $filePath = $_FILES['file']['tmp_name'];

    try {
        $spreadsheet = IOFactory::load($filePath);
        $sheetNames = $spreadsheet->getSheetNames();
        $sheetToLoad = $_POST['sheetName'] ?? $spreadsheet->getActiveSheet()->getTitle();
        $sheet = $spreadsheet->getSheetByName($sheetToLoad);

        // Convert sheet to array
        $rows = $sheet->toArray(null, true, true, true);
        $rowCount = count($rows);
        $colCount = count($rows[1] ?? []);

        // Auto-size baseline
        for ($c = 1; $c <= $colCount; $c++) {
            $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($c))->setAutoSize(true);
        }
        $sheet->calculateColumnWidths();

        // Column widths (improved with text measurement)
        $colWidths = [];
        for ($c = 1; $c <= $colCount; $c++) {
            $colLetter = Coordinate::stringFromColumnIndex($c);
            $baseWidth = round(
                $sheet->getColumnDimension($colLetter)->getWidth() * 7
            );
            if ($baseWidth < 60) $baseWidth = 60;

            // ensure wide enough for text
            $maxWidth = $baseWidth;
            for ($r = 1; $r <= $rowCount; $r++) {
                $val = trim((string)($rows[$r][$colLetter] ?? ''));
                if ($val !== '') {
                    $maxWidth = max($maxWidth, measureText($val));
                }
            }
            $colWidths[$c] = $maxWidth;
        }

        // Row heights
        $rowHeights = [];
        for ($r = 1; $r <= $rowCount; $r++) {
            $h = $sheet->getRowDimension($r)->getRowHeight();
            $rowHeights[$r] = round(($h ?: 18) * 1.3);
        }

        // Column labels
        $colLabels = '<div class="col-labels-inner">';
        for ($c = 1; $c <= $colCount; $c++) {
            $label = Coordinate::stringFromColumnIndex($c);
            $colLabels .= '<div class="col-label" style="width:' . $colWidths[$c] . 'px;">' . $label . '</div>';
        }
        $colLabels .= '</div>';

        // Row labels
        $rowLabels = '<div class="row-labels-inner">';
        for ($r = 1; $r <= $rowCount; $r++) {
            $rowLabels .= '<div class="row-label" data-row="' . $r . '" style="height:' . $rowHeights[$r] . 'px;">' . $r . '</div>';
        }
        $rowLabels .= '</div>';

        // Handle merged cells
        $mergedMap = [];
        $mergeSpans = [];
        foreach ($sheet->getMergeCells() as $range) {
            [$start, $end] = explode(':', $range);
            $startC = Coordinate::columnIndexFromString(preg_replace('/\d+/', '', $start));
            $startR = (int)preg_replace('/[A-Z]+/', '', $start);
            $endC   = Coordinate::columnIndexFromString(preg_replace('/\d+/', '', $end));
            $endR   = (int)preg_replace('/[A-Z]+/', '', $end);

            $mergedMap[$startR][$startC] = "root";
            $mergeSpans[$startR][$startC] = [
                'cols' => $endC - $startC + 1,
                'rows' => $endR - $startR + 1,
            ];
            for ($rr = $startR; $rr <= $endR; $rr++) {
                for ($cc = $startC; $cc <= $endC; $cc++) {
                    if ($rr == $startR && $cc == $startC) continue;
                    $mergedMap[$rr][$cc] = "skip";
                }
            }
        }

        // --- Calculate Total Grid Dimensions ---
        $totalWidth = array_sum($colWidths);
        $totalHeight = array_sum($rowHeights);

        // Dynamically create grid column and row definitions
        $colTemplates = [];
        for ($c = 1; $c <= $colCount; $c++) {
            $colTemplates[] = $colWidths[$c] . 'px';
        }
        $rowTemplates = [];
        for ($r = 1; $r <= $rowCount; $r++) {
            $rowTemplates[] = $rowHeights[$r] . 'px';
        }

        $gridStyle = 'grid-template-columns: ' . implode(' ', $colTemplates) . ';';
        $gridStyle .= ' grid-auto-rows: min-content;'; // Use min-content for auto-height
        $gridStyle .= ' grid-template-rows: ' . implode(' ', $rowTemplates) . ';';

        // Build grid HTML
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

                // Style CSS
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

                // Use CSS Grid properties for placement and spanning
                $css[] = "grid-column: " . $c . " / span " . $colspan;
                $css[] = "grid-row: " . $r . " / span " . $rowspan;

                $sheetHtml .= '<div class="' . implode(' ', $classes) . (isset($mergedMap[$r][$c]) && $mergedMap[$r][$c] === "root" ? ' merged-root' : '') . '" data-row="' . $r . '" data-col="' . $c . '" style="'
                    . implode(";", $css) . '">'
                    . '<span class="cell-content">' . $value . '</span></div>';
            }
        }
        $sheetHtml .= '</div>';

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
?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>PhpSpreadsheet Viewer</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
<div id="loader" class="loader-overlay" style="display:none;">
  <div class="spinner"></div>
  <div class="loader-text">Loading spreadsheet…</div>
</div>
<div class="wrap">
  <div class="card">
    <div class="top">
      <div class="title">🧾 PhpSpreadsheet Viewer</div>
      <div class="controls">
        <input id="file" class="control" type="file" accept=".xlsx,.xls,.csv" />
        <div class="search control" id="searchBox" style="display:none;">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor"
               stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
          <input id="query" placeholder="Search (Ctrl+/)" />
          <button id="clear" class="btn" style="padding:6px 8px;margin-left:8px">Clear</button>
        </div>
      </div>
    </div>

    <div class="sheet-area">
      <div class="grid-header">
        <div class="top-left">#</div>
        <div class="col-labels-container"><div class="col-labels" id="colLabels"></div></div>
      </div>
      <div class="grid-content">
        <div class="row-labels-container"><div class="row-labels" id="rowLabels"></div></div>
        <div class="viewport" id="viewport"></div>
      </div>
    </div>

    <div class="meta">
      <ul class="sheet-tabs" id="sheetTabs"></ul>
      <span id="sheetInfo">Load a file to display.</span>
    </div>
  </div>
</div>

<script>
const viewport = document.getElementById('viewport');
const queryInput = document.getElementById('query');
const clearBtn = document.getElementById('clear');
const searchBox = document.getElementById('searchBox');
const loader = document.getElementById('loader');
const sheetTabsContainer = document.getElementById('sheetTabs');
let currentFile = null;

async function loadFile(file, sheetName = null) {
  currentFile = file;
  const formData = new FormData();
  formData.append('file', file);
  if (sheetName) {
    formData.append('sheetName', sheetName);
  }
  loader.style.display = 'flex';

  try {
    let res = await fetch('index.php', { method:'POST', body: formData });
    let data = await res.json();

    if (data.error) {
      viewport.innerHTML = '<p style="color:red">' + data.error + '</p>';
    } else {
      viewport.innerHTML = data.html;
      document.getElementById('colLabels').innerHTML = data.colLabels;
      document.getElementById('rowLabels').innerHTML = data.rowLabels;
      document.getElementById('sheetInfo').textContent = data.info;
      searchBox.style.display = 'flex';

      // Render sheet tabs only on initial file load
      if (!sheetName) {
        sheetTabsContainer.innerHTML = ''; // Clear existing tabs
        data.sheets.forEach(sheet => {
          const tab = document.createElement('li');
          tab.className = 'sheet-tab';
          if (sheet === data.currentSheet) {
            tab.classList.add('active');
          }
          tab.textContent = sheet;
          tab.addEventListener('click', () => {
            document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            loadFile(currentFile, sheet);
          });
          sheetTabsContainer.appendChild(tab);
        });
      }

      // Adjust heights after rendering
      adjustRowHeights();
    }
  } catch (err) {
    viewport.innerHTML = '<p style="color:red">Error loading file</p>';
  } finally {
    loader.style.display = 'none';
  }
}

function adjustRowHeights() {
  const cells = document.querySelectorAll('.cell.word-wrap');
  const rowsToUpdate = {};

  cells.forEach(cell => {
    const row = cell.getAttribute('data-row');
    const content = cell.querySelector('.cell-content');
    if (content) {
      const requiredHeight = content.scrollHeight;
      if (requiredHeight > (rowsToUpdate[row] || 0)) {
        rowsToUpdate[row] = requiredHeight;
      }
    }
  });

  for (const row in rowsToUpdate) {
    const newHeight = rowsToUpdate[row] + 16; // Add padding

    // Update all cells in the row
    document.querySelectorAll(`[data-row="${row}"]`).forEach(cell => {
      cell.style.height = `${newHeight}px`;
    });

    // Update row label height
    const rowLabel = document.querySelector(`.row-label[data-row="${row}"]`);
    if (rowLabel) {
      rowLabel.style.height = `${newHeight}px`;
    }
  }
}

document.getElementById('file').addEventListener('change', function(e) {
  const file = e.target.files[0];
  if (!file) return;
  loadFile(file);
});

// Search
queryInput.addEventListener('input', function() {
  const query = this.value.trim().toLowerCase();
  const cells = viewport.querySelectorAll('.cell');
  let firstMatch = null;
  cells.forEach(cell => {
    let text = cell.textContent.toLowerCase();
    if (query && text.includes(query)) {
      if (!firstMatch) firstMatch = cell;
      const idx = text.indexOf(query);
      const original = cell.textContent;
      cell.innerHTML = '<span class="cell-content">' + original.substring(0, idx) +
                       '<mark>' + original.substring(idx, idx+query.length) + '</mark>' +
                       original.substring(idx+query.length) + '</span>';
    } else {
      cell.innerHTML = '<span class="cell-content">' + cell.textContent + '</span>';
    }
  });
  if (firstMatch) firstMatch.scrollIntoView({behavior:'smooth', block:'center'});
});

clearBtn.addEventListener('click', () => {
  queryInput.value = '';
  const cells = viewport.querySelectorAll('.cell');
  cells.forEach(cell => cell.innerHTML = '<span class="cell-content">' + cell.textContent + '</span>');
});
</script>
</body>
</html>
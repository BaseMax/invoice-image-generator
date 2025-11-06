<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Check if a row is entirely empty
 */
function isRowEmpty(array $row): bool {
    foreach ($row as $cell) {
        if (trim((string)$cell) !== '') return false;
    }
    return true;
}

/**
 * Get a lookup table of merged cell coordinates -> top-left coordinate
 */
function getMergedMap(Worksheet $sheet): array {
    $mergedMap = [];
    foreach ($sheet->getMergeCells() as $range) {
        [$start, $end] = explode(':', $range);
        $startCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(
            preg_replace('/\d+/', '', $start)
        );
        $endCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(
            preg_replace('/\d+/', '', $end)
        );
        $startRow = (int)preg_replace('/\D+/', '', $start);
        $endRow = (int)preg_replace('/\D+/', '', $end);

        for ($r = $startRow; $r <= $endRow; $r++) {
            for ($c = $startCol; $c <= $endCol; $c++) {
                $mergedMap["{$r}_{$c}"] = [$startRow, $startCol];
            }
        }
    }
    return $mergedMap;
}

/**
 * Remove empty columns and trim trailing empties
 */
function removeEmptyColumns(array $table): array {
    if (empty($table)) return $table;

    $colCount = max(array_map('count', $table));
    $colHasValue = array_fill(0, $colCount, false);

    foreach ($table as $row) {
        foreach ($row as $i => $cell) {
            if (trim((string)$cell) !== '') $colHasValue[$i] = true;
        }
    }

    $cleaned = [];
    foreach ($table as $row) {
        $newRow = [];
        foreach ($row as $i => $cell) {
            if (!empty($colHasValue[$i])) $newRow[] = trim((string)$cell);
        }
        while (!empty($newRow) && trim(end($newRow)) === '') array_pop($newRow);
        $cleaned[] = $newRow;
    }

    return $cleaned;
}

/**
 * Extract tables separated by empty rows (and handle merged cells)
 */
function extractTables(string $filePath, int $maxEmptyGap = 2): array {
    $reader = IOFactory::createReaderForFile($filePath);
    $reader->setReadDataOnly(false);
    $spreadsheet = $reader->load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    $mergedMap = getMergedMap($sheet);

    $rows = [];
    foreach ($sheet->getRowIterator() as $row) {
        $cells = [];
        foreach ($row->getCellIterator() as $cell) {
            $colIndex = $cell->getColumn();
            $rowIndex = $cell->getRow();
            $colNum = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($colIndex);

            $mergedKey = "{$rowIndex}_{$colNum}";
            if (isset($mergedMap[$mergedKey])) {
                [$topRow, $leftCol] = $mergedMap[$mergedKey];
                if ($topRow !== $rowIndex || $leftCol !== $colNum) {
                    continue;
                }
            }

            try {
                $value = $cell->getCalculatedValue();
            } catch (Exception $e) {
                $value = $cell->getValue();
            }

            $cells[] = trim((string)$value);
        }
        $rows[] = $cells;
    }

    $tables = [];
    $current = [];
    $emptyCount = 0;
    foreach ($rows as $row) {
        if (isRowEmpty($row)) {
            $emptyCount++;
        } else {
            if ($emptyCount >= $maxEmptyGap && !empty($current)) {
                $tables[] = removeEmptyColumns($current);
                $current = [];
            }
            $emptyCount = 0;
            $current[] = $row;
        }
    }
    if (!empty($current)) $tables[] = removeEmptyColumns($current);

    return $tables;
}

$filePath = 'input.xlsx';
$tables = extractTables($filePath);

foreach ($tables as $i => $table) {
    echo "=== Table " . ($i + 1) . " ===\n";
    foreach ($table as $row) {
        echo implode(' | ', $row) . "\n";
    }
    echo "\n";
}

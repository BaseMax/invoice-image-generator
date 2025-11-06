<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

function isRowEmpty(array $row): bool {
    foreach ($row as $cell) {
        if (trim((string)$cell) !== '') return false;
    }
    return true;
}

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
 * Remove *all* empty cells from rows (not just columns)
 */
function compactRows(array $table): array {
    $cleaned = [];
    foreach ($table as $row) {
        $filtered = array_values(array_filter($row, fn($v) => trim((string)$v) !== ''));
        $cleaned[] = $filtered;
    }
    return $cleaned;
}

/**
 * Extract tables separated by empty rows (handles merged + formulas)
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
                $tables[] = compactRows($current);
                $current = [];
            }
            $emptyCount = 0;
            $current[] = $row;
        }
    }
    if (!empty($current)) $tables[] = compactRows($current);

    return $tables;
}

// === Run ===
$filePath = 'input.xlsx';
$tables = extractTables($filePath);

foreach ($tables as $i => $table) {
    echo "=== Table " . ($i + 1) . " ===\n";
    foreach ($table as $row) {
        echo implode(' | ', $row) . "\n";
    }
    echo "\n";
}

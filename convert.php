<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Cell;

/**
 * Check if a row is completely empty.
 */
function isRowEmpty(array $row): bool {
    foreach ($row as $cell) {
        if (trim((string)$cell) !== '') {
            return false;
        }
    }
    return true;
}

/**
 * Remove columns that are empty in all rows.
 */
function removeEmptyColumns(array $table): array {
    if (empty($table)) return $table;

    $usedColumns = [];
    foreach ($table as $row) {
        foreach ($row as $index => $value) {
            if (trim((string)$value) !== '') {
                $usedColumns[$index] = true;
            }
        }
    }

    $filtered = [];
    foreach ($table as $row) {
        $newRow = [];
        foreach ($row as $index => $value) {
            if (isset($usedColumns[$index])) {
                $newRow[] = trim((string)$value);
            }
        }
        $filtered[] = $newRow;
    }

    return $filtered;
}

/**
 * Extract multiple tables from Excel separated by blank rows,
 * and evaluate Excel formulas automatically.
 */
function extractTables(string $filePath, int $maxEmptyGap = 2): array {
    $reader = IOFactory::createReaderForFile($filePath);
    $reader->setReadDataOnly(false);
    $spreadsheet = $reader->load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    $rows = [];
    foreach ($sheet->getRowIterator() as $row) {
        $cells = [];
        foreach ($row->getCellIterator() as $cell) {
            /** @var Cell $cell */
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
    $currentTable = [];
    $emptyCount = 0;

    foreach ($rows as $row) {
        if (isRowEmpty($row)) {
            $emptyCount++;
        } else {
            if ($emptyCount >= $maxEmptyGap && count($currentTable) > 0) {
                $tables[] = removeEmptyColumns($currentTable);
                $currentTable = [];
            }
            $emptyCount = 0;
            $currentTable[] = $row;
        }
    }

    if (count($currentTable) > 0) {
        $tables[] = removeEmptyColumns($currentTable);
    }

    return $tables;
}

$filePath = 'input.xlsx';
$tables = extractTables($filePath);

foreach ($tables as $index => $table) {
    echo "=== Table " . ($index + 1) . " ===\n";
    foreach ($table as $row) {
        echo implode(" | ", $row) . "\n";
    }
    echo "\n";
}

<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

function isRowEmpty(array $row): bool {
    foreach ($row as $cell) {
        if (trim((string)$cell) !== '') {
            return false;
        }
    }
    return true;
}

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

function extractTables(string $filePath, int $maxEmptyGap = 2): array {
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    $rows = [];
    foreach ($sheet->getRowIterator() as $row) {
        $cells = [];
        foreach ($row->getCellIterator() as $cell) {
            $cells[] = trim((string)$cell->getValue());
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
        $tables[] = $currentTable;
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

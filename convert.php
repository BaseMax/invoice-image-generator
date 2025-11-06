<?php
ini_set('memory_limit', '4095M');

require "fagd.php";
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

$products = json_decode(file_get_contents("all-products.json"), true);

function getNumberFromText(string $text): ?string
{
    preg_match_all('/\d+(?:\.\d+)?/', $text, $matches);
    if (empty($matches[0])) {
        return null;
    }
    return end($matches[0]);
}

function findProductBySku(string $sku): ?string {
    global $products;

    foreach ($products as $product_item) {
        if ($product_item["sku"] === $sku) {
            return $product_item["image"];
        }
    }

    return null;
}

function isRowEmpty(array $row): bool {
    foreach ($row as $cell) {
        if (trim((string)$cell) !== '') return false;
    }
    return true;
}

function reverseNumbersInText(string $text): string {
    return preg_replace_callback('/[0-9۰-۹٠-٩\/.,]+/u', function ($m) {
        $num = $m[0];

        $western = ['۰','۱','۲','۳','۴','۵','۶','۷','۸','۹','٠','١','٢','٣','٤','٥','٦','٧','٨','٩'];
        $latin   = ['0','1','2','3','4','5','6','7','8','9','0','1','2','3','4','5','6','7','8','9'];
        $num = str_replace($western, $latin, $num);

        $reversed = strrev($num);

        $reversed = toPersianDigits($reversed);
        return $reversed;
    }, $text);
}

function getMergedMap(Worksheet $sheet): array {
    $mergedMap = [];
    foreach ($sheet->getMergeCells() as $range) {
        [$start, $end] = explode(':', $range);
        $startCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(preg_replace('/\d+/', '', $start));
        $endCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(preg_replace('/\d+/', '', $end));
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

function compactRows(array $table): array {
    $cleaned = [];
    foreach ($table as $row) {
        $filtered = array_values(array_filter($row, fn($v) => trim((string)$v) !== ''));
        $cleaned[] = $filtered;
    }
    return $cleaned;
}

function toPersianDigits(string $text): string {
    $western = ['0','1','2','3','4','5','6','7','8','9'];
    $arabic  = ['٠','١','٢','٣','٤','٥','٦','٧','٨','٩'];
    $persian = ['۰','۱','۲','۳','۴','۵','۶','۷','۸','۹'];
    $text = str_replace($western, $persian, $text);
    $text = str_replace($arabic, $persian, $text);
    return $text;
}

function extractTables(string $filePath, int $maxEmptyGap = 2): array {
    $reader = IOFactory::createReaderForFile($filePath);
    $reader->setReadDataOnly(false);
    $spreadsheet = $reader->load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    $mergedMap = getMergedMap($sheet);
    $rows = [];

    foreach ($sheet->getRowIterator() as $rowObj) {
        $cells = [];
        $cellIterator = $rowObj->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        foreach ($cellIterator as $cell) {
            $colIndex = $cell->getColumn();
            $rowIndex = $cell->getRow();
            $colNum = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($colIndex);
            $mergedKey = "{$rowIndex}_{$colNum}";
            if (isset($mergedMap[$mergedKey])) {
                [$topRow, $leftCol] = $mergedMap[$mergedKey];
                if ($topRow !== $rowIndex || $leftCol !== $colNum) {
                    $cells[] = '';
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

function sizeText(int $fontSize, string $fontFile, string $text) {
    $angle = 0;
    $bbox = imagettfbbox($fontSize, $angle, $fontFile, $text);
    return $bbox;
}

function renderText($im, string $text, int $x, int $y, int $fontSize, $color, string $fontFile) {
    $text = fagd($text, 'fa', 'normal');
    imagettftext($im, $fontSize, 0, $x, $y, $color, $fontFile, $text);
}

$filePath = 'input.xlsx';
if (!file_exists($filePath)) die("input.xlsx not found\n");

$tables = extractTables($filePath);

$fontFile = __DIR__ . '/FreeFarsi.ttf';
if (!file_exists($fontFile)) die("Font not found: $fontFile\n");

$fontSize = 16;
$padding = 25;
$cellPadding = 16;
$lineHeight = (int)($fontSize * 2.2);

foreach ($tables as $tIndex => $table) {
    if (empty($table)) continue;

    print_r($table);

    $topLines = array_slice($table, 0, 2);
    $bottomLines = array_slice($table, -4);
    $headerIdx = null;
    foreach ($table as $ri => $row) {
        if (mb_strpos(implode(' ', $row), 'ردیف') !== false) {
            $headerIdx = $ri;
            break;
        }
    }
    if ($headerIdx === null) $headerIdx = 2;

    $tableLines = [];
    for ($ri = $headerIdx; $ri < count($table) - 4; $ri++) {
        $row = array_map('trim', $table[$ri]);
        if (empty(array_filter($row))) break;
        $tableLines[] = $row;
    }

    $maxCols = max(array_map('count', $tableLines));
    foreach ($tableLines as &$r) while (count($r) < $maxCols) $r[] = '';
    unset($r);

    $colWidths = array_fill(0, $maxCols, 0);
    foreach ($tableLines as $row) {
        foreach ($row as $ci => $cell) {
            $bbox = sizeText($fontSize, $fontFile, $cell ?: ' ');
            $w = abs($bbox[2] - $bbox[0]);
            $colWidths[$ci] = max($colWidths[$ci], $w + $cellPadding * 2);
        }
    }

    $totalWidth = array_sum($colWidths) + $padding * 2;
    $totalHeight = $padding * 4 + $lineHeight * (count($tableLines) + count($topLines) + count($bottomLines) + 4);

    $im = imagecreatetruecolor($totalWidth, $totalHeight);
    $white = imagecolorallocate($im, 255, 255, 255);
    $black = imagecolorallocate($im, 0, 0, 0);
    $gray  = imagecolorallocate($im, 240, 240, 240);
    imagefill($im, 0, 0, $white);

    $y = $padding + $fontSize;

    foreach ($topLines as $row) {
        $text = implode(' ', $row);
        $bbox = sizeText($fontSize, $fontFile, $text);
        $textW = abs($bbox[2] - $bbox[0]);
        renderText($im, $text, $totalWidth - $padding - $textW, $y, $fontSize, $black, $fontFile);
        $y += $lineHeight;
    }

    $y += $lineHeight / 2;

    $tableLeft = $totalWidth - $padding - array_sum($colWidths);
    $images = [];
    foreach ($tableLines as $rIndex => $row) {
        $x = $totalWidth - $padding;
        if ($rIndex === 0) {
            imagefilledrectangle($im, (int) $tableLeft, (int) $y - $fontSize - 6, (int) $totalWidth - $padding, (int) $y + $lineHeight - $fontSize, $gray);
        }

        if ($rIndex !== 0) {
            $product_name = $row[1];
            $product_sku = getNumberFromText($product_name);
            if ($product_sku === null) {
                print "Error: cannot get sku from product name.";
                exit();
            }
            $image = findProductBySku($product_sku);
            if ($image === null) {
                print "Error: cannot find the product details or image.";
                exit();
            }
            $images[] = $image;
        }
        foreach ($row as $ci => $cellText) {
            $colW = $colWidths[$ci];
            $bbox = sizeText($fontSize, $fontFile, $cellText ?: ' ');
            $textW = abs($bbox[2] - $bbox[0]);
            $textH = abs($bbox[5] - $bbox[3]);
            $textX = (int) ($x - $cellPadding - $textW);
            $textY = (int) ($y + ($lineHeight - $textH) / 2 - 4) + 10;
            renderText($im, $cellText, $textX, $textY, (int) $fontSize, $black, $fontFile);

            $x -= $colW;
            imageline($im, (int) $x, (int) $y - $lineHeight + 8 + 5, (int) $x, (int) $y + 8 + 20, $black);
        }

        $y += $lineHeight;
        imageline(
            $im,
            (int)$tableLeft,
            (int)($y - 6),
            (int)($totalWidth - $padding),
            (int)($y - 6),
            $black
        );
    }

    $y += $lineHeight / 2;

    foreach ($bottomLines as $row) {
        if ($row[0] === "3" || $row[0] === "4" || $row[0] === "2") {
            unset($row[0]);
            $row = array_values($row);
        }

        if (str_contains($row[0], "ارسال")) {
            if(isset($row[1])) unset($row[1]);
        }

        $text = implode(' ', $row);
        $bbox = sizeText($fontSize, $fontFile, $text);
        $textW = abs($bbox[2] - $bbox[0]);
        renderText($im, $text, $totalWidth - $padding - $textW, $y, $fontSize, $black, $fontFile);
        $y += $lineHeight;
    }

    $outFile = "table_{$tIndex}.png";
    imagepng($im, $outFile);
    imagedestroy($im);

    echo "Saved: $outFile\n";
    exit;
}

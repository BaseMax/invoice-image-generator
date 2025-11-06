<?php 
require "_base.php";

ini_set('memory_limit', '4095M');

require "fagd.php";
require 'vendor/autoload.php';

$products = json_decode(file_get_contents("all-products.json"), true);

$filePath = 'input.xlsx';
if (!file_exists($filePath)) die("input.xlsx not found\n");

$outDir = "tables/";
if (!is_dir($outDir)) {
    @mkdir($outDir, 0755, true);
}
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

    $outFile = $outDir . "table_{$tIndex}.png";
    if (file_exists($outFile)) continue;

    $topLines = array_slice($table, 0, 2);
    $bottomLines = array_slice($table, -4);
    $bottomLines[] =  ["از کل فروش این فاکتور یک درصد صرف امور خیریه می‌شود."];
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

        if (isset($row[0]) && str_contains($row[0], "ارسال")) {
            break;
        }
        else if (isset($row[1]) && str_contains($row[1], "ارسال")) {
            break;
        }

        if (isset($row[1]) && strlen($row[1]) < 4) {
            break;
        }

        if(! isset($row[4])) {
            print "Error: cannot find metraj in table items!\n";
            exit();
        }
        if ($ri !== $headerIdx) {
            $row[4] = strrev(str_replace(".", "/", $row[4]));
        }
        $tableLines[] = $row;
    }

    if (empty($tableLines)) continue;

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

    $tableContentWidth = array_sum($colWidths);
    $totalWidth = 1200;
    $totalHeight = 1200;

    $im = imagecreatetruecolor($totalWidth, $totalHeight);
    $white = imagecolorallocate($im, 255, 255, 255);
    $black = imagecolorallocate($im, 0, 0, 0);
    $gray  = imagecolorallocate($im, 255, 196, 217);
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

    $tableLeft = $totalWidth - $padding - $tableContentWidth;
    $tableRight = $totalWidth - $padding;
    $images = [];
    foreach ($tableLines as $rIndex => $row) {
        $x = $totalWidth - $padding;
        if ($rIndex === 0) {
            imagefilledrectangle($im, (int) $tableLeft, (int) $y - $fontSize - 6, (int) $totalWidth - $padding, (int) $y + $lineHeight - $fontSize + 8, $gray);
        }

        if ($rIndex !== 0) {
            $product_name = $row[1] ?? '';
            $product_sku = getNumberFromText($product_name);
            if ($product_sku === null) {
                print "Error: cannot get sku from product name.\n";
                print_r($row);
                exit();
            }
            $image = findProductBySku($product_sku);
            if ($image === null) {
                print "Error: cannot find the product details or image for SKU {$product_sku}.\n";
                print_r($row);
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
        if (!isset($row[0])) continue;
        if ($row[0] === "3" || $row[0] === "4" || $row[0] === "2") {
            unset($row[0]);
            $row = array_values($row);
        }

        if (isset($row[0]) && str_contains($row[0], "ارسال")) {
            if(isset($row[1])) unset($row[1]);
        }

        $text = implode(' ', $row);
        $bbox = sizeText($fontSize, $fontFile, $text);
        $textW = abs($bbox[2] - $bbox[0]);
        renderText($im, $text, $totalWidth - $padding - $textW, $y, $fontSize, $black, $fontFile);
        $y += $lineHeight;
    }

    $images = array_values(array_filter($images));
    if (count($images) > 3) $images = array_slice($images, 0, 3);

    $localImages = [];
    foreach ($images as $imgUrlOrPath) {
        if (filter_var($imgUrlOrPath, FILTER_VALIDATE_URL)) {
            $local = downloadImage($imgUrlOrPath);
            if ($local !== null) $localImages[] = $local;
        } else {
            if (is_file($imgUrlOrPath) && filesize($imgUrlOrPath) > 0) {
                $localImages[] = $imgUrlOrPath;
            } else {
                $possible = __DIR__ . '/' . ltrim($imgUrlOrPath, '/');
                if (is_file($possible) && filesize($possible) > 0) $localImages[] = $possible;
            }
        }
    }

    $imageAreaHeight = max(0, $totalHeight - $y);
    $imageAreaTop = $y + 10;
    drawImagesAtBottom($im, $localImages, 0, $totalWidth, $imageAreaTop, $imageAreaHeight);

    imagepng($im, $outFile);
    imagedestroy($im);

    echo "Saved: $outFile\n";
    // exit;
}

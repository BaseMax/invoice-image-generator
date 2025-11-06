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

function downloadImage(string $url, string $cacheDir = __DIR__ . '/cache_images'): ?string {
    if ($url === '') return null;
    if (!is_dir($cacheDir)) {
        @mkdir($cacheDir, 0755, true);
    }
    $hash = md5($url);
    foreach (glob("$cacheDir/$hash.*") as $f) {
        if (is_file($f) && filesize($f) > 0) {
            return $f;
        }
    }

    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_CONNECTTIMEOUT => 10,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_USERAGENT => 'Mozilla/5.0 (compatible; image-cache/1.0)',
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => 0,
    ]);
    $data = curl_exec($ch);
    $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $contentType = curl_getinfo($ch, CURLINFO_CONTENT_TYPE);
    $err = curl_error($ch);
    curl_close($ch);

    if ($data === false || $data === '' || $code !== 200) {
        return null;
    }

    $ext = null;
    if ($contentType) {
        if (str_contains($contentType, 'jpeg') || str_contains($contentType, 'jpg')) $ext = 'jpg';
        elseif (str_contains($contentType, 'png')) $ext = 'png';
        elseif (str_contains($contentType, 'gif')) $ext = 'gif';
        elseif (str_contains($contentType, 'webp')) $ext = 'webp';
    }
    if ($ext === null) {
        $finfo = finfo_open(FILEINFO_MIME_TYPE);
        $mime = finfo_buffer($finfo, $data);
        finfo_close($finfo);
        if (str_contains($mime, 'jpeg')) $ext = 'jpg';
        elseif (str_contains($mime, 'png')) $ext = 'png';
        elseif (str_contains($mime, 'gif')) $ext = 'gif';
        elseif (str_contains($mime, 'webp')) $ext = 'webp';
        else $ext = 'jpg';
    }

    $outPath = "$cacheDir/$hash.$ext";
    if (file_put_contents($outPath, $data) === false) {
        return null;
    }
    if (filesize($outPath) === 0) {
        @unlink($outPath);
        return null;
    }
    return $outPath;
}

function drawImagesAtBottom($im, array $imagePaths, int $tableLeft, int $tableRight, int $imageAreaTop, int $imageAreaHeight) {
    $count = count($imagePaths);
    if ($count === 0) return;

    $areaWidth = $tableRight - $tableLeft;
    $colWidth = (int)($areaWidth / $count);
    $paddingInside = 0;

    for ($i = 0; $i < $count; $i++) {
        $local = $imagePaths[$i];
        if (!is_file($local) || filesize($local) === 0) continue;
        $raw = file_get_contents($local);
        if ($raw === false) continue;
        $src = @imagecreatefromstring($raw);
        if ($src === false) continue;

        $srcW = imagesx($src);
        $srcH = imagesy($src);

        $cellLeft = $tableLeft + $i * $colWidth;
        $cellCenterX = (int)($cellLeft + $colWidth / 2);
        $maxW = $colWidth - $paddingInside * 2;
        $maxH = $imageAreaHeight - $paddingInside * 2;

        $ratio = max($maxW / $srcW, $maxH / $srcH); // fill fully, may crop

        $dstW = (int)ceil($srcW * $ratio);
        $dstH = (int)ceil($srcH * $ratio);

        // Center crop
        $srcX = max(0, (int)(($dstW - $maxW) / (2 * $ratio)));
        $srcY = max(0, (int)(($dstH - $maxH) / (2 * $ratio)));
        $cropW = (int)min($srcW - $srcX, $maxW / $ratio);
        $cropH = (int)min($srcH - $srcY, $maxH / $ratio);

        $dstX = (int)($cellCenterX - $dstW / 2);
        $dstY = (int)($imageAreaTop + ($imageAreaHeight - $dstH) / 2);

        $dst = imagecreatetruecolor($dstW, $dstH);
        imagealphablending($dst, false);
        imagesavealpha($dst, true);
        $transparent = imagecolorallocatealpha($dst, 0, 0, 0, 127);
        imagefilledrectangle($dst, 0, 0, $dstW, $dstH, $transparent);

        imagecopyresampled(
            $dst, $src,
            0, 0, 
            $srcX, $srcY,
            $dstW, $dstH,
            $cropW, $cropH
        );
        imagecopy($im, $dst, $dstX, $dstY, 0, 0, $dstW, $dstH);

        imagedestroy($src);
        imagedestroy($dst);
    }
}

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

        if (isset($row[0]) && str_contains($row[0], "ارسال")) {
            break;
        }
        else if (isset($row[1]) && str_contains($row[1], "ارسال")) {
            break;
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

    $tableLeft = $totalWidth - $padding - $tableContentWidth;
    $tableRight = $totalWidth - $padding;
    $images = [];
    foreach ($tableLines as $rIndex => $row) {
        $x = $totalWidth - $padding;
        if ($rIndex === 0) {
            imagefilledrectangle($im, (int) $tableLeft, (int) $y - $fontSize - 6, (int) $totalWidth - $padding, (int) $y + $lineHeight - $fontSize, $gray);
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

    $imageAreaHeight = max(0, $totalHeight - $y - $padding - 50);
    $imageAreaTop = $y + 20;
    drawImagesAtBottom($im, $localImages, 0, $totalWidth, $imageAreaTop, $imageAreaHeight);

    $outFile = $outDir . "table_{$tIndex}.png";
    imagepng($im, $outFile);
    imagedestroy($im);

    echo "Saved: $outFile\n";
    // exit;
}

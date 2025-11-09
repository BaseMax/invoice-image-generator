<?php
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

function num2en(string $text): string
{
    $persian = ['۰','۱','۲','۳','۴','۵','۶','۷','۸','۹'];
    $arabic  = ['٠','١','٢','٣','٤','٥','٦','٧','٨','٩'];
    $english = ['0','1','2','3','4','5','6','7','8','9'];
    
    $text = str_replace($persian, $english, $text);
    $text = str_replace($arabic, $english, $text);
    return $text;
}

function getNumberFromText(string $text): ?string
{
    $text = num2en($text);
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

function drawImagesAtBottom($im, array $imagePaths, int $tableLeft, int $tableRight, int $imageAreaTop, int $imageAreaHeight)
{
    $count = count($imagePaths);
    if ($count === 0) return;

    $areaWidth = $tableRight - $tableLeft;
    $colWidth = $areaWidth / $count;
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

        $cellLeft  = (int)round($tableLeft + $i * $colWidth);
        $cellRight = (int)round($tableLeft + ($i + 1) * $colWidth);
        $cellWidth = $cellRight - $cellLeft;

        $maxW = $cellWidth - $paddingInside * 2;
        $maxH = $imageAreaHeight - $paddingInside * 2;

        $ratio = max($maxW / $srcW, $maxH / $srcH);
        $dstW = (int)ceil($srcW * $ratio);
        $dstH = (int)ceil($srcH * $ratio);

        $srcX = max(0, (int)(($dstW - $maxW) / (2 * $ratio)));
        $srcY = max(0, (int)(($dstH - $maxH) / (2 * $ratio)));
        $cropW = (int)min($srcW - $srcX, $maxW / $ratio);
        $cropH = (int)min($srcH - $srcY, $maxH / $ratio);

        $dst = imagecreatetruecolor($maxW, $maxH);
        imagealphablending($dst, false);
        imagesavealpha($dst, true);
        $transparent = imagecolorallocatealpha($dst, 0, 0, 0, 127);
        imagefilledrectangle($dst, 0, 0, $maxW, $maxH, $transparent);

        imagecopyresampled(
            $dst, $src,
            0, 0,
            $srcX, $srcY,
            $maxW, $maxH,
            $cropW, $cropH
        );

        $dstX = $cellLeft + $paddingInside;
        $dstY = (int)($imageAreaTop + ($imageAreaHeight - $maxH) / 2);

        imagecopy($im, $dst, $dstX, $dstY, 0, 0, $maxW, $maxH);

        imagedestroy($src);
        imagedestroy($dst);
    }
}

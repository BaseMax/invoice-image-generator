<?php
require "_base.php";

// ---------------------------------------------------------
// Bootstrap WordPress
// ---------------------------------------------------------
$wp_load_path = __DIR__ . '/wp-load.php';
if (!file_exists($wp_load_path)) {
    $wp_load_path = __DIR__ . '/../wp-load.php';
}
if (!file_exists($wp_load_path)) {
    die("Could not find wp-load.php. Put this script in WP root or adjust \$wp_load_path.\n");
}
require_once $wp_load_path;
require_once ABSPATH . 'wp-admin/includes/image.php';

// ---------------------------------------------------------
// Autoload PhpSpreadsheet
// ---------------------------------------------------------
$autoload = __DIR__ . '/vendor/autoload.php';
if (!file_exists($autoload)) {
    die("Composer autoload (vendor/autoload.php) not found. Run: composer require phpoffice/phpspreadsheet\n");
}
require_once $autoload;

// ---------------------------------------------------------
// Configuration
// ---------------------------------------------------------
$filePath = "input4.xlsx";
$outDir = "tables4/";
$target_cat_id = 163; // پارچه‌های اوت لت
$tables = extractTables($filePath);
if (empty($tables)) die("No tables found in Excel file.\n");

$created = 0;
$log = [];

// ---------------------------------------------------------
// Helper: create WP attachment from given image
// ---------------------------------------------------------
function create_product_image(string $code, string $title, string $description, string $imagePath) {
    if (!file_exists($imagePath)) {
        return new WP_Error('missing_image', "Image not found: {$imagePath}");
    }

    $upload = wp_upload_dir();
    if (!isset($upload['path'])) {
        return new WP_Error('upload_error', 'Could not get upload dir');
    }

    // $filename = sanitize_file_name('outlet-product-' . $code . '-' . time() . '-' . wp_generate_password(4, false, false) . '.png');
    $filename = sanitize_file_name('outlet-product-' . $code . '.png');
    $filepath = $upload['path'] . '/' . $filename;

    if (!copy($imagePath, $filepath)) {
        return new WP_Error('copy_fail', "Failed to copy image: {$imagePath}");
    }

    $filetype = wp_check_filetype($filename, null);
    $attachment = [
        'guid'           => $upload['url'] . '/' . $filename,
        'post_mime_type' => $filetype['type'] ?? 'image/png',
        'post_title'     => sanitize_text_field($title),
        'post_content'   => $description,
        'post_status'    => 'inherit'
    ];

    $attach_id = wp_insert_attachment($attachment, $filepath);
    if (is_wp_error($attach_id) || !$attach_id) {
        return new WP_Error('attach_fail', 'wp_insert_attachment failed');
    }

    $attach_data = wp_generate_attachment_metadata($attach_id, $filepath);
    wp_update_attachment_metadata($attach_id, $attach_data);

    return $attach_id;
}

// ---------------------------------------------------------
// Iterate over extracted tables
// ---------------------------------------------------------
$number_sofar = 0;
foreach ($tables as $tIndex => $table) {
    $number_sofar++;
    if (empty($table)) continue;
    // if ($number_sofar > 60) {
    //     exit();
    // }
    if ($number_sofar > 120) {
        exit();
    }
    if ($number_sofar <= 60) {
        continue;
    }

    $image = $outDir . "table_{$tIndex}.png";
    if (!file_exists($image)) {
        $log[] = "Skipped table {$tIndex} - image not found.";
        continue;
    }

    $code = getNumberFromText($table[1][0] ?? "0");
    $title = "سبد پارچه تخفیفی شماره " . $code;
    $sku = 'OUT-' . $code;

    $existing = wc_get_product_id_by_sku($sku);
    if ($existing) {
        $product_id = $existing;
        $log[] = "Updating existing product with SKU: {$sku} (ID: {$product_id})";
    } else {
        $product = [
            'post_title'   => $title,
            'post_content' => '',
            'post_status'  => 'publish',
            'post_type'    => 'product',
        ];
        $product_id = wp_insert_post($product);
        if (is_wp_error($product_id) || !$product_id) {
            $log[] = "Failed to insert product for table {$tIndex}";
            continue;
        }
        $log[] = "Created new product: {$title} (ID: {$product_id}, SKU: {$sku})";
    }

    $description = "";
    foreach ($table as $row) {
        $row = array_map('trim', $row);
        if (!implode('', $row)) continue;
        $description .= implode(" | ", $row) . "\n";
    }
    $description = trim($description);

    wp_update_post([
        'ID'           => $product_id,
        'post_title'   => $title,
        'post_content' => $description,
        'post_status'  => 'publish',
    ]);

    wp_set_object_terms($product_id, 'simple', 'product_type');
    wp_set_object_terms($product_id, (int)$target_cat_id, 'product_cat');

    $attach_id = create_product_image($code, $title, $description, $image);
    if (!is_wp_error($attach_id)) {
        set_post_thumbnail($product_id, $attach_id);
    } else {
        $log[] = "Failed to attach image for product {$product_id}: " . $attach_id->get_error_message();
    }

    update_post_meta($product_id, '_manage_stock', 'yes');
    update_post_meta($product_id, '_stock', '1');
    update_post_meta($product_id, '_stock_status', 'instock');

    $fixed_price = 649000;
    update_post_meta($product_id, '_regular_price', $fixed_price);
    update_post_meta($product_id, '_price', $fixed_price);

    update_post_meta($product_id, '_visibility', 'visible');
    update_post_meta($product_id, '_sku', $sku);

    $created++;
    $log[] = "Created product: {$title} (ID: {$product_id}, SKU: {$sku})";
}

// ---------------------------------------------------------
// Output Summary
// ---------------------------------------------------------
$out = "Import finished. Created: {$created}\n";
foreach ($log as $l) $out .= $l . "\n";

if (php_sapi_name() === 'cli') {
    echo $out;
} else {
    header('Content-Type: text/plain; charset=UTF-8');
    echo $out;
}

<?php
$apikey = urlencode("...ghfgh3453fghfghdfg.f.gDJHIGDFFIgOhj4i3oj5s5i6f7j2i3o4234728965234dfgdfg><<<");
$api_url = "https://yasnachap.com/?yasna-api=products&limit=2000&secret_key=$apikey&all=true&page=";
$data_dir = "products/";
$merged_file = "all-products.json";

if (!is_dir($data_dir)) {
    mkdir($data_dir, 0777, true);
}

function fetch_url($url)
{
    $ch = curl_init();
    curl_setopt_array($ch, [
        CURLOPT_URL => $url,
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_CONNECTTIMEOUT => 10,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => false,
        CURLOPT_USERAGENT => "Mozilla/5.0 (PHP cURL bot)",
    ]);
    $response = curl_exec($ch);

    if (curl_errno($ch)) {
        echo "cURL error: " . curl_error($ch) . "\n";
        curl_close($ch);
        return false;
    }

    $status = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($status !== 200) {
        echo "HTTP Error: $status for URL: $url\n";
        return false;
    }

    return $response;
}

$all_products = [];
for ($i = 1; $i <= 50; $i++) {
    $new_api_url = $api_url . $i;
    echo "Fetching page $i ...\n";

    $data = fetch_url($new_api_url);
    if (!$data) {
        echo "Failed to fetch page $i, skipping...\n";
        continue;
    }

    $object = json_decode($data, true);
    if (empty($object)) {
        echo "No more data. End.\n";
        break;
    }

    $file = $data_dir . $i . ".json";
    file_put_contents($file, $data);
    echo "Saved: $file\n";

    $all_products = array_merge($all_products, $object);
}

file_put_contents($merged_file, json_encode($all_products, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));

echo "✅ Done! Merged " . count($all_products) . " products into $merged_file\n";

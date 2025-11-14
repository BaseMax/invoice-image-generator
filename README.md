# Invoice Image Generator

A PHP-based pipeline to generate product images from Excel tables and import them into WooCommerce. This repository handles the full workflow: fetching product data, generating images, and creating WooCommerce products with images automatically.

---

## Workflow Overview

1. **Fetch Products from API**
   `site-data.php` fetches product data from an external API (`yasnachap.com`) and saves all products locally in JSON format (`all-products.json`).

   * Handles pagination automatically (up to 50 pages).
   * Saves each page separately in `products/`.
   * Merges all pages into a single JSON file for later use.

2. **Generate Table Images from Excel**
   `convert.php` reads an Excel file (`input3.xlsx`) and extracts tables.

   * Uses `extractTables()` from `_base.php`.
   * Draws tables as PNG images in `tables3/`.
   * Supports Persian text formatting using `fagd.php`.
   * Adds product images from JSON based on SKU found in table.
   * Handles text, colors, layout, and footer lines automatically.

3. **Import to WooCommerce**
   `create.php` reads the generated PNGs and imports them as WooCommerce products.

   * Checks for existing product by SKU and updates if exists.
   * Creates new products otherwise.
   * Attaches the generated table image as product thumbnail.
   * Sets stock, price (default 649,000), visibility, category, and type.

---

## Directory Structure

```
/invoice-image-generator
├─ tables3/                # Generated table images (PNG)
├─ products/               # API-fetched JSON pages
├─ input3.xlsx             # Excel file containing product tables
├─ all-products.json       # Merged JSON file from API
├─ convert.php             # Script to convert Excel tables to images
├─ create.php              # Script to create WooCommerce products from images
├─ site-data.php           # Script to fetch products from API
├─ fagd.php                # Persian text shaping functions
├─ _base.php               # Helper functions (extractTables, renderText, etc.)
├─ FreeFarsi.ttf           # Font for image generation
├─ Vazirmatn-Regular.ttf   # Additional font
├─ composer.json           # PHP dependencies (PhpSpreadsheet)
└─ README.md
```

---

## Requirements

* PHP 8+
* GD Library (`php-gd`)
* Composer (`phpoffice/phpspreadsheet`)
* WooCommerce installed and active in WordPress
* Write permissions to (`tables1/`, or `tables2/`, or `tables3/`) and `products/`

---

## Installation

1. Clone repository:

```bash
git clone https://github.com/BaseMax/invoice-image-generator.git
cd invoice-image-generator
```

2. Install PHP dependencies:

```bash
composer install
```

3. Place your Excel file (`input3.xlsx`) in the root.

4. Make sure `FreeFarsi.ttf` is present in the root directory.

---

## Usage

### Step 1: Fetch Products

```bash
php site-data.php
```

This will populate `products/` and `all-products.json`.

### Step 2: Generate Table Images

```bash
php convert.php
```

This will generate table images in `tables3/`.

### Step 3: Import into WooCommerce

```bash
php create.php
```

This will create or update products in WooCommerce with attached images.

---

## Notes

* Images are automatically fetched using the SKU from `all-products.json`.
* Tables are drawn in PNG format with Persian text support.
* The workflow assumes `input3.xlsx` is properly formatted.

---

**Author:** Seyyed Ali Mohammadiyeh (Max Base)

**License:** MIT, Copyright 2025

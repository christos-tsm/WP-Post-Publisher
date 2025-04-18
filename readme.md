# TSM WP Post Publisher

**Version:** 2.2.1  
**Requires PHP:** 7.2+  
**Requires WordPress:** 5.6+  

## Description

**TSM WP Post Publisher** is a WordPress plugin that automates the creation of posts from external data sources on a daily schedule. It supports CSV, XLSX, Google Sheets, and DOCX files, converting rich DOCX content (including images, lists, tables, and styling) into valid HTML using PHPWord. You can schedule up to two runs per day, trigger manual imports on demand, and review detailed logs of each operation.

![TSM WP Post Publisher Settings](assets/screenshot-settings.png)

## Features

- **Remote File Import**: Fetch data from CSV, XLSX, Google Sheets, or DOCX files via URL.  
- **Scheduled Publishing**: Set up to two cron-based runs per day (server time).  
- **Manual Trigger**: One-click “Run Now” button in admin settings.  
- **DOCX → HTML Conversion**: Full conversion of DOCX documents to clean HTML, preserving images, bold/italic text, nested lists, and tables.  
- **List Rendering Fix**: Correct handling of Word lists via a custom PHPWord HTML writer and post-processing.  
- **Detailed Logging**: All steps and errors logged to `spp.log` for easy troubleshooting.  
- **Featured Images**: Auto-download and attach images specified in data rows.

## Requirements

- PHP 7.2 or higher  
- WordPress 5.6 or higher  
- PHP extensions: `zip`, `xml`, `gd` (for image processing)  
- Composer dependencies installed via `vendor/` (PhpSpreadsheet, PHPWord)

## Installation

1. **Download or Clone** this repository into your WordPress `TSM WP-content/plugins/` directory.  
2. Run `composer install` in the plugin folder to install dependencies.  
3. **Activate** the plugin through **Plugins** > **Installed Plugins** in your WordPress admin.  

## Configuration

1. Go to **Settings** > **Post Publisher** in the WordPress admin menu.  
2. Enter the **File URL** (CSV, XLSX, Google Sheets, or DOCX).  
3. Select up to two **Daily Run Times** (server time).  
4. (Optional) Click **Run Now** to trigger an immediate import.  
5. Review the log file at `TSM WP-content/plugins/TSM WP-post-publisher/spp.log` for details.

## Usage

- **Scheduled Import:** Posts will be created automatically at your specified times. Only rows matching the current date are published.  
- **Manual Import:** Use the **Run Now** button to fetch and publish immediately.  
- **Post Rows:** Each row in the source file should include these headers:  
  - `current_day` (YYYY-MM-DD)  
  - `post_title`  
  - `post_content`  
  - `post_thumbnail_url` (optional)  

## Customization

- **List Conversion Logic:** The file `includes/docx-html-fixers.php` contains the custom `ListItemRun` override and HTML post-processor. You can extend it to handle custom list styles or deeper indentation.  
- **Logging:** Modify `TSM WP_SPP_LOG_PATH` or the `log()` method in `TSM WP_Post_Publisher` to change log location or verbosity.

## Contributing

Contributions are welcome! Please fork the repo and submit a pull request for bug fixes or new features.

## Changelog

### 2.2.1
- Fixed list rendering for DOCX imports using a custom PHPWord HTML writer and post-processing.  

### 2.2.0
- Initial release with CSV, XLSX, and Google Sheets support; DOCX conversion via PHPWord.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.


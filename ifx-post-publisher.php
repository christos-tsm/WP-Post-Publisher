<?php

/**
 * Plugin Name: IFX Post Publisher
 * Description: Creates WordPress posts at custom times daily from a remote CSV, Excel, Google Sheet, or DOCX file, with manual run, detailed logging, and full DOCX→HTML conversion via PHPWord.
 * Version:     2.2.1
 * Author:      IronFX
 * Requires PHP: 7.2
 * Requires at least: 5.6
 */

if (! defined('ABSPATH')) {
    exit;
}

if (! defined('IFX_SPP_LOG_PATH')) {
    define('IFX_SPP_LOG_PATH', plugin_dir_path(__FILE__) . 'spp.log');
}

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Csv  as CsvReader;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Reader\Xls  as XlsReader;
use PhpOffice\PhpWord\IOFactory           as WordIO;

class IFX_Post_Publisher {
    const OPTION_URL   = 'ifx_spp_file_url';
    const OPTION_TIMES = 'ifx_spp_times';
    const CRON_HOOK    = 'ifx_spp_publish_posts';

    public function __construct() {
        add_action('admin_menu', [$this, 'add_admin_menu']);
        add_action('admin_init', [$this, 'register_settings']);
        add_action('update_option_' . self::OPTION_URL, [$this, 'reschedule_crons'], 10, 2);
        add_action('update_option_' . self::OPTION_TIMES, [$this, 'reschedule_crons'], 10, 2);
        add_action(self::CRON_HOOK, [$this, 'publish_posts']);
    }

    /* --------------------------------------------------------------------
	 * Generic helpers
	 * ------------------------------------------------------------------ */

    protected function log($message) {
        $msg = '[' . date('Y-m-d H:i:s') . '] IFX_SPP: ' . $message;
        if (defined('WP_DEBUG') && WP_DEBUG) {
            error_log($msg);
        }
        file_put_contents(IFX_SPP_LOG_PATH, $msg . "\n", FILE_APPEND);
    }

    /* --------------------------------------------------------------------
	 * Admin UI
	 * ------------------------------------------------------------------ */

    public function add_admin_menu() {
        add_options_page(
            'IFX Post Publisher',
            'Post Publisher',
            'manage_options',
            'ifx-spp-settings',
            [$this, 'settings_page']
        );
    }

    public function register_settings() {
        register_setting(
            'ifx-spp',
            self::OPTION_URL,
            [
                'type'              => 'string',
                'sanitize_callback' => 'esc_url_raw',
            ]
        );

        register_setting(
            'ifx-spp',
            self::OPTION_TIMES,
            [
                'type'              => 'array',
                'sanitize_callback' => [$this, 'sanitize_times'],
                'default'           => ['00:00', '17:00'],
            ]
        );
    }

    public function sanitize_times($times) {
        if (! is_array($times)) {
            return [];
        }

        $clean = [];
        foreach ($times as $time) {
            if (preg_match('/^(2[0-3]|[01][0-9]):[0-5][0-9]$/', trim($time))) {
                $clean[] = trim($time);
            }
        }

        return array_slice($clean, 0, 2); // keep max two
    }

    public function settings_page() {
        if (! current_user_can('manage_options')) {
            return;
        }

        /* Manual run --------------------------------------------------- */
        if (isset($_POST['ifx_spp_run_now'])) {
            check_admin_referer('ifx_spp_run_now_action');
            $this->log('Manual run triggered');
            $this->publish_posts();
            echo '<div class="updated"><p>Process completed. Check log for details.</p></div>';
        }
?>
        <div class="wrap">
            <h1>IFX Post Publisher Settings</h1>

            <form method="post" action="options.php">
                <?php
                settings_fields('ifx-spp');
                do_settings_sections('ifx-spp-settings');
                ?>
                <table class="form-table">
                    <tr>
                        <th>File URL (CSV, XLSX, DOCX &amp; Google Sheet)</th>
                        <td>
                            <input
                                type="url"
                                name="<?php echo self::OPTION_URL; ?>"
                                value="<?php echo esc_attr(get_option(self::OPTION_URL)); ?>"
                                class="regular-text"
                                required />
                        </td>
                    </tr>
                    <tr>
                        <th>Daily Run Times</th>
                        <td>
                            <?php
                            $times = get_option(self::OPTION_TIMES, []);
                            for ($i = 0; $i < 2; $i++) {
                                $val = $times[$i] ?? '';
                                echo '<input type="time" name="' . self::OPTION_TIMES . '[]" value="' . esc_attr($val) . '"><br>';
                            }
                            ?>
                            <p class="description">Up to two times daily (server time)</p>
                        </td>
                    </tr>
                </table>
                <?php submit_button(); ?>
            </form>

            <form method="post" style="margin-top:1em;">
                <?php wp_nonce_field('ifx_spp_run_now_action'); ?>
                <?php submit_button('Run Now', 'secondary', 'ifx_spp_run_now'); ?>
            </form>

            <h2>Log File</h2>
            <p><code><?php echo IFX_SPP_LOG_PATH; ?></code></p>
        </div>
<?php
    }

    /* --------------------------------------------------------------------
	 * Cron management
	 * ------------------------------------------------------------------ */

    public function reschedule_crons() {
        wp_clear_scheduled_hook(self::CRON_HOOK);

        $times = get_option(self::OPTION_TIMES, []);
        foreach ($times as $time) {
            [$h, $m] = explode(':', $time);

            $ts = mktime((int) $h, (int) $m, 0);
            if ($ts <= time()) {
                $ts = strtotime('tomorrow', $ts);
            }

            wp_schedule_event($ts, 'daily', self::CRON_HOOK);
            $this->log("Scheduled cron at $time (ts $ts)");
        }
    }

    /* --------------------------------------------------------------------
	 * Main worker
	 * ------------------------------------------------------------------ */

    public function publish_posts() {
        $this->log('Starting publish_posts');

        $url = trim(get_option(self::OPTION_URL, ''));
        if (! $url) {
            $this->log('No URL set');
            return;
        }
        $this->log("Original URL: $url");

        /* Google Sheets URL → CSV ------------------------------------ */
        if (strpos($url, 'docs.google.com/spreadsheets') !== false) {
            if (preg_match('@/d/([\w-]+)@', $url, $m)) {
                $gid = 0;
                parse_str(parse_url($url, PHP_URL_QUERY) ?: '', $q);
                if (! empty($q['gid'])) {
                    $gid = $q['gid'];
                }
                $url = "https://docs.google.com/spreadsheets/d/{$m[1]}/export?format=csv&gid={$gid}";
                $this->log("Converted to Sheets CSV: $url");
            } else {
                $this->log('Sheets parse failed');
            }
        }

        /* Download --------------------------------------------------- */
        $resp = wp_remote_get($url, ['timeout' => 60]);
        if (is_wp_error($resp)) {
            $this->log('Fetch error: ' . $resp->get_error_message());
            return;
        }

        $code = wp_remote_retrieve_response_code($resp);
        $this->log("HTTP code: $code");
        if (200 !== $code) {
            return;
        }

        $body = wp_remote_retrieve_body($resp);
        $this->log('Body length: ' . strlen($body));
        if (! strlen($body)) {
            return;
        }

        $tmp = wp_tempnam();
        file_put_contents($tmp, $body);

        $ext = strtolower(pathinfo($url, PATHINFO_EXTENSION));
        if (! $ext && stripos($url, 'format=csv') !== false) {
            $ext = 'csv';
        }
        $this->log("Temp file $tmp, ext=$ext");

        /* Parse rows ------------------------------------------------- */
        $rows = [];
        try {
            if ('csv' === $ext) {
                $reader = new CsvReader();
                $sheet  = $reader->load($tmp)->getActiveSheet();
                $data   = $sheet->toArray(null, true, true, true);

                $header = array_shift($data);
                foreach ($data as $row) {
                    $rows[] = array_combine($header, array_values($row));
                }
            } elseif (in_array($ext, ['xls', 'xlsx'], true)) {
                $reader = ('xls' === $ext) ? new XlsReader() : new XlsxReader();
                $sheet  = $reader->load($tmp)->getActiveSheet();
                $data   = $sheet->toArray(null, true, true, true);

                $header = array_shift($data);
                foreach ($data as $row) {
                    $entry = [];
                    foreach ($header as $col => $label) {
                        $key = trim((string) $label);
                        if ($key) {
                            $entry[$key] = $row[$col] ?? '';
                        }
                    }
                    $rows[] = $entry;
                }
            } elseif (preg_match('/\.docx$/i', $url)) {
                $html   = $this->convert_docx_to_html($tmp);
                $rows[] = [
                    'post_title'         => 'DOCX Import ' . date('Y-m-d H:i:s'),
                    'post_thumbnail_url' => '',
                    'post_content'       => $html,
                    'current_day'        => date('Y-m-d'),
                ];
            } else {
                $this->log("Unsupported extension: $ext");
            }
        } catch (Exception $e) {
            $this->log('Parse error: ' . $e->getMessage());
        }

        unlink($tmp);
        $this->log('Parsed rows: ' . count($rows));

        /* Insert posts ---------------------------------------------- */
        foreach ($rows as $i => $data) {
            $day   = trim($data['current_day'] ?? '');
            $title = sanitize_text_field($data['post_title'] ?? '');

            $this->log("Row$i: day=$day, title=$title");

            if ($day !== date('Y-m-d')) {
                $this->log("Skip$i: date mismatch");
                continue;
            }

            if (! $title || post_exists($title)) {
                $this->log("Skip$i: duplicate title");
                continue;
            }

            /* Content ------------------------------------------------ */
            $content = $data['post_content'] ?? '';
            if (preg_match('/\.docx$/i', $content)) {
                $tmp2    = wp_tempnam();
                $docBody = wp_remote_retrieve_body(wp_remote_get($content));
                file_put_contents($tmp2, $docBody);
                $content = $this->convert_docx_to_html($tmp2);
                unlink($tmp2);
            }

            $postId = wp_insert_post(
                [
                    'post_title'   => $title,
                    'post_content' => $content,
                    'post_status'  => 'publish',
                    'post_type'    => 'post',
                ]
            );

            if (is_wp_error($postId)) {
                $this->log("Insert error for row $i: " . $postId->get_error_message());
                continue;
            }
            $this->log("Inserted post $postId");

            /* Featured image ---------------------------------------- */
            $img = trim($data['post_thumbnail_url'] ?? '');
            if ($img) {
                $this->fetch_and_attach_image($img, $postId, $i);
            }
        }

        $this->log('publish_posts complete');
    }

    /* --------------------------------------------------------------------
	 * DOCX → HTML (rewritten, no stray code)
	 * ------------------------------------------------------------------ */
    /* --------------------------------------------------------------------
     * DOCX → HTML Conversion with Local Mammoth.js
     * ------------------------------------------------------------------ */

    protected function convert_docx_to_html($file_path) {
        $this->log('Starting DOCX conversion');

        // 1. Extract raw XML content for list processing
        $zip = new ZipArchive();
        if ($zip->open($file_path) !== true) {
            $this->log('Failed to open DOCX file');
            return '';
        }

        $xml_content = $zip->getFromName('word/document.xml');
        $zip->close();

        // 2. Process lists through XML
        $list_data = $this->process_lists_via_xml($xml_content);
        $this->log('List data extracted: ' . print_r($list_data, true));

        // 3. Generate base HTML with PHPWord
        $phpWord = WordIO::load($file_path);
        $htmlWriter = new \PhpOffice\PhpWord\Writer\HTML($phpWord);
        $html = $htmlWriter->getContent();

        // 4. Merge list structure into HTML
        return $this->inject_list_structure($html, $list_data);
    }

    private function process_lists_via_xml($xml_content) {
        $dom = new DOMDocument();
        $dom->loadXML($xml_content);
        $xpath = new DOMXPath($dom);
        $xpath->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');

        $lists = [];
        $current_level = 0;

        foreach ($xpath->query('//w:p') as $index => $paragraph) {
            $is_list = $xpath->query('.//w:numPr', $paragraph)->length > 0;
            $level = (int)$xpath->evaluate('number(.//w:ilvl/@w:val)', $paragraph) ?: 0;
            $text = $this->get_paragraph_text($xpath, $paragraph);

            $lists[] = [
                'index' => $index,
                'is_list' => $is_list,
                'level' => $level,
                'text' => $text
            ];
        }

        return $lists;
    }

    private function inject_list_structure($html, $list_data) {
        $dom = new DOMDocument();
        @$dom->loadHTML(
            mb_convert_encoding($html, 'HTML-ENTITIES', 'UTF-8'),
            LIBXML_HTML_NOIMPLIED | LIBXML_HTML_NODEFDTD
        );

        $xpath = new DOMXPath($dom);
        $paragraphs = $xpath->query('//p');
        $current_ul = null;

        foreach ($paragraphs as $index => $p) {
            if (!isset($list_data[$index])) continue;

            $list_item = $list_data[$index];
            $this->log("Processing paragraph {$index}: " . print_r($list_item, true));

            // Clean bullet characters and whitespace
            $clean_text = preg_replace('/^[•▪▶]\s*/u', '', $list_item['text']);

            if ($list_item['is_list'] && !empty(trim($clean_text))) {
                // Create new UL if no active list or after non-list item
                if (!$current_ul) {
                    $current_ul = $dom->createElement('ul');
                    $p->parentNode->insertBefore($current_ul, $p);
                    $this->log("Created new UL for list items");
                }

                // Create LI with cleaned text
                $li = $dom->createElement('li', htmlspecialchars($clean_text));
                $current_ul->appendChild($li);
                $this->log("Added LI: $clean_text");

                // Remove original paragraph
                $p->parentNode->removeChild($p);
            } else {
                // Reset UL context for non-list items
                if ($current_ul) {
                    $this->log("Closing UL after non-list item");
                    $current_ul = null;
                }
            }
        }

        // Final cleanup
        $html = $dom->saveHTML();
        $html = preg_replace('/<ul>\s*<li>:marker<\/li>\s*<\/ul>/i', '', $html); // Remove marker artifacts
        return preg_replace('/<p>\s*<\/p>/', '', $html); // Remove empty paragraphs
    }

    // Helper function with improved text extraction
    private function get_paragraph_text($xpath, $paragraph) {
        $text = '';
        foreach ($xpath->query('.//w:t', $paragraph) as $node) {
            $text .= $node->nodeValue;
        }
        return trim(str_replace(["\n", "\t"], ' ', $text));
    }

    /* --------------------------------------------------------------------
	 * Media helper
	 * ------------------------------------------------------------------ */

    protected function fetch_and_attach_image($url, $postId, $i) {
        $this->log("Row{$i}: fetching image {$url}");

        require_once ABSPATH . 'wp-admin/includes/file.php';
        require_once ABSPATH . 'wp-admin/includes/media.php';
        require_once ABSPATH . 'wp-admin/includes/image.php';

        $tmp = download_url($url);
        if (is_wp_error($tmp)) {
            $this->log('Image download error: ' . $tmp->get_error_message());
            return;
        }

        $file = ['name' => basename($url), 'tmp_name' => $tmp];
        $aid  = media_handle_sideload($file, $postId);

        if (is_wp_error($aid)) {
            $this->log('Media sideload error: ' . $aid->get_error_message());
            @unlink($tmp);
            return;
        }

        set_post_thumbnail($postId, $aid);
        $this->log("Row{$i}: featured image set {$aid}");
    }

    /* --------------------------------------------------------------------
	 * Activation / de‑activation
	 * ------------------------------------------------------------------ */

    public static function activation() {
        (new self())->reschedule_crons();
    }

    public static function deactivation() {
        wp_clear_scheduled_hook(self::CRON_HOOK);
    }
}

/* ------------------------------------------------------------------------
 * Bootstrapping
 * --------------------------------------------------------------------- */

register_activation_hook(__FILE__, ['IFX_Post_Publisher', 'activation']);
register_deactivation_hook(__FILE__, ['IFX_Post_Publisher', 'deactivation']);
new IFX_Post_Publisher();

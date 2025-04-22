<?php

/**
 * Plugin Name: TSM WP Post Publisher
 * Description: Creates WordPress posts at custom times daily from a remote CSV, Excel, Google Sheet, or DOCX file, with manual run, detailed logging, and full DOCX→HTML conversion via PHPWord.
 * Version:     2.2.1
 * Author:      Christos TSM
 * Requires PHP: 7.2
 * Requires at least: 5.6
 */

if (! defined('ABSPATH')) {
    exit;
}

if (! defined('TSM_WP_SPP_LOG_PATH')) {
    define('TSM_WP_SPP_LOG_PATH', plugin_dir_path(__FILE__) . 'spp.log');
}

require_once __DIR__ . '/vendor/autoload.php';
require_once __DIR__ . '/includes/docx-html-fixers.php';

use PhpOffice\PhpSpreadsheet\Reader\Csv  as CsvReader;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Reader\Xls  as XlsReader;
use PhpOffice\PhpWord\IOFactory           as WordIO;

class TSM_WP_Post_Publisher {
    const OPTION_URL   = 'TSM_WP_spp_file_url';
    const OPTION_TIMES = 'TSM_WP_spp_times';
    const CRON_HOOK    = 'TSM_WP_spp_publish_posts';

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
        $msg = '[' . date('Y-m-d H:i:s') . '] TSM_WP_SPP: ' . $message;
        if (defined('WP_DEBUG') && WP_DEBUG) {
            error_log($msg);
        }
        file_put_contents(TSM_WP_SPP_LOG_PATH, $msg . "\n", FILE_APPEND);
    }

    /* --------------------------------------------------------------------
	 * Admin UI
	 * ------------------------------------------------------------------ */

    public function add_admin_menu() {
        add_options_page(
            'TSM_WP Post Publisher',
            'Post Publisher',
            'manage_options',
            'TSM_WP-spp-settings',
            [$this, 'settings_page']
        );
    }

    public function register_settings() {
        register_setting(
            'TSM_WP-spp',
            self::OPTION_URL,
            [
                'type'              => 'string',
                'sanitize_callback' => 'esc_url_raw',
            ]
        );

        register_setting(
            'TSM_WP-spp',
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
        if (isset($_POST['TSM_WP_spp_run_now'])) {
            check_admin_referer('TSM_WP_spp_run_now_action');
            $this->log('Manual run triggered');
            $this->publish_posts();
            echo '<div class="updated"><p>Process completed. Check log for details.</p></div>';
        }
?>
        <div class="wrap">
            <h1>TSM WP Post Publisher Settings</h1>

            <form method="post" action="options.php">
                <?php
                settings_fields('TSM_WP-spp');
                do_settings_sections('TSM_WP-spp-settings');
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
                <?php wp_nonce_field('TSM_WP_spp_run_now_action'); ?>
                <?php submit_button('Run Now', 'secondary', 'TSM_WP_spp_run_now'); ?>
            </form>

            <h2>Log File</h2>
            <p><code><?php echo TSM_WP_SPP_LOG_PATH; ?></code></p>
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

            $meta_title = $data['meta_title'] ?? '';

            $meta_description = $data['meta_description'] ?? '';

            $postId = wp_insert_post(
                [
                    'post_title'   => $title,
                    'post_content' => $content,
                    'post_status'  => 'publish',
                    'post_type'    => 'post',
                ]
            );

            update_post_meta($postId, '_yoast_wpseo_title', $meta_title);
            update_post_meta($postId, '_yoast_wpseo_metadesc', $meta_description);

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
	 * DOCX → HTML (using PHPWord + list fixer)
	 * ------------------------------------------------------------------ */
    protected function convert_docx_to_html($file_path) {
        $this->log('Starting DOCX conversion');

        // Load with PHPWord
        $phpWord    = WordIO::load($file_path);
        $htmlWriter = new \PhpOffice\PhpWord\Writer\HTML($phpWord);
        $rawHtml    = $htmlWriter->getContent();

        // Rebuild valid lists
        return \TSM_WP_SPP\DocxHtml\ListHtmlFixer::rebuildLists($rawHtml);
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

register_activation_hook(__FILE__, ['TSM_WP_Post_Publisher', 'activation']);
register_deactivation_hook(__FILE__, ['TSM_WP_Post_Publisher', 'deactivation']);
new TSM_WP_Post_Publisher();

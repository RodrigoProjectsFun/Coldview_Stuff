"""
watcher.py — Teams-Integrated Grafana Monitoring Screenshot Tool

Architecture:
  - TriggerWatcher (background thread): Polls OneDrive/WatcherTriggers/ for
    trigger files created by Power Automate when a user sends !grafana
  - WebsiteMonitor (main thread): Manages DrissionPage browser, refreshes
    Grafana dashboard, takes section screenshots of panels matching
    target texts (POS, AFC3, AFC7)
  - Screenshots saved to OneDrive/WatcherResponses/ for Power Automate
    to pick up and post as a sharing link to the Teams channel
"""

import os
import sys
import json
import logging
import time
import threading
import queue
from datetime import datetime
from urllib.parse import urlparse
from DrissionPage import ChromiumPage, ChromiumOptions
from PIL import Image
import httpx

# ─── LOGGING ───────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('grafana_screens.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

CONFIG_PATH = 'watcher_config.json'


# ─── CONFIGURATION ────────────────────────────────────────────────────

def load_config(path=CONFIG_PATH):
    """Load and validate configuration from JSON file."""
    if not os.path.exists(path):
        logger.error(f"Config file not found: {path}")
        logger.info("Copy watcher_config.example.json to watcher_config.json and fill in your values")
        sys.exit(1)

    with open(path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # Validate required fields
    required = [
        ('website', 'url'),
        ('onedrive', 'triggers_folder'),
        ('onedrive', 'responses_folder'),
    ]
    for section, key in required:
        if section not in config or key not in config[section]:
            logger.error(f"Missing required config: {section}.{key}")
            sys.exit(1)

    return config


def validate_url(url: str) -> bool:
    """Validate URL format (accessibility check skipped for login-required sites)."""
    try:
        result = urlparse(url)
        if not all([result.scheme, result.netloc]):
            logger.error(f"Invalid URL format: {url}")
            return False
        if result.scheme not in ['http', 'https']:
            logger.error(f"Unsupported URL scheme: {result.scheme}")
            return False
        return True
    except Exception as e:
        logger.error(f"Error parsing URL: {e}")
        return False


# ─── TRIGGER WATCHER (Background Thread) ──────────────────────────────

class TriggerWatcher:
    """
    Background thread that polls the OneDrive triggers folder for new
    .json files created by Power Automate when a user sends !grafana.
    """

    def __init__(self, triggers_folder, poll_interval, trigger_queue):
        self.triggers_folder = triggers_folder
        self.poll_interval = poll_interval
        self.trigger_queue = trigger_queue
        self._stop_event = threading.Event()
        self._thread = None

    def start(self):
        """Start polling in a background thread."""
        os.makedirs(self.triggers_folder, exist_ok=True)
        self._thread = threading.Thread(
            target=self._poll_loop, daemon=True, name='TriggerWatcher'
        )
        self._thread.start()
        logger.info(
            f"TriggerWatcher started — polling '{self.triggers_folder}' "
            f"every {self.poll_interval}s"
        )

    def stop(self):
        """Signal the polling thread to stop."""
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=5)

    def _poll_loop(self):
        """Main polling loop."""
        while not self._stop_event.is_set():
            try:
                self._check_for_triggers()
            except Exception as e:
                logger.error(f"Error in trigger polling: {e}")
            self._stop_event.wait(self.poll_interval)

    def _check_for_triggers(self):
        """Check the triggers folder for new .json files."""
        if not os.path.exists(self.triggers_folder):
            return

        for filename in sorted(os.listdir(self.triggers_folder)):
            if not filename.lower().endswith('.json'):
                continue

            filepath = os.path.join(self.triggers_folder, filename)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    trigger_data = json.load(f)

                # Add processing metadata
                trigger_data['_filename'] = filename
                trigger_data['_received_at'] = datetime.now().isoformat()

                self.trigger_queue.put(trigger_data)
                logger.info(
                    f"🔔 Trigger detected: {filename} "
                    f"(from {trigger_data.get('sender', 'unknown')})"
                )

                # Delete processed trigger file
                os.remove(filepath)

            except json.JSONDecodeError as e:
                logger.warning(f"Invalid JSON in trigger file {filename}: {e}")
                bad_path = filepath + '.bad'
                os.rename(filepath, bad_path)
            except PermissionError:
                # File might still be syncing from OneDrive
                logger.debug(f"File {filename} is locked, will retry next cycle")
            except Exception as e:
                logger.error(f"Error processing trigger {filename}: {e}")


# ─── WEBSITE MONITOR (Main Thread) ────────────────────────────────────

class WebsiteMonitor:
    """
    Manages the DrissionPage browser instance. Navigates to the Grafana
    dashboard, handles login, and takes section screenshots on demand.
    """

    def __init__(self, config):
        self.config = config
        self.page = None
        self.is_logged_in = False
        self.screenshots_taken = 0
        self.screenshots_failed = 0

    def initialize(self):
        """Start browser and navigate to the dashboard."""
        co = ChromiumOptions()
        if self.config.get('browser', {}).get('headless', False):
            co.headless(True)

        logger.info("Starting Chromium browser (DrissionPage)...")
        self.page = ChromiumPage(co)
        logger.info("Browser started successfully")

        # Navigate to dashboard
        url = self.config['website']['url']
        logger.info(f"Navigating to {url}...")
        self.page.get(url, timeout=30)

        # Handle login if configured
        login_cfg = self.config['website'].get('login', {})
        if login_cfg.get('username') and login_cfg.get('password'):
            self._handle_login(login_cfg)
        else:
            logger.info("No login credentials configured, assuming no auth needed")

        logger.info("Dashboard loaded — ready for triggers")

    def _handle_login(self, login_cfg):
        """Handle Grafana login form."""
        try:
            username_sel = login_cfg.get('username_selector', "css:input[name='user']")
            password_sel = login_cfg.get('password_selector', "css:input[name='password']")
            button_sel = login_cfg.get('login_button_selector', "css:button[type='submit']")
            wait_after = login_cfg.get('wait_after_login_seconds', 5)

            # Check if login form is present
            username_field = self.page.ele(username_sel, timeout=5)
            if not username_field:
                logger.info("No login form found — may already be logged in")
                self.is_logged_in = True
                return

            logger.info("Login form detected, entering credentials...")
            username_field.clear()
            username_field.input(login_cfg['username'])

            password_field = self.page.ele(password_sel)
            password_field.clear()
            password_field.input(login_cfg['password'])

            login_button = self.page.ele(button_sel)
            login_button.click()

            # Wait for dashboard to load after login
            time.sleep(wait_after)
            self.is_logged_in = True
            logger.info("✅ Login successful")

        except Exception as e:
            logger.error(f"❌ Login failed: {e}")
            raise

    def _find_first_target_element(self):
        """
        Find the first element on the page matching any of the target texts
        (POS, AFC3, AFC7). Returns the element or None.
        """
        target_texts = self.config['website'].get('target_texts', ['POS', 'AFC3', 'AFC7'])

        for text in target_texts:
            try:
                element = self.page.ele(f'text:{text}', timeout=3)
                if element:
                    logger.info(f"Found target element with text '{text}'")
                    return element, text
            except Exception:
                continue

        logger.warning("No target elements found on the page")
        return None, None

    def _find_panel_parent(self, element):
        """
        Try to find the parent Grafana panel container for an element.
        Works with common Grafana panel CSS classes.
        """
        # Try various Grafana panel selectors (covers v7-v11)
        panel_selectors = [
            'css:div[class*="panel-container"]',
            'css:div[class*="react-grid-item"]',
            'css:div[class*="panel-wrapper"]',
            'css:div[data-panelid]',
        ]

        for selector in panel_selectors:
            try:
                panel = element.parent(selector)
                if panel:
                    return panel
            except Exception:
                continue

        # Fallback: go up several levels to find a reasonable container
        try:
            parent = element.parent('tag:div', index=5)
            if parent:
                return parent
        except Exception:
            pass

        return None

    def take_section_screenshot(self, trigger_data):
        """
        Refresh the page, find the section starting at the first target
        element (POS/AFC3/AFC7), and take a full-page screenshot cropped
        from that element downward. This captures "everything from there
        to below" as requested.
        """
        responses_folder = self.config['onedrive']['responses_folder']
        os.makedirs(responses_folder, exist_ok=True)

        # 1. Refresh the page to get latest monitoring data
        logger.info("Refreshing page for latest data...")
        self.page.refresh()
        wait_time = self.config['website'].get('wait_after_refresh_seconds', 5)
        time.sleep(wait_time)

        # 2. Find the first target element
        element, matched_text = self._find_first_target_element()
        if not element:
            logger.error("Cannot take screenshot — no target elements found")
            self.screenshots_failed += 1
            return None

        # 3. Try to find the parent panel for better positioning
        panel = self._find_panel_parent(element)
        anchor = panel if panel else element

        # 4. Get the anchor element's page-absolute Y position
        try:
            # Scroll element into view first
            anchor.scroll.to_see()
            time.sleep(1)  # Brief settle after scroll

            # Get element's absolute position on the page
            rect = anchor.rect
            y_start = int(rect.page_location[1])  # (x, y) tuple
            logger.info(f"Target section starts at Y={y_start}px")
        except Exception as e:
            logger.warning(f"Could not get element position: {e}. Taking full screenshot.")
            y_start = 0

        # 5. Clean up old screenshots — only keep the latest
        for old_file in os.listdir(responses_folder):
            if old_file.lower().endswith('.png'):
                try:
                    os.remove(os.path.join(responses_folder, old_file))
                except Exception:
                    pass

        # 6. Take full-page screenshot (fixed name = always overwrite)
        temp_path = os.path.join(responses_folder, '_temp_full.png')
        final_filename = 'grafana_latest.png'
        final_path = os.path.join(responses_folder, final_filename)

        try:
            self.page.get_screenshot(path=temp_path, full_page=True)
            logger.info("Full-page screenshot taken")
        except Exception as e:
            logger.error(f"Failed to take full-page screenshot: {e}")
            self.screenshots_failed += 1
            return None

        # 7. Crop from the target element's Y position to the bottom
        try:
            if y_start > 0:
                with Image.open(temp_path) as img:
                    # Account for device pixel ratio
                    try:
                        dpr = self.page.run_js('return window.devicePixelRatio || 1')
                        dpr = float(dpr) if dpr else 1.0
                    except Exception:
                        dpr = 1.0

                    crop_y = int(y_start * dpr)
                    crop_y = min(crop_y, img.height - 1)

                    cropped = img.crop((0, crop_y, img.width, img.height))
                    cropped.save(final_path, optimize=True)

                logger.info(f"📸 Cropped screenshot saved (from Y={crop_y}px, DPR={dpr})")
            else:
                # No cropping needed, just rename
                os.rename(temp_path, final_path)
                logger.info(f"📸 Full screenshot saved: {final_filename}")

            # Clean up temp file if it still exists
            if os.path.exists(temp_path):
                os.remove(temp_path)

            self.screenshots_taken += 1
            return final_path

        except Exception as e:
            logger.error(f"Failed to crop screenshot: {e}")
            if os.path.exists(temp_path):
                os.rename(temp_path, final_path)
                self.screenshots_taken += 1
                return final_path
            self.screenshots_failed += 1
            return None

    def shutdown(self):
        """Close the browser cleanly."""
        if self.page:
            try:
                self.page.quit()
                logger.info("Browser closed")
            except Exception as e:
                logger.warning(f"Error closing browser: {e}")


# ─── MAIN ──────────────────────────────────────────────────────────────

def run_test(config):
    """
    Run a quick self-test: validate config, check folders, open browser,
    find target elements, take a test screenshot.
    """
    logger.info("=== RUNNING SELF-TEST ===")

    # 1. Check folders
    for folder_key in ['triggers_folder', 'responses_folder']:
        folder = config['onedrive'][folder_key]
        os.makedirs(folder, exist_ok=True)
        logger.info(f"✅ Folder OK: {folder}")

    # 2. Validate URL
    url = config['website']['url']
    if not validate_url(url):
        logger.error("❌ URL validation failed")
        return False
    logger.info(f"✅ URL format valid: {url}")

    # 3. Open browser and navigate
    monitor = WebsiteMonitor(config)
    try:
        monitor.initialize()
        logger.info("✅ Browser initialized and page loaded")

        # 4. Check for target elements
        element, text = monitor._find_first_target_element()
        if element:
            logger.info(f"✅ Target element found: '{text}'")
        else:
            logger.warning("⚠️  No target elements found (page may need different selectors)")

        # 5. Take test screenshot
        test_trigger = {'sender': 'self-test', 'timestamp': datetime.now().isoformat()}
        result = monitor.take_section_screenshot(test_trigger)
        if result:
            logger.info(f"✅ Test screenshot saved: {result}")
        else:
            logger.warning("⚠️  Test screenshot failed")

        # 6. Test trigger watcher with a dummy file
        triggers_folder = config['onedrive']['triggers_folder']
        test_trigger_path = os.path.join(triggers_folder, 'test_trigger.json')
        with open(test_trigger_path, 'w') as f:
            json.dump({'sender': 'test', 'timestamp': datetime.now().isoformat()}, f)

        tq = queue.Queue()
        watcher = TriggerWatcher(triggers_folder, 1, tq)
        watcher.start()
        time.sleep(3)
        watcher.stop()

        if not tq.empty():
            logger.info("✅ TriggerWatcher picked up test trigger")
        else:
            logger.warning("⚠️  TriggerWatcher did not detect test trigger")

        logger.info("=== SELF-TEST COMPLETE ===")
        return True

    except Exception as e:
        logger.error(f"❌ Self-test failed: {e}")
        return False
    finally:
        monitor.shutdown()


def main():
    """Main entry point."""
    # Load configuration
    config = load_config()

    # Handle --test flag
    if '--test' in sys.argv:
        success = run_test(config)
        sys.exit(0 if success else 1)

    # Validate URL
    if not validate_url(config['website']['url']):
        logger.error("Invalid URL. Aborting.")
        return

    # Ensure OneDrive folders exist
    for folder_key in ['triggers_folder', 'responses_folder']:
        folder = config['onedrive'][folder_key]
        os.makedirs(folder, exist_ok=True)
        logger.info(f"OneDrive folder ready: {folder}")

    # Initialize browser and navigate to dashboard
    monitor = WebsiteMonitor(config)
    trigger_queue = queue.Queue()
    trigger_watcher = None

    try:
        monitor.initialize()

        # Start the trigger watcher (background thread)
        poll_interval = config.get('browser', {}).get('poll_interval_seconds', 3)
        trigger_watcher = TriggerWatcher(
            config['onedrive']['triggers_folder'],
            poll_interval,
            trigger_queue
        )
        trigger_watcher.start()

        logger.info("=" * 60)
        logger.info("🟢 WATCHER RUNNING — Waiting for !grafana triggers...")
        logger.info(f"   Triggers folder: {config['onedrive']['triggers_folder']}")
        logger.info(f"   Responses folder: {config['onedrive']['responses_folder']}")
        logger.info("   Press Ctrl+C to stop")
        logger.info("=" * 60)

        # Main loop: process triggers from the queue
        while True:
            try:
                trigger = trigger_queue.get(timeout=1)
                sender = trigger.get('sender', 'unknown')
                logger.info(f"⚡ Processing trigger from {sender}...")

                result = monitor.take_section_screenshot(trigger)
                if result:
                    logger.info(
                        f"✅ Screenshot delivered to responses folder "
                        f"(Total: {monitor.screenshots_taken})"
                    )
                else:
                    logger.warning(
                        f"⚠️  Screenshot failed "
                        f"(Failures: {monitor.screenshots_failed})"
                    )

            except queue.Empty:
                # No triggers pending, continue waiting
                continue

    except KeyboardInterrupt:
        logger.warning("Stopped by user (Ctrl+C)")
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
    finally:
        # Shutdown
        logger.info("=== SHUTDOWN ===")
        if trigger_watcher:
            trigger_watcher.stop()
        monitor.shutdown()

        logger.info(f"Screenshots taken: {monitor.screenshots_taken}")
        logger.info(f"Screenshots failed: {monitor.screenshots_failed}")
        logger.info("=== WATCHER STOPPED ===")


if __name__ == '__main__':
    main()
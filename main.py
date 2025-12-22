"""
CRM Desktop Automation Script
=============================
Automates login sequence for a legacy CRM desktop application using PyAutoGUI.

Design Principles:
- OOP architecture with CRMAutomator class managing session state
- Strict separation: all configurable values externalized to config.json
- Resilient design with configurable delays to handle UI timing variability
"""

import json
import os
import subprocess
import sys
import time

import pyautogui


class CRMAutomator:
    """
    Main automation controller for the CRM desktop application.
    
    Manages the complete automation lifecycle: initialization, application launch,
    and login sequence. Designed for extensibility with placeholder methods for
    future search() and download() functionality.
    """
    
    def __init__(self, config_path: str = "config.json"):
        """
        Initialize the CRM automator with configuration from JSON file.
        
        Args:
            config_path: Path to the JSON configuration file.
                         Defaults to 'config.json' in the script's directory.
        
        Raises:
            SystemExit: If config file is missing or invalid.
        """
        # Resolve config path relative to script location for portability
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.config_path = os.path.join(self.script_dir, config_path)
        
        # Load configuration first - everything depends on this
        self.config = self._load_config()
        
        # Configure PyAutoGUI safety and timing settings
        self._configure_pyautogui()
        
        # Session state tracking
        self.app_process = None
        self.is_logged_in = False
    
    def _load_config(self) -> dict:
        """
        Load and validate the configuration file.
        
        Returns:
            dict: Parsed configuration dictionary.
        
        Raises:
            SystemExit: If config file is missing or contains invalid JSON.
        """
        if not os.path.exists(self.config_path):
            print(f"ERROR: Configuration file not found at: {self.config_path}")
            print("Please create a config.json file with the required settings.")
            sys.exit(1)
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            print(f"Configuration loaded successfully from: {self.config_path}")
            return config
        except json.JSONDecodeError as e:
            print(f"ERROR: Invalid JSON in configuration file: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"ERROR: Failed to read configuration file: {e}")
            sys.exit(1)
    
    def _configure_pyautogui(self) -> None:
        """
        Configure PyAutoGUI with safety and timing settings from config.
        
        Safety Settings:
        - FAILSAFE=True: Moving mouse to top-left corner (0,0) aborts the script.
          This is CRITICAL for desktop automation - provides an emergency stop.
        
        - PAUSE: Adds a delay after EVERY PyAutoGUI call. This is essential for
          desktop automation because:
          1. UI elements need time to render and become responsive
          2. Prevents actions from executing faster than the app can process
          3. Makes debugging easier by slowing down visible actions
        """
        # FAILSAFE must always be True in production automation scripts
        # This allows emergency abort by moving mouse to screen corner
        pyautogui.FAILSAFE = self.config.get("failsafe", {}).get("enabled", True)
        
        # Global pause between ALL PyAutoGUI commands
        # 0.5s is a safe default - adjust based on target application's responsiveness
        global_pause = self.config.get("timing", {}).get("global_pause", 0.5)
        pyautogui.PAUSE = global_pause
        
        print(f"PyAutoGUI configured: FAILSAFE={pyautogui.FAILSAFE}, PAUSE={pyautogui.PAUSE}s")
    
    def launch_application(self) -> bool:
        """
        Launch the CRM application using subprocess.
        
        Uses subprocess.Popen for non-blocking execution, allowing the script
        to continue while the application loads. A configurable delay handles
        the application startup time.
        
        Returns:
            bool: True if application launched successfully, False otherwise.
        """
        app_path = self.config.get("application", {}).get("path", "")
        app_args = self.config.get("application", {}).get("arguments", "")
        app_name = self.config.get("application", {}).get("name", "Application")
        load_delay = self.config.get("timing", {}).get("app_load_delay", 5.0)
        
        if not app_path:
            print("ERROR: Application path not specified in config.")
            return False
        
        if not os.path.exists(app_path):
            print(f"ERROR: Application not found at: {app_path}")
            return False
        
        try:
            print(f"Launching {app_name}...")
            # Use Popen for non-blocking execution
            # shell=False is more secure and portable
            cmd = [app_path]
            if app_args:
                cmd.append(app_args)
            self.app_process = subprocess.Popen(cmd)
            
            # Wait for application to fully load
            # This delay is configurable because different systems have varying speeds
            print(f"Waiting {load_delay}s for application to initialize...")
            time.sleep(load_delay)
            
            print(f"{app_name} launched successfully (PID: {self.app_process.pid})")
            return True
            
        except FileNotFoundError:
            print(f"ERROR: Cannot find executable at: {app_path}")
            return False
        except PermissionError:
            print(f"ERROR: Permission denied to execute: {app_path}")
            return False
        except Exception as e:
            print(f"ERROR: Failed to launch application: {e}")
            return False
    
    def _clear_field(self) -> None:
        """
        Clear the current input field before typing.
        
        Uses Ctrl+A (select all) followed by Backspace to ensure the field
        is completely empty. This prevents concatenation errors where new
        text is appended to existing content.
        """
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
    
    def _click_at(self, coord_name: str) -> bool:
        """
        Click at coordinates defined in config by name.
        
        Args:
            coord_name: Key name in config["coordinates"] section.
        
        Returns:
            bool: True if coordinates found and clicked, False otherwise.
        """
        coords = self.config.get("coordinates", {}).get(coord_name, {})
        x = coords.get("x")
        y = coords.get("y")
        
        if x is None or y is None:
            print(f"ERROR: Coordinates for '{coord_name}' not found in config.")
            return False
        
        pyautogui.click(x, y)
        return True
    
    def _type_text(self, text: str) -> None:
        """
        Type text with configurable interval between keystrokes.
        
        A small interval between characters improves reliability with
        applications that process input slowly or have input event buffers.
        
        Args:
            text: The text string to type.
        """
        typing_interval = self.config.get("timing", {}).get("typing_interval", 0.05)
        pyautogui.typewrite(text, interval=typing_interval)
    
    def login(self) -> bool:
        """
        Execute the login sequence using credentials from config.
        
        Sequence:
        1. Click username field
        2. Clear any existing text
        3. Type username
        4. Click password field  
        5. Clear any existing text
        6. Type password
        7. Click login button
        8. Wait for transition
        
        Returns:
            bool: True if login sequence completed, False if error occurred.
        """
        credentials = self.config.get("credentials", {})
        username = credentials.get("username", "")
        password = credentials.get("password", "")
        login_delay = self.config.get("timing", {}).get("login_transition_delay", 3.0)
        
        if not username or not password:
            print("ERROR: Username or password not specified in config.")
            return False
        
        print("Starting login sequence...")
        
        # Step 1: Enter username
        print("  -> Clicking username field...")
        if not self._click_at("username_field"):
            return False
        
        self._clear_field()
        print("  -> Typing username...")
        self._type_text(username)
        
        # Step 2: Enter password
        print("  -> Clicking password field...")
        if not self._click_at("password_field"):
            return False
        
        self._clear_field()
        print("  -> Typing password...")
        # Note: typewrite() doesn't work with special characters
        # For passwords with special chars, consider pyautogui.write() or pyperclip
        self._type_text(password)
        
        # Step 3: Submit
        print("  -> Clicking login button...")
        if not self._click_at("login_button"):
            return False
        
        # Wait for login to process and main screen to load
        print(f"  -> Waiting {login_delay}s for login to complete...")
        time.sleep(login_delay)
        
        self.is_logged_in = True
        print("Login sequence completed successfully.")
        return True
    
    def search_and_extract(self, category_term: str = None, report_term: str = None) -> bool:
        """
        Execute the COLDview Hierarchy Search & Retrieval Workflow.
        
        Performs a two-step "Drill-Down" search by reusing the same search box
        twice to filter the hierarchy before opening the final file.
        
        Workflow Steps:
        Step 1 - Filter by Category (First Write):
            - Focus search box, sanitize, type category_term, ENTER
            - Wait 1.0s, click top result to expand category folders
        
        Step 2 - Filter by Report (Second Write):
            - Focus same search box, sanitize, type report_term, ENTER
            - Wait 1.0s, click top result to select specific report type
        
        Step 3 - Select & Open:
            - Wait 2.0s for grid to populate
            - Click target row, click "Abrir Emisión"
        
        Step 4 - Error Handling:
            - Wait 1.0s, dismiss error popup if present
        
        Args:
            category_term: Category to filter by (e.g., "BCP").
                          If not provided, uses config["search_workflow"]["inputs"]["category_term"].
            report_term: Report code to filter by (e.g., "CTAMAE").
                        If not provided, uses config["search_workflow"]["inputs"]["report_term"].
        
        Returns:
            bool: True if workflow completed successfully, False otherwise.
        """
        # Get search workflow configuration
        workflow_config = self.config.get("search_workflow", {})
        inputs = workflow_config.get("inputs", {})
        coords = workflow_config.get("coords", {})
        
        # Use provided terms or fallback to config
        cat_term = category_term or inputs.get("category_term", "")
        rep_term = report_term or inputs.get("report_term", "")
        
        if not cat_term or not rep_term:
            print("ERROR: Both category_term and report_term are required.")
            print(f"  category_term: '{cat_term}' | report_term: '{rep_term}'")
            return False
        
        typing_interval = self.config.get("timing", {}).get("typing_interval", 0.05)
        
        print("=" * 60)
        print("Module 3: Hierarchy Search & Retrieval")
        print(f"Category Term: {cat_term} -> Report Term: {rep_term}")
        print("=" * 60)
        
        # Get search box coordinates (reused for both steps)
        search_box = coords.get("search_box", {})
        sb_x, sb_y = search_box.get("x"), search_box.get("y")
        
        if sb_x is None or sb_y is None:
            print("ERROR: search_box coordinates not found in config.")
            return False
        
        # Get list result coordinates (reused for both steps)
        list_result = coords.get("list_result_1", {})
        lr_x, lr_y = list_result.get("x"), list_result.get("y")
        
        if lr_x is None or lr_y is None:
            print("ERROR: list_result_1 coordinates not found in config.")
            return False
        
        # ═══════════════════════════════════════════════════════════════════════
        # STEP 1: Filter by Category (First Write)
        # ═══════════════════════════════════════════════════════════════════════
        print("\n" + "─" * 60)
        print("STEP 1: Filter by Category (First Write)")
        print("─" * 60)
        
        # 1a: Focus the search box
        print(f"\n  1a. Clicking 'Catálogo' search box at ({sb_x}, {sb_y})...")
        pyautogui.click(sb_x, sb_y)
        
        # 1b: Sanitize - Clear any existing text
        print("  1b. Sanitizing: Ctrl+A, Backspace...")
        self._clear_field()
        
        # 1c: Type the category term
        print(f"  1c. Typing category term: '{cat_term}'...")
        pyautogui.write(cat_term, interval=typing_interval)
        
        # 1d: Press ENTER to filter
        print("  1d. Pressing ENTER to filter list...")
        pyautogui.press('enter')
        
        # 1e: Wait for list to filter
        print("  1e. Waiting 1.0s for list to filter...")
        time.sleep(1.0)
        
        # 1f: Select top result to expand category folders
        print(f"  1f. Clicking top result at ({lr_x}, {lr_y}) to expand category...")
        pyautogui.click(lr_x, lr_y)
        print("  -> Category expanded")
        
        # ═══════════════════════════════════════════════════════════════════════
        # STEP 2: Filter by Report (Second Write)
        # ═══════════════════════════════════════════════════════════════════════
        print("\n" + "─" * 60)
        print("STEP 2: Filter by Report (Second Write)")
        print("─" * 60)
        
        # 2a: Focus the SAME search box again
        print(f"\n  2a. Clicking 'Catálogo' search box again at ({sb_x}, {sb_y})...")
        pyautogui.click(sb_x, sb_y)
        
        # 2b: Sanitize - Clear category term
        print("  2b. Sanitizing: Ctrl+A, Backspace (removes category term)...")
        self._clear_field()
        
        # 2c: Type the report term
        print(f"  2c. Typing report term: '{rep_term}'...")
        pyautogui.write(rep_term, interval=typing_interval)
        
        # 2d: Press ENTER to filter
        print("  2d. Pressing ENTER to filter list...")
        pyautogui.press('enter')
        
        # 2e: Wait for list to filter again
        print("  2e. Waiting 1.0s for list to filter...")
        time.sleep(1.0)
        
        # 2f: Select top result (specific report type)
        print(f"  2f. Clicking top result at ({lr_x}, {lr_y}) to select report type...")
        pyautogui.click(lr_x, lr_y)
        print("  -> Report type selected, grid should load on right")
        
        # ═══════════════════════════════════════════════════════════════════════
        # STEP 3: Select & Open
        # ═══════════════════════════════════════════════════════════════════════
        print("\n" + "─" * 60)
        print("STEP 3: Select & Open")
        print("─" * 60)
        
        # 3a: Wait for grid to populate
        print("\n  3a. Waiting 2.0s for right-hand grid to populate...")
        time.sleep(2.0)
        print("  -> Grid populated")
        
        # 3b: Select target row in main grid
        grid_row = coords.get("grid_row_target", {})
        gr_x, gr_y = grid_row.get("x"), grid_row.get("y")
        
        if gr_x is None or gr_y is None:
            print("ERROR: grid_row_target coordinates not found in config.")
            return False
        
        print(f"  3b. Clicking target row at ({gr_x}, {gr_y})...")
        pyautogui.click(gr_x, gr_y)
        print("  -> Row selected")
        
        # 3c: Click "Abrir Emisión" button
        open_btn = coords.get("open_btn", {})
        ob_x, ob_y = open_btn.get("x"), open_btn.get("y")
        
        if ob_x is None or ob_y is None:
            print("ERROR: open_btn coordinates not found in config.")
            return False
        
        print(f"  3c. Clicking 'Abrir Emisión' button at ({ob_x}, {ob_y})...")
        pyautogui.click(ob_x, ob_y)
        print("  -> Open command executed")
        
        # ═══════════════════════════════════════════════════════════════════════
        # STEP 4: Error Handling
        # ═══════════════════════════════════════════════════════════════════════
        print("\n" + "─" * 60)
        print("STEP 4: Error Handling (Popup Watchdog)")
        print("─" * 60)
        
        # 4a: Wait for potential popup
        print("\n  4a. Waiting 1.0s for 'Error de configuración Dsn' popup...")
        time.sleep(1.0)
        
        # 4b: Attempt to dismiss popup if present
        popup_ok = coords.get("popup_ok_btn", {})
        po_x, po_y = popup_ok.get("x"), popup_ok.get("y")
        
        if po_x is not None and po_y is not None:
            try:
                print(f"  4b. Clicking 'Aceptar' at ({po_x}, {po_y}) to dismiss popup...")
                pyautogui.click(po_x, po_y)
                print("  -> Popup dismissed (if it was present)")
            except Exception as e:
                print(f"  -> Popup handler completed with note: {e}")
        else:
            print("  4b. No popup_ok_btn coordinates configured, skipping popup handler")
        
        # ═══════════════════════════════════════════════════════════════════════
        # WORKFLOW COMPLETE
        # ═══════════════════════════════════════════════════════════════════════
        print("\n" + "=" * 60)
        print("Hierarchy Search & Retrieval completed successfully!")
        print(f"  Category: {cat_term} -> Report: {rep_term}")
        print("  Report should now be opening.")
        print("=" * 60)
        
        return True
    
    def download(self, record_id: str, output_path: str) -> bool:
        """
        Placeholder for download/export functionality.
        
        This method is structured for future implementation without
        requiring changes to __init__ or login logic.
        
        Args:
            record_id: Identifier of the record to download.
            output_path: File path for the downloaded data.
        
        Returns:
            bool: True if download completed successfully.
        
        Raises:
            NotImplementedError: Method not yet implemented.
        """
        # TODO: Implement download functionality
        # Expected config additions:
        # - coordinates.download_button
        # - coordinates.save_dialog
        # - timing.download_delay
        raise NotImplementedError("download() method not yet implemented")
    
    def close(self) -> None:
        """
        Clean up resources and terminate the application if running.
        """
        if self.app_process is not None:
            try:
                self.app_process.terminate()
                self.app_process.wait(timeout=5)
                print("Application terminated.")
            except Exception as e:
                print(f"Warning: Could not terminate application cleanly: {e}")
        
        self.is_logged_in = False


def main():
    """
    Main entry point demonstrating the CRM automation workflow.
    """
    print("=" * 50)
    print("CRM Desktop Automation Script")
    print("=" * 50)
    
    # Initialize automator - this loads config and sets up PyAutoGUI
    automator = CRMAutomator()
    
    try:
        # Launch the CRM application
        if not automator.launch_application():
            print("Failed to launch application. Exiting.")
            return
        
        # Perform login
        if not automator.login():
            print("Login failed. Exiting.")
            return
        
        print("\n" + "=" * 50)
        print("Login completed. Starting hierarchy search...")
        print("=" * 50)
        
        # Execute the hierarchy search & retrieval workflow (Module 3)
        if not automator.search_and_extract():
            print("Search and extract failed. Exiting.")
            return
        
        # Future extensibility:
        # automator.search_and_extract("CUSTOM_CAT", "CUSTOM_RPT")  # Override config
        # automator.download("12345", "C:\\exports\\record.csv")
        
    except pyautogui.FailSafeException:
        print("\n!!! FAILSAFE TRIGGERED !!!")
        print("Mouse moved to corner - automation aborted for safety.")
    except KeyboardInterrupt:
        print("\nAutomation interrupted by user.")
    finally:
        # Uncomment to auto-close the app when script ends:
        # automator.close()
        pass


if __name__ == "__main__":
    main()

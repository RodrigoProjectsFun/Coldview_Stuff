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
    
    def search(self, query: str) -> bool:
        """
        Placeholder for search functionality.
        
        This method is structured for future implementation without
        requiring changes to __init__ or login logic.
        
        Args:
            query: Search term to look up in the CRM.
        
        Returns:
            bool: True if search completed successfully.
        
        Raises:
            NotImplementedError: Method not yet implemented.
        """
        # TODO: Implement search functionality
        # Expected config additions:
        # - coordinates.search_field
        # - coordinates.search_button
        # - timing.search_result_delay
        raise NotImplementedError("search() method not yet implemented")
    
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
        print("Automation completed successfully!")
        print("The CRM is now ready for use.")
        print("=" * 50)
        
        # Future extensibility example:
        # automator.search("customer name")
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

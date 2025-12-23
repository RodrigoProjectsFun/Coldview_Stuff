"""
Network Folder Watcher with Email Notification
===============================================
Monitors a remote network folder for new files and sends email
notifications via Outlook when files are detected.
"""

import os
import time
import sys
import json

# Import our email sender module
from email_sender import send_email_with_attachment


def load_watcher_config(config_path: str = "config.json") -> dict:
    """Load watcher configuration from config file."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(script_dir, config_path)
    
    with open(full_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    return config.get("watcher", {})


def watch_network_folder(network_path: str = None, check_interval: float = None):
    """
    Monitor a network folder for new files and send email notifications.
    
    Args:
        network_path: Path to watch (uses config if not provided).
        check_interval: Seconds between checks (uses config if not provided).
    """
    # Load config
    watcher_config = load_watcher_config()
    
    # Use provided values or fallback to config
    folder_path = network_path or watcher_config.get("network_folder", "")
    interval = check_interval or watcher_config.get("check_interval_seconds", 2)
    
    if not folder_path:
        print("ERROR: No network folder specified in config or method call.")
        return
    
    # Validate the path exists before starting
    if not os.path.exists(folder_path):
        print(f"ERROR: The path '{folder_path}' cannot be found.")
        print("Make sure the network drive is connected/mapped.")
        return

    print("=" * 50)
    print("Network Folder Watcher with Email Notification")
    print("=" * 50)
    print(f"Target: {folder_path}")
    print(f"Check Interval: {interval}s")
    print("Waiting for new files... (Ctrl+C to stop)\n")
    
    # Initialize the list of known files
    try:
        before = set(os.listdir(folder_path))
        print(f"Initial file count: {len(before)}")
    except OSError as e:
        print(f"Initial connection failed: {e}")
        return

    while True:
        try:
            time.sleep(interval)
            
            # Get the current list of files
            after = set(os.listdir(folder_path))
            
            # Check for additions
            added = after - before
            
            if added:
                for filename in added:
                    full_path = os.path.join(folder_path, filename)
                    
                    print(f"\n[NEW] File detected: {filename}")
                    print("-" * 40)
                    
                    # Send email with the new file attached
                    success = send_email_with_attachment(full_path)
                    
                    if success:
                        print(f"[OK] Email notification sent for: {filename}")
                    else:
                        print(f"[WARN] Email failed for: {filename}")
                    
                    print("-" * 40)
            
            # Update the reference list
            before = after

        except OSError:
            # Handle network drops (e.g., VPN disconnect, Server restart)
            print("\n(!) Network connection lost. Retrying in 5 seconds...")
            time.sleep(5)
            
            if os.path.exists(folder_path):
                print("(+) Connection restored.")
                try:
                    before = set(os.listdir(folder_path))
                except:
                    pass

        except KeyboardInterrupt:
            print("\n\nStopping monitor.")
            sys.exit()


if __name__ == "__main__":
    # Run the watcher - settings loaded from config.json
    watch_network_folder()
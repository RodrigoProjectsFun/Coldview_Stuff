import os
import time
import sys

def watch_network_folder(network_path):
    # 1. Validate the path exists before starting
    if not os.path.exists(network_path):
        print(f"ERROR: The path '{network_path}' cannot be found.")
        print("Make sure the network drive is connected/mapped.")
        return

    print(f"--- Monitoring Remote Folder ---")
    print(f"Target: {network_path}")
    print("Waiting for new files... (Ctrl+C to stop)\n")
    
    # Initialize the list of known files
    try:
        before = set(os.listdir(network_path))
    except OSError as e:
        print(f"Initial connection failed: {e}")
        return

    while True:
        try:
            time.sleep(2) # Check every 2 seconds
            
            # Get the current list of files
            after = set(os.listdir(network_path))
            
            # Check for additions
            added = after - before
            
            if added:
                # Loop through added files (in case multiple were pasted at once)
                for f in added:
                    print(f"[NEW] File added: {f}")
            
            # Update the reference list
            before = after

        except OSError:
            # This handles network drops (e.g., VPN disconnect, Server restart)
            print("(!) Network connection lost. Retrying in 5 seconds...")
            time.sleep(5)
            
            # Optional: Try to reconnect logic or just loop
            if os.path.exists(network_path):
                print("(+) Connection restored.")
                # We reset 'before' to avoid alerting on every existing file again
                # or we keep it to catch up. Usually, resetting is safer to avoid spam.
                try:
                    before = set(os.listdir(network_path))
                except:
                    pass

        except KeyboardInterrupt:
            print("\nStopping monitor.")
            sys.exit()

if __name__ == "__main__":
    # YOU CAN USE A UNC PATH OR A MAPPED DRIVE LETTER
    # Example 1: UNC Path (Recommended for servers)
    remote_folder = r"\\192.168.1.50\SharedDocs\Invoices"
    
    # Example 2: Mapped Drive
    # remote_folder = r"Z:\Invoices"

    # Make sure to update this variable!
    # remote_folder = r"C:\Users\YourName\Desktop\TestFolder" # for testing locally

    watch_network_folder(remote_folder)
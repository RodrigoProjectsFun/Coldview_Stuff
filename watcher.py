import os
import time

def watch_folder(path_to_watch):
    print(f"Monitoring '{path_to_watch}'...")
    
    # Create a set of current files
    before = set(os.listdir(path_to_watch))

    while True:
        time.sleep(2) # Check every 2 seconds
        
        # Get the current list of files
        after = set(os.listdir(path_to_watch))
        
        # Find what is in 'after' but not in 'before'
        added = after - before
        
        if added:
            print("New file added")
            
        # Update the 'before' list for the next check
        before = after

if __name__ == "__main__":
    # Watch the current directory
    watch_folder(".")
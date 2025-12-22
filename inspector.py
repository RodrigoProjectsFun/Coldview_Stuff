import uiautomation as auto
import keyboard
import json
import time
import os

def main():
    saved_locations = {}
    
    # Get the handle of the current console window so we can bring it to front later
    console_window = auto.GetConsoleWindow()
    console_elem = auto.ControlFromHandle(console_window)

    print("--- Coordinate Capture Tool (Focus Fixed) ---")
    print("1. Hover over an element in your app.")
    print("2. Press 'ENTER' to capture.")
    print("3. The script will switch focus here so you can type the name.")
    print("4. Press 'ENTER' to save, then 'ALT+TAB' back to your app.")
    print("5. Press 'ESC' to finish.\n")

    try:
        while True:
            if keyboard.is_pressed('esc'):
                break

            if keyboard.is_pressed('enter'):
                # 1. Capture element info immediately
                element = auto.ControlFromCursor()
                
                # Get coordinates safely
                rect = element.BoundingRectangle
                center_x = int((rect.left + rect.right) / 2)
                center_y = int((rect.top + rect.bottom) / 2)

                # 2. Visual feedback
                try:
                    element.DrawOutline(colour=0xFF0000, thickness=2)
                except:
                    pass

                # 3. CRITICAL: Switch focus back to this python console
                # This ensures your typing goes into the input() prompt, not the app
                if console_elem:
                    console_elem.SetFocus()
                
                # Wait for the 'enter' key to be released so it doesn't skip the input
                while keyboard.is_pressed('enter'):
                    time.sleep(0.05)
                
                # 4. Blocking Input
                # The script will completely pause here until you type a name and hit Enter
                print(f"\n[+] Captured Element at ({center_x}, {center_y})")
                print("    Waiting for name... (Type below)")
                
                key_name = input("    Name this location: ").strip()

                if key_name:
                    saved_locations[key_name] = {
                        "x": center_x,
                        "y": center_y
                    }
                    print(f"    Saved '{key_name}'!")
                else:
                    print("    Skipped (no name entered).")

                print("\n    ...Go back to your app now. (Waiting for next 'ENTER')")

            time.sleep(0.05)

    except KeyboardInterrupt:
        pass

    # Final Output
    print("\n" + "="*30)
    print("FINAL JSON OUTPUT:")
    print("="*30)
    print(json.dumps(saved_locations, indent=4))

    with open("locations.json", "w") as f:
        json.dump(saved_locations, f, indent=4)
    print("\nSaved to locations.json")

if __name__ == "__main__":
    main()
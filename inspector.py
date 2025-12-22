import uiautomation as auto
import keyboard
import json
import time

def main():
    saved_locations = {}
    
    print("--- Coordinate Capture Tool ---")
    print("1. Hover over an element in your app.")
    print("2. Press 'ENTER' to capture its coordinates.")
    print("3. Type the name for the location in this terminal when prompted.")
    print("4. Press 'ESC' to finish and generate the JSON.\n")

    try:
        while True:
            # Check if ESC is pressed to exit
            if keyboard.is_pressed('esc'):
                break

            # Check if ENTER is pressed to capture
            if keyboard.is_pressed('enter'):
                # 1. Capture the element under the mouse
                element = auto.ControlFromCursor()
                
                # 2. Get the center point of the element
                # Using the element center is safer than raw mouse position
                rect = element.GetBoundingRectangle()
                center_x = int((rect.left + rect.right) / 2)
                center_y = int((rect.top + rect.bottom) / 2)

                # 3. Visual feedback (draws a red box briefly)
                rect.Draw(0xFF0000, 1)

                # 4. Prompt user for the key name (e.g., "search_box")
                # We add a small sleep so the 'Enter' press doesn't skip the input prompt
                time.sleep(0.2) 
                
                # Clear any buffered input to prevent glitches
                while keyboard.is_pressed('enter'): pass
                
                print(f"\n[+] Capturing coordinates: x={center_x}, y={center_y}")
                key_name = input("    Enter name for this location (e.g., search_box): ").strip()

                if key_name:
                    saved_locations[key_name] = {
                        "x": center_x,
                        "y": center_y
                    }
                    print(f"    Saved '{key_name}'!")
                else:
                    print("    Skipped (no name provided).")

                print("\nResuming... (Press ESC to finish)")

            time.sleep(0.05)

    except KeyboardInterrupt:
        pass

    # Output the final JSON
    print("\n" + "="*30)
    print("FINAL JSON OUTPUT:")
    print("="*30)
    print(json.dumps(saved_locations, indent=4))

    # Optional: Save to file
    with open("locations.json", "w") as f:
        json.dump(saved_locations, f, indent=4)
    print("\nSaved to locations.json")

if __name__ == "__main__":
    main()
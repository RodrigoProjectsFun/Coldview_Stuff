import uiautomation as auto
import keyboard
import json
import time

def main():
    saved_locations = {}
    
    print("--- Coordinate Capture Tool (Final) ---")
    print("1. Hover over an element.")
    print("2. Press 'ENTER' to capture.")
    print("3. Type name, then press 'ENTER' again.")
    print("4. Press 'ESC' to finish.\n")

    try:
        while True:
            if keyboard.is_pressed('esc'):
                break
            if keyboard.is_pressed('enter'):
                # 1. Capture element
                element = auto.ControlFromCursor()
                # 2. Get coordinates safely
                # Access the rectangle properties directly
                rect = element.BoundingRectangle
                left = rect.left
                top = rect.top
                right = rect.right
                bottom = rect.bottom
                
                # Calculate center
                center_x = int((left + right) / 2)
                center_y = int((top + bottom) / 2)

                # 3. Visual feedback: Use the element's own method instead of the rect's
                try:
                    # Draw a red box around the element
                    element.DrawOutline(colour=0xFF0000, thickness=2)
                except AttributeError:
                    # If this specific control type doesn't support drawing, just skip it
                    pass

                # 4. Input handling
                time.sleep(0.2) 
                while keyboard.is_pressed('enter'): pass
                
                print(f"\n[+] Captured Element: x={center_x}, y={center_y}")
                
                key_name = input("    Name this location (e.g., search_box): ").strip()

                if key_name:
                    saved_locations[key_name] = {
                        "x": center_x,
                        "y": center_y
                    }
                    print(f"    Saved '{key_name}'!")
                else:
                    print("    Skipped.")

                print("\nResuming... (Press ESC to finish)")

            time.sleep(0.05)

    except KeyboardInterrupt:
        pass

    # Final Output
    print("\n" + "="*30)
    print("FINAL JSON OUTPUT:")
    print("="*30)
    print(json.dumps(saved_locations, indent=4))

    # Save to file
    with open("locations.json", "w") as f:
        json.dump(saved_locations, f, indent=4)
    print("\nSaved to locations.json")

if __name__ == "__main__":
    main()
# CRM Desktop Automator

A Python automation script for legacy CRM desktop applications using PyAutoGUI.

## Features

- **OOP Architecture**: Clean `CRMAutomator` class managing session state
- **Configurable**: All values externalized to `config.json`
- **Resilient**: Configurable delays to handle UI timing variability
- **Safe**: Failsafe enabled (move mouse to corner to abort)

## Setup

1. **Install dependencies:**
   ```bash
   pip install pyautogui
   ```

2. **Create your configuration:**
   ```bash
   cp config.example.json config.json
   ```

3. **Edit `config.json`** with your actual values:
   - `application.path`: Path to your CRM executable
   - `credentials`: Your login credentials
   - `coordinates`: Screen X/Y positions for input fields (use `pyautogui.position()` to find these)

## Usage

```bash
python main.py
```

## Finding Screen Coordinates

Open a Python REPL and run:
```python
import pyautogui
import time

time.sleep(3)  # Move your mouse to the target field
print(pyautogui.position())
```

## Safety Features

- **FAILSAFE**: Move mouse to top-left corner (0,0) to abort at any time
- **Global pause**: 0.5s delay between all actions (configurable)
- **Field sanitization**: Auto-clears fields before typing

## License

MIT

# File Watcher

Scheduled file checker that copies files from a watch directory to a destination.

## Quick Start

1. **Run setup:**
   ```
   setup.bat
   ```

2. **Edit config:**
   Open `watcher_config.json` and set your paths:
   ```json
   {
       "watch_dir": "C:\\Your\\Watch\\Path",
       "target_file": "yourfile.csv",
       "dest_dir": "C:\\Your\\Destination",
       "scheduled_times": ["09:00", "12:00", "17:00"],
       "run_on_startup": true
   }
   ```

3. **Run:**
   ```
   python watcher.py
   ```

## Commands

| Command | Description |
|---------|-------------|
| `python watcher.py` | Run scheduler (continuous) |
| `python watcher.py --once` | Single check, then exit |
| `python watcher.py --test` | Test all permissions |

## Run on Startup

### Option 1: Windows Startup Folder
1. Press `Win + R` → type `shell:startup` → Enter
2. Create shortcut to: `pythonw.exe "C:\path\to\watcher.py"`

### Option 2: Task Scheduler (requires admin)
```powershell
schtasks /create /tn "FileWatcher" /tr "pythonw C:\path\to\watcher.py" /sc onlogon
```

## Features

- ✅ Scheduled file checks (configurable times)
- ✅ Network path support with retry on failure
- ✅ Permission testing
- ✅ Logging to `logs/` directory

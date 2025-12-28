@echo off
echo ================================================
echo FILE WATCHER - SETUP
echo ================================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo Please install Python from https://python.org
    pause
    exit /b 1
)
echo [OK] Python found

REM Install dependencies
echo.
echo Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Dependencies installed

REM Check if config exists
if not exist "watcher_config.json" (
    echo.
    echo [WARNING] watcher_config.json not found
    echo Creating from template...
    copy watcher_config.example.json watcher_config.json >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Could not create config file
        echo Please create watcher_config.json manually
    ) else (
        echo [OK] Config file created - EDIT IT with your paths!
    )
) else (
    echo [OK] Config file exists
)

REM Run permission test
echo.
echo Running permission test...
python watcher.py --test
if errorlevel 1 (
    echo.
    echo [WARNING] Some permission tests failed - check config
) else (
    echo.
    echo [OK] All permissions OK
)

echo.
echo ================================================
echo SETUP COMPLETE
echo ================================================
echo.
echo To run manually:  python watcher.py --once
echo To run scheduler: python watcher.py
echo To test perms:    python watcher.py --test
echo.
pause

import os
import shutil
import logging
import socket
import ctypes
import sys
import time
import json
from dataclasses import dataclass
from datetime import datetime

try:
    import schedule
except ImportError:
    print("Missing dependency: schedule")
    sys.exit(1)


pending_retry = False


@dataclass
class Config:
    watch_dir: str
    target_file: str
    dest_dir: str
    scheduled_times: list[str]
    run_on_startup: bool = True
    
    @property
    def target_path(self) -> str:
        return os.path.join(self.watch_dir, self.target_file)
    
    @property
    def dest_path(self) -> str:
        return os.path.join(self.dest_dir, self.target_file)
    
    @staticmethod
    def load(path: str) -> "Config":
        with open(path, 'r') as f:
            data = json.load(f)
        return Config(
            watch_dir=data["watch_dir"],
            target_file=data["target_file"],
            dest_dir=data["dest_dir"],
            scheduled_times=data.get("scheduled_times", ["09:00"]),
            run_on_startup=data.get("run_on_startup", True)
        )


def setup_logging() -> logging.Logger:
    handlers = [logging.StreamHandler()]
    
    try:
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"watcher_{datetime.now().strftime('%Y%m%d')}.log")
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))
    except (PermissionError, OSError):
        pass
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers
    )
    return logging.getLogger(__name__)


logger = setup_logging()


def is_network_path(path: str) -> bool:
    path = os.path.normpath(path)
    if path.startswith("\\\\"):
        return True
    if os.name == 'nt' and len(path) >= 2 and path[1] == ':':
        try:
            return ctypes.windll.kernel32.GetDriveTypeW(path[0].upper() + ":\\") == 4
        except Exception:
            return False
    return False


def is_server_reachable(path: str) -> bool:
    path = os.path.normpath(path)
    if not path.startswith("\\\\"):
        return True
    parts = path.split("\\")
    if len(parts) < 3:
        return True
    server = parts[2]
    try:
        socket.setdefaulttimeout(5)
        for port in [445, 139]:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            if sock.connect_ex((server, port)) == 0:
                sock.close()
                return True
            sock.close()
        return False
    except socket.error:
        return False
    finally:
        socket.setdefaulttimeout(None)


def run_permission_test(config_file: str):
    print("\n" + "=" * 50)
    print("PERMISSION TEST")
    print("=" * 50 + "\n")
    
    all_passed = True
    
    # Test 1: Config file
    if os.path.exists(config_file) and os.access(config_file, os.R_OK):
        print("✅ Config file: readable")
        config = Config.load(config_file)
    else:
        print("❌ Config file: not found or not readable")
        print(f"   Path: {config_file}")
        return False
    
    # Test 2: Log directory
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    if os.path.exists(log_dir):
        if os.access(log_dir, os.W_OK):
            test_file = os.path.join(log_dir, ".write_test")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                print("✅ Log directory: writable")
            except (PermissionError, OSError):
                print("❌ Log directory: not writable (file creation failed)")
                all_passed = False
        else:
            print("❌ Log directory: not writable")
            all_passed = False
    else:
        try:
            os.makedirs(log_dir, exist_ok=True)
            print("✅ Log directory: created and writable")
        except OSError:
            print("❌ Log directory: cannot create")
            all_passed = False
    
    # Test 3: Watch directory - network check
    if is_network_path(config.watch_dir):
        if is_server_reachable(config.watch_dir):
            print("✅ Network server: reachable")
        else:
            print("❌ Network server: unreachable")
            print(f"   Path: {config.watch_dir}")
            all_passed = False
    
    # Test 4: Watch directory - access
    if os.path.exists(config.watch_dir):
        if os.access(config.watch_dir, os.R_OK):
            print("✅ Watch directory: readable")
        else:
            print("❌ Watch directory: not readable")
            all_passed = False
    else:
        print("❌ Watch directory: not found")
        print(f"   Path: {config.watch_dir}")
        all_passed = False
    
    # Test 5: Destination directory
    if os.path.exists(config.dest_dir):
        if os.access(config.dest_dir, os.W_OK):
            print("✅ Destination directory: writable")
        else:
            print("❌ Destination directory: not writable")
            all_passed = False
    else:
        try:
            os.makedirs(config.dest_dir, exist_ok=True)
            print("✅ Destination directory: created and writable")
        except OSError:
            print("❌ Destination directory: cannot create")
            print(f"   Path: {config.dest_dir}")
            all_passed = False
    
    # Summary
    print("\n" + "=" * 50)
    if all_passed:
        print("ALL TESTS PASSED ✅")
    else:
        print("SOME TESTS FAILED ❌")
    print("=" * 50 + "\n")
    
    return all_passed


def check_and_process(config: Config) -> bool:
    global pending_retry
    
    logger.info("=" * 50)
    logger.info(f"CHECK @ {datetime.now().strftime('%H:%M:%S')}")
    logger.info("=" * 50)
    
    # Check network path
    if is_network_path(config.watch_dir) and not is_server_reachable(config.watch_dir):
        logger.error(f"Server unreachable: {config.watch_dir}")
        logger.info("Will retry when network becomes available")
        pending_retry = True
        return False
    
    # Network is available - clear retry flag if it was set
    if pending_retry:
        logger.info("Network recovered - processing pending retry")
        pending_retry = False
    
    # Check watch directory
    if not os.path.exists(config.watch_dir):
        logger.error(f"Watch directory not found: {config.watch_dir}")
        return False
    
    if not os.access(config.watch_dir, os.R_OK):
        logger.error(f"No read access: {config.watch_dir}")
        return False
    
    logger.info(f"Watch directory OK: {config.watch_dir}")
    
    # Check if file exists
    if not os.path.isfile(config.target_path):
        logger.info(f"File not found: {config.target_file}")
        contents = os.listdir(config.watch_dir)[:10]
        if contents:
            logger.info(f"Directory contents: {contents}")
        return False
    
    size = os.path.getsize(config.target_path)
    logger.info(f"Found: {config.target_file} ({size} bytes)")
    
    # Ensure destination exists
    if not os.path.exists(config.dest_dir):
        try:
            os.makedirs(config.dest_dir, exist_ok=True)
            logger.info(f"Created: {config.dest_dir}")
        except OSError as e:
            logger.error(f"Cannot create destination: {e}")
            return False
    
    if not os.access(config.dest_dir, os.W_OK):
        logger.error(f"No write access: {config.dest_dir}")
        return False
    
    # Copy file
    try:
        if os.path.exists(config.dest_path):
            os.remove(config.dest_path)
        
        shutil.copy2(config.target_path, config.dest_path)
        
        if os.path.exists(config.dest_path):
            logger.info(f"Copied to: {config.dest_path}")
            logger.info("SUCCESS!")
            return True
        
        logger.error("Copy verification failed")
        return False
    except Exception as e:
        logger.error(f"Copy failed: {e}")
        return False


def run_with_network_retry(config: Config):
    global pending_retry
    
    # If network is available and there's a pending retry, run it
    if pending_retry and is_network_path(config.watch_dir):
        if is_server_reachable(config.watch_dir):
            logger.info("Network back online - running pending check")
            check_and_process(config)
            return
    
    # Run scheduled jobs
    schedule.run_pending()


def main():
    global pending_retry
    config_file = os.path.join(os.path.dirname(__file__), "watcher_config.json")
    
    # Handle --test flag (permission check only)
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        success = run_permission_test(config_file)
        sys.exit(0 if success else 1)
    
    # Handle --once flag (single run, no scheduler)
    if len(sys.argv) > 1 and sys.argv[1] == "--once":
        if not os.path.exists(config_file):
            print(f"Config file not found: {config_file}")
            sys.exit(1)
        config = Config.load(config_file)
        success = check_and_process(config)
        sys.exit(0 if success else 1)
    
    if not os.path.exists(config_file):
        logger.error(f"Config file not found: {config_file}")
        sys.exit(1)
    
    config = Config.load(config_file)
    
    logger.info("=" * 50)
    logger.info("FILE WATCHER STARTED")
    logger.info(f"Watch: {config.watch_dir}")
    logger.info(f"Target: {config.target_file}")
    logger.info(f"Dest: {config.dest_dir}")
    logger.info(f"Schedule: {config.scheduled_times}")
    logger.info("=" * 50)
    
    # Setup schedule
    for time_str in config.scheduled_times:
        schedule.every().day.at(time_str).do(lambda: check_and_process(config))
        logger.info(f"Scheduled: {time_str}")
    
    # Run on startup if configured
    if config.run_on_startup:
        check_and_process(config)
    
    # Run scheduler loop with network retry support
    logger.info("Scheduler running (Ctrl+C to stop)")
    try:
        while True:
            if pending_retry and is_network_path(config.watch_dir):
                if is_server_reachable(config.watch_dir):
                    logger.info("Network recovered - running pending check")
                    check_and_process(config)
            
            schedule.run_pending()
            time.sleep(60)
    except KeyboardInterrupt:
        logger.info("Stopped")


if __name__ == "__main__":
    main()
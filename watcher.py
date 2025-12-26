"""
File Checker Script - Designed to be run by cron/Task Scheduler
Checks if a target file exists and processes it, or logs why it couldn't be found.
"""
import os
import shutil
import logging
import socket
import ctypes
import sys
from pathlib import Path
from datetime import datetime


# --- CONFIGURATION ---
WATCH_DIRECTORY = r"C:\Users\Name\Downloads\InputFolder"
TARGET_FILENAME = "data.csv"  # The specific file to look for
DESTINATION_DIRECTORY = r"C:\Users\Name\Documents\Processed"

# Retry settings for network paths
NETWORK_RETRY_ATTEMPTS = 3
NETWORK_RETRY_DELAY_SECONDS = 5


# --- LOGGING CONFIGURATION ---
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

# Create logs directory if it doesn't exist
LOG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
os.makedirs(LOG_DIR, exist_ok=True)

# Create log file with timestamp
log_filename = os.path.join(LOG_DIR, f"watcher_{datetime.now().strftime('%Y%m%d')}.log")

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format=LOG_FORMAT,
    datefmt=LOG_DATE_FORMAT,
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler()  # Also print to console
    ]
)
logger = logging.getLogger(__name__)


class NetworkShareValidator:
    """Validates network share connectivity and accessibility."""
    
    @staticmethod
    def is_network_path(path: str) -> bool:
        """Check if the path is a UNC network path or mapped drive."""
        path = os.path.normpath(path)
        
        # Check for UNC path (\\server\share)
        if path.startswith("\\\\"):
            return True
        
        # Check if it's a mapped network drive on Windows
        if os.name == 'nt' and len(path) >= 2 and path[1] == ':':
            drive = path[0].upper() + ":"
            try:
                # Use Windows API to check drive type
                drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive + "\\")
                # 4 = DRIVE_REMOTE (network drive)
                return drive_type == 4
            except Exception as e:
                logger.warning(f"Could not determine drive type for {drive}: {e}")
                return False
        
        return False
    
    @staticmethod
    def extract_server_from_unc(path: str) -> str | None:
        """Extract server name from UNC path."""
        path = os.path.normpath(path)
        if path.startswith("\\\\"):
            parts = path.split("\\")
            if len(parts) >= 3:
                return parts[2]  # \\server\share -> server
        return None
    
    @staticmethod
    def check_server_reachable(server: str, timeout: int = 5) -> bool:
        """Check if the network server is reachable via socket."""
        try:
            socket.setdefaulttimeout(timeout)
            
            # Try SMB port (445) for Windows file sharing
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex((server, 445))
            sock.close()
            
            if result == 0:
                logger.info(f"Server '{server}' is reachable on SMB port 445")
                return True
            
            # Fallback: try NetBIOS port (139)
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex((server, 139))
            sock.close()
            
            if result == 0:
                logger.info(f"Server '{server}' is reachable on NetBIOS port 139")
                return True
            
            logger.warning(f"Server '{server}' is not reachable on standard SMB ports")
            return False
            
        except socket.gaierror as e:
            logger.error(f"DNS resolution failed for server '{server}': {e}")
            return False
        except socket.error as e:
            logger.error(f"Socket error when checking server '{server}': {e}")
            return False
        finally:
            socket.setdefaulttimeout(None)
    
    @staticmethod
    def validate_directory_access(path: str, check_write: bool = True) -> tuple[bool, str]:
        """
        Validate that a directory exists and is accessible.
        Returns (success, message) tuple.
        """
        try:
            # Check if path exists
            if not os.path.exists(path):
                return False, f"Path does not exist: {path}"
            
            # Check if it's a directory
            if not os.path.isdir(path):
                return False, f"Path is not a directory: {path}"
            
            # Check read access
            if not os.access(path, os.R_OK):
                return False, f"No read permission for: {path}"
            
            # Check write access if requested
            if check_write:
                if not os.access(path, os.W_OK):
                    return False, f"No write permission for: {path}"
            
            return True, f"Directory validated successfully: {path}"
            
        except PermissionError as e:
            return False, f"Permission denied accessing: {path} - {e}"
        except Exception as e:
            return False, f"Unexpected error validating directory: {path} - {e}"


class FileChecker:
    """
    File checker designed to run as a cron job.
    Checks for target file existence and processes it.
    """
    
    def __init__(
        self, 
        watch_directory: str, 
        target_filename: str, 
        destination_directory: str
    ):
        self.watch_directory = watch_directory
        self.target_filename = target_filename
        self.destination_directory = destination_directory
        self.network_validator = NetworkShareValidator()
        self.target_file_path = os.path.join(watch_directory, target_filename)
    
    def run(self) -> int:
        """
        Main entry point - run the file check.
        Returns exit code: 0 = success/file not found (normal), 1 = error
        """
        logger.info("=" * 60)
        logger.info("FILE CHECKER - CRON JOB EXECUTION")
        logger.info(f"Executed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        logger.info(f"Python PID: {os.getpid()}")
        logger.info("=" * 60)
        
        # Step 1: Validate watch directory is accessible
        if not self._validate_watch_directory():
            return 1
        
        # Step 2: Check if target file exists
        if not self._check_file_exists():
            # File not found - already logged, exit normally
            return 0
        
        # Step 3: Validate destination directory
        if not self._validate_destination_directory():
            return 1
        
        # Step 4: Process (move) the file
        if not self._process_file():
            return 1
        
        logger.info("=" * 60)
        logger.info("FILE CHECKER - EXECUTION COMPLETE (SUCCESS)")
        logger.info("=" * 60)
        return 0
    
    def _validate_watch_directory(self) -> bool:
        """Validate the watch directory is accessible."""
        logger.info("-" * 40)
        logger.info("Step 1: Validating watch directory...")
        
        is_network = self.network_validator.is_network_path(self.watch_directory)
        
        if is_network:
            logger.info(f"Detected network path: {self.watch_directory}")
            server = self.network_validator.extract_server_from_unc(self.watch_directory)
            
            if server:
                logger.info(f"Checking connectivity to server: {server}")
                if not self.network_validator.check_server_reachable(server):
                    logger.error(f"FAILED: Network server '{server}' is not reachable")
                    logger.error("Possible reasons:")
                    logger.error("  - Network connection is down")
                    logger.error("  - Server is offline or unavailable")
                    logger.error("  - Firewall blocking SMB ports (445/139)")
                    logger.error("  - VPN not connected")
                    return False
        
        success, message = self.network_validator.validate_directory_access(
            self.watch_directory, 
            check_write=False
        )
        
        if not success:
            logger.error(f"FAILED: {message}")
            logger.error("Possible reasons:")
            logger.error("  - Directory does not exist")
            logger.error("  - Insufficient permissions")
            logger.error("  - Path is incorrect or has typos")
            if is_network:
                logger.error("  - Network share credentials have expired")
                logger.error("  - Share has been removed or renamed")
            return False
        
        logger.info(f"SUCCESS: Watch directory is accessible: {self.watch_directory}")
        return True
    
    def _check_file_exists(self) -> bool:
        """Check if the target file exists. Log detailed info if not found."""
        logger.info("-" * 40)
        logger.info("Step 2: Checking for target file...")
        logger.info(f"Looking for: {self.target_file_path}")
        
        if os.path.exists(self.target_file_path):
            if os.path.isfile(self.target_file_path):
                file_size = os.path.getsize(self.target_file_path)
                file_mtime = datetime.fromtimestamp(os.path.getmtime(self.target_file_path))
                
                logger.info(f"SUCCESS: Target file FOUND!")
                logger.info(f"  File: {self.target_filename}")
                logger.info(f"  Size: {file_size} bytes")
                logger.info(f"  Last modified: {file_mtime}")
                return True
            else:
                logger.warning(f"Path exists but is not a file: {self.target_file_path}")
                logger.warning("The target is a directory, not a file")
                return False
        
        # File not found - log detailed information
        logger.info(f"File NOT found: {self.target_filename}")
        logger.info("-" * 40)
        logger.info("Diagnostic Information:")
        
        # List contents of watch directory
        try:
            contents = os.listdir(self.watch_directory)
            if contents:
                logger.info(f"Files currently in watch directory ({len(contents)} items):")
                for item in contents[:20]:  # Limit to first 20 items
                    item_path = os.path.join(self.watch_directory, item)
                    if os.path.isfile(item_path):
                        size = os.path.getsize(item_path)
                        logger.info(f"  [FILE] {item} ({size} bytes)")
                    else:
                        logger.info(f"  [DIR]  {item}/")
                if len(contents) > 20:
                    logger.info(f"  ... and {len(contents) - 20} more items")
            else:
                logger.info("Watch directory is EMPTY - no files present")
        except Exception as e:
            logger.warning(f"Could not list directory contents: {e}")
        
        # Check for similar filenames (case sensitivity, extensions)
        try:
            target_lower = self.target_filename.lower()
            target_stem = Path(self.target_filename).stem.lower()
            similar_files = []
            
            for item in os.listdir(self.watch_directory):
                item_lower = item.lower()
                item_stem = Path(item).stem.lower()
                
                # Check for exact match with different case
                if item_lower == target_lower and item != self.target_filename:
                    similar_files.append((item, "case mismatch"))
                # Check for same name but different extension
                elif item_stem == target_stem and item_lower != target_lower:
                    similar_files.append((item, "different extension"))
            
            if similar_files:
                logger.warning("Possible filename issues found:")
                for filename, reason in similar_files:
                    logger.warning(f"  Found '{filename}' ({reason})")
        except Exception as e:
            logger.debug(f"Could not check for similar filenames: {e}")
        
        logger.info("-" * 40)
        logger.info("Summary: Target file not present in watch directory")
        logger.info("This is normal if no new file has been dropped yet")
        logger.info("=" * 60)
        return False
    
    def _validate_destination_directory(self) -> bool:
        """Validate the destination directory is accessible and writable."""
        logger.info("-" * 40)
        logger.info("Step 3: Validating destination directory...")
        
        # Create destination if it doesn't exist
        if not os.path.exists(self.destination_directory):
            try:
                os.makedirs(self.destination_directory, exist_ok=True)
                logger.info(f"Created destination directory: {self.destination_directory}")
            except OSError as e:
                logger.error(f"FAILED: Could not create destination directory: {e}")
                return False
        
        success, message = self.network_validator.validate_directory_access(
            self.destination_directory, 
            check_write=True
        )
        
        if not success:
            logger.error(f"FAILED: {message}")
            logger.error("Possible reasons:")
            logger.error("  - Insufficient permissions to write")
            logger.error("  - Disk is full")
            logger.error("  - Read-only file system")
            return False
        
        logger.info(f"SUCCESS: Destination directory is ready: {self.destination_directory}")
        return True
    
    def _process_file(self) -> bool:
        """Move the file from watch directory to destination."""
        logger.info("-" * 40)
        logger.info("Step 4: Processing (moving) file...")
        
        destination_path = os.path.join(self.destination_directory, self.target_filename)
        
        try:
            # Log source file info
            source_stat = os.stat(self.target_file_path)
            logger.info(f"Moving file:")
            logger.info(f"  From: {self.target_file_path}")
            logger.info(f"  To:   {destination_path}")
            logger.info(f"  Size: {source_stat.st_size} bytes")
            
            # Handle existing file at destination
            if os.path.exists(destination_path):
                logger.warning(f"Destination file exists, will be overwritten")
                os.remove(destination_path)
                logger.info("Existing destination file removed")
            
            # Perform the move
            shutil.move(self.target_file_path, destination_path)
            
            # Verify the move was successful
            if os.path.exists(destination_path) and not os.path.exists(self.target_file_path):
                dest_stat = os.stat(destination_path)
                logger.info(f"SUCCESS: File moved successfully!")
                logger.info(f"  Final location: {destination_path}")
                logger.info(f"  Final size: {dest_stat.st_size} bytes")
                return True
            else:
                logger.error("FAILED: Move operation verification failed")
                return False
            
        except FileNotFoundError as e:
            logger.error(f"FAILED: File not found during move: {e}")
            logger.error("The file may have been moved or deleted by another process")
            return False
        except PermissionError as e:
            logger.error(f"FAILED: Permission denied during move: {e}")
            return False
        except shutil.Error as e:
            logger.error(f"FAILED: Shutil error during move: {e}")
            return False
        except Exception as e:
            logger.error(f"FAILED: Unexpected error during move: {type(e).__name__}: {e}")
            return False


def main():
    """Main entry point for cron job execution."""
    try:
        checker = FileChecker(
            watch_directory=WATCH_DIRECTORY,
            target_filename=TARGET_FILENAME,
            destination_directory=DESTINATION_DIRECTORY
        )
        exit_code = checker.run()
        sys.exit(exit_code)
    except Exception as e:
        logger.exception(f"CRITICAL: Unexpected error during execution: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
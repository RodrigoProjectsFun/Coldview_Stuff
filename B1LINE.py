import subprocess
import pandas as pd
import os
import sys

def get_resource_path(relative_path):
    """
    Gets the absolute path to a resource. 
    This is required for PyInstaller to find bundled files.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # If running as a normal Python script, use the current folder
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def run_executable(exe_path, args=None):
    """Runs the external executable."""
    command = [exe_path]
    if args:
        command.extend(args)
    
    print("Generating text files...")
    try:
        subprocess.run(command, check=True, capture_output=True, text=True)
        print("Executable finished successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error: The executable failed with return code {e.returncode}.")
        print(f"Error Message:\n{e.stderr}")
        sys.exit(1)
    except FileNotFoundError:
        print(f"Error: Could not find the executable at '{exe_path}'.")
        sys.exit(1)

def process_and_cleanup(txt_file1, txt_file2, output_excel, column_names):
    """Reads massive FWF files, merges them, saves to Excel, and cleans up."""
    if not os.path.exists(txt_file1) or not os.path.exists(txt_file2):
        print("Error: The text files were not generated properly.")
        sys.exit(1)

    try:
        print("Reading massive text files using PyArrow engine...")
        # read_fwf uses visual alignment. engine="pyarrow" makes it extremely fast.
        df1 = pd.read_fwf(txt_file1, names=column_names, engine="pyarrow") 
        df2 = pd.read_fwf(txt_file2, names=column_names, engine="pyarrow")

        print("Merging data...")
        # Stack the files vertically
        combined_df = pd.concat([df1, df2], axis=0, ignore_index=True)
        
        total_rows = len(combined_df)
        print(f"Writing {total_rows:,} rows to Excel... (This may take a few minutes for large files)")
        
        # Failsafe warning for Excel's hard row limit
        if total_rows > 1048576:
            print("WARNING: Row count exceeds Excel's 1,048,576 limit. The file may be truncated or corrupted.")

        # Write to Excel using Constant Memory Mode to prevent RAM crashes
        with pd.ExcelWriter(
            output_excel, 
            engine='xlsxwriter', 
            engine_kwargs={'options': {'constant_memory': True}}
        ) as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Combined Data')
            
        print(f"Success! Excel file created: {output_excel}")

        # --- CLEAN UP PHASE ---
        print("Cleaning up temporary text files...")
        os.remove(txt_file1)
        os.remove(txt_file2)
        print("Cleanup complete.")

    except Exception as e:
        print(f"An error occurred during conversion or cleanup: {e}")

if __name__ == "__main__":
    # --- CONFIGURATION ---
    # 1. Update this to the exact name of your executable
    EXE_NAME = "your_program.exe" 
    
    # 2. Map the correct path
    EXECUTABLE_PATH = get_resource_path(EXE_NAME) 
    EXECUTABLE_ARGS = [] # Add arguments here if your exe requires them
    
    # 3. Output file names
    TEXT_FILE_1 = "output1.txt" 
    TEXT_FILE_2 = "output2.txt" 
    EXCEL_OUTPUT = "Combined_Results.xlsx"
    
    # 4. Define your shared column names here
    # Must match the number of generated columns
    SHARED_COLUMNS = ["Column_A", "Column_B", "Column_C", "Column_D"] 
    # ---------------------

    # Execute the workflow
    run_executable(EXECUTABLE_PATH, EXECUTABLE_ARGS)
    process_and_cleanup(TEXT_FILE_1, TEXT_FILE_2, EXCEL_OUTPUT, SHARED_COLUMNS)

import pandas as pd
import glob
import os
import argparse
import concurrent.futures

# --- Worker Function ---
# This function runs in an isolated process. It must be defined at the top level
# of the module so that the multiprocessing library can easily serialize (pickle) it.
def read_single_excel(args):
    file_path, search_col, return_col = args
    filename = os.path.basename(file_path)
    
    try:
        # Strictly load only the configurable columns to save memory per worker
        df_temp = pd.read_excel(file_path, usecols=[search_col, return_col])
        df_temp = df_temp.dropna(subset=[search_col, return_col])
        print(f"Worker finished: {filename}")
        return df_temp
        
    except ValueError:
        print(f"Worker skipped {filename}: Columns '{search_col}' or '{return_col}' not found.")
        return None
    except Exception as e:
        print(f"Worker error reading {filename}: {e}")
        return None

def process_reconciliation(config):
    print(f"Loading target data from: {config.target_file}")
    
    try:
        df_targets = pd.read_excel(config.target_file)
    except FileNotFoundError:
        print(f"Error: Could not find target file at {config.target_file}")
        return

    print(f"Scanning directory: {config.source_folder} for source files...")
    all_files = glob.glob(os.path.join(config.source_folder, "*.xlsx"))
    
    if not all_files:
        print("No Excel files found in the source directory.")
        return

    # Package the arguments for the worker pool
    worker_args = [(file, config.search_col, config.return_col) for file in all_files]
    lookup_dataframes = []
    
    print(f"\nSpinning up workers to process {len(all_files)} files in parallel...")
    
    # ProcessPoolExecutor automatically utilizes the optimal number of CPU cores
    with concurrent.futures.ProcessPoolExecutor() as executor:
        # Map executes the worker function across the list of arguments concurrently
        results = executor.map(read_single_excel, worker_args)
        
        # Collect the valid DataFrames returned by the workers
        for df in results:
            if df is not None and not df.empty:
                lookup_dataframes.append(df)

    if not lookup_dataframes:
        print("\nNo valid source data found matching your column criteria. Exiting.")
        return
        
    print("\nConsolidating lookup data from all workers...")
    df_master_lookup = pd.concat(lookup_dataframes, ignore_index=True)
    df_master_lookup = df_master_lookup.drop_duplicates(subset=[config.search_col], keep='last')
    
    print(f"Merging data on '{config.search_col}' to retrieve '{config.return_col}'...")
    df_final = pd.merge(df_targets, df_master_lookup, on=config.search_col, how='left')
    
    df_final.to_excel(config.output_file, index=False)
    print(f"\nProcess complete! Results saved to: {config.output_file}")

if __name__ == "__main__":
    # The if __name__ == "__main__" block is STRICTLY REQUIRED on Windows
    # when using ProcessPoolExecutor to prevent infinite recursive child processes.
    
    parser = argparse.ArgumentParser(description="Extract and match specific fields across multiple Excel files in parallel.")
    
    # Configuration defaults
    parser.add_argument('--source_folder', type=str, default=r'C:\path\to\your\source\excels',
                        help='Folder containing the source data files')
    parser.add_argument('--target_file', type=str, default=r'C:\path\to\cards_to_search.xlsx',
                        help='Excel file containing the list of items to search for')
    parser.add_argument('--output_file', type=str, default=r'C:\path\to\output_results.xlsx',
                        help='Path to save the final merged Excel file')
    
    parser.add_argument('--search_col', type=str, default='Card Number',
                        help='The column name used to match records (e.g., Card Number)')
    parser.add_argument('--return_col', type=str, default='Account Number',
                        help='The column name containing the data to retrieve (e.g., Account Number)')
    
    args = parser.parse_args()
    process_reconciliation(args)
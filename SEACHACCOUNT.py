import pandas as pd
import glob
import os
import argparse

def process_reconciliation(config):
    print(f"Loading target data from: {config.target_file}")
    
    try:
        df_targets = pd.read_excel(config.target_file)
    except FileNotFoundError:
        print(f"Error: Could not find target file at {config.target_file}")
        return

    print(f"Scanning directory: {config.source_folder} for source files...")
    all_files = glob.glob(os.path.join(config.source_folder, "*.xlsx"))
    
    lookup_dataframes = []
    
    for file in all_files:
        try:
            # We strictly load only the configurable columns
            df_temp = pd.read_excel(file, usecols=[config.search_col, config.return_col])
            df_temp = df_temp.dropna(subset=[config.search_col, config.return_col])
            
            lookup_dataframes.append(df_temp)
            print(f"Successfully processed: {os.path.basename(file)}")
            
        except ValueError:
            print(f"Skipped {os.path.basename(file)}: Columns '{config.search_col}' or '{config.return_col}' not found.")
        except Exception as e:
            print(f"Error reading {os.path.basename(file)}: {e}")

    if not lookup_dataframes:
        print("\nNo valid source data found matching your column criteria. Exiting.")
        return
        
    print("\nConsolidating lookup data...")
    df_master_lookup = pd.concat(lookup_dataframes, ignore_index=True)
    df_master_lookup = df_master_lookup.drop_duplicates(subset=[config.search_col], keep='last')
    
    print(f"Merging data on '{config.search_col}' to retrieve '{config.return_col}'...")
    df_final = pd.merge(df_targets, df_master_lookup, on=config.search_col, how='left')
    
    df_final.to_excel(config.output_file, index=False)
    print(f"\nProcess complete! Results saved to: {config.output_file}")

if __name__ == "__main__":
    # --- Configuration Parser ---
    parser = argparse.ArgumentParser(description="Extract and match specific fields across multiple Excel files.")
    
    # You can change the 'default' values here for quick IDE execution
    parser.add_argument('--source_folder', type=str, default=r'C:\path\to\your\source\excels',
                        help='Folder containing the source data files')
    parser.add_argument('--target_file', type=str, default=r'C:\path\to\cards_to_search.xlsx',
                        help='Excel file containing the list of items to search for')
    parser.add_argument('--output_file', type=str, default=r'C:\path\to\output_results.xlsx',
                        help='Path to save the final merged Excel file')
    
    # --- Configurable Fields ---
    parser.add_argument('--search_col', type=str, default='Card Number',
                        help='The column name used to match records (e.g., Card Number)')
    parser.add_argument('--return_col', type=str, default='Account Number',
                        help='The column name containing the data to retrieve (e.g., Account Number)')
    
    # Parse the arguments and pass them to the main function
    args = parser.parse_args()
    process_reconciliation(args)
import pandas as pd
import os
import glob

def mass_conciliation():
    # --- CONFIGURATION ---
    # 1. Folder containing all your excel files (use '.' for current folder)
    folder_path = './accounting_files' 
    
    # 2. File Name Patterns (How we distinguish Debt from Credit files)
    debt_pattern = 'M2D-RECU*.xlsx'   # Looks for files starting with M2D-RECU
    credit_pattern = 'M6D-DEV*.xlsx'  # Looks for files starting with M6D-DEV
    
    # 3. Column Headers (Must match exactly)
    col_card = 'Card'               
    col_op = 'Operation Number'
    
    output_file = 'GLOBAL_CONCILIATION_REPORT.xlsx'
    # ---------------------

    print(f"Scanning folder: {folder_path} ...")

    # Helper function to load multiple files into one DataFrame
    def load_files(pattern, file_type_label):
        search_path = os.path.join(folder_path, pattern)
        files = glob.glob(search_path)
        
        if not files:
            print(f"WARNING: No files found matching {pattern}")
            return pd.DataFrame() # Return empty if nothing found
            
        all_data = []
        print(f"Found {len(files)} {file_type_label} files. Loading...")
        
        for file in files:
            try:
                # dtype=str is crucial to prevent losing leading zeros in card numbers
                df = pd.read_excel(file, dtype=str)
                
                # Add a column so we know exactly which file this row came from
                df['Origin_File'] = os.path.basename(file)
                
                # Clean keys immediately
                if col_card in df.columns and col_op in df.columns:
                    df[col_card] = df[col_card].str.strip()
                    df[col_op] = df[col_op].str.strip()
                    all_data.append(df)
                else:
                    print(f"  [Skipping] {os.path.basename(file)} - Missing required columns.")
            except Exception as e:
                print(f"  [Error] Could not read {os.path.basename(file)}: {e}")

        if all_data:
            return pd.concat(all_data, ignore_index=True)
        else:
            return pd.DataFrame()

    # --- EXECUTION ---
    
    # 1. Create the "Master Piles"
    print("--- Loading Debt Notes (M2D) ---")
    master_debt = load_files(debt_pattern, "Debt")
    
    print("--- Loading Credit Notes (M6D) ---")
    master_credit = load_files(credit_pattern, "Credit")

    if master_debt.empty or master_credit.empty:
        print("Error: One of the file lists is empty. Cannot conciliate.")
        return

    print(f"Total Debt Rows: {len(master_debt)} | Total Credit Rows: {len(master_credit)}")

    # 2. Perform the Global Match
    # This solves the cross-file issue because we are matching against *everything* at once.
    print("Performing Global Match...")
    
    matched_df = pd.merge(
        master_debt,
        master_credit,
        on=[col_card, col_op], # We can use 'on' if column names are identical in both files
        how='inner',
        suffixes=('_DEBT', '_CREDIT')
    )

    # 3. Export
    if not matched_df.empty:
        matched_df.to_excel(output_file, index=False)
        print("------------------------------------------------")
        print("SUCCESS.")
        print(f"Total Conciliated Transactions: {len(matched_df)}")
        print(f"Report saved to: {output_file}")
        
        # Quick Audit: Show an example of a Credit Note matching multiple Debt files?
        # We check if any Credit Origin File is associated with multiple Debt Origin Files
        multi_file_matches = matched_df.groupby('Origin_File_CREDIT')['Origin_File_DEBT'].nunique()
        complex_cases = multi_file_matches[multi_file_matches > 1]
        
        if not complex_cases.empty:
            print(f"\nNote: Found {len(complex_cases)} Credit Notes that bridged across multiple Debt files.")
        print("------------------------------------------------")
    else:
        print("No matches found across any files.")

if __name__ == "__main__":
    mass_conciliation()
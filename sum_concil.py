import pandas as pd
import os
import glob

def conciliation_breakdown_by_file():
    # --- CONFIGURATION ---
    folder_path = './accounting_files'
    
    # Patterns
    debt_pattern = 'M2D-RECU*.xlsx'
    credit_pattern = 'M6D-DEV*.xlsx'
    
    # Column Headers
    col_card = 'Card'               
    col_op = 'Operation Number'
    col_amount = 'Original Amount' 
    
    output_file = 'CONCILIATION_SUBTOTALS_REPORT.xlsx'
    # ---------------------

    print(f"Scanning {folder_path}...")

    # --- 1. LOAD & CLEAN DATA ---
    def load_pile(pattern, label):
        files = glob.glob(os.path.join(folder_path, pattern))
        all_dfs = []
        for f in files:
            try:
                df = pd.read_excel(f, dtype=str)
                df['Origin_File'] = os.path.basename(f)
                
                # Standardize Keys
                if col_card in df.columns and col_op in df.columns:
                    df[col_card] = df[col_card].str.strip()
                    df[col_op] = df[col_op].str.strip()
                
                # Clean Amount (Convert "$1,000.00" -> 1000.00)
                if col_amount in df.columns:
                    # Remove non-numeric chars except dot and minus
                    clean_amt = df[col_amount].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    df['Amt_Float'] = pd.to_numeric(clean_amt, errors='coerce').fillna(0)
                    all_dfs.append(df)
                else:
                    print(f"Skipping {os.path.basename(f)} - Missing Amount column")
                    
            except Exception as e:
                print(f"Error reading {os.path.basename(f)}: {e}")
        
        return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

    print("Loading Files...")
    df_debt = load_pile(debt_pattern, "Debt")
    df_credit = load_pile(credit_pattern, "Credit")

    if df_debt.empty or df_credit.empty:
        print("Error: Missing file data.")
        return

    # --- 2. MATCHING (INNER JOIN) ---
    print("Matching Transactions...")
    merged = pd.merge(
        df_debt, 
        df_credit, 
        on=[col_card, col_op], 
        how='inner', 
        suffixes=('_DEBT', '_CREDIT')
    )

    if merged.empty:
        print("No matches found.")
        return

    # --- 3. GENERATE BREAKDOWNS ---
    print("Calculating Subtotals...")

    # VIEW 1: DEBT PERSPECTIVE
    # "For this Debt File, where did the money come from?"
    # Group by Debt File + Credit File to separate the subtotals
    debt_breakdown = merged.groupby(['Origin_File_DEBT', 'Origin_File_CREDIT']).agg(
        Count_Matches=('Operation Number', 'count'),
        Subtotal_Amount=('Amt_Float_DEBT', 'sum') # Summing the Debt we cleared
    ).reset_index()
    
    # Sort for readability
    debt_breakdown = debt_breakdown.sort_values(['Origin_File_DEBT', 'Origin_File_CREDIT'])


    # VIEW 2: CREDIT PERSPECTIVE
    # "For this Credit File, which Debt Files did it pay off?"
    # Group by Credit File + Debt File
    credit_breakdown = merged.groupby(['Origin_File_CREDIT', 'Origin_File_DEBT']).agg(
        Count_Matches=('Operation Number', 'count'),
        Subtotal_Amount=('Amt_Float_DEBT', 'sum') # We sum Debt amount here to see "Value of Debt Covered"
    ).reset_index()

    # Sort
    credit_breakdown = credit_breakdown.sort_values(['Origin_File_CREDIT', 'Origin_File_DEBT'])


    # --- 4. EXPORT ---
    print(f"Exporting to {output_file}...")
    with pd.ExcelWriter(output_file) as writer:
        debt_breakdown.to_excel(writer, sheet_name='DEBT_Breakdown', index=False)
        credit_breakdown.to_excel(writer, sheet_name='CREDIT_Breakdown', index=False)

    print("Done.")

if __name__ == "__main__":
    conciliation_breakdown_by_file()
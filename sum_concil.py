import pandas as pd
import os
import glob

def generate_relationship_summary():
    # --- CONFIGURATION ---
    folder_path = './accounting_files'
    
    # Patterns
    debt_pattern = 'M2D-RECU*.xlsx'
    credit_pattern = 'M6D-DEV*.xlsx'
    
    # Keys
    col_card = 'Card'               
    col_op = 'Operation Number'
    col_amount = 'Original Amount' # Amount column name in BOTH files
    
    output_file = 'CONCILIATION_SUMMARY_VIEW.xlsx'
    # ---------------------

    print(f"Scanning {folder_path}...")

    # --- 1. LOADER ---
    def load_pile(pattern, label):
        files = glob.glob(os.path.join(folder_path, pattern))
        all_dfs = []
        for f in files:
            try:
                df = pd.read_excel(f, dtype=str)
                df['Origin_File'] = os.path.basename(f)
                
                # Standardize columns
                df[col_card] = df[col_card].str.strip()
                df[col_op] = df[col_op].str.strip()
                
                # Clean Amount (remove $ or ,)
                if col_amount in df.columns:
                    df[col_amount] = df[col_amount].str.replace('$', '', regex=False)
                    df[col_amount] = df[col_amount].str.replace(',', '', regex=False)
                    df['Amt_Float'] = pd.to_numeric(df[col_amount], errors='coerce').fillna(0)
                
                all_dfs.append(df)
            except Exception as e:
                print(f"Skipping {os.path.basename(f)}: {e}")
        
        return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

    # Load Data
    print("Loading all files...")
    df_debt = load_pile(debt_pattern, "Debt")
    df_credit = load_pile(credit_pattern, "Credit")

    if df_debt.empty or df_credit.empty:
        print("Missing files. Cannot proceed.")
        return

    # --- 2. MERGE ---
    print("Finding relationships...")
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

    # --- 3. CREATE SUMMARIES ---
    
    # Helper to format the list of files nicely
    def join_files(file_list):
        return " AND ".join(sorted(file_list.unique()))

    print("Summarizing Debt Side...")
    # Group by Debt File -> See which Credit Files it touched
    debt_summary = merged.groupby('Origin_File_DEBT').agg(
        Total_Matches=('Origin_File_CREDIT', 'count'),
        Total_Amount_Conciliated=('Amt_Float_DEBT', 'sum'),
        Matched_With_Credit_Files=('Origin_File_CREDIT', join_files)
    ).reset_index()

    print("Summarizing Credit Side...")
    # Group by Credit File -> See which Debt Files it touched
    # Note: On the credit side, we sum the CREDIT amount, but we must deduplicate first 
    # if one credit row matched multiple debt rows (to avoid inflating the credit total).
    
    # Create a unique list of credit rows that were used
    unique_credits_used = merged.drop_duplicates(subset=['Origin_File_CREDIT', col_card, col_op])
    
    # First aggregation: Amounts and Counts
    credit_stats = unique_credits_used.groupby('Origin_File_CREDIT').agg(
        Total_Amount_Conciliated=('Amt_Float_CREDIT', 'sum')
    ).reset_index()

    # Second aggregation: The file relationships (needs the full merge to see all connections)
    credit_relations = merged.groupby('Origin_File_CREDIT').agg(
        Total_Matches=('Origin_File_DEBT', 'count'),
        Matched_With_Debt_Files=('Origin_File_DEBT', join_files)
    ).reset_index()
    
    # Combine stats and relations
    credit_summary = pd.merge(credit_stats, credit_relations, on='Origin_File_CREDIT')

    # --- 4. EXPORT ---
    with pd.ExcelWriter(output_file) as writer:
        debt_summary.to_excel(writer, sheet_name='DEBT_View (M2D)', index=False)
        credit_summary.to_excel(writer, sheet_name='CREDIT_View (M6D)', index=False)

    print("------------------------------------------------")
    print(f"Report Generated: {output_file}")
    print("------------------------------------------------")

if __name__ == "__main__":
    generate_relationship_summary()
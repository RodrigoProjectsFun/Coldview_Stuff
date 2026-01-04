import pandas as pd
import os
import glob
import re

def robust_conciliation_duplicates_allowed():
    # --- CONFIGURATION ---
    folder_path = './accounting_files'
    
    # Patterns
    debt_pattern = '*m2d-recu*.xlsx'
    credit_pattern = '*m6d-dev*.xlsx'
    
    # Headers
    col_card = 'Card'               
    col_op = 'Operation Number'
    col_amount = 'Original Amount' 
    
    output_file = 'CONCILIATION_FINAL_REPORT.xlsx'
    # ---------------------

    print(f"--- Starting Conciliation (Duplicates Allowed) in {folder_path} ---")

    # --- HELPER: STANDARDIZE FILENAMES ---
    def get_standardized_name(filepath):
        """
        Converts filename to strict format: M2D-RECU <DATE> or M6D-DEV <DATE>
        """
        filename = os.path.basename(filepath)
        name_lower = filename.lower()
        
        # Regex to capture date (dots or dashes)
        date_match = re.search(r'(\d+[\.-]\d+[\.-]\d+)', name_lower)
        date_str = date_match.group(1) if date_match else "NO_DATE"
        
        if 'm2d-recu' in name_lower:
            return f"M2D-RECU {date_str}"
        elif 'm6d-dev' in name_lower:
            return f"M6D-DEV {date_str}"
        else:
            return f"UNKNOWN {filename}"

    # --- 1. LOADER ---
    def load_pile(pattern, label):
        files = glob.glob(os.path.join(folder_path, pattern))
        
        # Double check filter (glob can be broad)
        filter_keyword = 'm2d-recu' if label == "DEBT" else 'm6d-dev'
        files = [f for f in files if filter_keyword in os.path.basename(f).lower()]
        
        all_dfs = []
        print(f"Loading {len(files)} files for {label}...")

        for f in files:
            try:
                # Load as String to protect IDs from scientific notation
                df = pd.read_excel(f, dtype=str)
                
                # Create Standardized Reference Name
                std_name = get_standardized_name(f)
                df['Accounting_Ref'] = std_name
                
                # Clean Keys
                if col_card in df.columns and col_op in df.columns:
                    df[col_card] = df[col_card].str.strip()
                    df[col_op] = df[col_op].str.strip()
                else:
                    print(f"  [SKIP] {std_name} missing Card or Operation headers.")
                    continue
                
                # Clean Amount (Force to Float)
                if col_amount in df.columns:
                    clean_amt = df[col_amount].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    df['Amt_Float'] = pd.to_numeric(clean_amt, errors='coerce').fillna(0.0)
                    all_dfs.append(df)
                
            except Exception as e:
                print(f"  [ERROR] {os.path.basename(f)}: {e}")
        
        return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

    # Load Data
    df_debt = load_pile(debt_pattern, "DEBT")
    df_credit = load_pile(credit_pattern, "CREDIT")

    if df_debt.empty or df_credit.empty:
        print("Stopping: Missing data.")
        return

    # --- 2. MATCHING (The Critical Part) ---
    print("Matching Transactions (Allowing Duplicates)...")
    
    # We use an INNER JOIN. 
    # Because duplicates exist in Debt (and possibly Credit), this will create a Cartesian Product 
    # for those specific keys. This is EXPECTED behavior.
    # Ex: 2 Debts match 1 Credit -> Result is 2 rows.
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

    # --- 3. AGGREGATION (The Math Fix) ---
    # We must be very careful what we sum.
    # Since 1 Credit row might be repeated across 5 Debt rows, summing Credit column = WRONG.
    # But 5 Debt rows are distinct payments, so summing Debt column = CORRECT.
    
    print("Generating Accounting Breakdown...")

    # VIEW 1: DEBT FILE PERSPECTIVE
    # "Which Credit Files paid off this Debt File?"
    debt_breakdown = merged.groupby(['Accounting_Ref_DEBT', 'Accounting_Ref_CREDIT']).agg(
        Count_Operations=('Operation Number', 'count'),
        Total_Conciliated_Amount=('Amt_Float_DEBT', 'sum') # Summing the Debt side is safe
    ).reset_index()

    # VIEW 2: CREDIT FILE PERSPECTIVE
    # "Which Debt Files did this Credit File cover?"
    credit_breakdown = merged.groupby(['Accounting_Ref_CREDIT', 'Accounting_Ref_DEBT']).agg(
        Count_Operations=('Operation Number', 'count'),
        Total_Conciliated_Amount=('Amt_Float_DEBT', 'sum') # We still sum DEBT here.
        # Why? Because 'Amt_Float_DEBT' represents the actual individual transactions covered.
    ).reset_index()

    # --- 4. EXPORT ---
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            debt_breakdown.to_excel(writer, sheet_name='By_Debt_File', index=False)
            credit_breakdown.to_excel(writer, sheet_name='By_Credit_File', index=False)
            
            # Detailed Audit Sheet (Optional but recommended for tracing duplicates)
            merged.to_excel(writer, sheet_name='Detailed_Audit_Log', index=False)
            
        print(f"SUCCESS. Report saved to: {output_file}")
        print("NOTE: 'Total_Conciliated_Amount' is calculated based on the sum of DEBT notes found.")
        
    except PermissionError:
        print(f"ERROR: Please close {output_file} and run again.")

if __name__ == "__main__":
    robust_conciliation_duplicates_allowed()
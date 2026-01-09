import pandas as pd
import os
import shutil
import tempfile
import sys
import contextlib
import io
import sum_concil

def test_variance_analysis():
    print("--- Starting Variance Analysis Test ---")
    
    # 1. Setup
    test_dir = tempfile.mkdtemp()
    test_acc_dir = os.path.join(test_dir, 'accounting_files')
    os.makedirs(test_acc_dir)
    
    # 2. Create Validation Files
    # Scenario:
    # A. Perfect Match (1 Credit covers 2 Debts exactly)
    #    Credit A: $100 -> Covers Debt A1 ($50) + Debt A2 ($50)
    
    # B. Underpayment (1 Credit covers 1 Debt but amount is less)
    #    Credit B: $50  -> Covers Debt B1 ($60) -> Variance $10 (Debt > Credit)
    
    # C. Overpayment (1 Credit covers 1 Debt but amount is more)
    #    Credit C: $50  -> Covers Debt C1 ($40) -> Variance -$10 (Debt < Credit)
    
    # DEBT FILE
    df_debt = pd.DataFrame({
        'Card': ['CardA', 'CardA', 'CardB', 'CardC'],
        'Operation Number': ['OpA1', 'OpA2', 'OpB1', 'OpC1'],
        'Original Amount': ['50', '50', '60', '40'],
        'RECUPERAR': ['SI', 'SI', 'SI', 'SI']
    })
    df_debt.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.01.2023.xlsx'), index=False)
    
    # CREDIT FILE
    # Note: Credit Card/Op columns are usually matched 1-to-1 in the merge key logic.
    # But here we have 1 Credit Row matching Multiple Debt Rows.
    # How? 
    # Usually: unique key = Card + Op.
    # If Debt A1 and Debt A2 have DIFFERENT Ops (OpA1, OpA2), they need matching Credit Rows with OpA1, OpA2.
    # 
    # WAIT! The merge logic is on [Card, Op].
    # So for 1 Credit to match 2 Debts, either:
    # 1. Matches are Many-to-Many on same key (e.g. CardA + OpA matches both).
    # 2. OR -- the user implies that "Refund A covers multiple operations".
    #    But our current logic is strict merge on Op Number.
    #    
    #    If the Debt Notes have distinct Op Numbers, they generally need distinct Credit Rows to match.
    #    UNLESS the Credit File lists multiple Op Numbers? No, standard layout is 1 row per op.
    #
    #    Assumption: The "1 Credit covers multiple Debts" scenario usually means:
    #    "I issued one refined Refund Transaction (in the bank), but in my system I'm closing out 3 invoices".
    #    
    #    However, sum_concil.py merges on 'Operation Number'.
    #    If these 3 invoices have different Op Numbers, they WON'T match a single Credit row with a single Op Number.
    #
    #    Let's assume the "Duplicate/Many-to-Many" logic:
    #    Maybe all these debts share the SAME 'Operation Number' (e.g. a Batch ID), 
    #    but are distinct lines in the debt file.
    
    # Let's test the Same-Key scenario which is supported by current code.
    
    df_credit = pd.DataFrame({
        'Card': ['CardA', 'CardB', 'CardC'],
        'Operation Number': ['OpA_Shared', 'OpB1', 'OpC1'], # OpA_Shared matches multiple
        'Original Amount': ['100', '50', '50']
    })
    df_credit.to_excel(os.path.join(test_acc_dir, 'm6d-dev 01.05.2023.xlsx'), index=False)
    
    # Re-write Debt with shared key for A
    df_debt = pd.DataFrame({
        'Card': ['CardA', 'CardA', 'CardB', 'CardC'],
        'Operation Number': ['OpA_Shared', 'OpA_Shared', 'OpB1', 'OpC1'],
        'Original Amount': ['50', '50', '60', '40'],
        'RECUPERAR': ['SI', 'SI', 'SI', 'SI']
    })
    df_debt.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.01.2023.xlsx'), index=False)
    

    # 3. Run
    original_cwd = os.getcwd()
    try:
        os.chdir(test_dir)
        f = io.StringIO()
        with contextlib.redirect_stdout(f): 
            try:
                sum_concil.robust_conciliation_duplicates_allowed()
            except Exception as e:
                print(f"Crashed: {e}")
        
        # 4. Verify
        output_file = 'CONCILIATION_FINAL_REPORT.xlsx'
        if os.path.exists(output_file):
            with pd.ExcelFile(output_file) as xl:
                if 'Amount_Variances' in xl.sheet_names:
                    df = pd.read_excel(xl, 'Amount_Variances')
                    print("Found Amount_Variances Sheet. Contents:")
                    print(df[['Card', 'Operation', 'Refund_Amount', 'Total_Debts_Covered', 'Variance']].to_string())
                    
                    # We expect Case B and Case C. Case A should be hidden (perfect match).
                    
                    # Check Case A (Should NOT be there)
                    if not df[df['Card'] == 'CardA'].empty:
                         print("❌ FAILURE: CardA (Perfect Match) found in variance report!")
                    else:
                         print("✅ SUCCESS: CardA (Perfect Match) excluded.")
                         
                    # Check Case B (Underpaid)
                    row_b = df[df['Card'] == 'CardB']
                    if not row_b.empty:
                        var_b = row_b.iloc[0]['Variance']
                        # Credit 50 - Debt 60 = -10
                        if abs(var_b - (-10)) < 0.01:
                             print("✅ SUCCESS: CardB Variance correct (-10).")
                        else:
                             print(f"❌ FAILURE: CardB Variance wrong. Got {var_b}, expected -10.")
                    else:
                        print("❌ FAILURE: CardB missing.")

                    # Check Case C (Overpaid)
                    row_c = df[df['Card'] == 'CardC']
                    if not row_c.empty:
                        var_c = row_c.iloc[0]['Variance']
                        # Credit 50 - Debt 40 = +10
                        if abs(var_c - 10) < 0.01:
                             print("✅ SUCCESS: CardC Variance correct (+10).")
                        else:
                             print(f"❌ FAILURE: CardC Variance wrong. Got {var_c}, expected +10.")
                    else:
                        print("❌ FAILURE: CardC missing.")
                        
                else:
                    print("❌ FAILURE: Sheet 'Amount_Variances' missing.")
        else:
             print("❌ FAILURE: Output file missing.")
             print("--- CAPTURED OUTPUT ---")
             print(f.getvalue())
             
    finally:
        os.chdir(original_cwd)
        # shutil.rmtree(test_dir)

if __name__ == "__main__":
    test_variance_analysis()

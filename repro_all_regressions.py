import pandas as pd
import os
import shutil
import tempfile
import sys
import contextlib
import io
import sum_concil

def run_regression_tests():
    print("==================================================")
    print("      STARTING REFACORING REGRESSION SUITE        ")
    print("==================================================")
    
    # 1. Setup Environment
    test_dir = tempfile.mkdtemp()
    test_acc_dir = os.path.join(test_dir, 'accounting_files')
    os.makedirs(test_acc_dir)
    
    # ---------------------------------------------------------
    # SCENARIO 1: BASIC MATCHING & RECUPERAR LOGIC
    # ---------------------------------------------------------
    # Case 1A: Pending Claim (RECUPERAR='NO', Unmatched) -> Expected in Pending_Claims
    # Case 1B: Unexpected Refund (RECUPERAR='SI', Matched) -> Expected in Unexpected_Refunds
    # Case 1C: Normal Match (RECUPERAR='NO', Matched) -> Expected in Fully_Reconciled
    
    df_debt_1 = pd.DataFrame({
        'Card': ['C1A', 'C1B', 'C1C'],
        'Operation Number': ['OP1A', 'OP1B', 'OP1C'],
        'Original Amount': ['10.00', '20.00', '30.00'],
        'RECUPERAR': ['NO', 'SI', 'NO']
    })
    df_debt_1.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.01.2023.xlsx'), index=False)
    
    df_cred_1 = pd.DataFrame({
        'Card': ['C1B', 'C1C'],
        'Operation Number': ['OP1B', 'OP1C'],
        'Original Amount': ['20.00', '30.00']
    })
    df_cred_1.to_excel(os.path.join(test_acc_dir, 'm6d-dev 01.01.2023.xlsx'), index=False)

    # ---------------------------------------------------------
    # SCENARIO 2: VARIANCE ANALYSIS
    # ---------------------------------------------------------
    # Case 2A: Underpaid (Debt 100, Credit 90) -> Variance -10
    # Case 2B: Overpaid (Debt 100, Credit 110) -> Variance +10
    
    df_debt_2 = pd.DataFrame({
        'Card': ['C2A', 'C2B'],
        'Operation Number': ['OP2A', 'OP2B'],
        'Original Amount': ['100', '100'],
        'RECUPERAR': ['NO', 'NO']
    })
    df_debt_2.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.02.2023.xlsx'), index=False)
    
    df_cred_2 = pd.DataFrame({
        'Card': ['C2A', 'C2B'],
        'Operation Number': ['OP2A', 'OP2B'],
        'Original Amount': ['90', '110']
    })
    df_cred_2.to_excel(os.path.join(test_acc_dir, 'm6d-dev 01.02.2023.xlsx'), index=False)
    
    # ---------------------------------------------------------
    # SCENARIO 3: NET BALANCED LOGIC
    # ---------------------------------------------------------
    # Case 3A: Balanced (Pending 50 == Unexpected 50) -> Net Balanced
    # Case 3B: Unbalanced (Pending 50 != Unexpected 40) -> Excluded
    
    # File 3A (Balanced)
    df_debt_3a = pd.DataFrame({
        'Card': ['C3A1', 'C3A2'],
        'Operation Number': ['OP3A1', 'OP3A2'],
        'Original Amount': ['50', '50'],
        'RECUPERAR': ['NO', 'SI']
    })
    df_debt_3a.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.03.2023.xlsx'), index=False)
    
    df_cred_3a = pd.DataFrame({
        'Card': ['C3A2'],
        'Operation Number': ['OP3A2'],
        'Original Amount': ['50']
    })
    df_cred_3a.to_excel(os.path.join(test_acc_dir, 'm6d-dev 01.03.2023.xlsx'), index=False)

    # File 3B (Unbalanced)
    df_debt_3b = pd.DataFrame({
        'Card': ['C3B1', 'C3B2'],
        'Operation Number': ['OP3B1', 'OP3B2'],
        'Original Amount': ['50', '40'],
        'RECUPERAR': ['NO', 'SI']
    })
    df_debt_3b.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.04.2023.xlsx'), index=False)
    
    df_cred_3b = pd.DataFrame({
        'Card': ['C3B2'],
        'Operation Number': ['OP3B2'],
        'Original Amount': ['40']
    })
    df_cred_3b.to_excel(os.path.join(test_acc_dir, 'm6d-dev 01.04.2023.xlsx'), index=False)


    # ---------------------------------------------------------
    # EXECUTION
    # ---------------------------------------------------------
    original_cwd = os.getcwd()
    try:
        os.chdir(test_dir)
        # Redirect stdout to avoid noise
        f_buffer = io.StringIO()
        with contextlib.redirect_stdout(f_buffer):
            try:
                sum_concil.robust_conciliation_duplicates_allowed()
            except Exception as e:
                print(f"CRASH: {e}")
        
        # ---------------------------------------------------------
        # VERIFICATION
        # ---------------------------------------------------------
        output_file = 'CONCILIATION_FINAL_REPORT.xlsx'
        if not os.path.exists(output_file):
            print("❌ FAILURE: Output file not created.")
            print(f_buffer.getvalue())
            return
            
        print("✓ Output file created.")
        xls = pd.ExcelFile(output_file)
        
        # TEST 1: PENDING CLAIMS (From 1A, 3A, 3B)
        if 'Pending_Claims' in xls.sheet_names:
            df_p = pd.read_excel(xls, 'Pending_Claims')
            # Expect C1A (1A), C3A1 (3A), C3B1 (3B)
            # BUT 3A is Net Balanced, so it might disappear from Pending List depending on implementation?
            # User requirement: "Create a sheet... indicate which operations...". 
            # In current logic, Net Balanced logic adds to Net_Balanced_Sheet. 
            # Does it remove from Pending_Claims? 
            # The code I wrote simply dumps 'pending_claims' df. It doesn't drop the ones that went to Net Balanced.
            # Use logic: pending_claims.to_excel(...)
            # So they should still be there.
            
            cards = df_p['Card'].tolist()
            if 'C1A' in cards: print("✅ TEST 1A: Pending Claim found.")
            else: print("❌ TEST 1A: Pending Claim C1A missing.")
        else:
            print("❌ TEST 1: Pending_Claims sheet missing.")

        # TEST 2: UNEXPECTED REFUNDS (From 1B, 3A, 3B)
        if 'Unexpected_Refunds' in xls.sheet_names:
            df_u = pd.read_excel(xls, 'Unexpected_Refunds')
            cards = df_u['Card'].tolist()
            if 'C1B' in cards: print("✅ TEST 2A: Unexpected Refund found.")
            else: print("❌ TEST 2A: Unexpected Refund C1B missing.")
        else:
             print("❌ TEST 2: Unexpected_Refunds sheet missing.")

        # TEST 3: VARIANCE (From 2A, 2B)
        if 'Amount_Variances' in xls.sheet_names:
            df_v = pd.read_excel(xls, 'Amount_Variances')
            # Check 2A
            row_a = df_v[df_v['Card'] == 'C2A']
            if not row_a.empty and abs(row_a.iloc[0]['Variance'] + 10) < 0.01:
                print("✅ TEST 3A: Underpayment identified correctly.")
            else:
                print(f"❌ TEST 3A: Underpayment validation failed. {row_a.to_dict('records')}")
                
            # Check 2B
            row_b = df_v[df_v['Card'] == 'C2B']
            if not row_b.empty and abs(row_b.iloc[0]['Variance'] - 10) < 0.01:
                print("✅ TEST 3B: Overpayment identified correctly.")
            else:
                 print(f"❌ TEST 3B: Overpayment validation failed.")
        else:
            print("❌ TEST 3: Amount_Variances sheet missing.")

        # TEST 4: STRICT SUMMARY (From 1C)
        # Should include 1C. Should exclude 2A, 2B (Variance), 3A (Balanced but not strict), 3B (Unbalanced)
        if 'Fully_Reconciled_Notes' in xls.sheet_names:
            df_s = pd.read_excel(xls, 'Fully_Reconciled_Notes')
            
            # Use 'DEBTOR FILE' column name
            if 'DEBTOR FILE' in df_s.columns:
                files = df_s['DEBTOR FILE'].astype(str).tolist()
                
                # 1.01.2023 should be there
                if any('01.01.2023' in f for f in files):
                    print("✅ TEST 4A: Perfect file included.")
                else:
                    print("❌ TEST 4A: Perfect file missing.")
                    
                # 1.02.2023 (Variance) should be gone
                if any('01.02.2023' in f for f in files):
                    print("❌ TEST 4B: Variance file NOT excluded.")
                else:
                    print("✅ TEST 4B: Variance file correctly excluded.")
                    
                # Total Row check
                if 'TOTAL' in files:
                    print("✅ TEST 4C: Total row present.")
                else:
                    print("❌ TEST 4C: Total row missing.")
            else:
                print(f"❌ TEST 4: Wrong Headers: {df_s.columns.tolist()}")
        else:
            print("❌ TEST 4: Fully_Reconciled_Notes sheet missing.")
            
        # TEST 5: NET BALANCED (From 3A)
        # Should include 3A. Should exclude 3B.
        if 'Net_Balanced_Files' in xls.sheet_names:
            df_n = pd.read_excel(xls, 'Net_Balanced_Files')
            files = df_n['Debtor_File'].unique().astype(str)
            
            if any('01.03.2023' in f for f in files):
                 print("✅ TEST 5A: Net Balanced file found.")
            else:
                 print("❌ TEST 5A: Net Balanced file missing.")
                 
            if any('01.04.2023' in f for f in files):
                 print("❌ TEST 5B: Unbalanced file included incorrectly.")
            else:
                 print("✅ TEST 5B: Unbalanced file correctly excluded.")
                 
        else:
            print("❌ TEST 5: Net_Balanced_Files sheet missing.")

    finally:
        os.chdir(original_cwd)
        # shutil.rmtree(test_dir) # Clean up manually if needed

if __name__ == "__main__":
    run_regression_tests()

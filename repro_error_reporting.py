import pandas as pd
import os
from sum_concil import robust_conciliation_duplicates_allowed

# Mock the glob and pandas read_excel to simulate data without needing actual files
# However, sum_concil.py is a bit monolithic.
# Easier approach: Create a temporary directory with bad files, run the script, and check output.

import tempfile
import shutil
import contextlib
import io
import sys

def test_error_reporting():
    # 1. Setup temporary directory
    test_dir = tempfile.mkdtemp()
    test_acc_dir = os.path.join(test_dir, 'accounting_files')
    os.makedirs(test_acc_dir)
    
    # 2. Create a "Bad" file (Missing Card Number)
    # File 1: Good DEBT
    df_good = pd.DataFrame({
        'Card': ['1234'], 'Operation Number': ['OP01'], 'Original Amount': ['100']
    })
    df_good.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.01.2023.xlsx'), index=False)
    
    # File 2: Bad DEBT (Empty Card)
    df_bad = pd.DataFrame({
        'Card': [None], 'Operation Number': ['OP02'], 'Original Amount': ['200']
    })
    df_bad.to_excel(os.path.join(test_acc_dir, 'm2d-recu 01.02.2023.xlsx'), index=False)

    # File 3: Dummy CREDIT (to pass valid files check)
    df_credit = pd.DataFrame({
        'Card': ['1234'], 'Operation Number': ['OP01'], 'Original Amount': ['100']
    })
    df_credit.to_excel(os.path.join(test_acc_dir, 'm6d-dev 01.05.2023.xlsx'), index=False)

    # 3. Hijack functionality to point to this folder
    # We can't easily change the hardcoded folder in sum_concil.py without import override or changing the CWD.
    # Changing CWD is easiest for this script.
    
    original_cwd = os.getcwd()
    try:
        os.chdir(test_dir)
        
        # Capture stdout
        f = io.StringIO()
        with contextlib.redirect_stdout(f):
            try:
                robust_conciliation_duplicates_allowed()
            except Exception as e:
                print(f"Crashed: {e}")
        
        output = f.getvalue()
        
        print("--- CAPTURED OUTPUT ---")
        # print(output)
        print("-----------------------")
        
        # 4. Check for filename in error
        target_filename = "m2d-recu 01.02.2023" # The bad file
        if "Errors" in output or "ERRORS" in output:
            if target_filename in output and "empty/null Card" in output:
                print("\n✅ SUCCESS: Found filename in error report!")
            else:
                 print("\n❌ FAILURE: Filename NOT found in error report (or error not triggered).")
        else:
             print("\n⚠️ WARNING: No errors detected?")

    finally:
        os.chdir(original_cwd)
        shutil.rmtree(test_dir)

if __name__ == "__main__":
    test_error_reporting()

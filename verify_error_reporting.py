
import unittest
import pandas as pd
import io
import sys
from unittest.mock import patch
import sum_concil  # Assuming sum_concil.py is in the same directory

class TestErrorReporting(unittest.TestCase):
    
    def test_check_orphans_lists_problematic_files(self):
        """
        Test that check_orphans prints the specific filenames containing orphaned credits.
        """
        # Setup: 1 Debt File, 1 Credit File with an EXTRA (orphaned) credit
        df_debt = pd.DataFrame({
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
            'Accounting_Ref': ['DEBT_FILE_1']
        })
        
        df_credit = pd.DataFrame({
            'TARJETA': ['1234', '9999'], # 9999 is orphan
            'NUM OPE': ['OP-001', 'OP-ORPHAN'],
            'Accounting_Ref': ['CREDIT_FILE_WITH_ERROR', 'CREDIT_FILE_WITH_ERROR']
        })
        
        merged = pd.merge(df_debt, df_credit, on=['TARJETA', 'NUM OPE'], how='inner', suffixes=('_DEBT', '_CREDIT'))
        
        # Instantiate Conciliator
        c = sum_concil.Conciliator()
        c.df_debt = df_debt
        c.df_credit = df_credit
        c.merged = merged
        
        # Capture Stdout
        captured_output = io.StringIO()
        sys.stdout = captured_output
        
        try:
            # result = sum_concil.check_orphans(df_debt, df_credit, merged)
            # Use class method
            result = c._check_orphans()
        finally:
            sys.stdout = sys.__stdout__ # Restore
            
        output = captured_output.getvalue()
        
        print("\nCaptured Output during Test:")
        print(output)
        
        # Assertions
        self.assertFalse(result, "Should return False (Aborted) for orphaned credits")
        self.assertIn("CRITICAL ERROR: ORPHANED CREDITS DETECTED", output)
        self.assertIn("PROBLEMATIC FILES", output)
        self.assertIn("- CREDIT_FILE_WITH_ERROR", output)

if __name__ == '__main__':
    unittest.main()

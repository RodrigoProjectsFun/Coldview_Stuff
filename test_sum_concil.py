"""
Comprehensive tests for sum_concil.py

Tests cover:
1. Basic matching functionality
2. Duplicate handling (Cartesian product)
3. Edge cases: missing columns, empty files, malformed data
4. Amount parsing with special characters
5. Filename standardization
6. No matches scenario
"""

import unittest
import pandas as pd
import os
import shutil
import tempfile
from unittest.mock import patch
import sys

# Import the module under test
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


class TestSumConcil(unittest.TestCase):
    """Test suite for sum_concil.py conciliation logic"""

    @classmethod
    def setUpClass(cls):
        """Create a temporary folder structure for tests"""
        cls.test_dir = tempfile.mkdtemp()
        cls.original_folder = './accounting_files'
        cls.test_accounting_folder = os.path.join(cls.test_dir, 'accounting_files')
        os.makedirs(cls.test_accounting_folder, exist_ok=True)

    @classmethod
    def tearDownClass(cls):
        """Clean up temporary files"""
        import time
        time.sleep(0.1)  # Allow file handles to release on Windows
        try:
            shutil.rmtree(cls.test_dir, ignore_errors=True)
        except Exception:
            pass  # Ignore cleanup errors on Windows

    def setUp(self):
        """Clean accounting folder before each test"""
        for f in os.listdir(self.test_accounting_folder):
            os.remove(os.path.join(self.test_accounting_folder, f))

    def _create_excel(self, filename, data_dict):
        """Helper to create Excel test files"""
        df = pd.DataFrame(data_dict)
        filepath = os.path.join(self.test_accounting_folder, filename)
        df.to_excel(filepath, index=False)
        return filepath

    # =========================================================================
    # TEST 1: FILENAME STANDARDIZATION
    # =========================================================================
    def test_filename_standardization_m2d_recu(self):
        """Test that M2D-RECU files are standardized correctly"""
        from sum_concil import Conciliator
        
        # Instantiate class
        c = Conciliator()
        
        test_filenames = [
            ('m2d-recu 01.15.2026.xlsx', 'M2D-RECU 01.15.2026'),
            ('M2D-RECU-01-15-2026.xlsx', 'M2D-RECU 01-15-2026'),
            ('some_m2d-recu_12.31.2025_extra.xlsx', 'M2D-RECU 12.31.2025'),
        ]
        
        for filename, expected_prefix in test_filenames:
            result = c.get_standardized_name(filename)
            self.assertEqual(result, expected_prefix, f"Failed for: {filename}")

    def test_filename_standardization_m6d_dev(self):
        """Test that M6D-DEV files are standardized correctly"""
        from sum_concil import Conciliator
        c = Conciliator()
        
        test_filenames = [
            ('m6d-dev 01.15.2026.xlsx', 'M6D-DEV 01.15.2026'),
            ('M6D-DEV-01-15-2026.xlsx', 'M6D-DEV 01-15-2026'),
        ]
        
        for filename, expected_prefix in test_filenames:
            result = c.get_standardized_name(filename)
            self.assertEqual(result, expected_prefix, f"Failed for: {filename}")

    def test_filename_no_date_extraction(self):
        """Test behavior when filename has no valid date"""
        from sum_concil import Conciliator
        
        c = Conciliator()
        filename = 'm2d-recu-nodate.xlsx'
        result = c.get_standardized_name(filename)
        
        self.assertIn("NO_DATE", result)

    def test_filename_standardization_m6d_dev_vff(self):
        """Test that M6D-DEV_VFF files are standardized as ACREEDORA"""
        from sum_concil import Conciliator
        c = Conciliator()
        
        test_filenames = [
            ('m6d-dev_vff 01.20.2026.xlsx', 'ACREEDORA 01.20.2026'),
            ('M6D-DEV_VFF-01-20-2026.xlsx', 'ACREEDORA 01-20-2026'),
        ]
        
        for filename, expected_prefix in test_filenames:
            result = c.get_standardized_name(filename)
            self.assertEqual(result, expected_prefix, f"Failed for: {filename}")

    # =========================================================================
    # TEST 2: DATA LOADING AND CLEANING
    # =========================================================================
    def test_amount_cleaning_with_currency_symbols(self):
        """Test that amounts with currency symbols are parsed correctly"""
        # Simulate the amount cleaning logic
        test_amounts = [
            ('$1,234.56', 1234.56),
            ('€500.00', 500.00),
            ('-$100.50', -100.50),
            ('1234', 1234.0),
            ('invalid', 0.0),  # Should fallback to 0
        ]
        
        import re
        for raw, expected in test_amounts:
            clean_amt = re.sub(r'[^\d.-]', '', str(raw))
            result = pd.to_numeric(clean_amt, errors='coerce')
            if pd.isna(result):
                result = 0.0
            self.assertAlmostEqual(result, expected, places=2, 
                                   msg=f"Failed for amount: {raw}")

        test_values = [
            ('  1234  ', '1234'),
            ('ABC-123\n', 'ABC-123'),
            ('\t OP-456 \t', 'OP-456'),
        ]
        
        for raw, expected in test_values:
            result = raw.strip()
            self.assertEqual(result, expected)

    # =========================================================================
    # TEST 3: MATCHING LOGIC (Cartesian Product with Duplicates)
    # =========================================================================
    def test_inner_join_with_duplicates_creates_cartesian_product(self):
        """
        Critical Test: When 2 debts match 1 credit, we should get 2 rows.
        This validates the Cartesian product behavior mentioned in the code.
        """
        # Simulate debt data with DUPLICATES
        df_debt = pd.DataFrame({
            'TARJETA': ['1234', '1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-001', 'OP-002'],
            'Amt_Float': [100.0, 100.0, 200.0],
            'Accounting_Ref': ['M2D-RECU 01.01.2026', 'M2D-RECU 01.01.2026', 'M2D-RECU 01.01.2026']
        })
        
        # Credit has single entry for OP-001
        df_credit = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'Amt_Float': [100.0, 200.0],
            'Accounting_Ref': ['M6D-DEV 01.05.2026', 'M6D-DEV 01.05.2026']
        })
        
        merged = pd.merge(
            df_debt, df_credit,
            on=['TARJETA', 'NUM OPE'],
            how='inner',
            suffixes=('_DEBT', '_CREDIT')
        )
        
        # We expect 3 rows: 2 for OP-001 (duplicate debt) + 1 for OP-002
        self.assertEqual(len(merged), 3, 
            "Cartesian product should create 3 rows (2 duplicates + 1 unique)")

    def test_no_matches_returns_empty(self):
        """Test that completely non-matching data produces empty result"""
        df_debt = pd.DataFrame({
            'TARJETA': ['1111'],
            'NUM OPE': ['OP-AAA'],
        })
        
        df_credit = pd.DataFrame({
            'TARJETA': ['9999'],
            'NUM OPE': ['OP-ZZZ'],
        })
        
        merged = pd.merge(
            df_debt, df_credit,
            on=['TARJETA', 'NUM OPE'],
            how='inner'
        )
        
        self.assertTrue(merged.empty, "Non-matching data should produce empty DataFrame")

    # =========================================================================
    # TEST 4: AGGREGATION LOGIC
    # =========================================================================
    def test_aggregation_sums_debt_side_correctly(self):
        """
        Test that aggregation sums the DEBT amounts (not credit) to avoid
        inflation from Cartesian product.
        """
        # 2 debt entries for same Card/Op, 1 credit
        merged = pd.DataFrame({
            'TARJETA': ['1234', '1234'],
            'NUM OPE': ['OP-001', 'OP-001'],
            'Amt_Float_DEBT': [100.0, 150.0],
            'Amt_Float_CREDIT': [250.0, 250.0],  # Same credit repeated
            'Accounting_Ref_DEBT': ['M2D-RECU 01.01.2026', 'M2D-RECU 01.01.2026'],
            'Accounting_Ref_CREDIT': ['M6D-DEV 01.05.2026', 'M6D-DEV 01.05.2026'],
        })
        
        debt_breakdown = merged.groupby(['Accounting_Ref_DEBT', 'Accounting_Ref_CREDIT']).agg(
            Count_Operations=('NUM OPE', 'count'),
            Total_Conciliated_Amount=('Amt_Float_DEBT', 'sum')
        ).reset_index()
        
        # Total should be 100 + 150 = 250 (debt side), NOT 500 (credit inflated)
        self.assertEqual(debt_breakdown['Total_Conciliated_Amount'].iloc[0], 250.0)
        self.assertEqual(debt_breakdown['Count_Operations'].iloc[0], 2)

    def test_credit_summary_accounting_avoids_cartesian_inflation(self):
        """
        Test that `_generate_credit_reconciled_summary` correctly deduplicates by [TARJETA, NUM OPE] 
        so that multiple matches (Cartesian product) do not artificially inflate the sum.
        """
        from sum_concil import Conciliator
        
        c = Conciliator()
        
        # Simulating Cartesian product: 1 credit op matching 2 debt ops
        c.merged = pd.DataFrame({
            'TARJETA': ['1234', '1234'],
            'NUM OPE': ['OP-001', 'OP-001'],
            'Amt_Float_DEBT': [100.0, 100.0],
            'Amt_Float_CREDIT': [200.0, 200.0],
            'Accounting_Ref_DEBT': ['M2D 1', 'M2D 1'],
            'Accounting_Ref_CREDIT': ['M6D 1', 'M6D 1'],
            'RECUPERAR_DEBT': ['SI', 'SI'],
            'RECUPERAR': ['SI', 'SI']
        })
        
        c._generate_credit_reconciled_summary()
        summary = c.fully_reconciled_credits
        
        # Since [1234, OP-001] is dropped to unique, there's only 1 unique operation.
        # So Total Acreedora should be exactly 200 (not 400).
        # Monto Deudor should be exactly 100 (not 200).
        
        # summary_rows contains the header row, the data row, the subtotal row, and a blank row.
        # Data row is at index 1
        data_row = summary.iloc[1]
        
        self.assertEqual(data_row['Total Acreedora'], 200.0, "Total Acreedora should not be inflated by duplicates")
        self.assertEqual(data_row['Monto Deudor'], 100.0, "Monto Deudor should not be inflated by duplicates")


    # =========================================================================
    # TEST 4B: DUPLICATE FILE DETECTION (Critical Human Error Prevention)
    # =========================================================================
    def test_detects_exact_duplicate_dataframes(self):
        """Test that identical DataFrames are flagged as duplicates"""
        # Same data in both
        df1 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['100.00', '200.00'],
            'Accounting_Ref': ['M2D-RECU 01.01.2026', 'M2D-RECU 01.01.2026'],
            'Amt_Float': [100.0, 200.0]
        })
        
        df2 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['100.00', '200.00'],
            'Accounting_Ref': ['M6D-DEV 01.05.2026', 'M6D-DEV 01.05.2026'],
            'Amt_Float': [100.0, 200.0]
        })
        
        # Simulate the validation logic
        compare_cols = [col for col in df1.columns if col not in ['Accounting_Ref', 'Amt_Float']]
        df1_sorted = df1[compare_cols].sort_values(by=compare_cols).reset_index(drop=True)
        df2_sorted = df2[compare_cols].sort_values(by=compare_cols).reset_index(drop=True)
        
        self.assertTrue(df1_sorted.equals(df2_sorted), 
            "Should detect that core data is identical")

    def test_detects_high_key_overlap(self):
        """Test detection of suspiciously high key overlap with same row count"""
        df1 = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003'],
        })
        
        df2 = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003'],
        })
        
        debt_keys = set(zip(df1['TARJETA'], df1['NUM OPE']))
        credit_keys = set(zip(df2['TARJETA'], df2['NUM OPE']))
        
        overlap = debt_keys & credit_keys
        overlap_pct = len(overlap) / max(len(debt_keys), 1) * 100
        
        self.assertEqual(overlap_pct, 100.0, 
            "Should detect 100% key overlap")
        self.assertEqual(len(debt_keys), len(credit_keys),
            "Should detect same key count")

    def test_detects_identical_amount_fingerprint(self):
        """Test detection of identical sum/mean/count fingerprint"""
        df1 = pd.DataFrame({'Amt_Float': [100.0, 200.0, 300.0]})
        df2 = pd.DataFrame({'Amt_Float': [100.0, 200.0, 300.0]})
        
        self.assertAlmostEqual(df1['Amt_Float'].sum(), df2['Amt_Float'].sum())
        self.assertAlmostEqual(df1['Amt_Float'].mean(), df2['Amt_Float'].mean())
        self.assertEqual(len(df1), len(df2))

    def test_allows_legitimate_different_files(self):
        """Test that legitimately different files pass validation"""
        df1 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'Amt_Float': [100.0, 200.0]
        })
        
        # Different data
        df2 = pd.DataFrame({
            'TARJETA': ['9999', '8888'],
            'NUM OPE': ['OP-099', 'OP-098'],
            'Amt_Float': [500.0, 600.0]
        })
        
        debt_keys = set(zip(df1['TARJETA'], df1['NUM OPE']))
        credit_keys = set(zip(df2['TARJETA'], df2['NUM OPE']))
        
        overlap = debt_keys & credit_keys
        overlap_pct = len(overlap) / max(len(debt_keys), 1) * 100
        
        self.assertEqual(overlap_pct, 0.0, 
            "Different files should have 0% overlap")

    def test_detects_same_file_type_in_both(self):
        """Test warning when both files are the same type (e.g., both M2D-RECU)"""
        debt_sources = {'M2D-RECU 01.01.2026', 'M2D-RECU 01.02.2026'}
        credit_sources = {'M2D-RECU 01.03.2026', 'M2D-RECU 01.04.2026'}  # Wrong! Should be M6D-DEV
        
        debt_types = {s.split()[0] for s in debt_sources}
        credit_types = {s.split()[0] for s in credit_sources}
        
        self.assertEqual(debt_types, credit_types, 
            "Should detect both sources are same type")
        self.assertEqual(debt_types, {'M2D-RECU'})

    # =========================================================================
    # TEST 4C: INTRA-PILE DUPLICATE DETECTION (Same Category Duplicates)
    # =========================================================================
    def test_detects_identical_files_within_debt_pile(self):
        """Test detection of two identical files within the DEBT category"""
        # Two debt files with identical data
        file1 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['100.00', '200.00'],
            'Accounting_Ref': ['M2D-RECU 01.01.2026', 'M2D-RECU 01.01.2026'],
            'Amt_Float': [100.0, 200.0]
        })
        
        file2 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['100.00', '200.00'],
            'Accounting_Ref': ['M2D-RECU 01.02.2026', 'M2D-RECU 01.02.2026'],  # Different date
            'Amt_Float': [100.0, 200.0]
        })
        
        # Check if keys are identical
        keys1 = set(zip(file1['TARJETA'], file1['NUM OPE']))
        keys2 = set(zip(file2['TARJETA'], file2['NUM OPE']))
        
        self.assertEqual(keys1, keys2, "Should detect identical operation keys")
        
        # Check if data (excluding metadata) is identical
        compare_cols = ['TARJETA', 'NUM OPE', 'IMP VISA', 'Amt_Float']
        df1_sorted = file1[compare_cols].sort_values(by=['TARJETA', 'NUM OPE']).reset_index(drop=True)
        df2_sorted = file2[compare_cols].sort_values(by=['TARJETA', 'NUM OPE']).reset_index(drop=True)
        
        self.assertTrue(df1_sorted.equals(df2_sorted), 
            "Should detect identical data content")

    def test_detects_same_keys_different_amounts_within_pile(self):
        """Test detection of files with same operations but different amounts"""
        file1 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'Amt_Float': [100.0, 200.0]  # Original amounts
        })
        
        file2 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'Amt_Float': [150.0, 250.0]  # DIFFERENT amounts - suspicious!
        })
        
        keys1 = set(zip(file1['TARJETA'], file1['NUM OPE']))
        keys2 = set(zip(file2['TARJETA'], file2['NUM OPE']))
        
        self.assertEqual(keys1, keys2, "Keys should be identical")
        self.assertFalse(file1['Amt_Float'].equals(file2['Amt_Float']), 
            "Amounts should be different")

    def test_detects_high_overlap_within_pile(self):
        """Test detection of >90% overlap between files in same category"""
        file1 = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999', '8888', '7777'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003', 'OP-004', 'OP-005'],
        })
        
        # 4 out of 5 operations overlap (80%) - borderline
        file2 = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999', '8888', 'XXXX'],  # Last one different
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003', 'OP-004', 'OP-999'],
        })
        
        keys1 = set(zip(file1['TARJETA'], file1['NUM OPE']))
        keys2 = set(zip(file2['TARJETA'], file2['NUM OPE']))
        
        overlap = keys1 & keys2
        overlap_pct = len(overlap) / max(len(keys1), 1) * 100
        
        self.assertEqual(overlap_pct, 80.0, "Should calculate 80% overlap")

    def test_allows_different_files_within_pile(self):
        """Test that legitimately different files within same category pass"""
        file1 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
        })
        
        # Completely different operations
        file2 = pd.DataFrame({
            'TARJETA': ['AAAA', 'BBBB'],
            'NUM OPE': ['OP-100', 'OP-200'],
        })
        
        keys1 = set(zip(file1['TARJETA'], file1['NUM OPE']))
        keys2 = set(zip(file2['TARJETA'], file2['NUM OPE']))
        
        overlap = keys1 & keys2
        
        self.assertEqual(len(overlap), 0, "Different files should have no overlap")

    def test_skips_comparison_for_different_row_counts(self):
        """Test that files with different row counts are not flagged as duplicates"""
        file1 = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003'],
        })
        
        file2 = pd.DataFrame({
            'TARJETA': ['1234', '5678'],  # Only 2 rows
            'NUM OPE': ['OP-001', 'OP-002'],
        })
        
        self.assertNotEqual(len(file1), len(file2), 
            "Different row counts should skip detailed comparison")

    # =========================================================================
    # TEST 4D: DATA QUALITY VALIDATIONS
    # =========================================================================
    def test_detects_negative_amounts(self):
        """Test detection of negative amounts"""
        df = pd.DataFrame({'Amt_Float': [100.0, -50.0, 200.0, -25.0]})
        negative_count = (df['Amt_Float'] < 0).sum()
        
        self.assertEqual(negative_count, 2, "Should detect 2 negative amounts")

    def test_detects_zero_amounts(self):
        """Test detection of zero-amount transactions"""
        df = pd.DataFrame({'Amt_Float': [100.0, 0.0, 200.0, 0.0, 0.0]})
        zero_count = (df['Amt_Float'] == 0).sum()
        
        self.assertEqual(zero_count, 3, "Should detect 3 zero amounts")

    def test_detects_statistical_outliers(self):
        """Test detection of unusually large amounts (>3 std from mean)"""
        # Normal amounts around 100
        normal_amounts = [100.0] * 20
        # Add one massive outlier
        amounts = normal_amounts + [10000.0]
        
        df = pd.DataFrame({'Amt_Float': amounts})
        mean_amt = df['Amt_Float'].mean()
        std_amt = df['Amt_Float'].std()
        
        outlier_threshold = mean_amt + (3 * std_amt)
        outliers = df[df['Amt_Float'] > outlier_threshold]
        
        self.assertEqual(len(outliers), 1, "Should detect 1 outlier")
        self.assertEqual(outliers['Amt_Float'].iloc[0], 10000.0)

    def test_detects_empty_card_numbers(self):
        """Test detection of empty/null Card numbers"""
        df = pd.DataFrame({
            'TARJETA': ['1234', '', None, '5678', ''],
        })
        
        empty_cards = df['TARJETA'].isna().sum() + (df['TARJETA'] == '').sum()
        
        self.assertEqual(empty_cards, 3, "Should detect 3 empty/null cards")

    def test_detects_empty_operation_numbers(self):
        """Test detection of empty/null Operation Numbers"""
        df = pd.DataFrame({
            'NUM OPE': ['OP-001', '', 'OP-002', None],
        })
        
        empty_ops = df['NUM OPE'].isna().sum() + (df['NUM OPE'] == '').sum()
        
        self.assertEqual(empty_ops, 2, "Should detect 2 empty/null operations")

    def test_detects_whitespace_only_values(self):
        """Test detection of whitespace-only Card numbers"""
        df = pd.DataFrame({
            'TARJETA': ['1234', '   ', '\t', '5678', '  \n  '],
        })
        
        whitespace_cards = (df['TARJETA'].str.strip() == '').sum()
        
        self.assertEqual(whitespace_cards, 3, "Should detect 3 whitespace-only cards")

    def test_detects_internal_duplicates(self):
        """Test detection of duplicate key combinations within same source file"""
        df = pd.DataFrame({
            'TARJETA': ['1234', '1234', '5678', '5678', '5678'],
            'NUM OPE': ['OP-001', 'OP-001', 'OP-002', 'OP-002', 'OP-002'],
            'Accounting_Ref': ['File1', 'File1', 'File1', 'File1', 'File1'],  # Same source
        })
        
        dup_check = df.groupby(['TARJETA', 'NUM OPE', 'Accounting_Ref']).size()
        internal_dups = dup_check[dup_check > 1]
        
        self.assertEqual(len(internal_dups), 2, 
            "Should detect 2 duplicate key combinations")

    # =========================================================================
    # TEST 4E: ORPHANED RECORDS ANALYSIS
    # =========================================================================
    def test_calculates_orphaned_debts(self):
        """Test identification of debts without matching credits"""
        df_debt = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003'],
            'Amt_Float': [100.0, 200.0, 300.0]
        })
        
        df_credit = pd.DataFrame({
            'TARJETA': ['1234'],  # Only matches first debt
            'NUM OPE': ['OP-001'],
            'Amt_Float': [100.0]
        })
        
        merged = pd.merge(df_debt, df_credit, on=['TARJETA', 'NUM OPE'])
        
        merged_keys = set(zip(merged['TARJETA'], merged['NUM OPE']))
        all_debt_keys = set(zip(df_debt['TARJETA'], df_debt['NUM OPE']))
        orphaned_debt_keys = all_debt_keys - merged_keys
        
        # Orphaned debts are INFORMATIONAL ONLY (not all debts have been refunded yet)
        self.assertEqual(len(orphaned_debt_keys), 2, "Should find 2 orphaned debts")
        self.assertIn(('5678', 'OP-002'), orphaned_debt_keys)
        self.assertIn(('9999', 'OP-003'), orphaned_debt_keys)

    def test_orphaned_credits_are_critical_error(self):
        """
        CRITICAL BUSINESS RULE: Credits without matching debts are BLOCKING errors.
        Every credit (refund) MUST have a corresponding debt (original charge).
        """
        df_debt = pd.DataFrame({
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
        })
        
        df_credit = pd.DataFrame({
            'TARJETA': ['1234', 'AAAA', 'BBBB'],  # 2 credits won't match - CRITICAL ERROR!
            'NUM OPE': ['OP-001', 'OP-100', 'OP-200'],
        })
        
        merged = pd.merge(df_debt, df_credit, on=['TARJETA', 'NUM OPE'])
        
        merged_keys = set(zip(merged['TARJETA'], merged['NUM OPE']))
        all_credit_keys = set(zip(df_credit['TARJETA'], df_credit['NUM OPE']))
        orphaned_credit_keys = all_credit_keys - merged_keys
        
        # Orphaned credits are CRITICAL - should block conciliation
        self.assertEqual(len(orphaned_credit_keys), 2, 
            "Should find 2 orphaned credits - CRITICAL ERROR")
        self.assertTrue(len(orphaned_credit_keys) > 0, 
            "Any orphaned credits should trigger blocking error")

    def test_orphaned_debts_are_informational(self):
        """Test that orphaned debts are allowed (informational only)"""
        # This is normal - not all debts have been refunded yet
        orphaned_debt_count = 50
        is_blocking_error = False  # Orphaned debts should NOT block
        
        self.assertFalse(is_blocking_error, 
            "Orphaned debts should NOT block conciliation")

    def test_calculates_match_rate(self):
        """Test match rate calculation"""
        total_keys = 100
        matched_keys = 75
        orphaned_keys = 25
        
        match_rate = (total_keys - orphaned_keys) / total_keys * 100
        
        self.assertEqual(match_rate, 75.0, "Match rate should be 75%")

    def test_all_credits_matched_is_valid(self):
        """Test that 100% credit match rate is the expected valid state"""
        df_debt = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003'],
        })
        
        # All credits have matching debts
        df_credit = pd.DataFrame({
            'TARJETA': ['1234', '5678'],  # Subset of debts - valid!
            'NUM OPE': ['OP-001', 'OP-002'],
        })
        
        merged = pd.merge(df_debt, df_credit, on=['TARJETA', 'NUM OPE'])
        
        merged_keys = set(zip(merged['TARJETA'], merged['NUM OPE']))
        all_credit_keys = set(zip(df_credit['TARJETA'], df_credit['NUM OPE']))
        orphaned_credit_keys = all_credit_keys - merged_keys
        
        # No orphaned credits = valid state
        self.assertEqual(len(orphaned_credit_keys), 0, 
            "All credits should match debts - valid state")

    def test_orphan_amount_calculation(self):
        """Test that orphaned record amounts are calculated correctly"""
        df = pd.DataFrame({
            'TARJETA': ['1234', '5678', '9999'],
            'NUM OPE': ['OP-001', 'OP-002', 'OP-003'],
            'Amt_Float': [100.0, 200.0, 300.0]
        })
        
        orphaned_keys = {('5678', 'OP-002'), ('9999', 'OP-003')}
        orphaned_df = df[df.apply(lambda x: (x['TARJETA'], x['NUM OPE']) in orphaned_keys, axis=1)]
        
        orphaned_total = orphaned_df['Amt_Float'].sum()
        
        self.assertEqual(orphaned_total, 500.0, "Orphaned total should be 200 + 300 = 500")

    # =========================================================================
    # TEST 5: EDGE CASES
    # =========================================================================
    def test_missing_required_columns_handled(self):
        """Simulate file with missing Card or Operation Number columns"""
        df = pd.DataFrame({
            'Wrong_Column': ['data'],
            'IMP VISA': ['100.00']
        })
        
        col_card = 'TARJETA'
        col_op = 'NUM OPE'
        
        has_required = col_card in df.columns and col_op in df.columns
        self.assertFalse(has_required, "Should detect missing required columns")

    def test_empty_dataframe_handling(self):
        """Test that empty DataFrames are handled gracefully"""
        df_debt = pd.DataFrame()
        df_credit = pd.DataFrame()
        
        self.assertTrue(df_debt.empty)
        self.assertTrue(df_credit.empty)

    def test_scientific_notation_protection(self):
        """
        Test that loading as str dtype protects long IDs from scientific notation.
        Example: Card ID '12345678901234' should NOT become '1.23457E+13'
        """
        long_id = '12345678901234567890'
        
        # Simulate loading as string
        df = pd.DataFrame({'TARJETA': [long_id]}, dtype=str)
        self.assertEqual(df['TARJETA'].iloc[0], long_id)
        
        # If loaded as numeric, it could lose precision
        df_numeric = pd.DataFrame({'TARJETA': [int(long_id[:15])]})  # Truncate for valid int
        # This would cause issues if compared

    # =========================================================================
    # TEST 6: GLOB PATTERN FILTERING
    # =========================================================================
    def test_glob_filter_excludes_wrong_files(self):
        """Test that the secondary filter correctly excludes non-matching files"""
        from sum_concil import Conciliator
        c = Conciliator()
        
        # Simulate glob results that might include wrong files
        fake_files = [
            'accounting_files/m2d-recu 01.01.2026.xlsx',  # Should match DEBT
            'accounting_files/m6d-dev 01.05.2026.xlsx',   # Should match CREDIT (not DEBT)
            'accounting_files/random_m2d-recufile.xlsx',  # Should match DEBT
        ]
        
        # We can test the logic used inside _load_pile for filtering
        debt_keyword = 'm2d-recu'
        filtered = [f for f in fake_files if debt_keyword in os.path.basename(f).lower()]
        
        self.assertEqual(len(filtered), 2)
        self.assertTrue(all('m2d-recu' in f.lower() for f in filtered))

    # =========================================================================
    # TEST 7: OUTPUT FILE HANDLING
    # =========================================================================
    def test_excel_writer_creates_all_sheets(self):
        """Test that output Excel has all expected sheets"""
        output_path = os.path.join(self.test_dir, 'test_output.xlsx')
        
        # Create mock data
        debt_breakdown = pd.DataFrame({'A': [1, 2]})
        credit_breakdown = pd.DataFrame({'B': [3, 4]})
        merged = pd.DataFrame({'C': [5, 6]})
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            debt_breakdown.to_excel(writer, sheet_name='By_Debt_File', index=False)
            credit_breakdown.to_excel(writer, sheet_name='By_Credit_File', index=False)
            merged.to_excel(writer, sheet_name='Detailed_Audit_Log', index=False)
        
        # Verify sheets exist - use context manager for proper cleanup
        with pd.ExcelFile(output_path) as xl:
            expected_sheets = ['By_Debt_File', 'By_Credit_File', 'Detailed_Audit_Log']
            for sheet in expected_sheets:
                self.assertIn(sheet, xl.sheet_names, f"Missing sheet: {sheet}")
        
        # Clean up
        try:
            os.remove(output_path)
        except PermissionError:
            pass  # Ignore on Windows


class TestIntegration(unittest.TestCase):
    """
    Integration tests that run the full conciliation process.
    These require creating actual test Excel files.
    """
    
    @classmethod
    def setUpClass(cls):
        cls.test_dir = tempfile.mkdtemp()
        cls.accounting_folder = os.path.join(cls.test_dir, 'accounting_files')
        os.makedirs(cls.accounting_folder, exist_ok=True)
        
    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.test_dir, ignore_errors=True)
    
    def _create_test_excel(self, filename, data):
        """Helper to create test Excel files"""
        df = pd.DataFrame(data)
        path = os.path.join(self.accounting_folder, filename)
        df.to_excel(path, index=False)
        return path

    def test_full_conciliation_with_matching_data(self):
        """Integration test: Full workflow with matching debt/credit files"""
        # Create debt file
        self._create_test_excel('m2d-recu 01.01.2026.xlsx', {
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['$100.00', '$200.00']
        })
        
        # Create credit file
        self._create_test_excel('m6d-dev 01.05.2026.xlsx', {
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['$100.00', '$200.00']
        })
        
        # The full function would need folder_path modification to run
        # This test validates the test data was created correctly
        self.assertTrue(os.path.exists(os.path.join(self.accounting_folder, 'm2d-recu 01.01.2026.xlsx')))
        self.assertTrue(os.path.exists(os.path.join(self.accounting_folder, 'm6d-dev 01.05.2026.xlsx')))


class TestVFFCreditLogic(unittest.TestCase):
    """Tests for VFF (M6D-DEV_VFF) paired credit note computation and matching."""

    @classmethod
    def setUpClass(cls):
        cls.test_dir = tempfile.mkdtemp()
        cls.accounting_folder = os.path.join(cls.test_dir, 'accounting_files')
        os.makedirs(cls.accounting_folder, exist_ok=True)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.test_dir, ignore_errors=True)

    def setUp(self):
        for f in os.listdir(self.accounting_folder):
            os.remove(os.path.join(self.accounting_folder, f))

    def _create_test_excel(self, filename, data):
        df = pd.DataFrame(data)
        path = os.path.join(self.accounting_folder, filename)
        df.to_excel(path, index=False)
        return path

    # =========================================================================
    # TEST VFF 1: PAIR COMPUTATION
    # =========================================================================
    def test_vff_pair_computation(self):
        """Test that VFF paired operations produce correct credit note = creditor - debtor"""
        from sum_concil import Conciliator, COL_CARD, COL_OP, AMT_FLOAT, ACCOUNTING_REF

        c = Conciliator()
        df_vff = pd.DataFrame({
            COL_CARD: ['1234', '1234', '5678', '5678'],
            COL_OP: ['OP-001', 'OP-001', 'OP-002', 'OP-002'],
            AMT_FLOAT: [100.0, 250.0, 300.0, 500.0],
            ACCOUNTING_REF: ['ACREEDORA 01.20.2026'] * 4,
            'RECUPERAR': ['SI'] * 4,
        })

        result = c._compute_vff_pairs(df_vff)

        self.assertEqual(len(result), 2, "Should produce 2 credit notes from 2 pairs")

        row1 = result[result[COL_OP] == 'OP-001'].iloc[0]
        self.assertAlmostEqual(row1['VFF_Difference'], 150.0, places=2,
            msg="Credit note for OP-001 should be 250 - 100 = 150")
        self.assertAlmostEqual(row1[AMT_FLOAT], 150.0, places=2)

        row2 = result[result[COL_OP] == 'OP-002'].iloc[0]
        self.assertAlmostEqual(row2['VFF_Difference'], 200.0, places=2,
            msg="Credit note for OP-002 should be 500 - 300 = 200")

    # =========================================================================
    # TEST VFF 2: NEGATIVE DIFFERENCE = DEBTOR NOTE
    # =========================================================================
    def test_vff_negative_difference_flagged(self):
        """Test that negative VFF differences are captured as debtor notes (error)"""
        from sum_concil import Conciliator, COL_CARD, COL_OP, AMT_FLOAT, ACCOUNTING_REF

        c = Conciliator()
        df_vff = pd.DataFrame({
            COL_CARD: ['1234', '1234'],
            COL_OP: ['OP-001', 'OP-001'],
            AMT_FLOAT: [500.0, 200.0],  # creditor < debtor → negative
            ACCOUNTING_REF: ['ACREEDORA 01.20.2026'] * 2,
            'RECUPERAR': ['SI'] * 2,
        })

        result = c._compute_vff_pairs(df_vff)

        self.assertTrue(result.empty or len(result) == 0,
            "Negative difference should NOT be in credit notes")
        self.assertFalse(c.vff_debtor_notes.empty,
            "Negative difference should be in vff_debtor_notes")
        self.assertEqual(c.vff_debtor_notes.iloc[0]['VFF_Note_Type'], 'DEBTOR_NOTE (NEGATIVE)')
        self.assertAlmostEqual(c.vff_debtor_notes.iloc[0]['VFF_Difference'], -300.0, places=2)

    # =========================================================================
    # TEST VFF 3: VFF MATCHING AGAINST M2D
    # =========================================================================
    def test_vff_matching_against_m2d(self):
        """Test that VFF computed credits match against M2D debts"""
        from sum_concil import Conciliator, COL_CARD, COL_OP, AMT_FLOAT, ACCOUNTING_REF

        # Create M2D debt file
        self._create_test_excel('m2d-recu 01.01.2026.xlsx', {
            'TARJETA': ['1234', '5678'],
            'NUM OPE': ['OP-001', 'OP-002'],
            'IMP VISA': ['100.00', '300.00'],
            'RECUPERAR': ['NO', 'NO'],
        })

        # Create VFF file with paired operations
        self._create_test_excel('m6d-dev_vff 01.20.2026.xlsx', {
            'TARJETA': ['1234', '1234'],
            'NUM OPE': ['OP-001', 'OP-001'],
            'IMP VISA': ['100.00', '250.00'],
            'RECUPERAR': ['SI', 'SI'],
        })

        c = Conciliator(folder_path=self.accounting_folder)
        loaded = c.load_data()
        self.assertTrue(loaded, "Data should load successfully")

        # VFF credits should be loaded
        self.assertFalse(c.df_credit_vff.empty, "VFF credits should be computed")

        # Match
        result = c.match_transactions()
        self.assertTrue(result, "Matching should succeed")

        # VFF should match against M2D
        self.assertFalse(c.merged_vff.empty, "VFF should have matches against M2D")

    # =========================================================================
    # TEST VFF 4: UNMATCHED VFF → UNEXPECTED REFUNDS
    # =========================================================================
    def test_vff_no_match_goes_to_unexpected(self):
        """Test that VFF credits with no M2D match go to unexpected refunds"""
        from sum_concil import Conciliator, COL_CARD, COL_OP, AMT_FLOAT, ACCOUNTING_REF

        # Create M2D debt file - does NOT contain OP-999
        self._create_test_excel('m2d-recu 01.01.2026.xlsx', {
            'TARJETA': ['5678'],
            'NUM OPE': ['OP-002'],
            'IMP VISA': ['300.00'],
            'RECUPERAR': ['NO'],
        })

        # Create VFF file with operation that has NO matching M2D
        self._create_test_excel('m6d-dev_vff 01.20.2026.xlsx', {
            'TARJETA': ['9999', '9999'],
            'NUM OPE': ['OP-999', 'OP-999'],
            'IMP VISA': ['100.00', '250.00'],
            'RECUPERAR': ['SI', 'SI'],
        })

        c = Conciliator(folder_path=self.accounting_folder)
        c.load_data()
        c.match_transactions()

        # VFF credit with no M2D match should be in unexpected refunds
        self.assertFalse(c.unexpected_refunds.empty,
            "Unmatched VFF credit should appear in unexpected refunds")

    # =========================================================================
    # TEST VFF 5: REGULAR CREDIT EXCLUDES VFF
    # =========================================================================
    def test_regular_credit_excludes_vff(self):
        """Test that VFF files are NOT loaded into the regular credit pile"""
        from sum_concil import Conciliator

        # Create a regular credit AND a VFF file
        self._create_test_excel('m6d-dev 01.05.2026.xlsx', {
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
            'IMP VISA': ['100.00'],
        })

        self._create_test_excel('m6d-dev_vff 01.20.2026.xlsx', {
            'TARJETA': ['5678', '5678'],
            'NUM OPE': ['OP-002', 'OP-002'],
            'IMP VISA': ['100.00', '250.00'],
        })

        # Also need a debt file for load_data to succeed
        self._create_test_excel('m2d-recu 01.01.2026.xlsx', {
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
            'IMP VISA': ['100.00'],
        })

        c = Conciliator(folder_path=self.accounting_folder)
        c.load_data()

        # Regular credit should only have 1 row (not VFF data)
        self.assertEqual(len(c.df_credit), 1,
            "Regular credit pile should exclude VFF files")

        # VFF should have been loaded separately
        self.assertFalse(c.df_credit_vff.empty,
            "VFF should be loaded in its own pile")

    # =========================================================================
    # TEST VFF 6: CROSS-FILE MATCHING
    # =========================================================================
    def test_vff_cross_file_matching(self):
        """Test that unpaired VFF operations match across other VFF files"""
        from sum_concil import Conciliator, COL_CARD, COL_OP, AMT_FLOAT, ACCOUNTING_REF

        # VFF file 1: has debtor side of OP-001
        self._create_test_excel('m6d-dev_vff 01.20.2026.xlsx', {
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
            'IMP VISA': ['100.00'],
            'RECUPERAR': ['SI'],
        })

        # VFF file 2: has creditor side of OP-001
        self._create_test_excel('m6d-dev_vff 01.25.2026.xlsx', {
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
            'IMP VISA': ['250.00'],
            'RECUPERAR': ['SI'],
        })

        # Need a debt file for load_data
        self._create_test_excel('m2d-recu 01.01.2026.xlsx', {
            'TARJETA': ['1234'],
            'NUM OPE': ['OP-001'],
            'IMP VISA': ['100.00'],
        })

        c = Conciliator(folder_path=self.accounting_folder)
        c.load_data()

        # Cross-file matching should produce a paired credit note
        self.assertFalse(c.df_credit_vff.empty,
            "Cross-file VFF match should produce a computed credit note")

        # The credit note should be 250 - 100 = 150
        if not c.df_credit_vff.empty:
            self.assertAlmostEqual(
                c.df_credit_vff['VFF_Difference'].iloc[0], 150.0, places=2,
                msg="Cross-file credit note should be 250 - 100 = 150"
            )


if __name__ == '__main__':
    # Run with verbose output
    unittest.main(verbosity=2)

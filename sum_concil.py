import pandas as pd
import os
import glob
import re

# --- HEADERS CONFIGURATION ---
# --- HEADERS CONFIGURATION ---
DEFAULT_CONFIG = {
    # Input CSV/Excel Headers
    'COL_CARD': 'TARJETA',
    'COL_OP': 'NUM OPE',
    'COL_AMOUNT': 'IMP VISA',
    'COL_RECUPERAR': 'RECUPERAR',
    
    # Internal Technical Columns
    'AMT_FLOAT': 'Amt_Float',
    'ACCOUNTING_REF': 'Accounting_Ref',
    
    # Output Report Headers (Personalizable)
    'OUT_DEBTOR_FILE': 'DEBTOR FILE',
    'OUT_CREDIT_NOTE': 'CREDIT FILE NOTE',
    'OUT_MATCHED_AMT': 'AMOUNT THAT MATCHED',
    'OUT_STATUS': 'Status',
    'OUT_TYPE': 'Type',
    'OUT_DEBTOR_FILE_COL': 'Debtor_File' # Field name used in net_balanced dictionary
}

# Keep these for backward compatibility of external scripts referencing them, 
# but internally use self.config
COL_CARD = DEFAULT_CONFIG['COL_CARD']
COL_OP = DEFAULT_CONFIG['COL_OP']
COL_AMOUNT = DEFAULT_CONFIG['COL_AMOUNT']
COL_RECUPERAR = DEFAULT_CONFIG['COL_RECUPERAR']
AMT_FLOAT = DEFAULT_CONFIG['AMT_FLOAT']
ACCOUNTING_REF = DEFAULT_CONFIG['ACCOUNTING_REF']

DEBT_PATTERN = '*m2d-recu*.xlsx'
CREDIT_PATTERN = '*m6d-dev*.xlsx'
CREDIT_PATTERN_VFF = '*m6d-dev_vff*.xlsx'
DEFAULT_FOLDER_PATH = './accounting_files'

class Conciliator:
    """
    Handles the conciliation process between Debt (M2D-RECU) and Credit (M6D-DEV) files.
    """
    
    def __init__(self, folder_path=DEFAULT_FOLDER_PATH, config=None):
        self.folder_path = folder_path
        self.config = DEFAULT_CONFIG.copy()
        if config:
            self.config.update(config)
            
        self.df_debt = pd.DataFrame()
        self.df_credit = pd.DataFrame()
        self.df_credit_vff = pd.DataFrame()
        self.debt_files = {}
        self.credit_files = {}
        self.credit_vff_files = {}
        self.merged = pd.DataFrame()
        self.merged_vff = pd.DataFrame()
        
        # Results
        self.pending_claims = pd.DataFrame()
        self.unexpected_refunds = pd.DataFrame()
        self.variance_report = pd.DataFrame()
        self.fully_reconciled = pd.DataFrame()
        self.net_balanced = pd.DataFrame()
        self.vff_debtor_notes = pd.DataFrame()  # Negative VFF differences (error)
        self.vff_acreedoras = pd.DataFrame()     # Positive VFF differences (credit notes)
        self.m6d_sin_match = pd.DataFrame()      # Orphaned M6D credits (no matching M2D)
        
        # State
        self.merged_keys = set()
        self.bad_credit_keys = set()

    # =========================================================================
    # 1. HELPER FUNCTIONS
    # =========================================================================
    def get_standardized_name(self, filepath):
        """Standardizes filename to M2D-RECU <DATE> or M6D-DEV <DATE>"""
        filename = os.path.basename(filepath)
        name_lower = filename.lower()
        
        date_match = re.search(r'(\d+[\.-]\d+[\.-]\d+)', name_lower)
        date_str = date_match.group(1) if date_match else "NO_DATE"
    # =========================================================================
    # 1.1. STANDARDIZE NAME
    # =========================================================================
        if 'm2d-recu' in name_lower:
            return f"M2D-RECU {date_str}"
        elif 'm6d-dev_vff' in name_lower:
            return f"ACREEDORA {date_str}"
        elif 'm6d-dev' in name_lower:
            return f"M6D-DEV {date_str}"
        else:
            return f"UNKNOWN {filename}"

    # =========================================================================
    # 2. DATA LOADING
    # =========================================================================
    def load_data(self):
        """Loads all data from the configured folder path."""
        self.df_debt, self.debt_files = self._load_pile(DEBT_PATTERN, "DEBT")
        self.df_credit, self.credit_files = self._load_pile(CREDIT_PATTERN, "CREDIT")
        self.df_credit_vff, self.credit_vff_files = self._load_vff_pile()
        
        has_any_credit = not self.df_credit.empty or not self.df_credit_vff.empty
        if self.df_debt.empty or not has_any_credit:
            print("Stopping: Missing data.")
            return False
        return True

    def _load_pile(self, pattern, label):
        """Internal helper to load files matching a pattern."""
        files = glob.glob(os.path.join(self.folder_path, pattern))
        filter_keyword = 'm2d-recu' if label == "DEBT" else 'm6d-dev'
        files = [f for f in files if filter_keyword in os.path.basename(f).lower()]
        # Exclude VFF files from regular credit pile
        if label == "CREDIT":
            files = [f for f in files if 'm6d-dev_vff' not in os.path.basename(f).lower()]
        
        all_dfs = []
        individual_files = {}
        print(f"Loading {len(files)} files for {label}...")

        for f in files:
            try:
                df = self._process_single_file(f)
                if df is not None:
                    # Store result
                    std_name = df[ACCOUNTING_REF].iloc[0] # Use the name we just set
                    all_dfs.append(df)
                    individual_files[std_name] = df.copy()
            except Exception as e:
                print(f"  [ERROR] {os.path.basename(f)}: {e}")
        
        combined = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
        return combined, individual_files

    def _load_vff_pile(self):
        """
        Loads VFF (M6D-DEV_VFF) files and computes credit notes from paired operations.
        
        VFF Business Rules:
        - Each VFF file contains paired operations (same TARJETA + NUM OPE)
        - Within each pair: row 1 = debtor import, row 2 = creditor import
        - Credit note = creditor_amount - debtor_amount
        - If difference is negative → debtor note (stored as error in self.vff_debtor_notes)
        - Unpaired operations in one file are matched across other VFF files
        """
        files = glob.glob(os.path.join(self.folder_path, CREDIT_PATTERN_VFF))
        files = [f for f in files if 'm6d-dev_vff' in os.path.basename(f).lower()]
        
        all_dfs = []
        individual_files = {}
        print(f"Loading {len(files)} files for CREDIT_VFF...")
        
        for f in files:
            try:
                df = self._process_single_file(f)
                if df is not None:
                    std_name = df[ACCOUNTING_REF].iloc[0]
                    all_dfs.append(df)
                    individual_files[std_name] = df.copy()
            except Exception as e:
                print(f"  [ERROR] {os.path.basename(f)}: {e}")
        
        if not all_dfs:
            print("  No VFF files found.")
            return pd.DataFrame(), {}
        
        # Concatenate all VFF data for cross-file matching
        combined = pd.concat(all_dfs, ignore_index=True)
        computed_credits = self._compute_vff_pairs(combined)
        
        return computed_credits, individual_files

    def _compute_vff_pairs(self, df_vff):
        """
        Groups VFF rows by (TARJETA, NUM OPE) and computes credit notes.
        Row 1 of each pair = debtor, Row 2 = creditor.
        Returns a DataFrame of computed credit notes (one row per pair).
        """
        if df_vff.empty:
            return pd.DataFrame()
        
        credit_rows = []
        debtor_note_rows = []
        
        grouped = df_vff.groupby([COL_CARD, COL_OP])
        
        for (card, op), group in grouped:
            group = group.reset_index(drop=True)
            
            if len(group) == 2:
                # Normal pair: row 0 = debtor, row 1 = creditor
                debtor_amt = group[AMT_FLOAT].iloc[0]
                creditor_amt = group[AMT_FLOAT].iloc[1]
                difference = creditor_amt - debtor_amt
                
                row = {
                    COL_CARD: card,
                    COL_OP: op,
                    AMT_FLOAT: abs(difference),
                    ACCOUNTING_REF: group[ACCOUNTING_REF].iloc[0],
                    'VFF_Debtor_Amt': debtor_amt,
                    'VFF_Creditor_Amt': creditor_amt,
                    'VFF_Difference': difference,
                    'VFF_Source_Files': ', '.join(group[ACCOUNTING_REF].unique()),
                }
                # Copy RECUPERAR if available
                if COL_RECUPERAR in group.columns:
                    row[COL_RECUPERAR] = group[COL_RECUPERAR].iloc[0]
                
                if difference < 0:
                    row['VFF_Note_Type'] = 'DEBTOR_NOTE (NEGATIVE)'
                    debtor_note_rows.append(row)
                    print(f"  ⚠ VFF DEBTOR NOTE: {card}/{op} diff={difference:.2f}")
                else:
                    row['VFF_Note_Type'] = 'CREDIT_NOTE'
                    credit_rows.append(row)
                    
            elif len(group) == 1:
                # Single operation with no pair — should not normally happen
                print(f"  ⚠ VFF UNPAIRED: {card}/{op} in {group[ACCOUNTING_REF].iloc[0]} (matched cross-file)")
                # This was already handled by cross-file concatenation;
                # if still single after concat, it's truly unpaired
                row = {
                    COL_CARD: card,
                    COL_OP: op,
                    AMT_FLOAT: group[AMT_FLOAT].iloc[0],
                    ACCOUNTING_REF: group[ACCOUNTING_REF].iloc[0],
                    'VFF_Debtor_Amt': group[AMT_FLOAT].iloc[0],
                    'VFF_Creditor_Amt': 0.0,
                    'VFF_Difference': 0.0,
                    'VFF_Source_Files': group[ACCOUNTING_REF].iloc[0],
                    'VFF_Note_Type': 'UNPAIRED',
                }
                if COL_RECUPERAR in group.columns:
                    row[COL_RECUPERAR] = group[COL_RECUPERAR].iloc[0]
                debtor_note_rows.append(row)
            else:
                # More than 2 rows for same key — unexpected
                print(f"  ⚠ VFF ANOMALY: {card}/{op} has {len(group)} rows")
                # Take the first two as the pair
                debtor_amt = group[AMT_FLOAT].iloc[0]
                creditor_amt = group[AMT_FLOAT].iloc[1]
                difference = creditor_amt - debtor_amt
                row = {
                    COL_CARD: card,
                    COL_OP: op,
                    AMT_FLOAT: abs(difference),
                    ACCOUNTING_REF: group[ACCOUNTING_REF].iloc[0],
                    'VFF_Debtor_Amt': debtor_amt,
                    'VFF_Creditor_Amt': creditor_amt,
                    'VFF_Difference': difference,
                    'VFF_Source_Files': ', '.join(group[ACCOUNTING_REF].unique()),
                    'VFF_Note_Type': 'ANOMALY_MULTI_ROW',
                }
                if COL_RECUPERAR in group.columns:
                    row[COL_RECUPERAR] = group[COL_RECUPERAR].iloc[0]
                debtor_note_rows.append(row)
        
        # Store debtor notes (negative differences + unpaired + anomalies)
        self.vff_debtor_notes = pd.DataFrame(debtor_note_rows) if debtor_note_rows else pd.DataFrame()
        if not self.vff_debtor_notes.empty:
            print(f"  ⚠ Found {len(self.vff_debtor_notes)} VFF debtor notes / errors")
        
        # Store valid credit notes
        self.vff_acreedoras = pd.DataFrame(credit_rows) if credit_rows else pd.DataFrame()
        if not self.vff_acreedoras.empty:
            print(f"  ✓ Computed {len(self.vff_acreedoras)} VFF credit notes (acreedoras)")
        
        return self.vff_acreedoras

    def _process_single_file(self, filepath):
        """Reads and cleans a single Excel file."""
        # Load as String
        df = pd.read_excel(filepath, dtype=str)
        
        if COL_CARD in df.columns and COL_OP in df.columns:
            df[COL_CARD] = df[COL_CARD].replace(r'^\s*$', pd.NA, regex=True)
            df[COL_OP] = df[COL_OP].replace(r'^\s*$', pd.NA, regex=True)
            
            # Drop empty trailing rows
            df = df.dropna(subset=[COL_CARD, COL_OP], how='all')
        
        # Standardized Name
        std_name = self.get_standardized_name(filepath)
        df[ACCOUNTING_REF] = std_name
        
        # Clean Keys
        if COL_CARD in df.columns and COL_OP in df.columns:
            df[COL_CARD] = df[COL_CARD].str.strip()
            df[COL_OP] = df[COL_OP].str.strip()
        else:
            print(f"  [SKIP] {std_name} missing Card or Operation headers.")
            return None
        
        # Clean Amount
        if COL_AMOUNT in df.columns:
            clean_amt = df[COL_AMOUNT].astype(str).str.replace(r'[^\d.-]', '', regex=True)
            df[AMT_FLOAT] = pd.to_numeric(clean_amt, errors='coerce').fillna(0.0)
        
        # Clean RECUPERAR
        if COL_RECUPERAR in df.columns:
            df[COL_RECUPERAR] = df[COL_RECUPERAR].astype(str).str.strip().str.upper()
        else:
            df[COL_RECUPERAR] = 'SI'

        return df

    # =========================================================================
    # 3. VALIDATION
    # =========================================================================
    def validate(self):
        """Orchestrates validation checks."""
        print("Checking for duplicate files within each category...")
        intra_issues = self._check_intra_pile_duplicates(self.debt_files, "DEBT") + \
                       self._check_intra_pile_duplicates(self.credit_files, "CREDIT")
        
        if intra_issues:
            self._print_validation_errors("INTRA-CATEGORY DUPLICATE DETECTION", intra_issues)
            return False

        print("Validating files are not duplicates...")
        dups = self._validate_files_are_different(self.df_debt, self.df_credit)
        if dups:
            self._print_validation_errors("DUPLICATE FILE DETECTION", dups)
            return False

        print("Running data quality checks...")
        if not self._run_quality_checks():
            return False
            
        return True

    def _print_validation_errors(self, title, errors):
        print(f"\n{'='*60}\n⚠️  {title} ⚠️\n{'='*60}")
        for e in errors: print(f"  ❌ {e}")
        print("\nConciliation ABORTED.\n")

    def _check_intra_pile_duplicates(self, individual_files, label):
        issues = []
        # Optimization: Only compare files with same number of rows
        from collections import defaultdict
        files_by_len = defaultdict(list)
        
        for name, df in individual_files.items():
            files_by_len[len(df)].append(name)
            
        for length, names in files_by_len.items():
            if len(names) < 2: continue
            
            # Compare only within this group
            for i in range(len(names)):
                for j in range(i + 1, len(names)):
                    name1, name2 = names[i], names[j]
                    df1, df2 = individual_files[name1], individual_files[name2]
                    
                    keys1 = set(zip(df1[self.config['COL_CARD']], df1[self.config['COL_OP']]))
                    keys2 = set(zip(df2[self.config['COL_CARD']], df2[self.config['COL_OP']]))
                    
                    if keys1 == keys2:
                        compare_cols = [col for col in df1.columns if col not in [self.config['ACCOUNTING_REF']]]
                        df1_sorted = df1[compare_cols].sort_values(by=[self.config['COL_CARD'], self.config['COL_OP']]).reset_index(drop=True)
                        df2_sorted = df2[compare_cols].sort_values(by=[self.config['COL_CARD'], self.config['COL_OP']]).reset_index(drop=True)
                        
                        if df1_sorted.equals(df2_sorted):
                            issues.append(f"DUPLICATE {label} FILES: '{name1}' and '{name2}' contain IDENTICAL data!")
                        else:
                            issues.append(f"SUSPICIOUS {label} FILES: '{name1}' and '{name2}' have identical operations but different amounts!")
                    else:
                        overlap = keys1 & keys2
                        overlap_pct = len(overlap) / max(len(keys1), 1) * 100
                        if overlap_pct > 90:
                            issues.append(f"WARNING {label}: '{name1}' and '{name2}' share {overlap_pct:.1f}% of operations!")
        return issues

    def _validate_files_are_different(self, df1, df2):
        issues = []
        compare_cols = [col for col in df1.columns if col not in [ACCOUNTING_REF, AMT_FLOAT]]
        
        # 1. Exact DataFrame Match
        if set(compare_cols) == set([col for col in df2.columns if col not in [ACCOUNTING_REF, AMT_FLOAT]]):
            df1_cmp = df1[compare_cols].reset_index(drop=True)
            df2_cmp = df2[compare_cols].reset_index(drop=True)
            if len(df1_cmp) == len(df2_cmp):
                df1_s = df1_cmp.sort_values(by=compare_cols).reset_index(drop=True)
                df2_s = df2_cmp.sort_values(by=compare_cols).reset_index(drop=True)
                if df1_s.equals(df2_s):
                    issues.append("EXACT MATCH: DEBT and CREDIT files contain identical data!")

        # 2. Key Overlap
        debt_keys = set(zip(df1[COL_CARD], df1[COL_OP]))
        credit_keys = set(zip(df2[COL_CARD], df2[COL_OP]))
        overlap_pct = len(debt_keys & credit_keys) / max(len(debt_keys), 1) * 100
        if overlap_pct > 95 and len(debt_keys) == len(credit_keys):
            issues.append(f"SUSPICIOUS: {overlap_pct:.1f}% key overlap with same row count!")

        # 3. Amount Fingerprint
        if AMT_FLOAT in df1.columns and AMT_FLOAT in df2.columns:
            if (abs(df1[AMT_FLOAT].sum() - df2[AMT_FLOAT].sum()) < 0.01 and 
                abs(df1[AMT_FLOAT].mean() - df2[AMT_FLOAT].mean()) < 0.01 and
                len(df1) == len(df2)):
                issues.append("SUSPICIOUS: Identical sum, mean, and row count!")

        # 4. Source Type Check
        debt_sources = {s.split()[0] for s in df1[ACCOUNTING_REF].unique()}
        credit_sources = {s.split()[0] for s in df2[ACCOUNTING_REF].unique()}
        if debt_sources == credit_sources:
            issues.append(f"WARNING: Both sources are type '{debt_sources}' - expected different types!")
            
        return issues

    def _run_quality_checks(self):
        all_warnings, all_errors = [], []
        
        for f, df in self.debt_files.items():
            w, e = self._check_data_quality(df, f"DEBT ({f})")
            all_warnings.extend(w); all_errors.extend(e)
        for f, df in self.credit_files.items():
            w, e = self._check_data_quality(df, f"CREDIT ({f})")
            all_warnings.extend(w); all_errors.extend(e)

        if all_warnings:
            print(f"\n{'-'*60}\n⚠️  DATA QUALITY WARNINGS\n{'-'*60}")
            for w in all_warnings: print(f"  ⚠ {w}")
        
        if all_errors:
            self._print_validation_errors("DATA QUALITY ERRORS", all_errors)
            return False
        return True

    def _check_data_quality(self, df, label):
        warnings = []
        errors = []
        
        if AMT_FLOAT in df.columns:
            if (df[AMT_FLOAT] < 0).any(): warnings.append(f"{label}: Found negative amounts")
            if (df[AMT_FLOAT] == 0).any(): warnings.append(f"{label}: Found zero-amount transactions")
            
            # Outliers (>3 std)
            if len(df) > 10:
                mean, std = df[AMT_FLOAT].mean(), df[AMT_FLOAT].std()
                if std > 0 and not df[df[AMT_FLOAT] > mean + 3*std].empty:
                    warnings.append(f"{label}: Found unusually large amounts")

        if COL_CARD in df.columns:
            empty = df[COL_CARD].isna().sum() + (df[COL_CARD] == '').sum()
            if empty > 0: errors.append(f"{label}: {empty} rows with empty Card numbers")
            
        if COL_OP in df.columns:
            empty = df[COL_OP].isna().sum() + (df[COL_OP] == '').sum()
            if empty > 0: errors.append(f"{label}: {empty} rows with empty Operation numbers")
            
        # Duplicates within file
        if COL_CARD in df.columns and COL_OP in df.columns:
            dups = df.groupby([COL_CARD, COL_OP, ACCOUNTING_REF]).size()
            if (dups > 1).any():
                warnings.append(f"{label}: Found duplicate key combinations within same file")

        return warnings, errors

    # =========================================================================
    # 4. MATCHING & ANALYSIS
    # =========================================================================
    def match_transactions(self):
        print("Matching Transactions...")
        if not self.df_credit.empty:
            self.merged = pd.merge(
                self.df_debt, 
                self.df_credit, 
                on=[COL_CARD, COL_OP], 
                how='inner', 
                suffixes=('_DEBT', '_CREDIT')
            )
        else:
            self.merged = pd.DataFrame()
        
        # Match VFF credits against M2D debts
        self._match_vff_transactions()
        
        if self.merged.empty and self.merged_vff.empty:
            print("No matches found.")
            return False
        return self._check_orphans()

    def _match_vff_transactions(self):
        """Matches VFF computed credit notes against M2D debt files."""
        if self.df_credit_vff.empty:
            print("  No VFF credits to match.")
            return
        
        print("Matching VFF credits against M2D debts...")
        self.merged_vff = pd.merge(
            self.df_debt,
            self.df_credit_vff,
            on=[COL_CARD, COL_OP],
            how='inner',
            suffixes=('_DEBT', '_CREDIT')
        )
        
        if not self.merged_vff.empty:
            print(f"  ✓ Matched {len(self.merged_vff)} VFF credits to M2D debts")
        
        # Unmatched VFF credits → unexpected refunds
        if not self.df_credit_vff.empty:
            vff_keys = set(zip(self.df_credit_vff[COL_CARD], self.df_credit_vff[COL_OP]))
            matched_vff_keys = set(zip(self.merged_vff[COL_CARD], self.merged_vff[COL_OP])) if not self.merged_vff.empty else set()
            unmatched_vff_keys = vff_keys - matched_vff_keys
            
            if unmatched_vff_keys:
                print(f"  ⚠ {len(unmatched_vff_keys)} VFF credits with no M2D match → Unexpected Refunds")
                self.df_credit_vff['temp_key'] = list(zip(self.df_credit_vff[COL_CARD], self.df_credit_vff[COL_OP]))
                unmatched_vff = self.df_credit_vff[self.df_credit_vff['temp_key'].isin(unmatched_vff_keys)].copy()
                unmatched_vff.drop(columns=['temp_key'], inplace=True)
                # Add VFF source marker
                unmatched_vff['VFF_Source'] = 'VFF_UNMATCHED'
                self.unexpected_refunds = pd.concat([self.unexpected_refunds, unmatched_vff], ignore_index=True)
                self.df_credit_vff.drop(columns=['temp_key'], inplace=True)

    def _check_orphans(self):
        print("Analyzing unmatched records...")
        merged_keys = set(zip(self.merged[COL_CARD], self.merged[COL_OP])) if not self.merged.empty else set()
        
        # Also include VFF matched keys
        vff_matched_keys = set(zip(self.merged_vff[COL_CARD], self.merged_vff[COL_OP])) if not self.merged_vff.empty else set()
        all_matched_keys = merged_keys | vff_matched_keys
        
        # Check orphans for regular credits (NON-BLOCKING — alert + M6D SIN MATCH sheet)
        if not self.df_credit.empty:
            credit_keys = set(zip(self.df_credit[COL_CARD], self.df_credit[COL_OP]))
            orphaned_credit_keys = credit_keys - merged_keys
            
            if orphaned_credit_keys:
                print(f"\n{'='*60}")
                print(f"⚠️  M6D SIN MATCH: {len(orphaned_credit_keys)} credits with NO matching debt")
                print(f"{'='*60}")
                
                # Identify and log problematic files
                df_credit_chk = self.df_credit.copy()
                df_credit_chk['temp_key'] = list(zip(df_credit_chk[COL_CARD], df_credit_chk[COL_OP]))
                
                orphans_df = df_credit_chk[df_credit_chk['temp_key'].isin(orphaned_credit_keys)].copy()
                problem_files = orphans_df[ACCOUNTING_REF].unique()
                
                print("\n  📂 M6D FILES WITH UNMATCHED OPERATIONS:")
                for f in problem_files:
                    file_orphan_count = len(orphans_df[orphans_df[ACCOUNTING_REF] == f])
                    print(f"     - {f}: {file_orphan_count} unmatched operation(s)")
                
                # Store for export — include origin file name
                orphans_df.drop(columns=['temp_key'], inplace=True)
                orphans_df.rename(columns={ACCOUNTING_REF: 'Origin_File'}, inplace=True)
                self.m6d_sin_match = orphans_df
                
                print(f"\n  ℹ️  These will be exported to the 'M6D SIN MATCH' sheet.")
                print(f"  Conciliation will continue.\n")
        
        # VFF orphaned credits are NOT blocking — already funneled to unexpected refunds
        
        debt_keys = set(zip(self.df_debt[COL_CARD], self.df_debt[COL_OP]))
        orphaned_debt_keys = debt_keys - all_matched_keys
        if orphaned_debt_keys:
            print(f"\n📊 UNMATCHED DEBTS: {len(orphaned_debt_keys)} (Informational)")
        else:
            print("✓ All credits matched to debts (100% reconciliation).")
            
        return True

    def run_analysis(self):
        self._identify_recuperar_scenarios()
        self._analyze_variance()

    def _analyze_variance(self):
        print("Checking for amount variances...")
        variance_check = self.merged.groupby(
            [f"{self.config['ACCOUNTING_REF']}_CREDIT", self.config['COL_CARD'], self.config['COL_OP'], f"{self.config['AMT_FLOAT']}_CREDIT"]
        ).agg(
            Total_Debts_Covered=(f"{self.config['AMT_FLOAT']}_DEBT", 'sum')
        ).reset_index()
        
        variance_check['Variance'] = variance_check[f"{self.config['AMT_FLOAT']}_CREDIT"] - variance_check['Total_Debts_Covered']
        self.variance_report = variance_check[variance_check['Variance'].abs() > 0.01].copy()
        
        if not self.variance_report.empty:
            print(f"  ⚠ Found {len(self.variance_report)} MATCHES WITH VARIANCE")
            self.variance_report['Status'] = self.variance_report['Variance'].apply(
                lambda x: "OVERPAID (Refund > Debts)" if x > 0 else "UNDERPAID (Refund < Debts)"
            )
            # Collect bad keys to exclude from strict reconciliation
            for _, row in self.variance_report.iterrows():
                self.bad_credit_keys.add((row[f"{self.config['ACCOUNTING_REF']}_CREDIT"], row[self.config['COL_CARD']], row[self.config['COL_OP']]))
                
            self.variance_report.rename(columns={
                f"{self.config['ACCOUNTING_REF']}_CREDIT": 'Credit_File',
                f"{self.config['AMT_FLOAT']}_CREDIT": 'Refund_Amount'
            }, inplace=True)

    def _identify_recuperar_scenarios(self):
        print("Applying 'RECUPERAR' business logic...")
        
        # 1. Pending Claims (RECUPERAR='NO' and not matched)
        self.merged_keys = set(zip(self.merged[self.config['COL_CARD']], self.merged[self.config['COL_OP']]))
        self.df_debt['temp_key'] = list(zip(self.df_debt[self.config['COL_CARD']], self.df_debt[self.config['COL_OP']]))
        
        self.pending_claims = self.df_debt[
            (self.df_debt[self.config['COL_RECUPERAR']] == 'NO') & 
            (~self.df_debt['temp_key'].isin(self.merged_keys))
        ].copy()
        
        if not self.pending_claims.empty:
            print(f"  ⚠ Found {len(self.pending_claims)} PENDING CLAIMS")
            
        # 2. Unexpected Refunds (RECUPERAR!='NO' but matched)
        self.unexpected_refunds = self.merged[self.merged[f"{self.config['COL_RECUPERAR']}_DEBT"] != 'NO'].copy()
        
        if not self.unexpected_refunds.empty:
            print(f"  ℹ Found {len(self.unexpected_refunds)} UNEXPECTED REFUNDS")

    # =========================================================================
    # 5. REPORT GENERATION
    # =========================================================================
    def generate_reports(self):
        self._generate_fully_reconciled_summary()
        self._generate_net_balanced_summary()

    def _generate_fully_reconciled_summary(self):
        print("Generating Fully Reconciled Summary (Strict Mode)...")
        fully_reconciled_files = []
        debt_groups = self.df_debt.groupby(self.config['ACCOUNTING_REF'])
        
        # Optimization: Pre-compute exclusions sets for O(1) lookup
        pending_exclusion_set = set(self.pending_claims[self.config['ACCOUNTING_REF']].unique()) if not self.pending_claims.empty else set()
        unexpected_exclusion_set = set(self.unexpected_refunds[f"{self.config['ACCOUNTING_REF']}_DEBT"].unique()) if not self.unexpected_refunds.empty else set()
        
        for filename, group in debt_groups:
            # Check Exclusions first (fastest check)
            if filename in pending_exclusion_set: continue
            if filename in unexpected_exclusion_set: continue
            
            total_no = group[group[self.config['COL_RECUPERAR']] == 'NO']
            if total_no.empty: continue
            
            # Verify 100% Match
            matched_no = total_no[total_no['temp_key'].isin(self.merged_keys)]
            if len(total_no) != len(matched_no): continue
            
            # Verify Variance
            relevant_merged = self.merged[
                (self.merged[f"{self.config['ACCOUNTING_REF']}_DEBT"] == filename) & 
                (self.merged[f"{self.config['COL_RECUPERAR']}_DEBT"] == 'NO')
            ]
            
            # Set intersection for variance check (faster than loop)
            current_keys = set(zip(relevant_merged[f"{self.config['ACCOUNTING_REF']}_CREDIT"], relevant_merged[self.config['COL_CARD']], relevant_merged[self.config['COL_OP']]))
            if not current_keys.isdisjoint(self.bad_credit_keys):
                continue
            
            # Add to Summary
            creditor_breakdown = relevant_merged.groupby(f"{self.config['ACCOUNTING_REF']}_CREDIT").agg(
                Amount_Covered=(f"{self.config['AMT_FLOAT']}_DEBT", 'sum')
            ).reset_index()
            
            for _, row in creditor_breakdown.iterrows():
                fully_reconciled_files.append({
                    self.config['OUT_DEBTOR_FILE']: filename,
                    self.config['OUT_CREDIT_NOTE']: row[f"{self.config['ACCOUNTING_REF']}_CREDIT"],
                    self.config['OUT_MATCHED_AMT']: row['Amount_Covered']
                })
                
        self.fully_reconciled = pd.DataFrame(fully_reconciled_files)
        if not self.fully_reconciled.empty:
            total_val = self.fully_reconciled[self.config['OUT_MATCHED_AMT']].sum()
            total_row = pd.DataFrame([{
                self.config['OUT_DEBTOR_FILE']: 'TOTAL', 
                self.config['OUT_CREDIT_NOTE']: '', 
                self.config['OUT_MATCHED_AMT']: total_val
            }])
            self.fully_reconciled = pd.concat([self.fully_reconciled, total_row], ignore_index=True)

    def _generate_net_balanced_summary(self):
        print("Checking for Net Balanced files...")
        rows = []
        
        candidates = set(self.pending_claims[self.config['ACCOUNTING_REF']].unique()) | set(self.unexpected_refunds[f"{self.config['ACCOUNTING_REF']}_DEBT"].unique())
        
        if not self.fully_reconciled.empty:
            excluded = set(self.fully_reconciled[self.config['OUT_DEBTOR_FILE']].unique())
            candidates = candidates - excluded
        
        # Optimization: Pre-group to avoid O(N) filtering inside loop
        pending_map = {k: v for k, v in self.pending_claims.groupby(self.config['ACCOUNTING_REF'])} if not self.pending_claims.empty else {}
        unexpected_map = {k: v for k, v in self.unexpected_refunds.groupby(f"{self.config['ACCOUNTING_REF']}_DEBT")} if not self.unexpected_refunds.empty else {}
        
        for filename in candidates:
            if filename == 'TOTAL': continue
            
            file_pending = pending_map.get(filename, pd.DataFrame())
            file_unexpected = unexpected_map.get(filename, pd.DataFrame())
            
            sum_p = file_pending[self.config['COL_AMOUNT']].astype(float).sum() if not file_pending.empty else 0.0
            sum_u = file_unexpected[f"{self.config['AMT_FLOAT']}_CREDIT"].sum() if not file_unexpected.empty else 0.0
            
            if abs(sum_p - sum_u) < 0.01 and (sum_p > 0 or sum_u > 0):
                # IT IS BALANCED
                if not file_pending.empty:
                    for _, r in file_pending.iterrows():
                        rows.append({
                            self.config['OUT_DEBTOR_FILE_COL']: filename, 
                            self.config['OUT_STATUS']: 'NET BALANCED', 
                            self.config['OUT_TYPE']: 'PENDING_CLAIM', 
                            self.config['OUT_MATCHED_AMT']: float(r[self.config['COL_AMOUNT']])
                        })
                if not file_unexpected.empty:
                    for _, r in file_unexpected.iterrows():
                        rows.append({
                            self.config['OUT_DEBTOR_FILE_COL']: filename, 
                            self.config['OUT_STATUS']: 'NET BALANCED', 
                            self.config['OUT_TYPE']: 'UNEXPECTED_REFUND', 
                            self.config['OUT_MATCHED_AMT']: r[f"{self.config['AMT_FLOAT']}_CREDIT"]
                        })
                
                # Context lookup - could be optimized further but acceptable for now
                file_matched = self.merged[
                     (self.merged[f"{self.config['ACCOUNTING_REF']}_DEBT"] == filename) & 
                     (self.merged[f"{self.config['COL_RECUPERAR']}_DEBT"] == 'NO')
                ]
                for _, r in file_matched.iterrows():
                    rows.append({
                        self.config['OUT_DEBTOR_FILE_COL']: filename, 
                        self.config['OUT_STATUS']: 'NET BALANCED', 
                        self.config['OUT_TYPE']: 'CORRECTLY_MATCHED', 
                        self.config['OUT_MATCHED_AMT']: r[f"{self.config['AMT_FLOAT']}_DEBT"]
                    })
                    
        self.net_balanced = pd.DataFrame(rows)

    def export_results(self, output_file):
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Breakdowns
                if not self.merged.empty:
                    self.merged.groupby([f"{self.config['ACCOUNTING_REF']}_DEBT", f"{self.config['ACCOUNTING_REF']}_CREDIT"]).agg(
                        Count=(self.config['COL_OP'], 'count'), Amount=(f"{self.config['AMT_FLOAT']}_DEBT", 'sum')
                    ).reset_index().to_excel(writer, sheet_name='By_Debt_File', index=False)
                    
                    self.merged.groupby([f"{self.config['ACCOUNTING_REF']}_CREDIT", f"{self.config['ACCOUNTING_REF']}_DEBT"]).agg(
                        Count=(self.config['COL_OP'], 'count'), Amount=(f"{self.config['AMT_FLOAT']}_DEBT", 'sum')
                    ).reset_index().to_excel(writer, sheet_name='By_Credit_File', index=False)
                
                if not self.pending_claims.empty: self.pending_claims.to_excel(writer, sheet_name='Pending_Claims', index=False)
                if not self.unexpected_refunds.empty: self.unexpected_refunds.to_excel(writer, sheet_name='Unexpected_Refunds', index=False)
                if not self.fully_reconciled.empty: self.fully_reconciled.to_excel(writer, sheet_name='Fully_Reconciled_Notes', index=False)
                if not self.net_balanced.empty: self.net_balanced.to_excel(writer, sheet_name='Net_Balanced_Files', index=False)
                if not self.variance_report.empty: self.variance_report.to_excel(writer, sheet_name='Amount_Variances', index=False)
                
                # M6D SIN MATCH - orphaned M6D credits with origin filenames
                if not self.m6d_sin_match.empty:
                    self.m6d_sin_match.to_excel(writer, sheet_name='M6D SIN MATCH', index=False)
                
                # VFF Acreedoras sheet
                if not self.vff_acreedoras.empty:
                    self.vff_acreedoras.to_excel(writer, sheet_name='Acreedoras', index=False)
                
                # VFF Debtor Notes (errors - negative differences)
                if not self.vff_debtor_notes.empty:
                    self.vff_debtor_notes.to_excel(writer, sheet_name='VFF_Debtor_Notes', index=False)
                
                # VFF Matched transactions
                if not self.merged_vff.empty:
                    self.merged_vff.to_excel(writer, sheet_name='VFF_Matched', index=False)
                
                # Export separate sheets for Fully Reconciled Files
                if not self.fully_reconciled.empty:
                    self._export_individual_sheets(writer)

                if not self.merged.empty:
                    self.merged.to_excel(writer, sheet_name='Detailed_Audit_Log', index=False)
                
            print(f"SUCCESS. Report saved to: {output_file}")
        except PermissionError:
            print(f"ERROR: Close {output_file} and try again.")
            
    def _export_individual_sheets(self, writer):
        reconciled_debts = self.fully_reconciled[self.config['OUT_DEBTOR_FILE']].unique()
        reconciled_credits = self.fully_reconciled[self.config['OUT_CREDIT_NOTE']].unique()
        
        # Write Debt Files
        for fname in reconciled_debts:
            if fname == 'TOTAL': continue
            if fname in self.debt_files:
                clean_name = fname.replace(":", "").replace("/", "-")[:31]
                self.debt_files[fname].to_excel(writer, sheet_name=clean_name, index=False)
                
        # Write Credit Files
        for fname in reconciled_credits:
            if fname in self.credit_files:
                clean_name = fname.replace(":", "").replace("/", "-")[:31]
                self.credit_files[fname].to_excel(writer, sheet_name=clean_name, index=False)

    def run(self):
        print(f"--- Starting Conciliation in {self.folder_path} ---")
        if not self.load_data(): return
        if not self.validate(): return
        if not self.match_transactions(): return
        self.run_analysis()
        self.generate_reports()
        self.export_results('CONCILIATION_FINAL_REPORT.xlsx')


# =============================================================================
# BACKWARD COMPATIBILITY
# =============================================================================

# This function matches the original function name and signature (if any)
# allowing existing external scripts/tests to call it without breaking.

def robust_conciliation_duplicates_allowed():
    # Also expose internal helper functions for testing if needed, 
    # but primarily this function runs the full flow.
    # Note: Global FOLDER_PATH constant logic is now encapsulated in the Class default
    conciliator = Conciliator(folder_path=DEFAULT_FOLDER_PATH) # Uses the global default constant
    conciliator.run()

# Expose older functions by mapping them to class methods IF necessary for tests.
# Since current tests import specific functions like `check_orphans`, we might 
# need to create wrappers or update tests.
#
# STRATEGY: 
# The implementation below provides functional wrappers that re-use the logic 
# by instantiating a dummy class or just implementing them as standalone 
# helper functions again if we want to support partial testing.
#
# HOWEVER, for cleaner code, we should prefer updating the tests. 
# But to satisfy the "refactor without breaking" constraint as much as possible:

def check_orphans(df_debt, df_credit, merged):
    # Wrapper to support existing unit tests that call this function directly
    c = Conciliator()
    c.df_debt = df_debt
    c.df_credit = df_credit
    c.merged = merged
    return c._check_orphans()

# We can do similar wrappers for others if strictly needed by tests:
# check_intra_pile_duplicates
# validate_files_are_different
# check_data_quality

def check_intra_pile_duplicates(individual_files, label):
    return Conciliator()._check_intra_pile_duplicates(individual_files, label)

def validate_files_are_different(df1, df2):
    return Conciliator()._validate_files_are_different(df1, df2)

def check_data_quality(df, label):
    return Conciliator()._check_data_quality(df, label)


if __name__ == "__main__":
    robust_conciliation_duplicates_allowed()
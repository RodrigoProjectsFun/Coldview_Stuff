import pandas as pd
import os
import glob
import re

# --- HEADERS CONFIGURATION ---
COL_CARD = 'Card'               
COL_OP = 'Operation Number'
COL_AMOUNT = 'Original Amount' 
COL_RECUPERAR = 'RECUPERAR'
AMT_FLOAT = 'Amt_Float'
ACCOUNTING_REF = 'Accounting_Ref'

DEBT_PATTERN = '*m2d-recu*.xlsx'
CREDIT_PATTERN = '*m6d-dev*.xlsx'
FOLDER_PATH = './accounting_files'

# =============================================================================
# 1. HELPER FUNCTIONS
# =============================================================================

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

# =============================================================================
# 2. DATA LOADING & CLEANING
# =============================================================================

def load_pile(pattern, label):
    """
    Loads and cleans files matching the pattern.
    Returns: (combined_df, individual_files_dict)
    """
    files = glob.glob(os.path.join(FOLDER_PATH, pattern))
    
    # Double check filter (glob can be broad)
    filter_keyword = 'm2d-recu' if label == "DEBT" else 'm6d-dev'
    files = [f for f in files if filter_keyword in os.path.basename(f).lower()]
    
    all_dfs = []
    individual_files = {}  # Track individual files for duplicate detection
    print(f"Loading {len(files)} files for {label}...")

    for f in files:
        try:
            # Load as String to protect IDs from scientific notation
            df = pd.read_excel(f, dtype=str)
            
            # Drop empty rows (trailing rows Excel includes beyond actual data)
            if COL_CARD in df.columns and COL_OP in df.columns:
                # Replace empty strings/whitespace with NaN for proper dropna
                df[COL_CARD] = df[COL_CARD].replace(r'^\s*$', pd.NA, regex=True)
                df[COL_OP] = df[COL_OP].replace(r'^\s*$', pd.NA, regex=True)
                
                # Drop rows where BOTH key columns are empty (trailing rows)
                rows_before = len(df)
                df = df.dropna(subset=[COL_CARD, COL_OP], how='all')
                rows_dropped = rows_before - len(df)
                if rows_dropped > 0:
                    print(f"  [INFO] {os.path.basename(f)}: Dropped {rows_dropped} empty trailing rows")
            
            # Create Standardized Reference Name
            std_name = get_standardized_name(f)
            df[ACCOUNTING_REF] = std_name
            
            # Clean Keys
            if COL_CARD in df.columns and COL_OP in df.columns:
                df[COL_CARD] = df[COL_CARD].str.strip()
                df[COL_OP] = df[COL_OP].str.strip()
            else:
                print(f"  [SKIP] {std_name} missing Card or Operation headers.")
                continue
            
            # Clean Amount (Force to Float)
            if COL_AMOUNT in df.columns:
                clean_amt = df[COL_AMOUNT].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                df[AMT_FLOAT] = pd.to_numeric(clean_amt, errors='coerce').fillna(0.0)
            
            # Clean RECUPERAR (Default to 'SI' if missing, standardize to uppercase)
            if COL_RECUPERAR in df.columns:
                df[COL_RECUPERAR] = df[COL_RECUPERAR].astype(str).str.strip().str.upper()
            else:
                # Default to 'SI' if column missing (Assume valid charge)
                df[COL_RECUPERAR] = 'SI'

            # Store result
            all_dfs.append(df)
            individual_files[std_name] = df.copy()
            
        except Exception as e:
            print(f"  [ERROR] {os.path.basename(f)}: {e}")
    
    combined = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    return combined, individual_files

def load_all_data():
    df_debt, debt_files = load_pile(DEBT_PATTERN, "DEBT")
    df_credit, credit_files = load_pile(CREDIT_PATTERN, "CREDIT")
    return df_debt, df_credit, debt_files, credit_files

# =============================================================================
# 3. QUALITY CHECKS & VALIDATION
# =============================================================================

def check_intra_pile_duplicates(individual_files, label):
    """
    Check if any files within the same pile are duplicates of each other.
    """
    issues = []
    file_names = list(individual_files.keys())
    
    if len(file_names) < 2:
        return issues
    
    for i in range(len(file_names)):
        for j in range(i + 1, len(file_names)):
            name1, name2 = file_names[i], file_names[j]
            df1, df2 = individual_files[name1], individual_files[name2]
            
            if len(df1) != len(df2): continue
            
            keys1 = set(zip(df1[COL_CARD], df1[COL_OP]))
            keys2 = set(zip(df2[COL_CARD], df2[COL_OP]))
            
            if keys1 == keys2:
                compare_cols = [col for col in df1.columns if col not in [ACCOUNTING_REF]]
                df1_sorted = df1[compare_cols].sort_values(by=[COL_CARD, COL_OP]).reset_index(drop=True)
                df2_sorted = df2[compare_cols].sort_values(by=[COL_CARD, COL_OP]).reset_index(drop=True)
                
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

def validate_files_are_different(df1, df2):
    """
    Ensure files aren't duplicates (e.g. uploading Credit file as Debt).
    """
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

def check_data_quality(df, label):
    """
    Data quality checks for standard errors.
    """
    warnings = []
    errors = []
    
    if AMT_FLOAT in df.columns:
        if (df[AMT_FLOAT] < 0).any(): 
            warnings.append(f"{label}: Found negative amounts")
        if (df[AMT_FLOAT] == 0).any(): 
            warnings.append(f"{label}: Found zero-amount transactions")
        
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

def perform_validations(df_debt, df_credit, debt_files, credit_files):
    """
    Runs all validation logic. Returns True if valid, False if critical errors found.
    """
    print("Checking for duplicate files within each category...")
    intra_issues = check_intra_pile_duplicates(debt_files, "DEBT") + \
                   check_intra_pile_duplicates(credit_files, "CREDIT")
    if intra_issues:
        print("\n" + "="*60 + "\n⚠️  INTRA-CATEGORY DUPLICATE DETECTION ⚠️\n" + "="*60)
        for i in intra_issues: print(f"  ❌ {i}")
        print("\nConciliation ABORTED.\n")
        return False

    print("Validating files are not duplicates...")
    dups = validate_files_are_different(df_debt, df_credit)
    if dups:
        print("\n" + "="*60 + "\n⚠️  DUPLICATE FILE DETECTION ⚠️\n" + "="*60)
        for i in dups: print(f"  ❌ {i}")
        print("\nConciliation ABORTED.\n")
        return False

    print("Running data quality checks...")
    all_warnings, all_errors = [], []
    for f, df in debt_files.items():
        w, e = check_data_quality(df, f"DEBT ({f})")
        all_warnings.extend(w); all_errors.extend(e)
    for f, df in credit_files.items():
        w, e = check_data_quality(df, f"CREDIT ({f})")
        all_warnings.extend(w); all_errors.extend(e)

    if all_warnings:
        print("\n" + "-"*60 + "\n⚠️  DATA QUALITY WARNINGS\n" + "-"*60)
        for w in all_warnings: print(f"  ⚠ {w}")
    
    if all_errors:
        print("\n" + "="*60 + "\n❌  DATA QUALITY ERRORS\n" + "="*60)
        for e in all_errors: print(f"  ❌ {e}")
        print("\nConciliation ABORTED.\n")
        return False

    return True

# =============================================================================
# 4. MATCHING & ANALYSIS
# =============================================================================

def perform_matching(df_debt, df_credit):
    print("Matching Transactions...")
    merged = pd.merge(
        df_debt, 
        df_credit, 
        on=[COL_CARD, COL_OP], 
        how='inner', 
        suffixes=('_DEBT', '_CREDIT')
    )
    return merged

def check_orphans(df_debt, df_credit, merged):
    print("Analyzing unmatched records...")
    merged_keys = set(zip(merged[COL_CARD], merged[COL_OP]))
    credit_keys = set(zip(df_credit[COL_CARD], df_credit[COL_OP]))
    orphaned_credit_keys = credit_keys - merged_keys
    
    if orphaned_credit_keys:
        print("\n" + "="*60 + "\n❌  CRITICAL ERROR: ORPHANED CREDITS DETECTED\n" + "="*60)
        print(f"  Found {len(orphaned_credit_keys)} credits with NO matching debt!")
        print("  Every credit MUST have a corresponding debt.")
        print("\nConciliation ABORTED.")
        return False
    
    debt_keys = set(zip(df_debt[COL_CARD], df_debt[COL_OP]))
    orphaned_debt_keys = debt_keys - merged_keys
    if orphaned_debt_keys:
        print(f"\n📊 UNMATCHED DEBTS: {len(orphaned_debt_keys)} (Informational)")
    else:
        print("✓ All credits matched to debts (100% reconciliation).")
        
    return True

def analyze_variance(merged):
    print("Checking for amount variances...")
    variance_check = merged.groupby(
        [f'{ACCOUNTING_REF}_CREDIT', COL_CARD, COL_OP, f'{AMT_FLOAT}_CREDIT']
    ).agg(
        Total_Debts_Covered=(f'{AMT_FLOAT}_DEBT', 'sum')
    ).reset_index()
    
    variance_check['Variance'] = variance_check[f'{AMT_FLOAT}_CREDIT'] - variance_check['Total_Debts_Covered']
    variance_report = variance_check[variance_check['Variance'].abs() > 0.01].copy()
    
    bad_credit_keys = set()
    
    if not variance_report.empty:
        print(f"  ⚠ Found {len(variance_report)} MATCHES WITH VARIANCE")
        variance_report['Status'] = variance_report['Variance'].apply(
            lambda x: "OVERPAID (Refund > Debts)" if x > 0 else "UNDERPAID (Refund < Debts)"
        )
        # Collect bad keys to exclude from strict reconciliation
        for _, row in variance_report.iterrows():
            bad_credit_keys.add((row[f'{ACCOUNTING_REF}_CREDIT'], row[COL_CARD], row[COL_OP]))
            
        variance_report.rename(columns={
            f'{ACCOUNTING_REF}_CREDIT': 'Credit_File',
            f'{AMT_FLOAT}_CREDIT': 'Refund_Amount'
        }, inplace=True)
        
    return variance_report, bad_credit_keys

def identify_recuperar_scenarios(df_debt, merged):
    print("Applying 'RECUPERAR' business logic...")
    
    # 1. Pending Claims (RECUPERAR='NO' and not matched)
    merged_keys = set(zip(merged[COL_CARD], merged[COL_OP]))
    df_debt['temp_key'] = list(zip(df_debt[COL_CARD], df_debt[COL_OP]))
    
    pending_claims = df_debt[
        (df_debt[COL_RECUPERAR] == 'NO') & 
        (~df_debt['temp_key'].isin(merged_keys))
    ].copy()
    
    if not pending_claims.empty:
        print(f"  ⚠ Found {len(pending_claims)} PENDING CLAIMS")
        
    # 2. Unexpected Refunds (RECUPERAR!='NO' but matched)
    unexpected_refunds = merged[merged[f'{COL_RECUPERAR}_DEBT'] != 'NO'].copy()
    
    if not unexpected_refunds.empty:
        print(f"  ℹ Found {len(unexpected_refunds)} UNEXPECTED REFUNDS")
        
    return pending_claims, unexpected_refunds, merged_keys

# =============================================================================
# 5. REPORT GENERATION
# =============================================================================

def generate_fully_reconciled_summary(df_debt, merged, pending_claims, unexpected_refunds, bad_credit_keys, merged_keys):
    print("Generating Fully Reconciled Summary (Strict Mode)...")
    fully_reconciled_files = []
    
    debt_groups = df_debt.groupby(ACCOUNTING_REF)
    
    for filename, group in debt_groups:
        total_no = group[group[COL_RECUPERAR] == 'NO']
        if total_no.empty: continue
        
        # Check Exclusions
        if filename in pending_claims[ACCOUNTING_REF].values: continue
        if filename in unexpected_refunds[f'{ACCOUNTING_REF}_DEBT'].values: continue
        
        # Verify 100% Match
        matched_no = total_no[total_no['temp_key'].isin(merged_keys)]
        if len(total_no) != len(matched_no): continue
        
        # Verify Variance
        relevant_merged = merged[
            (merged[f'{ACCOUNTING_REF}_DEBT'] == filename) & 
            (merged[f'{COL_RECUPERAR}_DEBT'] == 'NO')
        ]
        
        has_variance = False
        for _, row in relevant_merged.iterrows():
            key = (row[f'{ACCOUNTING_REF}_CREDIT'], row[COL_CARD], row[COL_OP])
            if key in bad_credit_keys:
                has_variance = True; break
        if has_variance: continue
        
        # Add to Summary
        creditor_breakdown = relevant_merged.groupby(f'{ACCOUNTING_REF}_CREDIT').agg(
            Amount_Covered=(f'{AMT_FLOAT}_DEBT', 'sum')
        ).reset_index()
        
        for _, row in creditor_breakdown.iterrows():
            fully_reconciled_files.append({
                'DEBTOR FILE': filename,
                'CREDIT FILE NOTE': row[f'{ACCOUNTING_REF}_CREDIT'],
                'AMOUNT THAT MATCHED': row['Amount_Covered']
            })
            
    df_full = pd.DataFrame(fully_reconciled_files)
    if not df_full.empty:
         # Add Total Row
        total_val = df_full['AMOUNT THAT MATCHED'].sum()
        total_row = pd.DataFrame([{
            'DEBTOR FILE': 'TOTAL', 
            'CREDIT FILE NOTE': '', 
            'AMOUNT THAT MATCHED': total_val
        }])
        df_full = pd.concat([df_full, total_row], ignore_index=True)
        
    return df_full

def generate_net_balanced_summary(pending_claims, unexpected_refunds, merged, df_fully_reconciled):
    print("Checking for Net Balanced files...")
    rows = []
    
    candidates = set(pending_claims[ACCOUNTING_REF].unique()) | set(unexpected_refunds[f'{ACCOUNTING_REF}_DEBT'].unique())
    
    # Exclude strictly reconciled
    if not df_fully_reconciled.empty:
        excluded = set(df_fully_reconciled['DEBTOR FILE'].unique())
        candidates = candidates - excluded
        
    for filename in candidates:
        if filename == 'TOTAL': continue
        
        file_pending = pending_claims[pending_claims[ACCOUNTING_REF] == filename]
        file_unexpected = unexpected_refunds[unexpected_refunds[f'{ACCOUNTING_REF}_DEBT'] == filename]
        
        sum_p = file_pending[COL_AMOUNT].astype(float).sum()
        sum_u = file_unexpected[f'{AMT_FLOAT}_CREDIT'].sum()
        
        if abs(sum_p - sum_u) < 0.01 and (sum_p > 0 or sum_u > 0):
            # IT IS BALANCED
            # Add Pending
            for _, r in file_pending.iterrows():
                rows.append({'Debtor_File': filename, 'Status': 'NET BALANCED', 'Type': 'PENDING_CLAIM', 'Amount': float(r[COL_AMOUNT])})
            # Add Unexpected
            for _, r in file_unexpected.iterrows():
                rows.append({'Debtor_File': filename, 'Status': 'NET BALANCED', 'Type': 'UNEXPECTED_REFUND', 'Amount': r[f'{AMT_FLOAT}_CREDIT']})
            # Add Context
            file_matched = merged[
                 (merged[f'{ACCOUNTING_REF}_DEBT'] == filename) & 
                 (merged[f'{COL_RECUPERAR}_DEBT'] == 'NO')
            ]
            for _, r in file_matched.iterrows():
                rows.append({'Debtor_File': filename, 'Status': 'NET BALANCED', 'Type': 'CORRECTLY_MATCHED', 'Amount': r[f'{AMT_FLOAT}_DEBT']})
                
    return pd.DataFrame(rows)

def export_results(merged, pending, unexpected, variance, full, net_balanced, output_file):
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Breakdowns
            merged.groupby([f'{ACCOUNTING_REF}_DEBT', f'{ACCOUNTING_REF}_CREDIT']).agg(
                Count=(COL_OP, 'count'), Amount=(f'{AMT_FLOAT}_DEBT', 'sum')
            ).reset_index().to_excel(writer, sheet_name='By_Debt_File', index=False)
            
            merged.groupby([f'{ACCOUNTING_REF}_CREDIT', f'{ACCOUNTING_REF}_DEBT']).agg(
                Count=(COL_OP, 'count'), Amount=(f'{AMT_FLOAT}_DEBT', 'sum')
            ).reset_index().to_excel(writer, sheet_name='By_Credit_File', index=False)
            
            if not pending.empty: pending.to_excel(writer, sheet_name='Pending_Claims', index=False)
            if not unexpected.empty: unexpected.to_excel(writer, sheet_name='Unexpected_Refunds', index=False)
            if not full.empty: full.to_excel(writer, sheet_name='Fully_Reconciled_Notes', index=False)
            if not net_balanced.empty: net_balanced.to_excel(writer, sheet_name='Net_Balanced_Files', index=False)
            if not variance.empty: variance.to_excel(writer, sheet_name='Amount_Variances', index=False)
            
            merged.to_excel(writer, sheet_name='Detailed_Audit_Log', index=False)
            
        print(f"SUCCESS. Report saved to: {output_file}")
    except PermissionError:
        print(f"ERROR: Close {output_file} and try again.")

# =============================================================================
# 6. MAIN ORCHESTRATOR
# =============================================================================

def robust_conciliation_duplicates_allowed():
    print(f"--- Starting Conciliation in {FOLDER_PATH} ---")
    
    # 1. Load
    df_debt, df_credit, debt_files, credit_files = load_all_data()
    if df_debt.empty or df_credit.empty:
        print("Stopping: Missing data.")
        return

    # 2. Validate
    if not perform_validations(df_debt, df_credit, debt_files, credit_files):
        return

    # 3. Match
    merged = perform_matching(df_debt, df_credit)
    if merged.empty:
        print("No matches found.")
        return
        
    if not check_orphans(df_debt, df_credit, merged):
        return

    # 4. Analyze Logic
    pending_claims, unexpected_refunds, merged_keys = identify_recuperar_scenarios(df_debt, merged)
    variance_report, bad_credit_keys = analyze_variance(merged)
    
    # 5. Generate Summaries
    df_full = generate_fully_reconciled_summary(df_debt, merged, pending_claims, unexpected_refunds, bad_credit_keys, merged_keys)
    df_net = generate_net_balanced_summary(pending_claims, unexpected_refunds, merged, df_full)
    
    # 6. Export
    export_results(merged, pending_claims, unexpected_refunds, variance_report, df_full, df_net, 'CONCILIATION_FINAL_REPORT.xlsx')

if __name__ == "__main__":
    robust_conciliation_duplicates_allowed()
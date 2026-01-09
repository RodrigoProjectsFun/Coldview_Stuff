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
    col_op = 'Operation Number'
    col_amount = 'Original Amount' 
    col_recuperar = 'RECUPERAR'  # New Column
    
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
        individual_files = {}  # Track individual files for duplicate detection
        print(f"Loading {len(files)} files for {label}...")

        for f in files:
            try:
                # Load as String to protect IDs from scientific notation
                df = pd.read_excel(f, dtype=str)
                
                # Drop empty rows (trailing rows Excel includes beyond actual data)
                # A valid row MUST have both Card and Operation Number
                if col_card in df.columns and col_op in df.columns:
                    # Replace empty strings/whitespace with NaN for proper dropna
                    df[col_card] = df[col_card].replace(r'^\s*$', pd.NA, regex=True)
                    df[col_op] = df[col_op].replace(r'^\s*$', pd.NA, regex=True)
                    
                    # Drop rows where BOTH key columns are empty (these are trailing rows)
                    rows_before = len(df)
                    df = df.dropna(subset=[col_card, col_op], how='all')
                    rows_dropped = rows_before - len(df)
                    if rows_dropped > 0:
                        print(f"  [INFO] {os.path.basename(f)}: Dropped {rows_dropped} empty trailing rows")
                
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
                    df['Amt_Float'] = pd.to_numeric(clean_amt, errors='coerce').fillna(0.0)
                
                # Clean RECUPERAR (Default to 'SI' if missing, standardize to uppercase)
                if col_recuperar in df.columns:
                    df[col_recuperar] = df[col_recuperar].astype(str).str.strip().str.upper()
                else:
                    # Default to 'SI' if column missing (Assume valid charge)
                    df[col_recuperar] = 'SI'

                # Store result
                all_dfs.append(df)
                individual_files[std_name] = df.copy()
                
            except Exception as e:
                print(f"  [ERROR] {os.path.basename(f)}: {e}")
        
        combined = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
        return combined, individual_files

    # --- INTRA-PILE DUPLICATE DETECTION ---
    def check_intra_pile_duplicates(individual_files, label):
        """
        Check if any files within the same pile are duplicates of each other.
        Returns list of issues found.
        """
        issues = []
        file_names = list(individual_files.keys())
        
        if len(file_names) < 2:
            return issues  # Need at least 2 files to compare
        
        # Compare each pair of files
        for i in range(len(file_names)):
            for j in range(i + 1, len(file_names)):
                name1, name2 = file_names[i], file_names[j]
                df1, df2 = individual_files[name1], individual_files[name2]
                
                # Skip if different row counts (quick filter)
                if len(df1) != len(df2):
                    continue
                
                # Check 1: Key set comparison
                keys1 = set(zip(df1[col_card], df1[col_op]))
                keys2 = set(zip(df2[col_card], df2[col_op]))
                
                if keys1 == keys2:
                    # Keys are identical - high suspicion
                    # Check 2: Full data comparison (excluding metadata)
                    compare_cols = [col for col in df1.columns if col not in ['Accounting_Ref']]
                    df1_sorted = df1[compare_cols].sort_values(by=[col_card, col_op]).reset_index(drop=True)
                    df2_sorted = df2[compare_cols].sort_values(by=[col_card, col_op]).reset_index(drop=True)
                    
                    if df1_sorted.equals(df2_sorted):
                        issues.append(
                            f"DUPLICATE {label} FILES: '{name1}' and '{name2}' contain IDENTICAL data!"
                        )
                    else:
                        # Same keys but different amounts - still suspicious
                        issues.append(
                            f"SUSPICIOUS {label} FILES: '{name1}' and '{name2}' have identical operations but different amounts!"
                        )
                else:
                    # Check overlap percentage
                    overlap = keys1 & keys2
                    overlap_pct = len(overlap) / max(len(keys1), 1) * 100
                    
                    if overlap_pct > 90:
                        issues.append(
                            f"WARNING {label}: '{name1}' and '{name2}' share {overlap_pct:.1f}% of operations!"
                        )
        
        return issues

    # Load Data
    df_debt, debt_files = load_pile(debt_pattern, "DEBT")
    df_credit, credit_files = load_pile(credit_pattern, "CREDIT")

    # Check for duplicates within each pile
    print("Checking for duplicate files within each category...")
    intra_issues = []
    intra_issues.extend(check_intra_pile_duplicates(debt_files, "DEBT"))
    intra_issues.extend(check_intra_pile_duplicates(credit_files, "CREDIT"))
    
    if intra_issues:
        print("\n" + "="*60)
        print("⚠️  INTRA-CATEGORY DUPLICATE DETECTION ⚠️")
        print("="*60)
        for issue in intra_issues:
            print(f"  ❌ {issue}")
        print("="*60)
        print("\nSame file may have been uploaded multiple times with different names.")
        print("Conciliation ABORTED to prevent incorrect results.\n")
        return
    
    print("✓ No intra-category duplicates found.")


    if df_debt.empty or df_credit.empty:
        print("Stopping: Missing data.")
        return

    # --- CRITICAL VALIDATION: Detect Duplicate Files (Human Error Prevention) ---
    print("Validating files are not duplicates...")
    
    def validate_files_are_different(df1, df2, label1="DEBT", label2="CREDIT"):
        """
        Comprehensive check to ensure files aren't accidentally the same.
        Returns True if files are different (valid), False if duplicates detected.
        """
        issues = []
        
        # Check 1: Exact DataFrame Match (excluding metadata columns)
        compare_cols = [col for col in df1.columns if col not in ['Accounting_Ref', 'Amt_Float']]
        compare_cols2 = [col for col in df2.columns if col not in ['Accounting_Ref', 'Amt_Float']]
        
        if set(compare_cols) == set(compare_cols2):
            df1_compare = df1[compare_cols].reset_index(drop=True)
            df2_compare = df2[compare_cols2].reset_index(drop=True)
            
            if len(df1_compare) == len(df2_compare):
                # Sort both for comparison
                df1_sorted = df1_compare.sort_values(by=compare_cols).reset_index(drop=True)
                df2_sorted = df2_compare.sort_values(by=compare_cols2).reset_index(drop=True)
                
                if df1_sorted.equals(df2_sorted):
                    issues.append(f"EXACT MATCH: {label1} and {label2} contain identical data!")
        
        # Check 2: Row count + key overlap analysis
        debt_keys = set(zip(df1[col_card], df1[col_op]))
        credit_keys = set(zip(df2[col_card], df2[col_op]))
        
        overlap = debt_keys & credit_keys
        overlap_pct = len(overlap) / max(len(debt_keys), 1) * 100
        
        if overlap_pct > 95 and len(debt_keys) == len(credit_keys):
            issues.append(f"SUSPICIOUS: {overlap_pct:.1f}% key overlap with same row count!")
        
        # Check 3: Amount distribution fingerprint
        if 'Amt_Float' in df1.columns and 'Amt_Float' in df2.columns:
            debt_sum = df1['Amt_Float'].sum()
            credit_sum = df2['Amt_Float'].sum()
            debt_mean = df1['Amt_Float'].mean()
            credit_mean = df2['Amt_Float'].mean()
            
            if (abs(debt_sum - credit_sum) < 0.01 and 
                abs(debt_mean - credit_mean) < 0.01 and
                len(df1) == len(df2)):
                issues.append(f"SUSPICIOUS: Identical sum ({debt_sum:.2f}), mean ({debt_mean:.2f}), and row count!")
        
        # Check 4: Source file reference check
        debt_sources = set(df1['Accounting_Ref'].unique())
        credit_sources = set(df2['Accounting_Ref'].unique())
        
        # Normalize for comparison (remove date, compare base type)
        debt_types = {s.split()[0] for s in debt_sources}  # e.g., "M2D-RECU"
        credit_types = {s.split()[0] for s in credit_sources}  # e.g., "M6D-DEV"
        
        if debt_types == credit_types:
            issues.append(f"WARNING: Both sources are type '{debt_types}' - expected different types!")
        
        return issues
    
    validation_issues = validate_files_are_different(df_debt, df_credit)
    
    if validation_issues:
        print("\n" + "="*60)
        print("⚠️  DUPLICATE FILE DETECTION - HUMAN ERROR LIKELY ⚠️")
        print("="*60)
        for issue in validation_issues:
            print(f"  ❌ {issue}")
        print("="*60)
        print("\nPlease verify you haven't uploaded the same file twice.")
        print("Conciliation ABORTED to prevent incorrect results.\n")
        return
    
    print("✓ Files validated as distinct.")

    # --- DATA QUALITY VALIDATIONS ---
    print("Running data quality checks...")
    
    def check_data_quality(df, label):
        """
        Comprehensive data quality checks for financial data.
        Returns (warnings, errors) - errors are critical, warnings are informational.
        """
        warnings = []
        errors = []
        
        # Check 1: Negative Amounts
        if 'Amt_Float' in df.columns:
            negative_count = (df['Amt_Float'] < 0).sum()
            if negative_count > 0:
                warnings.append(f"{label}: Found {negative_count} negative amounts (might be legitimate refunds)")
        
        # Check 2: Zero Amounts
        if 'Amt_Float' in df.columns:
            zero_count = (df['Amt_Float'] == 0).sum()
            if zero_count > 0:
                warnings.append(f"{label}: Found {zero_count} zero-amount transactions")
        
        # Check 3: Very Large Amounts (Statistical Outliers - >3 std from mean)
        if 'Amt_Float' in df.columns and len(df) > 10:
            mean_amt = df['Amt_Float'].mean()
            std_amt = df['Amt_Float'].std()
            if std_amt > 0:
                outlier_threshold = mean_amt + (3 * std_amt)
                outliers = df[df['Amt_Float'] > outlier_threshold]
                if len(outliers) > 0:
                    max_outlier = outliers['Amt_Float'].max()
                    warnings.append(f"{label}: Found {len(outliers)} unusually large amounts (max: {max_outlier:,.2f})")
        
        # Check 4: Missing/Empty Card Numbers
        if col_card in df.columns:
            empty_cards = df[col_card].isna().sum() + (df[col_card] == '').sum()
            if empty_cards > 0:
                errors.append(f"{label}: {empty_cards} rows have empty/null Card numbers!")
        
        # Check 5: Missing/Empty Operation Numbers
        if col_op in df.columns:
            empty_ops = df[col_op].isna().sum() + (df[col_op] == '').sum()
            if empty_ops > 0:
                errors.append(f"{label}: {empty_ops} rows have empty/null Operation numbers!")
        
        # Check 6: Duplicate Rows Within Single DataFrame
        if col_card in df.columns and col_op in df.columns:
            # Group by card+op and check if any combination appears more than expected
            dup_check = df.groupby([col_card, col_op, 'Accounting_Ref']).size()
            internal_dups = dup_check[dup_check > 1]
            if len(internal_dups) > 0:
                warnings.append(f"{label}: Found {len(internal_dups)} duplicate key combinations within same source file")
        
        # Check 7: Whitespace-only values
        if col_card in df.columns:
            whitespace_cards = (df[col_card].str.strip() == '').sum()
            if whitespace_cards > 0:
                errors.append(f"{label}: {whitespace_cards} Card numbers contain only whitespace!")
        
        return warnings, errors
    
    # Run quality checks on individual files to pinpoint errors
    all_warnings = []
    all_errors = []

    print(f"  Checking {len(debt_files)} DEBT files...")
    for filename, df_single in debt_files.items():
        w, e = check_data_quality(df_single, f"DEBT ({filename})")
        all_warnings.extend(w)
        all_errors.extend(e)

    print(f"  Checking {len(credit_files)} CREDIT files...")
    for filename, df_single in credit_files.items():
        w, e = check_data_quality(df_single, f"CREDIT ({filename})")
        all_warnings.extend(w)
        all_errors.extend(e)

    
    # Print warnings (non-blocking)
    if all_warnings:
        print("\n" + "-"*60)
        print("⚠️  DATA QUALITY WARNINGS (Review Recommended)")
        print("-"*60)
        for warning in all_warnings:
            print(f"  ⚠ {warning}")
        print("-"*60 + "\n")
    
    # Print errors (blocking)
    if all_errors:
        print("\n" + "="*60)
        print("❌  DATA QUALITY ERRORS (Critical)")
        print("="*60)
        for error in all_errors:
            print(f"  ❌ {error}")
        print("="*60)
        print("\nConciliation ABORTED due to data quality issues.\n")
        return
    
    print("✓ Data quality checks passed.")

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

    # --- ORPHANED RECORDS ANALYSIS ---
    # BUSINESS RULE: All credits MUST match debts (can't have refund without original charge)
    # But debts without credits are okay (not all charges have been refunded yet)
    print("Analyzing unmatched records...")
    
    # Find orphaned debts (debts with no matching credit) - INFORMATIONAL ONLY
    merged_debt_keys = set(zip(merged[col_card], merged[col_op]))
    all_debt_keys = set(zip(df_debt[col_card], df_debt[col_op]))
    orphaned_debt_keys = all_debt_keys - merged_debt_keys
    
    # Find orphaned credits (credits with no matching debt) - CRITICAL ERROR
    all_credit_keys = set(zip(df_credit[col_card], df_credit[col_op]))
    orphaned_credit_keys = all_credit_keys - merged_debt_keys
    
    # Calculate orphaned amounts
    orphaned_debts = df_debt[df_debt.apply(lambda x: (x[col_card], x[col_op]) in orphaned_debt_keys, axis=1)]
    orphaned_credits = df_credit[df_credit.apply(lambda x: (x[col_card], x[col_op]) in orphaned_credit_keys, axis=1)]
    
    # CRITICAL: Check for orphaned credits FIRST (blocking error)
    if len(orphaned_credit_keys) > 0:
        orphaned_credit_total = orphaned_credits['Amt_Float'].sum() if 'Amt_Float' in orphaned_credits.columns else 0
        
        print("\n" + "="*60)
        print("❌  CRITICAL ERROR: ORPHANED CREDITS DETECTED")
        print("="*60)
        print(f"  Found {len(orphaned_credit_keys):,} credit(s) with NO matching debt!")
        print(f"  Total orphaned credit amount: ${orphaned_credit_total:,.2f}")
        print("")
        print("  BUSINESS RULE VIOLATION:")
        print("  Every credit (refund) MUST have a corresponding debt (original charge).")
        print("  Credits without matching debts indicate data integrity issues.")
        print("")
        print("  Sample orphaned credits (first 5):")
        for i, (card, op) in enumerate(list(orphaned_credit_keys)[:5]):
            amt = orphaned_credits[(orphaned_credits[col_card] == card) & 
                                   (orphaned_credits[col_op] == op)]['Amt_Float'].iloc[0] if 'Amt_Float' in orphaned_credits.columns else 'N/A'
            print(f"    {i+1}. Card: {card}, Op: {op}, Amount: ${amt:,.2f}" if isinstance(amt, float) else f"    {i+1}. Card: {card}, Op: {op}")
        if len(orphaned_credit_keys) > 5:
            print(f"    ... and {len(orphaned_credit_keys) - 5} more")
        print("="*60)
        print("\nConciliation ABORTED. Please verify credit file data.\n")
        return
    
    # Informational: Report orphaned debts (non-blocking)
    if len(orphaned_debt_keys) > 0:
        orphaned_debt_total = orphaned_debts['Amt_Float'].sum() if 'Amt_Float' in orphaned_debts.columns else 0
        match_rate_debt = (len(all_debt_keys) - len(orphaned_debt_keys)) / max(len(all_debt_keys), 1) * 100
        
        print("\n" + "-"*60)
        print("📊 UNMATCHED DEBTS (Informational - Not Yet Refunded)")
        print("-"*60)
        print(f"  • Unmatched DEBT operations: {len(orphaned_debt_keys):,} (Total: ${orphaned_debt_total:,.2f})")
        print(f"  • DEBT match rate: {match_rate_debt:.1f}%")
        print("-"*60)
        print("  ℹ️  These debts have no matching credits yet (normal if not refunded).\n")
    else:
        print("✓ All credits matched to debts. All debts have corresponding credits (100% reconciliation).")


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

    # --- 4. LOGIC: RECUPERAR (New Logic) ---
    print("Applying 'RECUPERAR' business logic...")
    
    # 4a. Pending Claims (RECUPERAR = 'NO' AND Unmatched)
    # Filter DEBT keys that are NOT in matched
    merged_keys = set(zip(merged[col_card], merged[col_op]))
    
    # helper for keys
    df_debt['temp_key'] = list(zip(df_debt[col_card], df_debt[col_op]))
    
    # Filter: Has NO, Key NOT in merged
    pending_claims = df_debt[
        (df_debt[col_recuperar] == 'NO') & 
        (~df_debt['temp_key'].isin(merged_keys))
    ].copy()
    
    if not pending_claims.empty:
        print(f"  ⚠ Found {len(pending_claims)} PENDING CLAIMS (Debtor Notes without Refunds)")
    
    # 4b. Unexpected Refunds (RECUPERAR != 'NO' AND Matched)
    # Filter MERGED: RECUPERAR_DEBT != 'NO'
    unexpected_refunds = merged[
        merged[f'{col_recuperar}_DEBT'] != 'NO'
    ].copy()
    
    if not unexpected_refunds.empty:
        print(f"  ℹ Found {len(unexpected_refunds)} UNEXPECTED REFUNDS (Standard charges that were refunded)")

    # Clean up temp key (Moved to end)
    # df_debt.drop(columns=['temp_key'], inplace=True, errors='ignore')

    # --- 5. VARIANCE ANALYSIS (Smart Amount Check) ---
    print("Checking for amount variances (Many-to-One Validation)...")
    
    # Track credits with variance to exclude them from "Fully Reconciled" status
    bad_credit_keys = set() 
    
    # Group by Unique Credit Operation
    # Key: Credit Ref, Card, Op Number, Credit Amount
    # Sum: Debt Amount
    variance_check = merged.groupby(
        ['Accounting_Ref_CREDIT', col_card, col_op, 'Amt_Float_CREDIT']
    ).agg(
        Total_Debts_Covered=('Amt_Float_DEBT', 'sum'),
        Count_Debts=('Amt_Float_DEBT', 'count')
    ).reset_index()
    
    # Calculate Variance: Credit - Debt
    # Positive Variance = Credit > Debt (Overpaid/Surplus)
    # Negative Variance = Credit < Debt (Underpaid/Partial Refund)
    variance_check['Variance'] = variance_check['Amt_Float_CREDIT'] - variance_check['Total_Debts_Covered']
    
    # Filter for significant variance (> 0.01)
    variance_report = variance_check[variance_check['Variance'].abs() > 0.01].copy()
    
    if not variance_report.empty:
        print(f"  ⚠ Found {len(variance_report)} MATCHES WITH VARIANCE (Amount Mismatches)")
        
        # Add Status Column
        def get_status(v):
            if v > 0: return "OVERPAID (Refund > Debts)"
            return "UNDERPAID (Refund < Debts)"
            
        variance_report['Status'] = variance_report['Variance'].apply(get_status)
        
        # Collect "Bad Credits" to block full reconciliation
        # Key needs to match what we can look up from merged
        for _, row in variance_report.iterrows():
            # Tuple: (CreditFile, Card, Op)
            bad_key = (row['Accounting_Ref_CREDIT'], row[col_card], row[col_op])
            bad_credit_keys.add(bad_key)
        
        # Rename for export
        variance_report.rename(columns={
            'Accounting_Ref_CREDIT': 'Credit_File',
            'Amt_Float_CREDIT': 'Refund_Amount'
        }, inplace=True)


    # --- 4c. Fully Reconciled Debtor Notes Summary (Moved after Variance for validation) ---
    # Check if all 'NO' items in a file are matched AND clean.
    print("Generating Fully Reconciled Summary (Strict Mode)...")
    fully_reconciled_files = []
    
    # Group by file
    debt_groups = df_debt.groupby('Accounting_Ref')
    
    for filename, group in debt_groups:
        total_no = group[group[col_recuperar] == 'NO']
        if total_no.empty:
            continue # Skip files with no debtor notes
            
        # CHECK 1: Must have NO Pending Claims (Unmatched Items)
        # If this file appears in the pending_claims list, it's disqualified
        if filename in pending_claims['Accounting_Ref'].values:
            continue

        # CHECK 2: Must have NO Unexpected Refunds
        # If this file appears in unexpected_refunds (as the DEBT source), it's disqualified
        if filename in unexpected_refunds['Accounting_Ref_DEBT'].values:
             continue
            
        # Check overlaps (Should be 100% since check 1 passed, but verify)
        matched_no = total_no[total_no['temp_key'].isin(merged_keys)]
        
        if len(total_no) == len(matched_no):
            # 100% Match!
            # Get the creditors for this file from merged
            relevant_merged = merged[
                (merged['Accounting_Ref_DEBT'] == filename) & 
                (merged[f'{col_recuperar}_DEBT'] == 'NO')
            ]
            
            # CHECK 3: None of the matching credits can have Variance
            has_bad_credit = False
            for _, row in relevant_merged.iterrows():
                # Check if this specific credit match has a variance
                check_key = (row['Accounting_Ref_CREDIT'], row[col_card], row[col_op])
                if check_key in bad_credit_keys:
                    has_bad_credit = True
                    break
            
            if has_bad_credit:
                continue # Disqualified due to variance
            
            # Group by Creditor to get breakdown of how much each covered
            creditor_breakdown = relevant_merged.groupby('Accounting_Ref_CREDIT').agg(
                Amount_Covered=('Amt_Float_DEBT', 'sum'),
                Count_Ops=('Operation Number', 'count')
            ).reset_index()
            
            for _, row_breakdown in creditor_breakdown.iterrows():
                fully_reconciled_files.append({
                    'Debtor_Note_File': filename,
                    'Creditor_File': row_breakdown['Accounting_Ref_CREDIT'],
                    'Amount_Covered': row_breakdown['Amount_Covered'],
                    'Count_Ops': row_breakdown['Count_Ops'],
                    'Status': 'FULLY RECONCILED'
                })
            
    df_full_reconciled = pd.DataFrame(fully_reconciled_files)
    
    if not df_full_reconciled.empty:
        # Keep only required columns and rename
        df_full_reconciled = df_full_reconciled[['Debtor_Note_File', 'Creditor_File', 'Amount_Covered']]
        df_full_reconciled.columns = ['DEBTOR FILE', 'CREDIT FILE NOTE', 'AMOUNT THAT MATCHED']
        
        # Add Total Row
        total_matched = df_full_reconciled['AMOUNT THAT MATCHED'].sum()
        total_row = pd.DataFrame([{
            'DEBTOR FILE': 'TOTAL', 
            'CREDIT FILE NOTE': '', 
            'AMOUNT THAT MATCHED': total_matched
        }])
        
        df_full_reconciled = pd.concat([df_full_reconciled, total_row], ignore_index=True)

    # --- 6. EXPORT ---
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            debt_breakdown.to_excel(writer, sheet_name='By_Debt_File', index=False)
            credit_breakdown.to_excel(writer, sheet_name='By_Credit_File', index=False)
            
            if not pending_claims.empty:
                pending_claims.to_excel(writer, sheet_name='Pending_Claims', index=False)
                
            if not unexpected_refunds.empty:
                unexpected_refunds.to_excel(writer, sheet_name='Unexpected_Refunds', index=False)
                
            if not df_full_reconciled.empty:
                df_full_reconciled.to_excel(writer, sheet_name='Fully_Reconciled_Notes', index=False)
            
            if not variance_report.empty:
                variance_report.to_excel(writer, sheet_name='Amount_Variances', index=False)
            
            # Detailed Audit Sheet (Optional but recommended for tracing duplicates)
            merged.to_excel(writer, sheet_name='Detailed_Audit_Log', index=False)
            
        print(f"SUCCESS. Report saved to: {output_file}")
        print("NOTE: 'Total_Conciliated_Amount' is calculated based on the sum of DEBT notes found.")
        
    except PermissionError:
        print(f"ERROR: Please close {output_file} and run again.")

if __name__ == "__main__":
    robust_conciliation_duplicates_allowed()
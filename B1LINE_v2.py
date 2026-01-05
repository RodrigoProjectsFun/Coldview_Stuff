import pandas as pd
import numpy as np
import re
import os
import time
from datetime import datetime, timedelta

# =============================================================================
# FIELD CONFIGURATION
# =============================================================================
FIELD_CONFIG = {
    'line1_fields': {
        'OPERAC':           (0, 6),
        'RS':               (8, 10),
        'MOVIM':            (12, 17),
        'MONEDA ORIGINAL':  (19, 22),
        'IMPORTE ORIGINAL': (22, 37),
        'MONEDA VISA':      (37, 40),
        'IMPORT VISA':      (40, 55),
        'MONEDA AFECTADO':  (55, 58),
        'IMPORTE AFECTADO': (58, 73),
        'TIPO CUENTA':      (73, 77),
        'CUENTA AFECTADA':  (77, 97),
        'FECOPE':           (97, 106),
        'HORA':             (106, 113),
        'FBASE1':           (113, 122),
        'EXPIRACION':       (122, 128),
    },
    'line2_fields': {
        'TERMINAL':       (0, 12),
        'TIPO':           (12, 17),
        'IDENTIFICACION': (17, 32),
        'ESTABLECIMIENTO':(32, 58),
        'CIUDAD':         (58, 72),
        'PAIS':           (72, 78),
        'BIN ADQUIR.':    (78, 91),
        'PIN':            (91, 96),
        'VIS.REFER':      (96, 108),
        'TRNX':           (108, 113),
        'CAVV':           (113, 119),
        'POS.C.CODE':     (119, 140),
    },
}

IMPORTE_FIELDS = ['IMPORTE ORIGINAL', 'IMPORT VISA', 'IMPORTE AFECTADO']

def get_last_business_day():
    today = datetime.now()
    offset = 1
    if today.weekday() == 0: offset = 3
    elif today.weekday() == 6: offset = 2
    return today - timedelta(days=offset)

def clean_importe(series):
    """Vectorized cleanup of numeric columns."""
    # Remove non-numeric/non-dot/non-minus
    cleaned = series.str.replace(r'[^\d.-]', '', regex=True)
    # Convert to numeric, coercing errors to NaN
    return pd.to_numeric(cleaned, errors='coerce')

def parse_fast(file_path):
    print(f"Reading {file_path}...")
    start_time = time.time()

    # 1. Read entire file into a Series
    # 'sep' is None or a non-existent char to force reading lines
    try:
        # Read as fixed width (width=1000) or just read_csv with single column
        # quoting=3 (QUOTE_NONE) is important to avoid parsing quotes
        df = pd.read_csv(file_path, header=None, names=['raw'], 
                         sep='\0', quoting=3, engine='c', 
                         encoding='utf-8', encoding_errors='replace')
    except Exception as e:
        # Fallback for truly unstructured text if read_csv fails
        with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
        df = pd.DataFrame(lines, columns=['raw'])
        # Strip newlines
        df['raw'] = df['raw'].str.rstrip('\n\r')

    # 2. Identify Metadata Rows
    
    # Card Header: Starts with "- TARJETA"
    # We strip() first to handle indentation safely, though - TARJETA usually starts at 0
    df['stripped'] = df['raw'].str.strip()
    
    is_card = df['stripped'].str.startswith('- TARJETA', na=False)
    
    # 3. Extract Card Info and Propagate (ffill)
    # Extract "TARJETA ..." and "NOMBRE ..."
    # Format: "- TARJETA <card_num>     NOMBRE <name>"
    # We can use regex extract on the strictly identified card rows
    card_info_df = df.loc[is_card, 'raw'].str.extract(
        r'- TARJETA\s+(?P<TARJETA>\S+).*?NOMBRE\s+(?P<NOMBRE>.*)'
    )
    
    # Assign these to the main df
    df['TARJETA'] = np.nan
    df['NOMBRE'] = np.nan
    
    df.loc[is_card, ['TARJETA', 'NOMBRE']] = card_info_df.values
    
    # Forward fill to propagate card info to subsequent transaction lines
    df[['TARJETA', 'NOMBRE']] = df[['TARJETA', 'NOMBRE']].ffill()

    # 4. Filter Data Rows
    # Exclude: Card headers, Page Headers (***), Separators (---), Empty lines
    # Data Rule: 
    #   Line 1: stripped starts with DIGIT
    #   Line 2: raw starts with space (and is not empty)
    # To be robust like B1LINE.py which pairs them:
    # We will identify potential data lines and assume pairs.
    
    # Filter out known non-data
    is_separator = df['stripped'].str.contains(r'^\*+|^-+$', regex=True)
    is_empty = df['stripped'] == ''
    
    # Data Candidates must have Card Info and not be headers/separators
    mask_candidates = (~is_card) & (~is_separator) & (~is_empty) & (df['TARJETA'].notna())
    candidates = df[mask_candidates].copy()
    
    # 5. Split into Pairs (Line 1 / Line 2)
    # Reset index to allow clean slicing
    candidates = candidates.reset_index(drop=True)
    
    if len(candidates) % 2 != 0:
        print(f"Warning: Odd number of data lines ({len(candidates)}). Checking for alignment...")
        # Optional: Add logic to align based on content if strict pairing fails
        # For now, strict pairing is the B1LINE behavior equivalent logic
        # Truncate last one to avoid crash or investigate? 
        # B1LINE drops the last pending if no pairing found.
        candidates = candidates.iloc[:-1]

    # Assume Even indices = Line 1, Odd indices = Line 2
    # Verify assumption: Line 1 usually starts with Digit (after strip)
    # Check first few?
    # line1_check = candidates.iloc[::2]['stripped'].str[0].str.isdigit().mean()
    # print(f"Heuristic: {line1_check*100:.1f}% of Line 1s start with digit")
    
    line1_df = candidates.iloc[::2].reset_index(drop=True)
    line2_df = candidates.iloc[1::2].reset_index(drop=True)
    
    # 6. Extract Fields
    # Helper to slice fixed width from 'raw'
    def extract_fields(source_df, config):
        extracted = pd.DataFrame(index=source_df.index)
        for field, (start, end) in config.items():
            # Slice, strip, and assign
            # PAD row with spaces if too short? str.slice handles it gracefully (returns empty or shorter)
            extracted[field] = source_df['raw'].str.slice(start, end).str.strip()
        return extracted

    extracted_l1 = extract_fields(line1_df, FIELD_CONFIG['line1_fields'])
    extracted_l2 = extract_fields(line2_df, FIELD_CONFIG['line2_fields'])
    
    # 7. Merge
    final_df = pd.concat([
        line1_df[['TARJETA', 'NOMBRE']], 
        extracted_l1, 
        extracted_l2
    ], axis=1)
    
    # 8. Type Conversion
    for col in IMPORTE_FIELDS:
        if col in final_df.columns:
            final_df[col] = clean_importe(final_df[col])
            
    end_time = time.time()
    print(f"Parsed {len(final_df)} records in {end_time - start_time:.4f} seconds")
    return final_df

def run_test():
    input_file = 'test_repro.txt'
    if not os.path.exists(input_file):
        print(f"{input_file} not found.")
        return

    df = parse_fast(input_file)
    print(df.head().T)
    
    output_path = f"test_output_fast_{datetime.now().strftime('%H%M%S')}.xlsx"
    df.to_excel(output_path, index=False)
    print(f"Saved to {output_path}")

if __name__ == "__main__":
    run_test()

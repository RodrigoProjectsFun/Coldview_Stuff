import pandas as pd
import numpy as np
import re
import os
import time
from datetime import datetime, timedelta
from tqdm import tqdm

# Standard input filename
INPUT_FILENAME = 'reporte.txt'

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
    """Get the last business day (Monday-Friday), skipping weekends."""
    today = datetime.now()
    offset = 1
    if today.weekday() == 0:  # Monday
        offset = 3
    elif today.weekday() == 6:  # Sunday
        offset = 2
    return today - timedelta(days=offset)

def generate_output_filename(output_dir=None):
    """
    Generate the standard output filename with last business day date.
    Format: BASE 1 PENDIENTES DE CONCILIAR LINEALIZADO (DD-MM-YYYY).xlsx
    """
    last_bday = get_last_business_day()
    date_str = last_bday.strftime("%d-%m-%Y")
    filename = f"BASE 1 PENDIENTES DE CONCILIAR LINEALIZADO ({date_str}).xlsx"
    if output_dir:
        return os.path.join(output_dir, filename)
    return filename

def clean_importe(series):
    """Vectorized cleanup of numeric columns."""
    # Remove any non-numeric characters except . and -
    # Note: Regex replaces everything NOT (^) in [\d.-] with empty string
    cleaned = series.astype(str).str.replace(r'[^\d.-]', '', regex=True)
    return pd.to_numeric(cleaned, errors='coerce')

def get_page_skip_mask(df):
    """
    Generate a boolean mask for lines to skip based on the legacy state machine.
    
    Logic:
    - ENTER skip mode when line contains "*****".
    - EXIT skip mode when 2nd "-----" line is seen after entering.
    """
    # 1. Identify Event Locations
    # Using specific markers from legacy logic
    star_idxs = df.index[df['stripped'].str.contains(r'\*{5,}', regex=True)].tolist()
    dash_idxs = df.index[df['stripped'].str.contains(r'-{5,}', regex=True)].tolist()
    
    # Combined events: (index, type)
    events = sorted(
        [(i, 'STAR') for i in star_idxs] + 
        [(i, 'DASH') for i in dash_idxs]
    )
    
    skip_mask = np.zeros(len(df), dtype=bool)
    
    if not events:
        return skip_mask

    # 2. Process Events to build ranges
    current_start = -1
    in_skip_mode = False
    dash_count = 0
    
    exclusion_ranges = []
    
    for idx, event_type in events:
        if event_type == 'STAR':
            # Always reset / start skipping on STAR
            # If we were already skipping, we just reset the start point effectively 
            # (though practically simpler to just treat as start of new block or continuation)
            if not in_skip_mode:
                in_skip_mode = True
                current_start = idx
            else:
                # Re-trigger: Reset dash count logic, but we are already skipping from previous start
                # Legacy code: "dash_lines_seen = 0"
                dash_count = 0
                
        elif event_type == 'DASH':
            if in_skip_mode:
                dash_count += 1
                if dash_count >= 2:
                    in_skip_mode = False
                    # End skipping at this line (inclusive)
                    exclusion_ranges.append((current_start, idx))
                    current_start = -1
                    dash_count = 0

    # Handle case where file ends while skipping
    if in_skip_mode and current_start != -1:
        exclusion_ranges.append((current_start, len(df) - 1))

    # 3. Fill Mask
    for start, end in exclusion_ranges:
        skip_mask[start : end + 1] = True
        
    return skip_mask


def parse_cobol_vectorized(file_path, output_path):
    print(f"Processing {file_path}...")
    start_time = time.time()

    # 1. Read entire file into a Series
    # Use read_csv with special separator to read whole lines quickly
    try:
        # engine='c' is faster. sep='\0' or similar avoids splitting. 
        # quoting=3 (csv.QUOTE_NONE) ensures no quoting logic runs.
        # encoding='utf-8-sig' handles BOM (ï»¿) automatically
        df = pd.read_csv(
            file_path, 
            header=None, 
            names=['raw'], 
            sep='\0', 
            quoting=3, 
            engine='c', 
            encoding='utf-8-sig', 
            encoding_errors='replace'
        )
    except Exception:
        # Fallback for systems where read_csv might behave differently on text files
        with open(file_path, 'r', encoding='utf-8-sig', errors='replace') as f:
            lines = f.readlines()
        df = pd.DataFrame(lines, columns=['raw'])
        df['raw'] = df['raw'].str.rstrip('\n\r')

    print(f"Loaded {len(df):,} lines.")

    # 2. Identify Metadata Rows and Structure
    df['stripped'] = df['raw'].str.strip()
    
    # Identify Card Headers
    is_card = df['stripped'].str.startswith('- TARJETA', na=False)
    
    # Identify Separators / Metadata to exclude
    is_separator = df['stripped'].str.contains(r'^\*+|^-+$', regex=True)
    is_empty = df['stripped'] == ''

    # BLOCK-BASED FILTERING (Replaces regex)
    is_page_header = get_page_skip_mask(df)
    
    # 3. Extract Card Info
    card_info_df = df.loc[is_card, 'raw'].str.extract(
        r'- TARJETA\s+(?P<TARJETA>\S+).*?NOMBRE\s+(?P<NOMBRE>.*)'
    )
    
    df['TARJETA'] = np.nan
    df['NOMBRE'] = np.nan
    
    if not card_info_df.empty:
        df.loc[is_card, ['TARJETA', 'NOMBRE']] = card_info_df.values
    
    df[['TARJETA', 'NOMBRE']] = df[['TARJETA', 'NOMBRE']].ffill()

    # 4. Filter Data Rows
    # A generic data row is one that:
    #   - Has card info (TARJETA is not null)
    #   - Is not inside a Page Header block
    #   - Is not a card header itself
    #   - Is not a separator or empty line
    mask_candidates = (~is_card) & (~is_separator) & (~is_empty) & (~is_page_header) & (df['TARJETA'].notna())
    candidates = df[mask_candidates].copy()
    
    
    if candidates.empty:
        print("Warning: No data rows found.")
        return 0
    
    # 5. Split into Pairs (Line 1 / Line 2)
    # Logic: The file format strictly alternates Line 1 / Line 2 for transactions.
    # Line 1 usually starts with a digit/code. Line 2 contains Terminal/Establecimiento.
    
    # Reset index to operate on 0..N indices
    candidates = candidates.reset_index(drop=True)
    
    # Check parity
    if len(candidates) % 2 != 0:
        print(f"Warning: Odd number of data lines ({len(candidates)}). Dropping last orphan line.")
        candidates = candidates.iloc[:-1]

    # Split using array slicing
    line1_df = candidates.iloc[::2].reset_index(drop=True)
    line2_df = candidates.iloc[1::2].reset_index(drop=True)
    
    # 6. Extract Fixed Width Fields
    def extract_fields(source_df, field_config):
        extracted = pd.DataFrame(index=source_df.index)
        # Handle leading whitespace by stripping it before fixed-width extraction
        # This ensures that even if lines are indented, the indices (0, 6, etc.) apply to the data.
        cleaned_raw = source_df['raw'].str.lstrip()
        
        for field, (start, end) in field_config.items():
            # Slice the cleaned string. 
            ext_col = cleaned_raw.str.slice(start, end).str.strip()
            extracted[field] = ext_col
        return extracted

    extracted_l1 = extract_fields(line1_df, FIELD_CONFIG['line1_fields'])
    extracted_l2 = extract_fields(line2_df, FIELD_CONFIG['line2_fields'])

    # Validation: Filter invalid RS
    # RS must be numeric. If it's not, it's likely a parsing artifact (e.g. alignment issue).
    if 'RS' in extracted_l1.columns:
        valid_rs = pd.to_numeric(extracted_l1['RS'], errors='coerce').notna()
        n_dropped = (~valid_rs).sum()
        if n_dropped > 0:
            print(f"Warning: Dropped {n_dropped} records with invalid (non-numeric) RS values.")
            extracted_l1 = extracted_l1[valid_rs]
            extracted_l2 = extracted_l2[valid_rs]
            line1_df = line1_df[valid_rs]
    
    # 7. Concatenate Final DataFrame
    final_df = pd.concat([
        line1_df[['TARJETA', 'NOMBRE']], 
        extracted_l1, 
        extracted_l2
    ], axis=1)
    
    # 8. Type Conversion
    for col in IMPORTE_FIELDS:
        if col in final_df.columns:
            final_df[col] = clean_importe(final_df[col])

    # 9. Write Output
    print(f"Parsing complete. Found {len(final_df):,} records.")
    print("Writing Excel...")
    
    # Use pandas ExcelWriter with tqdm (optional visually, but nice)
    final_df.to_excel(output_path, index=False)
    
    print(f"Success! Output saved to {output_path}")
    print(f"Total time: {time.time() - start_time:.2f} seconds.")
    
    return len(final_df)

def run(input_file=None, output_dir=None):
    """
    Parse a report with auto-generated output filename.
    
    Args:
        input_file: Path to input file. If None, looks for reporte.txt in script directory.
        output_dir: Optional output directory (defaults to script directory)
    
    Returns:
        Tuple of (output_path, record_count) or (None, 0) if input not found
    """
    # Default to reporte.txt in script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if input_file is None:
        input_file = os.path.join(script_dir, INPUT_FILENAME)
    
    # Check if input exists
    if not os.path.exists(input_file):
        print(f"Input file not found: {input_file}")
        return None, 0
    
    output_path = generate_output_filename(output_dir)
    record_count = parse_cobol_vectorized(input_file, output_path)
    return output_path, record_count

# --- RUN THE SCRIPT ---
if __name__ == "__main__":
    import sys
    input_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run(input_arg)
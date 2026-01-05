import pandas as pd
import re
import time
from tqdm import tqdm

# =============================================================================
# FIELD CONFIGURATION
# Define the character positions (start, end) for each field in fixed-width format.
# Adjust these values based on your production data format.
# Set to None to use delimiter-based parsing instead of fixed-width.
# =============================================================================

FIELD_CONFIG = {
    # Set parsing mode: 'fixed' for fixed-width, 'delimiter' for space-delimited
    'parsing_mode': 'delimiter',  # Change to 'fixed' when you know the exact positions
    
    # Delimiter pattern (used when parsing_mode='delimiter')
    'delimiter_pattern': r'\s{2,}',  # 2+ spaces
    
    # Line 1 fields: (start_col, end_col) - 0-indexed, end is exclusive
    # These are EXAMPLE positions - adjust based on actual production data
    'line1_fields': {
        'OPERAC':           (0, 6),
        'RS':               (8, 10),
        'MOVIM':            (12, 17),
        'IMPORTE ORIGINAL': (18, 32),
        'MONEDA':           (33, 36),
        'IMPORT VISA':      (38, 52),
        'IMPORTE AFECTADO': (53, 67),
        'CUENTA AFECTADA':  (68, 88),
        'FECOPE':           (90, 98),
        'HORA':             (99, 105),
        'FBASE1':           (106, 114),
        'EXPIRACION':       (115, 120),
    },
    
    # Line 2 fields: (start_col, end_col) - 0-indexed, end is exclusive
    'line2_fields': {
        'TERMINAL':       (0, 10),
        'TIPO':           (10, 15),
        'IDENTIFICACION': (16, 22),
        'ESTABLECIMIENTO':(23, 53),
        'CIUDAD':         (54, 68),
        'PAIS':           (69, 71),
        'BIN ADQUIR.':    (73, 79),
        'PIN':            (81, 83),
        'VIS.REFER':      (85, 97),
        'TRNX':           (98, 100),
        'CAVV':           (101, 103),
        'POS.C.CODE':     (104, 120),
    },
}


def count_lines(file_path):
    """Quick line count for progress bar."""
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        return sum(1 for _ in f)


def extract_fixed_width(line, field_config):
    """Extract fields using fixed-width column positions."""
    result = {}
    for field_name, (start, end) in field_config.items():
        if len(line) > start:
            result[field_name] = line[start:min(end, len(line))].strip()
        else:
            result[field_name] = ""
    return result


def extract_delimiter(line, field_names, delimiter_pattern):
    """Extract fields using delimiter-based splitting."""
    parts = delimiter_pattern.split(line.strip())
    safe_parts = parts + [""] * (len(field_names) - len(parts))
    return {name: safe_parts[i] for i, name in enumerate(field_names)}


def parse_cobol_dynamic(file_path, output_path, config=None):
    """
    Parse COBOL-style report file and export to Excel.
    
    Args:
        file_path: Input text file path
        output_path: Output Excel file path
        config: Optional field configuration dict (uses FIELD_CONFIG if None)
    """
    if config is None:
        config = FIELD_CONFIG
    
    print(f"Processing {file_path}...")
    start_time = time.time()
    
    # Setup parsing based on mode
    parsing_mode = config.get('parsing_mode', 'delimiter')
    delimiter_pattern = re.compile(config.get('delimiter_pattern', r'\s{2,}'))
    
    # Get field configs
    line1_fields = config.get('line1_fields', {})
    line2_fields = config.get('line2_fields', {})
    line1_names = list(line1_fields.keys())
    line2_names = list(line2_fields.keys())
    
    data_rows = []
    
    # State variables
    current_card_info = None
    current_person_info = None
    pending_record = None 
    
    # Strategy flags
    skipping_mode = False
    dash_lines_seen = 0
    
    # Count lines for progress bar
    total_lines = count_lines(file_path)
    
    # --- READING PHASE ---
    print(f"Reading {total_lines:,} lines...")
    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        for line in tqdm(f, total=total_lines, desc="Reading", unit="lines"):
            stripped_line = line.strip()
            
            # 1. CHECK START OF HEADER (The Asterisks)
            if "*****" in stripped_line:
                skipping_mode = True
                dash_lines_seen = 0
                continue
            
            # 2. HANDLE SKIPPING MODE
            if skipping_mode:
                if "-----" in stripped_line:
                    dash_lines_seen += 1
                    if dash_lines_seen >= 2:
                        skipping_mode = False
                continue

            if not stripped_line: 
                continue
            
            # PATTERN 1: Card Header (- TARJETA)
            if stripped_line.startswith("- TARJETA"):
                content = stripped_line.replace("- TARJETA", "").strip()
                parts = delimiter_pattern.split(content)
                
                current_card_info = parts[0] if len(parts) >= 1 else "UNKNOWN"
                
                if len(parts) >= 2:
                    name_part = parts[1]
                    if name_part.upper().startswith("NOMBRE "):
                        current_person_info = name_part[7:]
                    else:
                        current_person_info = name_part
                else:
                    current_person_info = "UNKNOWN"
                
                pending_record = None
                continue

            # PATTERN 2: Line 1 - Transaction Data (Starts with digit)
            if line[0].isdigit() and current_card_info is not None:
                if parsing_mode == 'fixed':
                    fields = extract_fixed_width(line, line1_fields)
                else:
                    fields = extract_delimiter(stripped_line, line1_names, delimiter_pattern)
                
                pending_record = {
                    'TARJETA': current_card_info,
                    'NOMBRE': current_person_info,
                    **fields
                }
                continue
            
            # PATTERN 3: Line 2 - Terminal/Merchant Info (Starts with space)
            if line.startswith(" ") and pending_record is not None:
                if parsing_mode == 'fixed':
                    fields = extract_fixed_width(line, line2_fields)
                else:
                    fields = extract_delimiter(stripped_line, line2_names, delimiter_pattern)
                
                pending_record.update(fields)
                data_rows.append(pending_record)
                pending_record = None
                continue

    # --- WRITING PHASE ---
    print(f"\nParsing complete. Found {len(data_rows):,} records.")
    
    if data_rows:
        print("Writing Excel...")
        df = pd.DataFrame(data_rows)
        
        # Write with progress tracking
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # tqdm wrapper for write operation
            with tqdm(total=len(df), desc="Writing", unit="rows") as pbar:
                df.to_excel(writer, index=False)
                pbar.update(len(df))
        
        print(f"Success! Output saved to {output_path}")
    else:
        print("Warning: No data found. Check your text file formatting.")

    print(f"Total time: {time.time() - start_time:.2f} seconds.")
    return len(data_rows)


# --- RUN THE SCRIPT ---
if __name__ == "__main__":
    # Example usage:
    # parse_cobol_dynamic('large_report.txt', 'output.xlsx')
    pass
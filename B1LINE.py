import pandas as pd
import re
import time
import os
from datetime import datetime, timedelta
from tqdm import tqdm

# =============================================================================
# FIELD CONFIGURATION
# Define the character positions (start, end) for each field in fixed-width format.
# Adjust these values based on your production data format.
# Set to None to use delimiter-based parsing instead of fixed-width.
# =============================================================================

FIELD_CONFIG = {
    # Set parsing mode: 'fixed' for fixed-width, 'delimiter' for space-delimited
    'parsing_mode': 'fixed',  # Fixed-width is more reliable for COBOL reports
    
    # Delimiter pattern (used when parsing_mode='delimiter')
    'delimiter_pattern': r'\s{2,}',  # 2+ spaces
    
    # Line 1 fields: (start_col, end_col) - 0-indexed, end is exclusive
    # Column positions verified against sample data
    # Pattern: MONEDA then IMPORTE for each section
    'line1_fields': {
        'OPERAC':           (0, 6),
        'RS':               (8, 10),
        'MOVIM':            (12, 17),
        # ORIGINAL section: 604 (moneda) then 23.00 (importe)
        'MONEDA ORIGINAL':  (19, 22),       # 604
        'IMPORTE ORIGINAL': (22, 37),       # 23.00
        # VISA section: SOL (moneda) then 23.00 (importe)
        'MONEDA VISA':      (37, 40),       # SOL
        'IMPORT VISA':      (40, 55),       # 23.00
        # AFECTADO section: SOL (moneda) then 23.00 (importe)
        'MONEDA AFECTADO':  (55, 58),       # SOL
        'IMPORTE AFECTADO': (58, 73),       # 23.00
        # Account type and number: AHO (tipo) then 194-36830982-0-10
        'TIPO CUENTA':      (73, 77),       # AHO
        'CUENTA AFECTADA':  (77, 97),       # 194-36830982-0-10
        # Dates: 14062025, 234248, 14062025, 06-27
        'FECOPE':           (97, 106),
        'HORA':             (106, 113),
        'FBASE1':           (113, 122),
        'EXPIRACION':       (122, 128),
    },
    
    # Line 2 fields: (start_col, end_col) - 0-indexed, end is exclusive
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


def get_last_business_day():
    """Get the last business day (Monday-Friday), skipping weekends."""
    today = datetime.now()
    offset = 1
    # If today is Monday, go back to Friday (3 days)
    if today.weekday() == 0:  # Monday
        offset = 3
    # If today is Sunday, go back to Friday (2 days)
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


# Fields that should be converted to numeric values
IMPORTE_FIELDS = ['IMPORTE ORIGINAL', 'IMPORT VISA', 'IMPORTE AFECTADO']


def parse_importe(value):
    """
    Convert IMPORTE string to numeric value.
    Handles empty strings, whitespace, and various number formats.
    """
    if not value or not value.strip():
        return None
    
    cleaned = value.strip()
    # Remove any non-numeric characters except . and -
    cleaned = ''.join(c for c in cleaned if c.isdigit() or c in '.-')
    
    if not cleaned:
        return None
    
    try:
        return float(cleaned)
    except ValueError:
        return None


def clean_record(record):
    """Clean a record by converting IMPORTE fields to numeric."""
    for field in IMPORTE_FIELDS:
        if field in record:
            record[field] = parse_importe(record[field])
    return record


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
                data_rows.append(clean_record(pending_record))
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


def run(input_file, output_dir=None):
    """
    Convenience function to parse a report with auto-generated output filename.
    
    Args:
        input_file: Path to the input text file
        output_dir: Optional output directory (defaults to current directory)
    
    Returns:
        Tuple of (output_path, record_count)
    """
    output_path = generate_output_filename(output_dir)
    record_count = parse_cobol_dynamic(input_file, output_path)
    return output_path, record_count


# --- RUN THE SCRIPT ---
if __name__ == "__main__":
    # Example usage:
    # run('large_report.txt')  # Auto-generates filename with last business day
    # parse_cobol_dynamic('input.txt', 'custom_output.xlsx')  # Custom filename
    pass
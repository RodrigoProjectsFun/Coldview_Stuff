import pandas as pd
import numpy as np
import os
import time
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import List, Tuple
import tkinter as tk
from tkinter import filedialog

# =============================================================================
# CONFIGURATION DEFINITIONS
# =============================================================================

@dataclass
class FieldConfig:
    name: str
    start: int
    end: int
    type: str = "string"  # Options: 'string', 'amount', 'numeric'

# Configuration for the first line of each record
LINE_1_CONFIG = [
    FieldConfig('OPERAC',           0,   6),
    FieldConfig('RS',               8,   10,  type='numeric'), # Must be numeric
    FieldConfig('MOVIM',            12,  17),
    FieldConfig('MONEDA ORIGINAL',  19,  22),
    FieldConfig('IMPORTE ORIGINAL', 22,  37,  type='amount'), # Cleaned as currency
    FieldConfig('MONEDA VISA',      37,  40),
    FieldConfig('IMPORT VISA',      40,  55,  type='amount'),
    FieldConfig('MONEDA AFECTADO',  55,  58),
    FieldConfig('IMPORTE AFECTADO', 58,  73,  type='amount'),
    FieldConfig('TIPO CUENTA',      73,  77),
    FieldConfig('CUENTA AFECTADA',  77,  97),
    FieldConfig('FECOPE',           97,  106),
    FieldConfig('HORA',             106, 113),
    FieldConfig('FBASE1',           113, 122),
    FieldConfig('EXPIRACION',       122, 128),
]

# Configuration for the second line of each record
LINE_2_CONFIG = [
    FieldConfig('TERMINAL',         0,   12),
    FieldConfig('TIPO',             12,  17),
    FieldConfig('IDENTIFICACION',   17,  32),
    FieldConfig('ESTABLECIMIENTO',  32,  58),
    FieldConfig('CIUDAD',           58,  72),
    FieldConfig('PAIS',             72,  78),
    FieldConfig('BIN ADQUIR.',      78,  91),
    FieldConfig('PIN',              91,  96),
    FieldConfig('VIS.REFER',        96,  108),
    FieldConfig('TRNX',             108, 113),
    FieldConfig('CAVV',             113, 119),
    FieldConfig('POS.C.CODE',       119, 140),
]

# Standard input filename
INPUT_FILENAME = 'reporte.txt'

# =============================================================================
# UTILITIES
# =============================================================================

def get_last_business_day() -> datetime:
    """Get the last business day (Monday-Friday), skipping weekends."""
    today = datetime.now()
    offset = 1
    if today.weekday() == 0:  # Monday
        offset = 3
    elif today.weekday() == 6:  # Sunday
        offset = 2
    return today - timedelta(days=offset)

def generate_output_filename(output_dir: str = None) -> str:
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

# =============================================================================
# PARSER LOGIC
# =============================================================================

class B1LineParser:
    """Class to parse COLDview BASE 1 reports, converting fixed-width text to Excel."""
    
    def __init__(self, line1_config: List[FieldConfig], line2_config: List[FieldConfig]):
        self.line1_config = line1_config
        self.line2_config = line2_config

    def parse(self, file_path: str, output_path: str) -> int:
        """Main orchestrator for parsing the file."""
        print(f"Processing {file_path}...")
        start_time = time.time()

        # 1. Load data
        df = self._load_file(file_path)
        print(f"Loaded {len(df):,} lines.")
        
        # 2. Extract metadata and filter relevant rows
        df = self._extract_metadata_and_filter(df)
        if df.empty:
            print("Warning: No data rows found.")
            return 0
        
        # 3. Split alternating lines into Line 1 and Line 2
        line1_df, line2_df = self._split_lines(df)
        
        # 4. Extract fields according to fixed-width config
        extracted_l1, extracted_l2 = self._extract_all_fields(line1_df, line2_df)
        
        # 5. Validate mandatory numeric formats
        line1_df, extracted_l1, extracted_l2 = self._validate_records(
            line1_df, extracted_l1, extracted_l2
        )
        
        # 6. Assemble final DataFrame
        final_df = pd.concat([
            line1_df[['TARJETA', 'NOMBRE']], 
            extracted_l1, 
            extracted_l2
        ], axis=1)
        
        # 7. Apply types (like formatting amounts)
        final_df = self._apply_types(final_df)
        
        # 8. Export to Excel
        print(f"Parsing complete. Found {len(final_df):,} records.")
        print("Writing Excel...")
        final_df.to_excel(output_path, index=False)
        
        print(f"Success! Output saved to {output_path}")
        print(f"Total time: {time.time() - start_time:.2f} seconds.")
        
        return len(final_df)

    def _load_file(self, file_path: str) -> pd.DataFrame:
        """Read the entire text file efficiently into a DataFrame."""
        try:
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
            with open(file_path, 'r', encoding='utf-8-sig', errors='replace') as f:
                lines = f.readlines()
            df = pd.DataFrame(lines, columns=['raw'])
            df['raw'] = df['raw'].str.rstrip('\n\r')
        return df

    def _get_page_skip_mask(self, df: pd.DataFrame) -> np.ndarray:
        """Generate a boolean mask for header/footer blocks that should be parsed out."""
        star_idxs = df.index[df['stripped'].str.contains(r'\*{5,}', regex=True)].tolist()
        dash_idxs = df.index[df['stripped'].str.contains(r'-{5,}', regex=True)].tolist()
        
        events = sorted([(i, 'STAR') for i in star_idxs] + [(i, 'DASH') for i in dash_idxs])
        skip_mask = np.zeros(len(df), dtype=bool)
        
        if not events:
            return skip_mask

        current_start = -1
        in_skip_mode = False
        dash_count = 0
        exclusion_ranges = []
        
        for idx, event_type in events:
            if event_type == 'STAR':
                if not in_skip_mode:
                    in_skip_mode = True
                    current_start = idx
                else:
                    dash_count = 0
            elif event_type == 'DASH':
                if in_skip_mode:
                    dash_count += 1
                    if dash_count >= 2:
                        in_skip_mode = False
                        exclusion_ranges.append((current_start, idx))
                        current_start = -1
                        dash_count = 0

        if in_skip_mode and current_start != -1:
            exclusion_ranges.append((current_start, len(df) - 1))

        for start, end in exclusion_ranges:
            skip_mask[start : end + 1] = True
            
        return skip_mask

    def _extract_metadata_and_filter(self, df: pd.DataFrame) -> pd.DataFrame:
        """Find card sections and isolate actual transaction rows."""
        df['stripped'] = df['raw'].str.strip()
        
        is_card = df['stripped'].str.startswith('- TARJETA', na=False)
        is_separator = df['stripped'].str.contains(r'^\*+|^-+$', regex=True)
        is_empty = df['stripped'] == ''
        is_page_header = self._get_page_skip_mask(df)
        
        # Extract 'TARJETA' and 'NOMBRE' context headers
        card_info_df = df.loc[is_card, 'raw'].str.extract(
            r'- TARJETA\s+(?P<TARJETA>\S+).*?NOMBRE\s+(?P<NOMBRE>.*)'
        )
        
        df['TARJETA'] = np.nan
        df['NOMBRE'] = np.nan
        if not card_info_df.empty:
            df.loc[is_card, ['TARJETA', 'NOMBRE']] = card_info_df.values
        df[['TARJETA', 'NOMBRE']] = df[['TARJETA', 'NOMBRE']].ffill()

        # Data rows are what is left over
        mask_candidates = (
            (~is_card) & 
            (~is_separator) & 
            (~is_empty) & 
            (~is_page_header) & 
            (df['TARJETA'].notna())
        )
        return df[mask_candidates].copy()

    def _split_lines(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Split the dataset by alternating transaction lines (Line 1 vs Line 2)."""
        df = df.reset_index(drop=True)
        if len(df) % 2 != 0:
            print(f"Warning: Odd number of data lines ({len(df)}). Dropping last orphan line.")
            df = df.iloc[:-1]

        line1_df = df.iloc[::2].reset_index(drop=True)
        line2_df = df.iloc[1::2].reset_index(drop=True)
        return line1_df, line2_df

    def _extract_fields(self, source_df: pd.DataFrame, field_config: List[FieldConfig], clean_raw_series: pd.Series) -> pd.DataFrame:
        """Slice the raw strings based on the field configurations."""
        extracted = pd.DataFrame(index=source_df.index)
        for field in field_config:
            extracted[field.name] = clean_raw_series.str.slice(field.start, field.end).str.strip()
        return extracted

    def _extract_all_fields(self, line1_df: pd.DataFrame, line2_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Extract fields using the coupled indentation logic for Line 1 and Line 2."""
        # Line 1 indentation defines the relative offset
        line1_len = line1_df['raw'].str.len()
        line1_clean = line1_df['raw'].str.lstrip()
        line1_indent_counts = line1_len - line1_clean.str.len()
        
        # Apply Line 1's indentation to Line 2 to preserve inner spacing
        l2_raw = line2_df['raw'].tolist()
        l1_indents = line1_indent_counts.tolist()
        l2_cleaned_list = [s[i:] for s, i in zip(l2_raw, l1_indents)]
        line2_clean = pd.Series(l2_cleaned_list, index=line2_df.index)

        extracted_l1 = self._extract_fields(line1_df, self.line1_config, line1_clean)
        extracted_l2 = self._extract_fields(line2_df, self.line2_config, line2_clean)
        return extracted_l1, extracted_l2

    def _validate_records(self, line1_df: pd.DataFrame, l1: pd.DataFrame, l2: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """Check for errors like invalid RS non-numeric values. Drops invalid rows across both lines."""
        valid_mask = pd.Series(True, index=l1.index)
        
        for config, data in [(self.line1_config, l1), (self.line2_config, l2)]:
            for field in config:
                if field.type == 'numeric' and field.name in data.columns:
                    # numeric fields must not end up as nan
                    is_valid = pd.to_numeric(data[field.name], errors='coerce').notna()
                    valid_mask = valid_mask & is_valid
                    
        n_dropped = (~valid_mask).sum()
        if n_dropped > 0:
            print(f"Warning: Dropped {n_dropped} records strictly failing validation rules.")
            
        return line1_df[valid_mask], l1[valid_mask], l2[valid_mask]

    def _apply_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean amount fields to valid numeric representations."""
        for config in [self.line1_config, self.line2_config]:
            for field in config:
                if field.type == 'amount' and field.name in df.columns:
                    cleaned_amount = df[field.name].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    df[field.name] = pd.to_numeric(cleaned_amount, errors='coerce')
        return df


def run(input_file: str = None, output_dir: str = None):
    """
    Parse a report with auto-generated output filename.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if input_file is None:
        root = tk.Tk()
        root.withdraw()
        input_file = filedialog.askopenfilename(
            title="Seleccionar reporte a linealizar",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialdir=script_dir
        )
        if not input_file:
            print("No se seleccionó ningún archivo. Operación cancelada.")
            return None, 0
    
    if not os.path.exists(input_file):
        print(f"Input file not found: {input_file}")
        return None, 0
    
    output_path = generate_output_filename(output_dir)
    parser = B1LineParser(LINE_1_CONFIG, LINE_2_CONFIG)
    record_count = parser.parse(input_file, output_path)
    return output_path, record_count

if __name__ == "__main__":
    import sys
    input_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run(input_arg)
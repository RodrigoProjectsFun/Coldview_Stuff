import pandas as pd
import numpy as np
import os
import time
import json
from datetime import datetime, timedelta
from dataclasses import dataclass
from typing import List, Tuple, Dict
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

# =============================================================================
# CONFIGURATION DEFINITIONS
# =============================================================================

@dataclass
class FieldConfig:
    name: str
    start: int
    end: int
    type: str = "string"  # Options: 'string', 'amount', 'numeric'

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

def generate_output_filename(report_type: str, output_dir: str = None) -> str:
    """
    Generate the standard output filename with last business day date.
    Format: PENDIENTES DE CONCILIAR LINEALIZADO B1 (DD-MM-YYYY).xlsx
    """
    last_bday = get_last_business_day()
    date_str = last_bday.strftime("%d-%m-%Y")
    filename = f"PENDIENTES DE CONCILIAR LINEALIZADO {report_type.upper()} ({date_str}).xlsx"
    if output_dir:
        return os.path.join(output_dir, filename)
    return filename

# =============================================================================
# PARSER LOGIC
# =============================================================================

class ReportParser:
    """Class to parse COLDview fixed-width reports dynamically based on JSON config."""
    
    def __init__(self, lines_per_record: int, line_configs: List[List[FieldConfig]]):
        self.lines_per_record = lines_per_record
        self.line_configs = line_configs

    def parse(self, file_path: str, output_path: str) -> int:
        """Main orchestrator for parsing the file."""
        print(f"Procesando {file_path}...")
        start_time = time.time()

        # 1. Load data
        df = self._load_file(file_path)
        print(f"Cargadas {len(df):,} líneas crudas.")
        
        # 2. Extract metadata and filter relevant rows
        df = self._extract_metadata_and_filter(df)
        if df.empty:
            print("Advertencia: No se detectaron líneas de datos válidas (posiblemente esté vacío o mal formateado).")
            return 0
        
        # 3. Split alternating lines dynamically into N groups
        line_dfs = self._split_lines(df)
        if not line_dfs:
            return 0
        
        # 4. Extract fields according to fixed-width config
        extracted_dfs = self._extract_all_fields(line_dfs)
        
        # 5. Validate mandatory numeric formats
        line_dfs, extracted_dfs = self._validate_records(line_dfs, extracted_dfs)
        if not extracted_dfs:
            return 0
            
        # 6. Apply formatting and types to amounts
        extracted_dfs = self._apply_types(extracted_dfs)
        
        # 7. Assemble final DataFrame
        # We always attach TARJETA and NOMBRE from the first line DataFrame context.
        final_pieces = [line_dfs[0][['TARJETA', 'NOMBRE']]] + extracted_dfs
        final_df = pd.concat(final_pieces, axis=1)
        
        # 8. Export to Excel
        print(f"Parseo completado. Encontrados {len(final_df):,} registros correctos.")
        print("Guardando Excel...")
        try:
            final_df.to_excel(output_path, index=False)
            print(f"¡Éxito! Archivo guardado en {output_path}")
        except Exception as e:
            print(f"Error guardando Excel. ¿Está abierto? Detalles: {e}")
            if 'tk' in globals():
                messagebox.showerror("Error", f"No se pudo guardar el archivo Excel.\n{e}")
            return 0

        print(f"Tiempo total: {time.time() - start_time:.2f} segundos.")
        
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

    def _split_lines(self, df: pd.DataFrame) -> List[pd.DataFrame]:
        """Split the dataset by dynamically separating transaction lines."""
        df = df.reset_index(drop=True)
        N = self.lines_per_record
        rem = len(df) % N
        if rem != 0:
            print(f"Advertencia: El número final de líneas a procesar ({len(df)}) no es divisible por {N}. Ignorando las últimas {rem} líneas huérfanas.")
            df = df.iloc[:-rem]

        line_dfs = []
        for i in range(N):
            line_df = df.iloc[i::N].reset_index(drop=True)
            line_dfs.append(line_df)
            
        return line_dfs

    def _extract_fields(self, source_df: pd.DataFrame, field_config: List[FieldConfig], clean_raw_series: pd.Series) -> pd.DataFrame:
        """Slice the raw strings based on the field configurations."""
        extracted = pd.DataFrame(index=source_df.index)
        for field in field_config:
            extracted[field.name] = clean_raw_series.str.slice(field.start, field.end).str.strip()
        return extracted

    def _extract_all_fields(self, line_dfs: List[pd.DataFrame]) -> List[pd.DataFrame]:
        """Extract fields using the coupled indentation logic relative to Line 1 for all subsequent lines."""
        if not line_dfs:
            return []
            
        line1_df = line_dfs[0]
        # Calculate horizontal relative shift from line 1
        line1_len = line1_df['raw'].str.len()
        line1_clean = line1_df['raw'].str.lstrip()
        line1_indent_counts = (line1_len - line1_clean.str.len()).tolist()
        
        extracted_dfs = []
        
        # 1st line extraction
        config1 = self.line_configs[0] if len(self.line_configs) > 0 else []
        extracted_dfs.append(self._extract_fields(line1_df, config1, line1_clean))
        
        # Subsequent lines
        for i in range(1, len(line_dfs)):
            line_n_df = line_dfs[i]
            ln_raw = line_n_df['raw'].tolist()
            # Apply Line 1's indentation down to Line N to preserve inner spacing
            ln_cleaned_list = [s[ind:] for s, ind in zip(ln_raw, line1_indent_counts)]
            line_n_clean = pd.Series(ln_cleaned_list, index=line_n_df.index)
            
            config_n = self.line_configs[i] if i < len(self.line_configs) else []
            extracted_dfs.append(self._extract_fields(line_n_df, config_n, line_n_clean))
            
        return extracted_dfs

    def _validate_records(self, line_dfs: List[pd.DataFrame], extracted_dfs: List[pd.DataFrame]) -> Tuple[List[pd.DataFrame], List[pd.DataFrame]]:
        """Check for errors like invalid RS non-numeric values. Drops invalid rows across all lines."""
        if not extracted_dfs:
            return line_dfs, extracted_dfs
            
        valid_mask = pd.Series(True, index=extracted_dfs[0].index)
        
        for i, config in enumerate(self.line_configs):
            if i >= len(extracted_dfs):
                break
            data = extracted_dfs[i]
            for field in config:
                if field.type == 'numeric' and field.name in data.columns:
                    # numeric fields must not end up as nan or empty
                    is_valid = pd.to_numeric(data[field.name], errors='coerce').notna()
                    valid_mask = valid_mask & is_valid
                    
        n_dropped = (~valid_mask).sum()
        if n_dropped > 0:
            print(f"Advertencia: Excluidos {n_dropped} registros que reprobaron reglas de formato estricto (ej. el RS no era un número).")
            
        return [df[valid_mask] for df in line_dfs], [df[valid_mask] for df in extracted_dfs]

    def _apply_types(self, extracted_dfs: List[pd.DataFrame]) -> List[pd.DataFrame]:
        """Clean amount fields to valid numeric representations."""
        for i, config in enumerate(self.line_configs):
            if i >= len(extracted_dfs):
                break
            df = extracted_dfs[i]
            for field in config:
                if field.type == 'amount' and field.name in df.columns:
                    cleaned_amount = df[field.name].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                    df[field.name] = pd.to_numeric(cleaned_amount, errors='coerce')
        return extracted_dfs

def load_config(config_path: str) -> dict:
    if not os.path.exists(config_path):
        return None
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def build_field_configs(raw_line_configs: List[List[dict]]) -> List[List[FieldConfig]]:
    parsed_configs = []
    for line_conf in raw_line_configs:
        parsed_line = []
        for field in line_conf:
            parsed_line.append(FieldConfig(
                name=field['name'],
                start=field['start'],
                end=field['end'],
                type=field.get('type', 'string')
            ))
        parsed_configs.append(parsed_line)
    return parsed_configs

def run(input_file: str = None, output_dir: str = None):
    """
    Parse a report with auto-generated output filename based on JSON layout configs.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(script_dir, 'b1_config.json')
    config_data = load_config(config_path)
    
    # Needs a parent root window for messageboxes
    root = tk.Tk()
    root.withdraw()
    
    if not config_data:
        print(f"Error crítico: Configuración JSON no encontrada en {config_path}")
        messagebox.showerror("Archivo No Encontrado", f"No se encontró el archivo de configuración:\n{config_path}")
        return None, 0

    layout_keys = list(config_data.keys())
    if not layout_keys:
        messagebox.showerror("Error", "El archivo de configuración JSON está vacío o sin reportes.")
        return None, 0
        
    # 1. Ask User for Layout (B1, B2, etc.)
    opciones_str = " / ".join(layout_keys)
    selected_layout = simpledialog.askstring(
        "Formato del Reporte", 
        f"¿Qué formato de reporte vas a procesar?\n\nOpciones detectadas:\n{opciones_str}",
        initialvalue=layout_keys[0],
        parent=root
    )
    
    if selected_layout is None:
        print("Operación cancelada por el usuario.")
        return None, 0
        
    # Make it case-insensitive
    match = [k for k in layout_keys if k.upper() == selected_layout.strip().upper()]
    if not match:
        messagebox.showerror("Error", f"'{selected_layout}' no es una opción válida.")
        return None, 0
        
    actual_layout_key = match[0]
    layout_conf = config_data[actual_layout_key]
    
    lines_per_record = layout_conf.get('lines_per_record', 2)
    raw_line_configs = layout_conf.get('line_configs', [])
    line_configs = build_field_configs(raw_line_configs)

    # 2. Open File Dialog
    if input_file is None:
        input_file = filedialog.askopenfilename(
            title=f"Seleccionar reporte {actual_layout_key} a linealizar",
            filetypes=[("Archivos de Texto", "*.txt"), ("Todos los archivos", "*.*")],
            initialdir=script_dir
        )
        if not input_file:
            print("No se seleccionó ningún archivo. Operación cancelada.")
            return None, 0
    
    if not os.path.exists(input_file):
        print(f"Archivo no encontrado: {input_file}")
        messagebox.showerror("Error de Archivo", f"El archivo seleccionado ya no existe:\n{input_file}")
        return None, 0
    
    # 3. Parse and Export
    output_path = generate_output_filename(actual_layout_key, output_dir)
    parser = ReportParser(lines_per_record, line_configs)
    record_count = parser.parse(input_file, output_path)
    
    return output_path, record_count

if __name__ == "__main__":
    import sys
    input_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run(input_arg)
    
    print("\n" + "-"*40)
    input("Presione ENTER para salir...")
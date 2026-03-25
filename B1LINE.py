import pandas as pd
import numpy as np
import logging
import json
import time
import sys
from pathlib import Path
from datetime import datetime, timedelta
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Optional, Any
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

# =============================================================================
# CONFIGURATION DEFINITIONS
# =============================================================================

# Setup rudimentary logger
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

@dataclass
class FieldConfig:
    """
    Data structure representing the configuration for a single field extraction.
    
    Attributes:
        name (str): The column name for this field in the output DataFrame.
        start (int): The starting character index (0-based) for slicing.
        end (int): The ending character index (exclusive) for slicing.
        type (str): The expected data type of the field ('string', 'amount', 'numeric').
    """
    name: str
    start: int
    end: int
    type: str = "string"

# Standard input filename to be used if no file dialog is presented
INPUT_FILENAME = "reporte.txt"

# =============================================================================
# UTILITIES
# =============================================================================

def get_last_business_day() -> datetime:
    """
    Calculate and return the date of the last business day (Monday-Friday).
    If today is Monday, it returns the previous Friday (offset 3 days).
    If today is Sunday, it returns Friday (offset 2 days).
    Otherwise, it returns yesterday (offset 1 day).
    """
    today = datetime.now()
    offset = {0: 3, 6: 2}.get(today.weekday(), 1)
    return today - timedelta(days=offset)

def generate_output_filename(report_type: str, output_dir: Optional[Path] = None) -> Path:
    """
    Generate a standardized output filename containing the report type and the last business day's date.
    
    Args:
        report_type (str): The type of report being processed (e.g., 'B1', 'B2').
        output_dir (Path, optional): The directory where the file should be saved.
        
    Returns:
        Path: The full path or filename for the output Excel file.
    """
    date_str = get_last_business_day().strftime("%d-%m-%Y")
    filename = Path(f"PENDIENTES DE CONCILIAR LINEALIZADO {report_type.upper()} ({date_str}).xlsx")
    return output_dir / filename if output_dir else filename

def extract_smart_field(text: Any, start: int, end: int) -> str:
    """
    Extracts a fixed-width field from text, ensuring continuous words 
    are kept intact and assigned only to the most appropriate field.
    """
    if not isinstance(text, str):
        return ""
    
    length = len(text)
    if start >= length:
        return ""
    
    s, e = max(0, start), min(length, end)
    
    # Scenario A: The START boundary slices a continuous word in half.
    if s > 0 and s < length and not text[s].isspace() and not text[s-1].isspace():
        while s < length and not text[s].isspace():
            s += 1
            
    # Scenario B: The END boundary slices a continuous word in half.
    while e > 0 and e < length and not text[e-1].isspace() and not text[e].isspace():
        e += 1
        
    return text[s:e].strip() if s < e else ""

# =============================================================================
# PARSER LOGIC
# =============================================================================

class ReportParser:
    """
    A class responsible for parsing COLDview fixed-width text reports, extracting
    relevant transaction lines, and applying layout configurations to structure the data.
    """
    
    def __init__(self, lines_per_record: int, line_configs: List[List[FieldConfig]]):
        self.lines_per_record = lines_per_record
        self.line_configs = line_configs

    def parse(self, file_path: Path, output_path: Path) -> int:
        """
        Main orchestration method that reads the file, cleans the data,
        extracts fields, and saves the resulting structured data into an Excel file.
        """
        logger.info(f"Procesando {file_path}...")
        start_time = time.time()

        # 1. Load data
        df = self._load_file(file_path)
        logger.info(f"Cargadas {len(df):,} líneas crudas.")
        
        # 2. Extract metadata and filter relevant rows
        df = self._extract_metadata_and_filter(df)
        if df.empty:
            logger.warning("No se detectaron líneas de datos válidas (posiblemente esté vacío o mal formateado).")
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
        final_pieces = [line_dfs[0][['TARJETA', 'NOMBRE']]] + extracted_dfs
        final_df = pd.concat(final_pieces, axis=1)
        
        # 8. Export to Excel
        logger.info(f"Parseo completado. Encontrados {len(final_df):,} registros correctos.")
        logger.info("Guardando Excel...")
        
        try:
            final_df.to_excel(output_path, index=False)
            logger.info(f"¡Éxito! Archivo guardado en {output_path}")
        except Exception as e:
            logger.error(f"Error guardando Excel. ¿Está abierto? Detalles: {e}")
            if 'tk' in sys.modules:
                messagebox.showerror("Error", f"No se pudo guardar el archivo Excel.\n{e}")
            return 0

        logger.info(f"Tiempo total: {time.time() - start_time:.2f} segundos.")
        return len(final_df)

    def _load_file(self, file_path: Path) -> pd.DataFrame:
        """Reads the entire text file efficiently into a single-column DataFrame named 'raw'."""
        try:
            return pd.read_csv(
                file_path, header=None, names=['raw'], sep='\0', quoting=3, 
                engine='c', encoding='utf-8-sig', encoding_errors='replace'
            )
        except Exception:
            with file_path.open('r', encoding='utf-8-sig', errors='replace') as f:
                lines = [line.rstrip('\n\r') for line in f]
            return pd.DataFrame(lines, columns=['raw'])

    def _get_page_skip_mask(self, df: pd.DataFrame) -> np.ndarray:
        """Identify page headers/footers to create an exclusion mask."""
        is_star = df['stripped'].str.contains(r'\*{5,}', regex=True, na=False)
        is_dash = df['stripped'].str.contains(r'-{5,}', regex=True, na=False)
        
        events = sorted([(i, 'STAR') for i in df.index[is_star]] + 
                        [(i, 'DASH') for i in df.index[is_dash]])
        
        skip_mask = np.zeros(len(df), dtype=bool)
        if not events:
            return skip_mask

        current_start, dash_count, in_skip_mode = -1, 0, False
        exclusion_ranges = []
        
        for idx, event_type in events:
            if event_type == 'STAR':
                if not in_skip_mode:
                    in_skip_mode, current_start = True, idx
                else:
                    dash_count = 0
            elif event_type == 'DASH' and in_skip_mode:
                dash_count += 1
                if dash_count >= 2:
                    exclusion_ranges.append((current_start, idx))
                    in_skip_mode, current_start, dash_count = False, -1, 0

        if in_skip_mode and current_start != -1:
            exclusion_ranges.append((current_start, len(df) - 1))

        for start, end in exclusion_ranges:
            skip_mask[start:end + 1] = True
            
        return skip_mask

    def _extract_metadata_and_filter(self, df: pd.DataFrame) -> pd.DataFrame:
        """Isolates metadata ('TARJETA', 'NOMBRE') and filters out junk lines."""
        df['stripped'] = df['raw'].str.strip()
        
        is_card = df['stripped'].str.startswith('- TARJETA', na=False)
        is_separator = df['stripped'].str.contains(r'^\*+|^-+$', regex=True, na=False)
        is_empty = df['stripped'] == ''
        is_page_header = self._get_page_skip_mask(df)
        
        # Extract 'TARJETA' and 'NOMBRE' context headers
        card_info_df = df.loc[is_card, 'raw'].str.extract(
            r'- TARJETA\s+(?P<TARJETA>\S+).*?NOMBRE\s+(?P<NOMBRE>.*)'
        )
        
        df = df.assign(TARJETA=np.nan, NOMBRE=np.nan)
        if not card_info_df.empty:
            df.loc[is_card, ['TARJETA', 'NOMBRE']] = card_info_df.values
            
        df[['TARJETA', 'NOMBRE']] = df[['TARJETA', 'NOMBRE']].ffill()

        mask_candidates = ~is_card & ~is_separator & ~is_empty & ~is_page_header & df['TARJETA'].notna()
        return df[mask_candidates].copy()

    def _split_lines(self, df: pd.DataFrame) -> List[pd.DataFrame]:
        """Split dataframe into separate dataframes for Line 1, Line 2, etc."""
        df = df.reset_index(drop=True)
        n = self.lines_per_record
        if (rem := len(df) % n) != 0:
            logger.warning(f"El número total de líneas ({len(df)}) no es divisible por {n}. Ignorando {rem} huérfanas.")
            df = df.iloc[:-rem]

        return [df.iloc[i::n].reset_index(drop=True) for i in range(n)]

    def _extract_fields(self, source_df: pd.DataFrame, field_config: List[FieldConfig], clean_raw_series: pd.Series) -> pd.DataFrame:
        """Slices raw strings using smart boundary adjustment."""
        extracted = pd.DataFrame(index=source_df.index)
        for field in field_config:
            extracted[field.name] = clean_raw_series.apply(
                lambda x: extract_smart_field(x, field.start, field.end)
            )
        return extracted

    def _extract_all_fields(self, line_dfs: List[pd.DataFrame]) -> List[pd.DataFrame]:
        """Extracts fields across all line instances holding correct indentation."""
        if not line_dfs:
            return []
            
        line1_df = line_dfs[0]
        line1_len = line1_df['raw'].str.len()
        line1_clean = line1_df['raw'].str.lstrip()
        line1_indents = (line1_len - line1_clean.str.len()).tolist()
        
        extracted_dfs = []
        for i, line_df in enumerate(line_dfs):
            clean_series = (
                line1_clean if i == 0 
                else pd.Series([s[ind:] for s, ind in zip(line_df['raw'], line1_indents)], index=line_df.index)
            )
            config = self.line_configs[i] if i < len(self.line_configs) else []
            extracted_dfs.append(self._extract_fields(line_df, config, clean_series))
            
        return extracted_dfs

    def _validate_records(self, line_dfs: List[pd.DataFrame], extracted_dfs: List[pd.DataFrame]) -> Tuple[List[pd.DataFrame], List[pd.DataFrame]]:
        """Drops entire records containing textual garbage in strictly numeric fields."""
        if not extracted_dfs:
            return line_dfs, extracted_dfs
            
        valid_mask = pd.Series(True, index=extracted_dfs[0].index)
        
        for config, data in zip(self.line_configs, extracted_dfs):
            for field in config:
                if field.type == 'numeric' and field.name in data.columns:
                    is_valid = pd.to_numeric(data[field.name], errors='coerce').notna()
                    valid_mask &= is_valid
                    
        if (n_dropped := (~valid_mask).sum()) > 0:
            logger.warning(f"Excluidos {n_dropped} registros que reprobaron reglas de formato (ej. RS no numérico).")
            
        return [df[valid_mask] for df in line_dfs], [df[valid_mask] for df in extracted_dfs]

    def _apply_types(self, extracted_dfs: List[pd.DataFrame]) -> List[pd.DataFrame]:
        """Sanitizes 'amount' columns to valid floating-point numeric types."""
        for config, df in zip(self.line_configs, extracted_dfs):
            amount_fields = [f.name for f in config if f.type == 'amount' and f.name in df.columns]
            for field in amount_fields:
                cleaned = df[field].astype(str).str.replace(r'[^\d.-]', '', regex=True)
                df[field] = pd.to_numeric(cleaned, errors='coerce')
        return extracted_dfs

def load_config(config_path: Path) -> Optional[Dict]:
    """Reads JSON configuration file."""
    if not config_path.exists():
        return None
    with config_path.open('r', encoding='utf-8') as f:
        return json.load(f)

def build_field_configs(raw_line_configs: List[List[Dict]]) -> List[List[FieldConfig]]:
    """Converts JSON dictionary representation into FieldConfig lists."""
    return [
        [FieldConfig(name=f['name'], start=f['start'], end=f['end'], type=f.get('type', 'string')) for f in line]
        for line in raw_line_configs
    ]

def run(input_file: Optional[str] = None, output_dir: Optional[str] = None) -> Tuple[Optional[Path], int]:
    """
    Entry point for the application execution.
    """
    script_dir = Path(__file__).resolve().parent
    config_path = script_dir / 'b1_config.json'
    
    root = tk.Tk()
    root.withdraw()
    
    if not (config_data := load_config(config_path)):
        logger.critical(f"Configuración JSON no encontrada en {config_path}")
        messagebox.showerror("Archivo No Encontrado", f"No se encontró:\n{config_path}")
        return None, 0

    if not (layout_keys := list(config_data.keys())):
        messagebox.showerror("Error", "El archivo de configuración JSON está vacío o sin reportes.")
        return None, 0
        
    opciones_str = " / ".join(layout_keys)
    selected_layout = simpledialog.askstring(
        "Formato del Reporte", 
        f"¿Qué formato de reporte vas a procesar?\n\nOpciones:\n{opciones_str}",
        initialvalue=layout_keys[0], parent=root
    )
    
    if not selected_layout:
        logger.info("Operación cancelada por el usuario.")
        return None, 0
        
    actual_layout_key = next((k for k in layout_keys if k.casefold() == selected_layout.strip().casefold()), None)
    if not actual_layout_key:
        messagebox.showerror("Error", f"'{selected_layout}' no es una opción válida.")
        return None, 0
        
    layout_conf = config_data[actual_layout_key]
    line_configs = build_field_configs(layout_conf.get('line_configs', []))
    lines_per_record = layout_conf.get('lines_per_record', 2)

    input_path = Path(input_file) if input_file else Path(filedialog.askopenfilename(
        title=f"Seleccionar reporte {actual_layout_key} a linealizar",
        filetypes=[("Archivos de Texto", "*.txt"), ("Todos los archivos", "*.*")],
        initialdir=script_dir
    ))
    
    if not input_path.name:
        logger.info("No se seleccionó ningún archivo. Cancelado.")
        return None, 0
    
    if not input_path.exists():
        logger.error(f"Archivo no encontrado: {input_path}")
        messagebox.showerror("Error", f"El archivo ya no existe:\n{input_path}")
        return None, 0
    
    out_dir_path = Path(output_dir) if output_dir else None
    output_path = generate_output_filename(actual_layout_key, out_dir_path)
    
    parser = ReportParser(lines_per_record, line_configs)
    record_count = parser.parse(input_path, output_path)
    
    return output_path, record_count

if __name__ == "__main__":
    input_arg = sys.argv[1] if len(sys.argv) > 1 else None
    run(input_arg)
    
    print("\n" + "-"*40)
    input("Presione ENTER para salir...")
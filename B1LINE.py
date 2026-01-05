import pandas as pd
import re
import time

def parse_cobol_dynamic(file_path, output_path):
    print(f"Processing {file_path}...")
    start_time = time.time()
    
    # --- CONFIGURATION ---
    # Matches 2 or more spaces acting as a delimiter
    delimiter_pattern = re.compile(r'\s{2,}')
    
    data_rows = []
    
    # State variables
    current_card_info = "UNKNOWN"
    current_person_info = "UNKNOWN"
    pending_record = None 
    
    # STRATEGY FLAGS
    skipping_mode = False
    dash_lines_seen = 0

    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
        for line in f:
            stripped_line = line.strip()
            
            # 1. CHECK START OF HEADER (The Asterisks)
            # We look for a long string of asterisks to trigger "Skip Mode"
            if "*****" in stripped_line:
                skipping_mode = True
                dash_lines_seen = 0 # Reset counter
                continue
            
            # 2. HANDLE SKIPPING MODE
            if skipping_mode:
                # We look for the separator lines "------"
                if "-----" in stripped_line:
                    dash_lines_seen += 1
                    
                    # We only stop skipping after the SECOND dash line 
                    # (This ensures we skip the 'OPERAC' column headers too)
                    if dash_lines_seen >= 2:
                        skipping_mode = False
                
                # While in skipping mode, we ignore EVERYTHING (Merchant names, Dates, etc.)
                continue

            # ---------------------------------------------------------
            # DATA EXTRACTION (Only happens when NOT in skipping mode)
            # ---------------------------------------------------------
            
            if not stripped_line: continue
            
            # Capture Card Info (Group Header)
            if stripped_line.startswith("- TARJETA"):
                parts = delimiter_pattern.split(stripped_line.replace("- TARJETA", "").strip())
                if len(parts) >= 1: current_card_info = parts[0]
                if len(parts) >= 2: current_person_info = parts[1]
                pending_record = None # Reset
                continue

            # Line 1: Main Transaction Data (Starts with a digit)
            if line[0].isdigit():
                parts = delimiter_pattern.split(stripped_line)
                # Ensure list is long enough to avoid crashes
                safe_parts = parts + [""] * (15 - len(parts))
                
                pending_record = {
                    'Card_Number': current_card_info,
                    'Card_Holder': current_person_info,
                    'Operac': safe_parts[0],
                    'RS': safe_parts[1],
                    'Movim': safe_parts[2],
                    'Importe_Orig': safe_parts[3],
                    'Moneda': safe_parts[4],
                    'Importe_Visa': safe_parts[5],
                    'Importe_Afec': safe_parts[6],
                    'Cuenta': safe_parts[7] if len(parts) > 7 else "",
                    # Mapping dates/times from the end of the list
                    'Fec_Ope': safe_parts[-4] if len(parts) > 4 else "",
                    'Hora': safe_parts[-3] if len(parts) > 3 else "",
                    'F_Base': safe_parts[-2] if len(parts) > 2 else "",
                    'Expiracion': safe_parts[-1] if len(parts) > 1 else "",
                }
            
            # Line 2: Terminal/Merchant Info (Starts with space, and we have a Line 1 ready)
            elif line.startswith(" ") and pending_record is not None:
                parts = delimiter_pattern.split(stripped_line)
                safe_parts = parts + [""] * (10 - len(parts))
                
                pending_record.update({
                    'Terminal': safe_parts[0],
                    'MCC': safe_parts[1] if len(parts) > 1 else "",
                    'Establecimiento': safe_parts[2] if len(parts) > 2 else "",
                    'Ciudad': safe_parts[3] if len(parts) > 3 else "",
                    'Pais': safe_parts[4] if len(parts) > 4 else "",
                    'Ref_Num': safe_parts[-3] if len(parts) > 3 else "",
                    'Auth_Code': safe_parts[-2] if len(parts) > 2 else "",
                })
                
                data_rows.append(pending_record)
                pending_record = None # Record complete

    # --- EXPORT TO EXCEL ---
    print(f"Parsing complete. Found {len(data_rows)} records. Writing Excel...")
    
    if data_rows:
        df = pd.DataFrame(data_rows)
        # Using xlsxwriter for speed
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        print(f"Success! Output saved to {output_path}")
    else:
        print("Warning: No data found. Check your text file formatting.")

    print(f"Total time: {time.time() - start_time:.2f} seconds.")

# --- RUN THE SCRIPT ---
# parse_cobol_dynamic('large_report.txt', 'output.xlsx')
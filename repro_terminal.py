import pandas as pd
import B1LINE
import os

def test_empty_terminal():
    input_file = 'repro_terminal_data.txt'
    output_file = 'output_terminal.xlsx'
    
    # Line 1 (Valid)
    line1 = "005695021201010001USD0000000000100.00USD0000000000100.00USD0000000000100.0011110000000000000012342023010112000020230101000000"
    
    # Line 2 with EMPTY TERMINAL
    # TERMINAL is 0-12 (12 chars). 
    # TIPO is 12-17 (5 chars). Let's set TIPO = '09999'
    # If correct, TERMINAL should be '' or spaces, TIPO should be '09999'.
    # If incorrect (lstrip assumes spaces are indentation), TIPO might end up in TERMINAL.
    
    # Using "." to represent spaces for visualization, then replace.
    # 12 spaces for terminal, then 09999 for TIPO.
    # "            09999..."
    line2_empty_terminal_no_indent = " " * 12 + "09999" + "X" * 50
    
    # Line 2 with INDENTATION AND EMPTY TERMINAL
    # 5 spaces indentation + 12 spaces terminal + 09999
    # Total 17 spaces + 09999
    line2_indented_empty_terminal = " " * 5 + " " * 12 + "09999" + "X" * 50

    # User Request: 2 spaces L1, 4 spaces L2
    line1_2spaces = "  " + "005695021201010001USD0000000000100.00USD0000000000100.00USD0000000000100.0011110000000000000012342023010112000020230101000000"
    # TERMINAL = "123456789012" (12 chars). 4 Leading spaces.
    line2_4spaces = "    " + "123456789012" + "X"*50

    line1_indented_5 = " " * 5 + line1.lstrip()

    content = [
        "- TARJETA 123456 LINEA NOMBRE TEST USER",
        line1,
        line2_empty_terminal_no_indent,
        line1_indented_5,
        line2_indented_empty_terminal,
        line1_2spaces,
        line2_4spaces
    ]
    
    with open(input_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(content))
        
    print(f"Created {input_file}...")
    
    if os.path.exists(output_file):
        os.remove(output_file)

    try:
        count = B1LINE.parse_cobol_vectorized(input_file, output_file)
        print(f"Parsed {count} records.")
        
        df = pd.read_excel(output_file, dtype=str)
        
        # Check Record 1 (No indent, Empty Terminal)
        rec1 = df.iloc[0]
        term1 = str(rec1.get('TERMINAL', '')).strip()
        tipo1 = str(rec1.get('TIPO', '')).strip()
        
        print(f"Rec 1 - TERMINAL: '{term1}' (Expected ''), TIPO: '{tipo1}' (Expected '09999')")
        
        # Check Record 2 (Indent 5, Empty Terminal)
        rec2 = df.iloc[1]
        term2 = str(rec2.get('TERMINAL', '')).strip()
        tipo2 = str(rec2.get('TIPO', '')).strip()
        
        
        print(f"Rec 2 - TERMINAL: '{term2}' (Expected ''), TIPO: '{tipo2}' (Expected '09999')")

        # Check Record 3 (User Request: 2 leading spaces L1, 4 leading spaces L2)
        # L1: 2 spaces. L2: 4 spaces.
        # If we use L1 indent (2), we strip 2 from L2. L2 starts with 2 spaces.
        # TERMINAL (0-12) should capture those 2 spaces + first 10 chars of data.
        # Data used: "123456789012"
        # L2 constructed: "    " + "123456789012"
        # Result expected: "  1234567890" (First 2 spaces + 10 chars)
        rec3 = df.iloc[2]
        term3 = str(rec3.get('TERMINAL', ''))
        # Don't strip term3 for checking exact content
        
        print(f"Rec 3 - TERMINAL Raw: '{term3}'")
        
        if kind_of_failure(term1, tipo1) or kind_of_failure(term2, tipo2):
             print("FAILURE: Parsing logic incorrect for empty TERMINAL.")
        else:
             print("SUCCESS: Parsing logic correct.")

    except Exception as e:
        print(f"Error: {e}")

def kind_of_failure(term, tipo):
    # Failure if TIPO is empty (meaning it got shifted into TERMINAL)
    # OR if TERMINAL has the TIPO value
    if tipo != '09999':
        return True
    return False

if __name__ == "__main__":
    test_empty_terminal()

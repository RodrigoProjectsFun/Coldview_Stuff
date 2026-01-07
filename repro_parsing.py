import pandas as pd
import B1LINE
import os

def test_parsing():
    input_file = 'sample_shifted.txt'
    output_file = 'output_repro.xlsx'
    
    if os.path.exists(output_file):
        os.remove(output_file)

    print(f"Testing parsing with {input_file}...")
    try:
        count = B1LINE.parse_cobol_vectorized(input_file, output_file)
        
        if count == 0:
            print("FAILURE: No records parsed.")
            return

        df = pd.read_excel(output_file, dtype=str)
        first_row = df.iloc[0]
        
        # Check if OPERAC is correct. It should be '005695'
        # If shifting occurred, it will likely be empty or '00569' etc depending on extraction
        operac = str(first_row.get('OPERAC', '')).strip()
        print(f"Extracted OPERAC: '{operac}'")
        
        if operac == '005695':
            print("SUCCESS: OPERAC matches expected value.")
        else:
            print(f"FAILURE: OPERAC '{operac}' does not match expected '005695'. Likely shifting issue.")
            
    except Exception as e:
        print(f"FAILURE: Exception occurred: {e}")

if __name__ == "__main__":
    test_parsing()

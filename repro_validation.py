import pandas as pd
import B1LINE
import os

def test_validation():
    input_file = 'repro_validation_data.txt'
    output_file = 'output_validation.xlsx'
    
    # Create valid line 1 & 2
    # RS is pos 8-10 (length 2). "12" is valid.
    # Line 1 length must be sufficient.
    # OPERAC(0-6), RS(8-10) ...
    # "005695  12  ..."
    line1_valid = "005695021201010001USD0000000000100.00USD0000000000100.00USD0000000000100.0011110000000000000012342023010112000020230101000000"
    # invalid RS "XX"
    line1_invalid = "00569502XX01010001USD0000000000100.00USD0000000000100.00USD0000000000100.0011110000000000000012342023010112000020230101000000"
    
    line2 = "TERMINAL00100123456789012345ESTABLECIMIENTO TEST      CIUDAD TEST   PAIS  US000000000000012345REF12345678900001123456000000000000000000000"
    
    content = [
        "- TARJETA 123456 LINEA NOMBRE TEST USER",
        line1_valid,
        line2,
        line1_invalid,
        line2
    ]
    
    with open(input_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(content))
        
    print(f"Created {input_file} with 1 valid and 1 invalid record.")
    
    if os.path.exists(output_file):
        os.remove(output_file)

    try:
        count = B1LINE.parse_cobol_vectorized(input_file, output_file)
        
        print(f"Parsed {count} records.")
        
        # We expect 1 record if validation works, 2 if it doesn't
        if count == 1:
            print("SUCCESS: Only 1 record parsed (invalid RS filtered).")
        elif count == 2:
            print("FAILURE: 2 records parsed (invalid RS NOT filtered).")
        else:
            print(f"UNEXPECTED: {count} records parsed.")
            
    except Exception as e:
        print(f"Error: {e}")
        
if __name__ == "__main__":
    test_validation()

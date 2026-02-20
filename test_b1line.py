import unittest
from unittest.mock import patch
import os
import pandas as pd
import B1LINE

class TestB1LineParser(unittest.TestCase):
    def setUp(self):
        # Create test input files
        self.b1_file = "test_b1.txt"
        with open(self.b1_file, "w") as f:
            f.write("- TARJETA 1111222233334444 NOMBRE JUAN PEREZ\n")
            f.write("000001  20  MOV01ARS0  000000000200.00USD0  00000000100.00ARS0  000000000200.00CTA 000000000000000001240101 123000 FBASE1230 120124\n")
            f.write("TERM00001   PAGOSID1234567890123ESTABLECIMIENTO MUESTRA   CIUDAD MUESTRAAR    BIN1234567890PIN12VISREFER123TRNX0CAVV12POSCCODE1234567890\n")
            f.write("000002  20  MOV02ARS0  000000000300.00USD0  00000000150.00ARS0  000000000300.00CTA 000000000000000002240101 123000 FBASE1230 120124\n")
            f.write("TERM00002   PAGOSID1234567890123ESTABLECIMIENTO MUESTRA   CIUDAD MUESTRAAR    BIN1234567890PIN12VISREFER123TRNX0CAVV12POSCCODE1234567890\n")

        self.b2_file = "test_b2.txt"
        with open(self.b2_file, "w") as f:
            f.write("- TARJETA 5555666677778888 NOMBRE MARIA GOMEZ\n")
            f.write("000003  20  MOV03ARS0  000000000400.00USD0  00000000200.00ARS0  000000000400.00CTA 000000000000000003240101 123000 FBASE1230 120124\n")
            f.write("TERM00003   PAGOSID1234567890123ESTABLECIMIENTO MUESTRA   CIUDAD MUESTRAAR    BIN1234567890PIN12VISREFER123TRNX0CAVV12POSCCODE1234567890\n")
            f.write("CAMPOXTRA1          CAMPOXTRA2          \n")

    def tearDown(self):
        if os.path.exists(self.b1_file): os.remove(self.b1_file)
        if os.path.exists(self.b2_file): os.remove(self.b2_file)
        # Clear out generated excels
        for file in os.listdir("."):
            if file.startswith("PENDIENTES DE CONCILIAR LINEALIZADO"):
                os.remove(file)

    @patch('tkinter.simpledialog.askstring')
    @patch('tkinter.filedialog.askopenfilename')
    def test_b1_format(self, mock_askopenfilename, mock_askstring):
        mock_askstring.return_value = 'B1'
        mock_askopenfilename.return_value = self.b1_file
        
        output_path, record_count = B1LINE.run()
        self.assertEqual(record_count, 2)
        
        # Verify excel content
        df = pd.read_excel(output_path)
        self.assertEqual(len(df), 2)
        self.assertIn("TARJETA", df.columns)
        self.assertEqual(df.iloc[0]["TARJETA"], 1111222233334444)
        self.assertIn("PIN", df.columns) # specific to line 2

    @patch('tkinter.simpledialog.askstring')
    @patch('tkinter.filedialog.askopenfilename')
    def test_b2_format(self, mock_askopenfilename, mock_askstring):
        mock_askstring.return_value = 'B2'
        mock_askopenfilename.return_value = self.b2_file
        
        output_path, record_count = B1LINE.run()
        self.assertEqual(record_count, 1)
        
        # Verify excel content
        df = pd.read_excel(output_path)
        self.assertEqual(len(df), 1)
        self.assertEqual(df.iloc[0]["TARJETA"], 5555666677778888)
        self.assertIn("CAMPO_EXTRA_1", df.columns) # specific to line 3 B2
        self.assertEqual(df.iloc[0]["CAMPO_EXTRA_1"].strip(), "CAMPOXTRA1")
        self.assertEqual(df.iloc[0]["CAMPO_EXTRA_2"].strip(), "CAMPOXTRA2")


if __name__ == '__main__':
    unittest.main()

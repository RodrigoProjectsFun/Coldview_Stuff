from sum_concil import Conciliator
import unittest

class TestStandardization(unittest.TestCase):
    def test_filename_standardization_m6d_dev_vff(self):
        c = Conciliator()
        test_filenames = [
            ('m6d-dev_vff 01.20.2026.xlsx', 'ACREEDORA 01.20.2026'),
            ('M6D-DEV_VFF-01-20-2026.xlsx', 'ACREEDORA 01-20-2026'),
        ]
        
        for filename, expected_prefix in test_filenames:
            result = c.get_standardized_name(filename)
            print(f"File: {filename} -> Result: '{result}' | Expected: '{expected_prefix}'")
            self.assertEqual(result, expected_prefix)

if __name__ == '__main__':
    unittest.main()

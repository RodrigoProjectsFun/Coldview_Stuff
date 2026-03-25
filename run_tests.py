import unittest
import sys
import os
import test_sum_concil

sys.stdout.reconfigure(encoding='utf-8')

if __name__ == '__main__':
    suite = unittest.TestLoader().loadTestsFromModule(test_sum_concil)
    res = unittest.TextTestRunner(stream=sys.stdout, verbosity=2).run(suite)
    if not res.wasSuccessful():
        for test, err in res.failures + res.errors:
            print("="*60)
            print("FAILED TEST:", test)
            print("-" * 60)
            print(err)
        sys.exit(1)

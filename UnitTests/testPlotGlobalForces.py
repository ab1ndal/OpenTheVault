import unittest
import sys
import os

sys.path.append(os.path.join(os.path.dirname(__file__), "..", "scripts"))

from plotGlobalForces import getCutHeight

class TestGetCutHeight(unittest.TestCase):
    def test_getCutHeight(self):
        self.assertEqual(getCutHeight('TOTAL - Z= -10m'), -10)
        self.assertEqual(getCutHeight('TOTAL - Z= 11m'), 11)
        self.assertEqual(getCutHeight('CORE 4 - Z= 0m'), 0)
        self.assertEqual(getCutHeight('CORE 1 - Z= 1.5m'), 1.5)
        self.assertEqual(getCutHeight('CORE 5s - X= -1.45m'), -1.45)

    def test_get_cut_height_invalid(self):
        # Test cases for invalid input
        with self.assertRaises(IndexError):
            getCutHeight('CORE 1')
        with self.assertRaises(IndexError):
            getCutHeight('CORE 2=')
        with self.assertRaises(IndexError):
            getCutHeight('')
        with self.assertRaises(IndexError):
            getCutHeight('CORE 2=m')
        with self.assertRaises(IndexError):
            getCutHeight('CORE 2= m')

if __name__ == '__main__':
    unittest.main()


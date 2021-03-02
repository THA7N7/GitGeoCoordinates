import unittest
import sys, os

sys.path.append(os.getcwd())
from testing.main import *

class TestIngredient(unittest.TestCase):
    def test_ingredient1_returns_i(self):
        self.assertEqual(calc_ingredient(1000), 10.5)
        self.assertEqual(calc_ingredient(1500), 15.75)
        self.assertEqual(calc_ingredient(800), 8.4)

    def test_ingredient1_returns_integer(self):
        self.assertIsInstance(calc_ingredient(1000), float)

    def test_receives_string_returns_float(self):
        self.assertEqual(convert_input('1000'), 1000)

    def test_receives_alpha_string_returns_zero(self):
        self.assertEqual(convert_input('gadsgag'), 0)



if __name__ == '__main__':

    unittest.main()

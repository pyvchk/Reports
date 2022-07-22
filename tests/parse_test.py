import unittest
from parse import parse
import os


class ParseTest(unittest.TestCase):
    def test_ideal_1(self):
        self.assertEqual(parse(dir_name='tests/parse_test/',
                               express_name=r'Экспресс-отчет КЦ-1 КС Калач Калачеевского ЛПУМГ.xlsx'),
                         False)


if __name__ == '__main__':
    unittest.main()

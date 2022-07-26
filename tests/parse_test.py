import unittest
from parse import parse
from configparser import ConfigParser

cfg = ConfigParser()
cfg.read('tests/parse_test/output.ini')


class ParseTest(unittest.TestCase):
    def test_title_1(self):
        self.assertEqual(parse(dir_name='tests/parse_test/',
                               express_name=r'Экспресс-отчет без категории.xlsx'), False)


if __name__ == '__main__':
    unittest.main()

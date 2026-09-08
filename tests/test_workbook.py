import unittest

from calc import Workbook


class TestWorkbook(unittest.TestCase):

    def setUp(self):
        self.wb = Workbook()

    def tearDown(self):
        self.wb.close()

    def test_class(self):
        self.assertEqual(type(self.wb), Workbook)


if __name__ == "__main__":
    unittest.main()

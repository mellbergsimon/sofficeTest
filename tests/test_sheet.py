import unittest

from calc import Workbook, Sheet


class TestSheet(unittest.TestCase):

    def setUp(self):
        self.wb = Workbook()
        self.sheet = self.wb.create_sheet("Test")

    def tearDown(self):
        self.wb.close()

    def test_name(self):
        name = "Sheet"
        sheet = self.wb.create_sheet(name)
        self.assertEqual(sheet.name, name)

    def test_class(self):
        sheet = self.wb.create_sheet("Test")
        self.assertEqual(type(sheet), Sheet)


if __name__ == "__main__":
    unittest.main()

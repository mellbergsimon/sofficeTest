import unittest

from calc import Workbook


class TestCells(unittest.TestCase):

    def setUp(self):
        self.wb = Workbook()
        self.sheet = self.wb.create_sheet("Test")

    def tearDown(self):
        self.wb.close()

    def test_set_scalar(self):
        self.sheet.cells("A1:A3").set(42)

        self.assertEqual(
            self.sheet.cells("A1").get(),
            42,
        )
        self.assertEqual(
            self.sheet.cells("A2").get(),
            42,
        )
        self.assertEqual(
            self.sheet.cells("A3").get(),
            42,
        )

    def test_set_list(self):
        self.sheet.cells("B2:B4").set([1, 2, 3])

        self.assertEqual(
            self.sheet.cells("B2:B4").get(),
            [1, 2, 3],
        )

    def test_set_tuple(self):
        self.sheet.cells("B2:B4").set((1, 2, 3))

        self.assertEqual(
            self.sheet.cells("B2:B4").get(),
            [1, 2, 3],
        )

    def test_set_2d_list(self):
        self.sheet.cells("A1:B2").set(
            [
                [1, 2],
                [3, 4],
            ]
        )

        self.assertEqual(
            self.sheet.cells("A1:B2").get(),
            [
                [1, 2],
                [3, 4],
            ],
        )

    def test_set_wrong_length(self):
        with self.assertRaises(ValueError):
            self.sheet.cells("B2:B4").set([1, 2])

    def test_set_wrong_length_2d_range(self):
        with self.assertRaises(ValueError):
            self.sheet.cells("B2:C3").set([1, 2, 3])


if __name__ == "__main__":
    unittest.main()

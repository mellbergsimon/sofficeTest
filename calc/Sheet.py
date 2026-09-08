import uno
from .Cells import Cells


class Sheet:
    """Wrapper around a LibreOffice Calc sheet."""

    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

    @property
    def name(self):
        return self.sheet.getName()

    def uno(self):
        """Get the raw uno class representation"""
        return self.sheet

    def cells(self, cell_range):
        """Return a Cells object for the specified address."""
        uno_cell = self._get_uno_range(cell_range)

        return Cells(self, uno_cell)

    def show(self):
        """
        Show the sheet for the user.
        """

        self.workbook.controller.setActiveSheet(self.sheet)
        self.workbook.show()
        return self

    def select(self, cell):
        """
        Select a cell or range of cells in the sheet and make it active.

        Show must be called after this to make the selection visible for the user.
        """

        controller = self.workbook.controller

        controller.setActiveSheet(self.sheet)

        cell_range = self.sheet.getCellRangeByName(cell)
        controller.select(cell_range)
        return self

    def set_left_border(self, cell_range, color=0x00B050):
        """
        Set the left border of a cell range.
        """
        cells = self.sheet.getCellRangeByName(cell_range)

        border = cells.getPropertyValue("TableBorder2")

        line = uno.createUnoStruct("com.sun.star.table.BorderLine2")

        line.Color = color
        line.OuterLineWidth = 20
        line.InnerLineWidth = 0
        line.LineDistance = 0

        border.LeftLine = line
        border.IsLeftLineValid = True

        cells.setPropertyValue(
            "TableBorder2",
            border,
        )
        return self

    def delete(self):
        """Delete this current sheet from the Workbook"""
        sheets = self.workbook.doc.getSheets()
        sheets.removeByName(self.name)

    def _get_uno_range(self, cell_range):
        """
        Convert str range to pyuno range if needed.
        """

        if isinstance(cell_range, str):
            return self.sheet.getCellRangeByName(cell_range)
        else:
            return cell_range

    def freeze(self, row=0):
        """Freeze row so scrolling motion still shows this row"""
        controller = self.workbook.doc.getCurrentController()
        controller.setActiveSheet(self.sheet)
        controller.freezeAtPosition(0, row)
        return self

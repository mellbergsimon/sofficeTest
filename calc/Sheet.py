import uno
from .Cells import Cells


class Sheet:
    """Wrapper around a LibreOffice Calc sheet."""

    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

    def name(self):
        return self.sheet.getName()

    def path_str(self):
        return str(self.path)

    def _raw(self):
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

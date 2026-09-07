import uno
from pathlib import Path
import time


class Calc:
    def __init__(self, ctx, path=None):
        self.ctx = ctx
        self.desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx
        )

        if path and Path(path).exists():
            url = uno.systemPathToFileUrl(str(Path(path).absolute()))

            self.doc = self.desktop.loadComponentFromURL(
                url,
                "_blank",
                0,
                (),
            )
        else:
            self.doc = self.desktop.loadComponentFromURL(
                "private:factory/scalc",
                "_blank",
                0,
                (),
            )

        self.window = self.doc.getCurrentController().getFrame().getContainerWindow()

        self.window.setVisible(False)

    def show(self):
        self.window.setVisible(True)
        import subprocess

        subprocess.run(
            [
                "osascript",
                "-e",
                'tell application "LibreOffice" to activate',
            ],
            check=False,
        )

        self.window.setFocus()

    def set_enum(self, sheet, cell_name, values):
        cell = sheet.getCellRangeByName(cell_name)
        validation = cell.getPropertyValue("Validation")

        validation.Type = uno.Enum(
            "com.sun.star.sheet.ValidationType",
            "LIST",
        )

        validation.Formula1 = ";".join(f'"{str(v)}"' for v in values)

        validation.ShowErrorMessage = True
        validation.ErrorTitle = "Invalid value"
        validation.ErrorMessage = "Choose a value from the list."

        validation.ShowList = 1

        cell.setPropertyValue("Validation", validation)

    def select(self, sheet, cell):
        controller = self.doc.getCurrentController()

        controller.setActiveSheet(sheet)

        cell_range = sheet.getCellRangeByName(cell)
        controller.select(cell_range)

    def set_left_border(self, sheet, cell_range, color=0x00B050):
        cells = sheet.getCellRangeByName(cell_range)

        border = cells.getPropertyValue("TableBorder2")

        line = uno.createUnoStruct("com.sun.star.table.BorderLine2")

        line.Color = color
        line.OuterLineWidth = 20
        line.InnerLineWidth = 0
        line.LineDistance = 0

        border.LeftLine = line
        border.IsLeftLineValid = True

        cells.setPropertyValue("TableBorder2", border)

    def merge(self, sheet, cell_range):
        cells = sheet.getCellRangeByName(cell_range)
        cells.merge(True)

    def rotate(self, sheet, cell_range, degrees):
        cells = sheet.getCellRangeByName(cell_range)
        cells.setPropertyValue(
            "RotateAngle",
            int(degrees * 100),
        )

    def autofit_columns(self, sheet, cell_range, min_width=2000):
        columns = sheet.getCellRangeByName(cell_range).getColumns()

        # Let Calc calculate optimal widths
        columns.OptimalWidth = True

        # Enforce minimum width
        for i in range(columns.getCount()):
            column = columns.getByIndex(i)

            if column.Width < min_width:
                column.Width = min_width

    def hide(self):
        self.window.setVisible(False)

    def save(self):
        self.doc.store()

    def save_as_xlsm(self, path):
        path = Path(path).absolute()

        print("exists:", path.exists())
        print("location:", self.doc.getLocation())
        print("has location:", self.doc.hasLocation())

        self.doc.storeAsURL(
            uno.systemPathToFileUrl(str(path)),
            (
                self.prop("FilterName", "Calc MS Excel 2007 VBA XML"),
                self.prop("Overwrite", True),
            ),
        )

    @staticmethod
    def prop(name, value):
        p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        p.Name = name
        p.Value = value
        return p

    def close(self):
        self.doc.close(True)

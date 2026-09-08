import uno
from pathlib import Path
from .soffice import connect
from .Sheet import Sheet


class Workbook:
    def __init__(self, path=None):
        # Connect to LibreOffice
        self.ctx = connect()
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

        isNew = False
        # Load or create a new spreadsheet
        if path and Path(path).exists():
            url = uno.systemPathToFileUrl(str(Path(path).absolute()))

            self.doc = self.desktop.loadComponentFromURL(
                url,
                "_default",
                0,
                (),
            )
        else:
            isNew = True
            self.doc = self.desktop.loadComponentFromURL(
                "private:factory/scalc",
                "_default",
                0,
                (),
            )
        self.window = self.doc.getCurrentController().getFrame().getContainerWindow()

        if isNew:
            self.window.setVisible(False)

        self.controller = self.doc.getCurrentController()
        self.path = Path(path).absolute() if path else None

    @property
    def name(self):
        return self.path.name

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

    def sheet(self, name=None):
        sheets = self.doc.getSheets()

        if name is None:
            return Sheet(self, sheets.getByIndex(0))

        if isinstance(name, int):
            return Sheet(self, sheets.getByIndex(0))

        if name not in sheets.getElementNames():
            raise ValueError(f"Sheet '{name}' does not exist.")

        return Sheet(self, sheets.getByName(name))

    def create_sheet(self, name):
        sheets = self.doc.getSheets()

        if name in sheets.getElementNames():
            return self.sheet(name)

        sheets.insertNewByName(name, sheets.getCount())

        return self.sheet(name)

    def rename_sheet(self, old_name, new_name):
        sheets = self.doc.getSheets()
        if old_name not in sheets.getElementNames():
            raise ValueError(f"Sheet '{old_name}' does not exist.")
        if new_name in sheets.getElementNames():
            raise ValueError(f"Sheet '{new_name}' already exists.")
        sheets.getByName(old_name).setName(new_name)

    def delete_sheets(self, names):
        sheets = self.doc.getSheets()
        for name in names:
            if name not in sheets.getElementNames():
                raise ValueError(f"Sheet '{name}' does not exist.")
            sheets.removeByName(name)

    def hide(self):
        self.window.setVisible(False)

    def save(self):
        self.doc.store()

    def save_as(self, path, extension="xlsm"):
        path = Path(path).absolute()

        match (extension.lower()):
            case "xlsm":
                filter_name = "Calc MS Excel 2007 VBA XML"
            case "xlsx":
                filter_name = "Calc MS Excel 2007 XML"
            case "ods":
                filter_name = "calc8"
            case _:
                raise ValueError(f"Unsupported extension: {extension}")

        self.doc.storeAsURL(
            uno.systemPathToFileUrl(str(path)),
            (
                self.prop("FilterName", filter_name),
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

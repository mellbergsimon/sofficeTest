import uno
from pathlib import Path
from .soffice import connect
from .Sheet import Sheet


class Workbook:
    """Wrapper around a LibreOffice Calc Workbook."""

    def __init__(self, path="Untitled"):
        self.ctx = connect()
        self.desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )

        self.path = Path(path).absolute()
        self._ensure_extension()

        # Load or create a new spreadsheet
        self._open()
        self.window = self.doc.getCurrentController().getFrame().getContainerWindow()

        # Hide by default
        self.hide()
        self.controller = self.doc.getCurrentController()

    @property
    def name(self):
        return self.path.name

    def rename(self, new_name):
        """rename deletes the old if it exists and saves the new name"""
        old_name = self.path
        self.path = new_name
        self._ensure_extension()
        self.save_as(new_name)

        # Delete old name
        if old_name.exists():
            old_name.unlink()

    def sheets(self):
        sheets = self.doc.getSheets()
        return [self.sheet(sheet.getName()) for sheet in sheets]

    def _ensure_extension(self):
        if self.path.suffix:
            return self.path

        return self.path.with_suffix(".xlsm")

    def _open(self):
        path = self.path
        hidden = self.prop("Hidden", True)
        if path and Path(path).exists():
            url = uno.systemPathToFileUrl(str(Path(path).absolute()))

            self.doc = self.desktop.loadComponentFromURL(
                url,
                "_default",
                0,
                (hidden,),
            )
        else:
            self.doc = self.desktop.loadComponentFromURL(
                "private:factory/scalc",
                "_default",
                0,
                (hidden,),
            )
        return self

    def show(self):
        self.window.setVisible(True)

        self.window.setFocus()
        return self

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
        return self.sheet(new_name)

    def delete_sheets(self, names):
        sheets = self.doc.getSheets()
        for name in names:
            if name not in sheets.getElementNames():
                raise ValueError(f"Sheet '{name}' does not exist.")
            sheets.removeByName(name)

    def hide(self):
        self.window.setVisible(False)
        return self

    def save(self):
        self.save_as()
        return self

    def save_as(self, path=None, extension="xlsm"):
        if path is not None:
            self.path = path

        extension = extension.strip(".")

        match (extension.lower()):
            case "xlsm":
                filter_name = "Calc MS Excel 2007 VBA XML"
            case "xlsx":
                filter_name = "Calc MS Excel 2007 XML"
            case "ods":
                filter_name = "calc8"
            case _:
                raise ValueError(f"Unsupported extension: {extension}")

        # Check if extension matches the saved one. Otherwise switch to the new and save
        if self.path.suffix.strip(".") != extension:
            self.path = self.path.with_suffix("." + extension)

        self.doc.storeAsURL(
            uno.systemPathToFileUrl(str(self.path)),
            (
                self.prop("FilterName", filter_name),
                self.prop("Overwrite", True),
            ),
        )
        return self

    @staticmethod
    def prop(name, value):
        p = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        p.Name = name
        p.Value = value
        return p

    def start(self):
        self._open()
        return self

    def close(self):
        self.doc.close(True)
        return self

import uno


class Cells:
    """Represents one or more cells in a Calc sheet."""

    def __init__(self, sheet, uno_range):
        """Create a Cells wrapper around a UNO cell range."""
        self.sheet = sheet
        self._raw = uno_range
        self.rows = self._raw.getRows()
        self.cols = self._raw.getColumns()

    def _raw(self):
        """Return the raw uno representation of the cells"""
        return self._raw

    def set(self, value):
        """Set values to cell range, value will fill range if its a scalar"""
        nr_rows = self.rows.getCount()
        nr_cols = self.cols.getCount()

        # Scalar -> broadcast to every cell
        if isinstance(value, (str, int, float)):
            return self._set_scalar(value)

        # Sequence -> must match range
        if isinstance(value, (list, tuple)):
            return self._set_sequence(value, nr_rows, nr_cols)

        raise TypeError(f"Unsupported value type: {type(value).__name__}")

    def _set_scalar(self, value):
        nr_rows = self.rows.getCount()
        nr_cols = self.cols.getCount()

        data = tuple(tuple(value for _ in range(nr_cols)) for _ in range(nr_rows))

        self._raw.setDataArray(data)
        return self

    def _set_sequence(self, value, nr_rows, nr_cols):
        # 1D sequence
        if not any(isinstance(v, (list, tuple)) for v in value):
            if len(value) != nr_rows * nr_cols:
                raise ValueError(
                    f"Expected {nr_rows * nr_cols} values, " f"got {len(value)}."
                )

            # Treat a 1D sequence as row-major data
            data = []
            i = 0

            for _ in range(nr_rows):
                row = []
                for _ in range(nr_cols):
                    row.append(value[i])
                    i += 1
                data.append(tuple(row))

            self._raw.setDataArray(tuple(data))
            return self

        # 2D sequence
        if len(value) != nr_rows:
            raise ValueError(f"Expected {nr_rows} rows, got {len(value)}.")

        if any(len(row) != nr_cols for row in value):
            raise ValueError(f"Expected a {nr_rows}x{nr_cols} array.")

        self._raw.setDataArray(tuple(tuple(row) for row in value))

        return self

    def _set_string(self, str):
        nr_rows = self.rows.getCount()
        nr_cols = self.cols.getCount()

        data = tuple(tuple(str for _ in range(nr_cols)) for _ in range(nr_rows))

        self._raw.setDataArray(data)

        return self

    def get(self):
        data = self._raw.getDataArray()

        nr_rows = self.rows.getCount()
        nr_cols = self.cols.getCount()

        # Single cell
        if nr_rows == 1 and nr_cols == 1:
            return data[0][0]

        # Single row
        if nr_rows == 1:
            return list(data[0])

        # Single column
        if nr_cols == 1:
            return [row[0] for row in data]

        # 2D range
        return [list(row) for row in data]

    def clear(self):
        self._raw.clearContents(1023)
        return self

    def rotate(self, degrees):
        """
        Rotate the text in a cell range by the given number of degrees.
        """
        cells = self._raw

        cells.setPropertyValue(
            "RotateAngle",
            int(degrees * 100),
        )
        return self

    def merge(self):
        """Clump togheter cells"""
        self._raw.merge(True)
        return self

    def unmerge(self):
        """Unclump cells"""
        self.cells.merge(False)
        return self

    def set_enum(self, values):
        """Set a cell range to have a drop-down list of values."""
        cells = self._raw

        validation = cells.getPropertyValue("Validation")

        validation.Type = uno.Enum(
            "com.sun.star.sheet.ValidationType",
            "LIST",
        )

        validation.Formula1 = ";".join(f'"{str(value)}"' for value in values)

        validation.ShowErrorMessage = True
        validation.ErrorTitle = "Invalid value"
        validation.ErrorMessage = "Choose a value from the list."
        validation.ShowList = 1

        cells.setPropertyValue("Validation", validation)
        return self

    def select(self):
        """
        Select the current cell range
        """
        controller = self.sheet.workbook.controller

        controller.setActiveSheet(self.sheet._raw())
        controller.select(self._raw)
        return self

    def set_left_border(self, color=0x00B050):
        """
        Set the left border of a cell range.
        """
        cells = self._raw

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

    def autofit_columns(self, min_width=2000):
        """Autofit columns in the range."""
        columns = self._raw.getColumns()

        columns.OptimalWidth = True

        for i in range(columns.getCount()):
            column = columns.getByIndex(i)

            if column.Width < min_width:
                column.Width = min_width

import os

from soffice import connect
from calc import Calc
import time

now = time.time()
ctx = connect()
print("Connected in", time.time() - now, "seconds")

now = time.time()
calc = Calc(ctx, "test.xlsm")

currentSheet = calc.doc.getCurrentController().getActiveSheet()

# Work with the spreadsheet:
calc.set_enum(
    currentSheet,
    "D1:D3",
    ["Option 1", "Option 2", "Option 3"],
)

calc.select(currentSheet, "D10:G13")

# Write to cell D10
currentSheet.getCellRangeByName("D10").setString("Hello, worldasdfasdfasdf!")

calc.rotate(currentSheet, "A1:Z1", 45)

# Write value to cell A10
currentSheet.getCellRangeByName("A10").setValue(42)

calc.autofit_columns(currentSheet, "A:Z")

calc.set_left_border(currentSheet, "B1:B10", color=0xFF0000)

calc.merge(currentSheet, "E1:F1")

# Now show it
calc.show()
print("Shown in", time.time() - now, "seconds")

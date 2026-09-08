# Table of Contents

* [workbook](#workbook)
* [sheet](#sheet)
  * [Sheet](#sheet.Sheet)
    * [uno](#sheet.Sheet.uno)
    * [cells](#sheet.Sheet.cells)
    * [show](#sheet.Sheet.show)
    * [select](#sheet.Sheet.select)
    * [set\_left\_border](#sheet.Sheet.set_left_border)
    * [delete](#sheet.Sheet.delete)
    * [freeze](#sheet.Sheet.freeze)
* [cells](#cells)
  * [Cells](#cells.Cells)
    * [\_\_init\_\_](#cells.Cells.__init__)
    * [uno](#cells.Cells.uno)
    * [set](#cells.Cells.set)
    * [rotate](#cells.Cells.rotate)
    * [merge](#cells.Cells.merge)
    * [unmerge](#cells.Cells.unmerge)
    * [set\_enum](#cells.Cells.set_enum)
    * [autofit\_columns](#cells.Cells.autofit_columns)
* [soffice](#soffice)

<a id="workbook"></a>

# workbook

<a id="sheet"></a>

# sheet

<a id="sheet.Sheet"></a>

## Sheet Objects

```python
class Sheet()
```

Wrapper around a LibreOffice Calc sheet.

<a id="sheet.Sheet.uno"></a>

#### uno

```python
def uno()
```

Get the raw uno class representation

<a id="sheet.Sheet.cells"></a>

#### cells

```python
def cells(cell_range)
```

Return a Cells object for the specified address.

<a id="sheet.Sheet.show"></a>

#### show

```python
def show()
```

Show the sheet for the user.

<a id="sheet.Sheet.select"></a>

#### select

```python
def select(cell)
```

Select a cell or range of cells in the sheet and make it active.

Show must be called after this to make the selection visible for the user.

<a id="sheet.Sheet.set_left_border"></a>

#### set\_left\_border

```python
def set_left_border(cell_range, color=0x00B050)
```

Set the left border of a cell range.

<a id="sheet.Sheet.delete"></a>

#### delete

```python
def delete()
```

Delete this current sheet from the Workbook

<a id="sheet.Sheet.freeze"></a>

#### freeze

```python
def freeze(row=0)
```

Freeze row so scrolling motion still shows this row

<a id="cells"></a>

# cells

<a id="cells.Cells"></a>

## Cells Objects

```python
class Cells()
```

Represents one or more cells in a Calc sheet.

<a id="cells.Cells.__init__"></a>

#### \_\_init\_\_

```python
def __init__(sheet, uno_range)
```

Create a Cells wrapper around a UNO cell range.

<a id="cells.Cells.uno"></a>

#### uno

```python
@property
def uno()
```

Return the raw uno representation of the cells

<a id="cells.Cells.set"></a>

#### set

```python
def set(value)
```

Set values to cell range, value will fill range if its a scalar

<a id="cells.Cells.rotate"></a>

#### rotate

```python
def rotate(degrees)
```

Rotate the text in a cell range by the given number of degrees.

<a id="cells.Cells.merge"></a>

#### merge

```python
def merge()
```

Clump togheter cells

<a id="cells.Cells.unmerge"></a>

#### unmerge

```python
def unmerge()
```

Unclump cells

<a id="cells.Cells.set_enum"></a>

#### set\_enum

```python
def set_enum(values)
```

Set a cell range to have a drop-down list of values.

<a id="cells.Cells.autofit_columns"></a>

#### autofit\_columns

```python
def autofit_columns(min_width=2000)
```

Autofit columns in the range.

<a id="soffice"></a>

# soffice


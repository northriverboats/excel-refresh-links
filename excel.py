#!/usr/bin/env python

from win32com.client import constants, Dispatch, DispatchEx, gencache
import pythoncom
import os

borderTop = 3
borderBottom = 4
borderLeft = 1
borderRight = 2
borderSolid = 1
borderDashed = 2
borderDotted = 3
colorBlack = 1
directionUp = -4162
directionDown = -4121
directionLeft = -4131
directionRight = -4152
xlshiftdown = -4121
xlshifttoright = -4161
xlformatfromleftorabove = 0  # row insert from above / column inserst from left
xlformatfromrightorbelow = 1 # row insert from below / column inserst from right
xlWhole = 1
xlPart = 2
xlValues = -4163
xlByRows = 1
xlByColumns = 2


class ExcelDocument(object):
  """
  Some convenience methods for Excel documents accessed
  through COM.
  """
  def __init__(self, visible=False):
    self.app = DispatchEx("Excel.Application")
    self.app.Visible = visible
    self.sheet = 1
    
  def new(self, filename=None):
    """
    Create a new Excel workbook. If 'filename' specified,
    use the file as a template.
    """
    self.app.Workbooks.Add(filename)
    
  def set_visible(self, visible):
    """
    Sheet Excel Visibility
    """
    self.app.Visible = visible
    
  def display_alerts(self, status):
    """
    Enable or Disable alert popups.
    """
    self.app.DisplayAlerts = status
    
  def open(self, filename, updatelinks=0):
    """
    Open an existing Excel workbook for editing.
    """
    self.app.Workbooks.Open(filename, UpdateLinks=updatelinks)
    
  def _get_app_(self):
    """
    Low Level function for debugging and dev purporses
    """
    return self.app

  def _get_function_(self):
    """
    Return the worksheet function object
    """
    return self.app.WorksheetFunction
    
  def get_sheet(self):
    """
    Get the active worksheet.
    """
    return self.app.ActiveWorkbook.Sheets(self.sheet)
    
  def set_sheet(self, sheet):
    """
    Set the active worksheet.
    """
    self.sheet = sheet
  
  def num_to_ref(self, row, col):
    """
	Convert 1 based column and row number into an alpha/num reference
	Example: column 89, row 90 yields "CK90"
    """
    output = str(row)
    while True:
      col -= 1
      output = chr(65 + col % 26) + output
      col = int(col / 26)
      if col == 0:
        break
    return output    
    
  def get_range(self, range):
    """
    Get a range object for the specified range or single cell.
    """
    return self.app.ActiveWorkbook.Sheets(self.sheet).Range(range)

  def get_cell(self, row, col):
    """
    Get a single cell range for the specified cell
    """
    return self.app.ActiveWorkbook.Sheets(self.sheet).Cells(row, col)
    
  def set_value(self, cell, value=''):
    """
    Set the value of 'cell' to 'value'.
    """
    self.get_range(cell).Value = value
    
  def get_value(self, cell):
    """
    Get the value of 'cell'.
    """
    value = self.get_range(cell).Value
    if isinstance(value, tuple):
      value = [v[0] for v in value]
    return value
    
  def set_border(self, range, side, line_style=borderSolid, color=colorBlack):
    """
    Set a border on the specified range of cells or single cell.
    'range' = range of cells or single cell
    'side' = one of borderTop, borderBottom, borderLeft, borderRight
    'line_style' = one of borderSolid, borderDashed, borderDotted, others?
    'color' = one of colorBlack, others?
    """
    range = self.get_range(range).Borders(side)
    range.LineStyle = line_style
    range.Color = color
    
  def sort(self, range, key_cell):
    """
    Sort the specified 'range' of the activeworksheet by the
    specified 'key_cell'.
    """
    range.Sort(Key1=self.get_range(key_cell), Order1=1, Header=0, OrderCustom=1, MatchCase=False, Orientation=1)
    
  def hide_row(self, row, hide=True):
    """
    Hide the specified 'row'.
    Specify hide=False to show the row.
    """
    self.get_range('a%s' % row).EntireRow.Hidden = hide
    
  def hide_column(self, column, hide=True):
    """
    Hide the specified 'column'.
    Specify hide=False to show the column.
    """
    self.get_range('%s1' % column).EntireColumn.Hidden = hide
    
  def delete_row(self, row, shift=directionUp):
    """
    Delete the entire 'row'.
    """
    self.get_range('a%s' % row).EntireRow.Delete(Shift=shift)
    
  def delete_column(self, column, shift=directionLeft):
    """
    Delete the entire 'column'.
    """
    self.get_range('%s1' % column).EntireColumn.Delete(Shift=shift)
    
  def fit_column(self, column):
    """
    Resize the specified 'column' to fit all its contents.
    """
    self.get_range('%s1' % column).EntireColumn.AutoFit()

  def goto(self, cell):
    """
    Goto the specified cell
    """
    range = self.get_range(cell)
    range.Activate()

  def get_last_row(self):
    """
    """
    return self.app.ActiveWorkbook.Sheets(self.sheet).UsedRange.Rows.Count

  def get_last_column(self):
    """
    """
    return self.app.ActiveWorkbook.Sheets(self.sheet).UsedRange.Columns.Count

  def get_countif(self, range, text):
    """
    Return count of string found in range
    """
    return self.app.WorksheetFunction.CountIf(range, text)

  def save(self):
    """
    Save the active workbook.
    """
    self.app.ActiveWorkbook.Save()
    
  def save_as(self, filename, delete_existing=False):
    """
    Save the active workbook as a different filename.
    If 'delete_existing' is specified and the file already
    exists, it will be deleted before saving.
    """
    if delete_existing and os.path.exists(filename):
      os.remove(filename)
    self.app.ActiveWorkbook.SaveAs(filename)
    
  def print_out(self):
    """
    Print the active workbook.
    """
    self.app.Application.PrintOut()
    
  def close(self):
    """
    Close the active workbook.
    """
    self.app.ActiveWorkbook.Close()
    
  def quit(self):
    """
    Quit Excel.
    """
    return self.app.Quit()

if __name__ == "__main__":
    print ("Tests go here...")
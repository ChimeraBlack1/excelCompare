import math
import xlrd
import xlwt

def OpenSheet(wb):
  """
  opens a workbook and returns the first sheet
  """
  workbook = xlrd.open_workbook(wb)
  sheet = workbook.sheet_by_index(0)
  return sheet 

def GetValue(sheet, row, col):
  serial = sheet.cell_value(row,col)
  return serial

def FindLastPopulatedRow(wb, row=0, col=0):
  """
  Find the number of populated rows in an excel workbook
  """  
  workbook = xlrd.open_workbook(wb)
  sheet = workbook.sheet_by_index(0)
  content = sheet.cell_value(row, col)
  rowCount = 0

  while content != "":
    try:
      content = sheet.cell_value(row + rowCount, col)
    except:
      rowCount = rowCount +1
      return rowCount
    rowCount = rowCount + 1
  return rowCount

def GetStatusDetails(wb, row=0, col=0):
  """
  Collect Status Details from Manager Workbook
  """
  workbook = xlrd.open_workbook(wb)
  sheet = workbook.sheet_by_index(0)
  status = sheet.cell_value(row, col)
  notes = sheet.cell_value(row, col + 1)
  renewalDate = sheet.cell_value(row, col + 2)
  acctStatus = {
    "status": status,
    "notes": notes,
    "renewalDate": renewalDate,
  }
  return acctStatus
import math
import xlrd
import xlwt

def CheckSerial(wb1, wb2, col1=0, col2=0, row1=0, row2=0):
  """
  open two workbooks, set a variable for each column + row position. 
  compare the two and return a boolean.
  """  
  # assign workbooks and sheets
  wb1 = xlrd.open_workbook(wb1)
  wb2 = xlrd.open_workbook(wb2)
  sheet1 = wb1.sheet_by_index(0)
  sheet2 = wb2.sheet_by_index(0)

  serial1 = sheet1.cell_value(row1, col1)
  serial2 = sheet2.cell_value(row2, col2)

  if serial1 == serial2:
    return True
  else:
    return False

def FindLastPopulatedRow(wb, row=0, col=0):
  workbook = xlrd.open_workbook(wb)
  sheet = workbook.sheet_by_index(0)
  content = sheet.cell_value(row, col)
  rowCount = 0
  contentList = []

  while content != "":
    try:
      content = sheet.cell_value(row + rowCount, col)
      contentList.append(content)
      print(str(len(contentList)) + " - " + str(content))
    except:
      rowCount = rowCount +1
      return rowCount
    rowCount = rowCount + 1
  
  return rowCount, contentList


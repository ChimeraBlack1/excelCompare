import excelCompare

wb1 = "test.xlsm"
wb2 = "test2.xlsx"
col1 = 10
col2 = 12
row1 = 1
row2 = 1

torF = excelCompare.CheckSerial(wb1, wb2, col1, col2, row1, row2)

print(str(torF))

rowCount = excelCompare.FindLastPopulatedRow(wb1, row1, col1)
print(str(rowCount))

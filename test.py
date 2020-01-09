import excelCompare as ec

wb1 = "Lease Expiration Tracking Sheet Jan 2020.xlsm"
wb2 = "Team Ron Updates (1-7-2020).xlsx"
col1 = 10
col2 = 10
row1 = 1
row2 = 1

rowCount = ec.FindLastPopulatedRow(wb1, row1)
rowCount2 = ec.FindLastPopulatedRow(wb2, row2)

sheet1 = ec.OpenSheet(wb1)
sheet2 = ec.OpenSheet(wb2)

for i in range(row1, rowCount):
  serial1 = ec.GetValue(sheet1, row1, col1)
  for x in range(row1, rowCount2):
    serial2 = ec.GetValue(sheet2, row2, col2)
    if serial1 == serial2:
      
    else:
      continue
    write status, notes, renewalDate

print("rowCount " + str(rowCount))
print("rowCount2 " + str(rowCount2))
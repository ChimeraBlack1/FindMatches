from findMatches import FindMatchesInTwoBooks, WriteToWorkbook, SaveWorkbook, CreateWorkbook
import math
import xlrd
import xlwt

# rowStart, rowStart2, col1, col2, End1, End2, Sheet1, Sheet2
wb1 = xlrd.open_workbook("WellsPortfolio(dec2019).xlsx")
wb2 = xlrd.open_workbook("NewLeasesFound.xls")
rowStart = 1
rowStart2 = 1
col1 = 4
col2 = 0
End1 = 2206
End2 = 75
Sheet1 = wb1.sheet_by_index(0)
Sheet2 = wb2.sheet_by_index(0)

FindMatchesInTwoBooks(rowStart, rowStart2, col1, col2, End1, End2, Sheet1, Sheet2)
mywb = CreateWorkbook("newWb")
SaveWorkbook(mywb[1])
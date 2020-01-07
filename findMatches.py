import math
import xlrd
import xlwt

def SaveWorkbook(workbook, newWorkBookName="Matches.xls"):
  workbook.save(newWorkBookName)
  print("saved: " + str(newWorkBookName))

def CreateWorkbook(newWorkBookName="Matches.xls"):
  workbook = xlwt.Workbook()
  worksheet = workbook.add_sheet('Matches')
  return worksheet, workbook

def WriteToWorkbook(worksheet, workbook):
  worksheet.write()

def FindMatchesInTwoBooks(rowStart, rowStart2, col1, col2, End1, End2, Sheet1, Sheet2):
  matches = []

  for x in range(rowStart, End1):
    serial = Sheet1.cell_value(x,col1)
    for y in range(rowStart2, End2):
      serialToTest = Sheet2.cell_value(y,col2)
      if serial == serialToTest:
        matches.append(serial)

  print(str(len(matches)))
  



# PREVIOUS Month's Report
# goodFile = False

# while goodFile == False:
#   fileToRead = input("Please enter the name of the PREVIOUS month's report)> ")
#   if fileToRead == "exit" or fileToRead == "quit":
#     print("ok, bye!")
#     exit()
#   else:
#     prevReport = fileToRead + ".xlsm"
#     try:
#       prevwb = xlrd.open_workbook(prevReport)
#       prevSheet = prevwb.sheet_by_index(0)
#       goodFile = True
#     except:
#       print("I can't find that file, try again...")
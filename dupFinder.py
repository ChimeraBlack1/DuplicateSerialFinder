import math
import xlrd
import xlwt

# Get Report
goodFile = False

while goodFile == False:
  fileToRead = input("Please enter the name of the report in question)> ")
  if fileToRead == "exit" or fileToRead == "quit":
    print("ok, bye!")
    exit()
  else:
    excelFile = fileToRead + ".xlsm"
    try:
      wb = xlrd.open_workbook(excelFile)
      thisSheet = wb.sheet_by_index(0)
      goodFile = True
    except:
      print("I can't find that file, try again...")

# Sheet row count
sheetNumberGood = False

while sheetNumberGood == False:
  endofSheet = input("How many cells are in the sheet?)> ")
  if endofSheet == "exit" or endofSheet == "quit":
    print("ok, bye!")
    exit()
  else:
    try:
      endofSheet = int(endofSheet)
      sheetNumberGood = True
    except:
      print("Not the right number, try again...)>")

sample = []
array1 = []
duplicates = []
sampleSize = 100

# for i in range(0,sampleSize):
# serial = prevSheet.cell_value(x,SherpaReportSerialCol)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Duplicates')
NewWorkbookName = "Duplicates.xls"

workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))
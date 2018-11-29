import xlrd
import xlwt
import os
import sys

cwd = os.getcwd()

# idxSheet = int(sys.argv[1])
# idxColumn = int(sys.argv[2])

idxWr = 0
outputFilename = 'merged_file.xls'
outFile = xlwt.Workbook('UFT-8')
sheet0 = outFile.add_sheet('Sheet1', False)

listFiles = os.listdir(cwd) # list out all files under working folder
for xls in listFiles:
    # print(xls)
    filename, fileext = os.path.splitext(xls)
    # print(filename)
    # print(fileext)
    if (fileext == '.xls' or fileext == '.xlsx'):
        workbook = xlrd.open_workbook(xls) # open excel file
        sheet = workbook.get_sheet(0) # get first sheet
        col = sheet.col_values(0) # get first column

        for i in range(0, len(col)):
            sheet0.write(0, idxWr, col[i])
            idxWr += 1

if (idxWr > 0):
    outFile.save(outputFilename)
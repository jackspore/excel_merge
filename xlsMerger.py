import xlrd
import xlwt
import os
import sys

# cwd = os.getcwd()
# idxSheet = int(sys.argv[1])
# idxColumn = int(sys.argv[2])
wdir = input('Folder path holds excel files:')
idxSheet = int(input('Number of sheet to read from each file:')) -1
idxColumn = int(input('Number of column to read from each sheet:')) -1

rowsWr = 0 # how many rows written
outputFilename = wdir + '\\merged_file.xls'
outFile = xlwt.Workbook('UFT-8')
sheet0 = outFile.add_sheet('Sheet1', False) # add a new sheet1 into output file

listFiles = os.listdir(wdir) # list out all files under working folder
for fl in listFiles:
    print('processing file',fl, '...')
    filename, fileext = os.path.splitext(fl)

    if (fileext == '.xls' or fileext == '.xlsx'):
        fp = os.path.join(wdir, fl)
        workbook = xlrd.open_workbook(fp) # open excel file
        sheet = workbook.sheet_by_index(idxSheet) # get target sheet
        col = sheet.col_values(idxColumn) # get target column

        for i in range(0, len(col)):
            sheet0.write(r=rowsWr, c=0, label=col[i]) # write data into 1st column
            rowsWr += 1

if (rowsWr > 0):
    outFile.save(outputFilename)
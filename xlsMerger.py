import xlrd
import xlwt
import os
import sys
import  time 
from  datetime  import  *  

wdir = input('Folder path holds excel files:')

rowsWr = 0 # how many rows written
outputFilename = os.path.join(wdir, 'merged_file.xls')
outFile = xlwt.Workbook('UFT-8')
sheet0 = outFile.add_sheet('Sheet1', False) # add a new sheet1 into output file

listFiles = os.listdir(wdir) # list out all files under working folder
for fl in listFiles:
    print('processing file',fl, '...')
    filename, fileext = os.path.splitext(fl)

    if (fileext == '.xls' or fileext == '.xlsx'):
        fp = os.path.join(wdir, fl)
        workbook = xlrd.open_workbook(fp) # open excel file
        sheet = workbook.sheet_by_index(0) # get target sheet

        # get loanee name
        name = sheet.cell(0,0).value

        # get number of months of loan
        numRow = int(sheet.cell(1,7).value)

        # get target columns
        col2 = sheet.col_values(2) # 3rd column in input sheet, interest
        col6 = sheet.col_values(6) # 7th column in input sheet, date

        rowsRd = 0
        for i in range(3, 3 + numRow):
            sheet0.write(r=rowsWr, c=0, label=name)
            
            interestValue = sheet.cell(3+rowsRd, 2).value
            sheet0.write(r=rowsWr, c=1, label=round(interestValue, 2))
            
            if (sheet.cell(3+rowsRd, 6).ctype == 3): # cell value is date type
                dateValue = xlrd.xldate_as_tuple(sheet.cell(3+rowsRd, 6).value, workbook.datemode)
                dateString = date(*dateValue[:3]).strftime('%Y/%m/%d')
            else:
                dateString = sheet.cell(3+rowsRd, 6).value
            sheet0.write(r=rowsWr, c=2, label=dateString)

            rowsWr += 1
            rowsRd += 1

if (rowsWr > 0):
    print('Process complete, save output file as merged_file.xls')
    outFile.save(outputFilename)
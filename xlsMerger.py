# author: jackspore
# ref: https://github.com/jackspore/excel_merge.git
# 
# This source code is under DO WHATEVER THE FUCK YOU LIKE license
# You can use, distribute, modify copies of this file with or
# without author's consent

import xlrd
import xlwt
import os
import sys
import  time 
from  datetime  import  date 
from ExcelUtil import *
import re

wdir = input('Folder path holds excel files:')
outputFilename = os.path.join(wdir, 'merged_file.xls')
if os.path.isfile(outputFilename):
    os.remove(outputFilename) # remove output file if already exist

outFile = xlwt.Workbook('UFT-8')
sheet0 = outFile.add_sheet('Sheet1', False) # add a new sheet1 into output file

rowsWr = 1 # how many rows written
sheet0.write(r=0, c=0, label='借款人', style=titleStyle)
sheet0.write(r=0, c=1, label='还贷方式', style=titleStyle)
sheet0.write(r=0, c=2, label='当期应还利息', style=titleStyle)
sheet0.write(r=0, c=3, label='还款日', style=titleStyle)

listFiles = os.listdir(wdir) # list out all files under working folder
for fl in listFiles:
    print('processing file',fl, '...')
    filename, fileext = os.path.splitext(fl)

    if (fileext == '.xls' or fileext == '.xlsx'):
        fp = os.path.join(wdir, fl)
        workbook = xlrd.open_workbook(fp) # open excel file
        sheet = workbook.sheet_by_index(0) # get target sheet

        # get loanee name
        nameStr = sheet.cell(0,0).value
        searchObj = re.search(u'(.*)贷款还款明细表（(.*)）', nameStr)
        name = searchObj.group(1)
        loanType = searchObj.group(2)

        # get number of months of loan
        numRow = int(sheet.cell(1,7).value)

        # get target columns
        col2 = sheet.col_values(2) # 3rd column in input sheet, interest
        col6 = sheet.col_values(6) # 7th column in input sheet, date

        rowsRd = 0
        for i in range(3, 3 + numRow):
            sheet0.write(r=rowsWr, c=0, label=name)
            sheet0.write(r=rowsWr, c=1, label=loanType)
            
            interestValue = sheet.cell(3+rowsRd, 2).value
            sheet0.write(r=rowsWr, c=2, label=round(interestValue, 2), style=rmbStyle)
            
            if (sheet.cell(3+rowsRd, 6).ctype == 3): # cell value is date type
                dateValue = xlrd.xldate_as_tuple(sheet.cell(3+rowsRd, 6).value, workbook.datemode)
                dateString = toDateCellStr(dateValue)
            else:
                dateString = sheet.cell(3+rowsRd, 6).value
            sheet0.write(r=rowsWr, c=3, label=dateString)

            rowsWr += 1
            rowsRd += 1

if (rowsWr > 1):
    print('Process complete, save output file as merged_file.xls')
    outFile.save(outputFilename)
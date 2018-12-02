# Utility functions of Excel spreadsheet operations

import xlrd
import xlwt
from  datetime  import  date 

# cell write style for RMB currency
rmbStyle = xlwt.Style.easyxf(num_format_str='￥#,##0.00')

# cell write style for column title
titleStyle = xlwt.Style.easyxf(strg_to_parse='font: bold on')

def toDateCellStr(_dateValue):
    return date(*_dateValue[:3]).strftime('%Y/%m/%d')
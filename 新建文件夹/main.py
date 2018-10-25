
# 这是一个操作excel的类

import xlrd
import xlwt
from xlutils.copy import copy

def set_style(name,height,bold=False):  
  style = xlwt.XFStyle() # 初始化样式  
  
  font = xlwt.Font() # 为样式创建字体  
  font.name = name # 'Times New Roman'  
  font.bold = bold  
  font.color_index = 4  
  font.height = height  
  
  # borders= xlwt.Borders()  
  # borders.left= 6  
  # borders.right= 6  
  # borders.top= 6  
  # borders.bottom= 6  
  
  style.font = font  
  # style.borders = borders  
  
  return style  
  
# 打开文件  
workbook = xlrd.open_workbook("G:/huangwen/code/excel/text_1.xlsx")  
# 获取所有sheet  
print(workbook.sheet_names())

# 根据sheet索引或者名称获取sheet内容  
sheet = workbook.sheet_by_index(0) # sheet索引从0开始
  
# sheet的名称，行数，列数
print(sheet.name,sheet.nrows,sheet.ncols)

rows = sheet.row_values(3) # 获取第四行内容
cols = sheet.col_values(1) # 获取第三列内容

# 遍历全部内容
for row in range(0, sheet.nrows):
    for col in range(0, sheet.ncols):
        # 找到sum就做求和
        if sheet.cell_value(row,col) == "sum":
            rows = sheet.row_values(row + 1)
            sum = 0
            for index in range(0, col):
                sum += sheet.cell_value(row + 1, index)

            workbooknew = copy(workbook)
            ws = workbooknew.get_sheet(0)
            ws.write(row+1, col+1, sum)
            workbooknew.save(u'copy.xls')
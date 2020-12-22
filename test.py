#!/usr/bin/env python 
# -*- coding:utf-8 -*-
# import pandas as pd
from pathlib import Path

import win32com.client as wc
#启动Excel应用
excel_app = wc.Dispatch('Excel.Application')
excel_app.Visible = False
#连接excel
workbook = excel_app.Workbooks.Open((Path.cwd()) / 'Lib' / '物料库.xlsx' )
mySheet = workbook.Worksheets(1)
#写入数据
# workbook.Worksheets('Sheet1').Cells(1,1).Value = 'data'
print(mySheet.Cells(3,1).Value)
#关闭并保存
LastRow = mySheet.usedrange.columns.count
print("该sheet页目前已经存在", LastRow, "行")
# workbook.SaveAs('newexcel.xlsx')
excel_app.Application.Quit()
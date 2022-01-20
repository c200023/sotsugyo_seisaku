import pandas as pd
import openpyxl
from openpyxl import Workbook
import win32com.client as win32
import os
from openpyxl.utils.dataframe import dataframe_to_rows
win32c = win32.constants

Excel = win32.gencache.EnsureDispatch("Excel.Application")
Excel.Visible = True

# ソースファイルフォルダへ移動
os.chdir('../csv/')

wb = Workbook()
ws = wb.active
ws.title = 'Sheet1'
df = pd.read_csv('test_file.csv', encoding='utf-8')

df['date'] = pd.to_datetime(df['date'])

for row in dataframe_to_rows(df, index=None, header=True):
    ws.append(row)

# 保存フォルダへ移動
os.chdir('../output/')

# 仮保存
wb.save('save_test_file.xlsx')

# ブックを開く
fpath = os.path.join(os.getcwd(), 'save_test_file.xlsx')
wb = Excel.Workbooks.Open(fpath)

# Sheet 1 指定し､フィルターを有効にする
wbs0 = wb.Sheets('Sheet1')
wbs0.Columns.AutoFilter(1)

##################
# ピボットテーブルの作成1
wbs1_name = 'pivot1'
wb.Sheets.Add().Name = wbs1_name

wbs1 = wb.Sheets(wbs1_name)
pvt_name = 'pvt'
pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=wbs0.UsedRange)
pc.CreatePivotTable(TableDestination='{sheet}!R1C1'.format(
    sheet=wbs1_name), TableName=pvt_name)

# 行
wbs1.PivotTables(pvt_name).PivotFields(
    'shop').Orientation = win32c.xlRowField

# データ
wbs1.PivotTables(pvt_name).AddDataField(wbs1.PivotTables(pvt_name).PivotFields(
    'amount'), 'Sum/amount', win32c.xlSum).NumberFormat = '0'

#############
# ピボットテーブルの作成2
wbs2_name = 'pivot2'
wb.Sheets.Add().Name = wbs2_name

wbs2 = wb.Sheets(wbs2_name)
pvt_name = 'pvt'
pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=wbs0.UsedRange)
pc.CreatePivotTable(TableDestination='{sheet}!R1C1'.format(
    sheet=wbs2_name), TableName=pvt_name)

# 行
wbs2.PivotTables(pvt_name).PivotFields(
    'date').Orientation = win32c.xlRowField

# データ
wbs2.PivotTables(pvt_name).AddDataField(wbs2.PivotTables(pvt_name).PivotFields(
    'amount'), 'Sum/amount', win32c.xlSum).NumberFormat = '0'
wbs2.Cells(2, 1).Select()
Excel.Selection.Group(Start=True, End=True, Periods=(
    False, False, False, False, False, False, True))


##################
# ピボットテーブルの作成3
wbs3_name = 'pivot3'
wb.Sheets.Add().Name = wbs3_name

wbs3 = wb.Sheets(wbs3_name)
pvt_name = 'pvt'
pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=wbs0.UsedRange)
pc.CreatePivotTable(TableDestination='{sheet}!R1C1'.format(
    sheet=wbs3_name), TableName=pvt_name)

# 行
wbs3.PivotTables(pvt_name).PivotFields(
    'product').Orientation = win32c.xlRowField

# データ
wbs3.PivotTables(pvt_name).AddDataField(wbs3.PivotTables(pvt_name).PivotFields(
    'amount'), 'Sum/amount', win32c.xlSum).NumberFormat = '0'


# VBA呼び出し
# マクロフォルダへ移動
os.chdir('../macro/')

# マクロ実行
macro_filename = 'macro1.xlsm'
macro_fullpath = os.path.join(os.getcwd(), macro_filename)
macro_wb = Excel.Workbooks.Open(Filename=macro_fullpath)
Excel.Application.Run(macro_filename + '!macro_format')
macro_wb.Close()

# 保存フォルダへ移動
os.chdir('../output/')

# 上書き保存
wb.Save()

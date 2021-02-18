# ref : https://qiita.com/yuta-38/items/337059e1eafab3582851

import win32com.client
from pathlib import Path
import locale
import datetime as dt

# *************** Extract diary from Excel *************** #
locale.setlocale(locale.LC_ALL, '')
abspath = str(Path(r"./日報.xlsx").resolve())

excel = win32com.client.Dispatch("Excel.Application")
workbook = excel.Workbooks.Open(abspath, UpdateLinks=0, ReadOnly=True)
worksheet = workbook.Worksheets("Sheet1")

# 検索する文字列(日付)を作成
today = dt.datetime.today()
# %#m, %#dで0埋めしないで出力できる
text_today = today.strftime("%Y/%#m/%#d")
print(text_today)

for row_count in range(1, 16):
    if text_today in worksheet.Cells.Item(row_count, 1).Text:
        print("日付が見つかりました")
        print(str(row_count))
        print(worksheet.Cells.Item(row_count, 1).Value)
        print(worksheet.Cells.Item(row_count, 1).Text)
        print(worksheet.Cells.Item(row_count, 1).Formula)
        print(worksheet.Cells.Item(row_count, 1).Address)

        splitted = worksheet.Cells.Item(row_count, 1).Address.split('$')
        rng = "{}:{}".format(worksheet.Cells.Item(row_count, 1).Address, '$J$' + str(int(splitted[2]) + 10))
        print(rng)


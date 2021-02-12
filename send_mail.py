import win32com.client
from pathlib import Path
import locale
import date2text
import extract_diary

# *************** Extract diary from Excel *************** #
locale.setlocale(locale.LC_ALL, '')
abspath = str(Path(r"./日報.xlsx").resolve())

excel = win32com.client.Dispatch("Excel.Application")
workbook = excel.Workbooks.Open(abspath, UpdateLinks=0, ReadOnly=True)
worksheet = workbook.Worksheets("Sheet1")

rng = "A3:I10"
diary = worksheet.Range(rng)
table = list()

for row_cnt in range(1, diary.Rows.Count + 1):
    row = []
    for col_cnt in range(1, diary.Columns.Count + 1):
        # row.append(diary(row_cnt, col_cnt).Value)
        row.append(diary(row_cnt, col_cnt).Text)
    table.append(row)

# *************** Make diary mail *************** #
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

textData = date2text.getTextData()
kanri_tasks = extract_diary.get_kanri_list(table)
other_tasks = extract_diary.get_other_list(table)

# Text : 1 / HTML : 2 / RichText : 3
mail.BodyFormat = 3
mail.To = "shin5884@softbank.ne.jp"
# mail.cc = "xxx@yyy.com"
# mail.Bcc = "xxx@yyy.com"
mail.Subject = "【報告】日報(" + textData + " 小林)"
mail.Body = "Aさん\n" \
            + "CC：ADメンバ\n\n" \
            + "お疲れ様です、小林です。\n\n" \
            + "本日(" + textData + ")の日報を送付いたします。\n" \
            + "*******************************************************\n" \
            + "【作業実績】合計実働時間[]\n\n" \
            + "■管理[]\n" \
            + kanri_tasks + "\n" \
            + other_tasks + "\n" \
            + "【作業予定】\n\n" \
            + "【不在予定】\n\n" \
            + "【出社予定】\n\n" \
            + "【問題点・懸念点・不明点など】\n\n" \
            + "【月末時点の勤務時間見込】\n\n" \
            + "【標準勤務時間】\n\n" \
            + "*******************************************************\n" \
            + "以上になります。"

# mail.Attachments.Add("C:/Users/user/pic.jpg")
mail.Display(True)
# mail.Send()

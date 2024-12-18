import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Border, Side

settlement_file = r"C:\Users\lenovo02\Desktop\自動化\精算書作成\精算書.xlsx"
sheet = "【請求書】(原本)"

input_file = r"C:\Users\lenovo02\Desktop\自動化\精算書作成\2024-12-ライブ結果.xlsx"
sheet_name = "原本"

comp_name = "JTL"

df = pd.read_excel(input_file, sheet_name)
df = df.dropna(subset=['出品者'])  # '出品業者'列のNaNを削除
df = df.dropna(subset=['購入者'])  # '購入者'列のNaNを削除

df = df[df['購入者'] == comp_name] # 精算書(購入)を作成する業者でフィルターをかける

df['通し番号'] = df['通し番号'].apply(lambda x: int(x) if x.is_integer() else x)
df['販売価格（税込）'] = df['販売価格（税込）'].apply(lambda x: int(x) if x.is_integer() else x)


# '出品者'と'通し番号'を結合して新しい列を作成
df['出品者+通し番号'] = df['出品者'].astype(str) + df['通し番号'].astype(str)

# '出品者'と'通し番号'列を削除
df = df.drop(columns=['出品者', '通し番号'])

# 必要な列だけを選択
df = df[['出品者+通し番号', '販売価格（税込）']]

wb = load_workbook(settlement_file)

# 新しいシートに表を追加
new_sheet_name = '集計表'
if new_sheet_name not in wb.sheetnames:
    ws = wb.create_sheet(new_sheet_name)
else:
    ws = wb[new_sheet_name]

# 最下行を取得
last_row = ws.max_row  # 最後の行の番号を取得

# 通貨書式を設定
currency_style = NamedStyle(name="currency_style", number_format='"¥"#,##0')

# 罫線の設定
border_style = Border(
    left=Side(style='medium'),
    right=Side(style='medium'),
    top=Side(style='medium'),
    bottom=Side(style='medium')
)

# 新しいデータを最下行の次の行に追加
for row_num, row in enumerate(df.values, start=last_row):  # 最下行の次の行からデータを追加
    for col_num, cell_value in enumerate(row, start=1):  # 各列にデータを挿入
        if col_num == 1:
            ws.cell(row=row_num, column=1, value=row_num-1)
            ws.cell(row=row_num, column=1).border = border_style
            ws.cell(row=row_num, column=2, value=cell_value)
        else :
            ws.cell(row=row_num, column=col_num+1, value=cell_value)

        # 販売価格列に通貨書式を適用
        ws.cell(row=row_num, column=3).style = currency_style 

        # 罫線を適用
        ws.cell(row=row_num, column=col_num+1).border = border_style



# 2x2の新しい表を追加
start_row = last_row + len(df)  # 新しい表の開始行（既存の表の下から2行下）

# AB列を結合
ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=2)
ws.merge_cells(start_row=start_row+1, start_column=1, end_row=start_row+1, end_column=2)


# "販売価格合計(税込)"を左上のセルに入力
ws.cell(row=start_row, column=1, value="販売価格合計(税込)")
# 合計の数式を3列目（column=3）に入力
ws.cell(row=start_row, column=3, value="=SUM(C{}:C{})".format(last_row, last_row + len(df) - 1))  # 合計の数式

# "内消費税金額(10%)"を左上のセルに入力
ws.cell(row=start_row + 1, column=1, value="内消費税金額(10%)")
# 消費税の計算式を3列目（column=3）に入力
ws.cell(row=start_row + 1, column=3, value="=C{}*10/110".format(start_row))  # 消費税の計算式

ws.cell(row=row_num, column=3).style = currency_style

# 2x2表に罫線を適用
for row in range(start_row, start_row + 2):
    ws.cell(row=row, column=3).style = currency_style
    for col in range(1, 4): 
        cell = ws.cell(row=row, column=col)
        cell.border = border_style

# 保存
output_file = r"C:\Users\lenovo02\Desktop\自動化\精算書作成\集計結果.xlsx"  # 出力先のファイル名
wb.save(output_file)

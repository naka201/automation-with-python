import pandas as pd
import pyodbc
import pprint
import datetime
import os

def fetch_filtered_data(db_file, start, end, output_dir):
    # コネクションの設定
    connection_string = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_file + ';'
    connection = pyodbc.connect(connection_string)

    # データの取得
    query = 'SELECT * FROM 箱番'
    df = pd.read_sql(query, connection)

    # データのフィルタリング処理
    df = df.iloc[:, 1:]  # 最初のカラムを削除
    df['業者コード'] = df['業者コード'].astype('Int64')  # 型変換

    # 欠番処理
    make_output_excel(df[df['欠番'] == True], output_dir, "欠番リスト.xlsx")
    df = df[df['欠番'] != True]  # 欠番チェック

    # 検品済み処理
    df = df[df['検品済み'] == True]  # 検品済みのみ
    make_output_excel(df, output_dir, "検品済みリスト.xlsx")

    df = df.drop(columns=['欠番'])  # 欠番カラムの削除
    df = df.dropna(subset=['タイムスタンプ'])  # タイムスタンプがNaNの行を削除
    df['タイムスタンプ'] = pd.to_datetime(df['タイムスタンプ'])  # タイムスタンプの型変換

    # 指定した時間内の行をフィルタリング
    start_time = pd.Timestamp(start)
    end_time = pd.Timestamp(end)
    df = df[(df['タイムスタンプ'] >= start_time) & (df['タイムスタンプ'] <= end_time)]
    make_output_excel(df, output_dir, "チェック分リスト.xlsx")

    return df

def make_output_excel(df, output_dir, output_filename):
    df.to_excel(f"{output_dir}/{output_filename}", index=False, engine='openpyxl')


def ensure_output_directory(output_dir):
    """
    出力ディレクトリが存在しない場合は作成する
    """
    os.makedirs(output_dir, exist_ok=True)

def make_check_set(df, check):
    check = set()
    for _, row in df.iterrows():
        num = row["箱番"]
        check.add(num)

    #print(check)
    return check
    

def excel_to_set(file_path, box):
    df = pd.read_excel(file_path, engine='openpyxl')

    # E2から値のある範囲を取得
    # E2はインデックスとして1, 4に相当（0始まりのため）
    start_row = 1  # E2の行
    start_col = 4  # E列（0始まり）

    # 値のある範囲を取得
    # 最初に行を取得
    values = df.iloc[start_row:, start_col]

    # 空でない値をセットに追加
    box = set(int(v) for v in values.dropna() if pd.notna(v))

    # デバッグ用
    #print(box)

    return box

def main():
    file_path = r"C:\Users\lenovo02\Desktop\自動化\検品チェック\入力済み.xlsx"
    db_file = r"C:\Users\lenovo02\Desktop\自動化\検品チェック\database.accdb"

    ##################################
    ############ 変更部分 ############ 
    start_time = '2024-11-16 17:00:00'
    end_time = '2024-11-22 12:00:00'
    ##################################
    ##################################

    dt_now = datetime.datetime.now()
    def_year = dt_now.strftime("%Y")
    def_month = dt_now.strftime("%m")  # 2桁の月
    def_day = dt_now.strftime("%d")     # 2桁の日

    print(f"{def_year}-{def_month}-{def_day}のチェックを開始します。")

    output_dir = f"検品チェック/チェック結果/{def_year}-{def_month}-{def_day}"
    # 出力ディレクトリを確認・作成
    ensure_output_directory(output_dir)

    df = fetch_filtered_data(db_file, start_time, end_time, output_dir)

    box_num = set()
    box_num = excel_to_set(file_path, box_num)

    check = set()
    check = make_check_set(df, check)

    result = set()
    for c in check:
        if c not in box_num:
            result.add(c)
    make_output_excel(df[df['箱番'].isin(result)], output_dir, "未入力リスト.xlsx")

    result = sorted(result)

    print("チェックが完了しました。")

    #print(result)


if __name__ == "__main__":
    main()

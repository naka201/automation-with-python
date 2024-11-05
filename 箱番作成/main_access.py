import random
import pyodbc
import pandas as pd
import pprint

def sort_key(item):
    company, box_cnt, start, end = item
    if pd.isna(start) and pd.isna(end):
        return (float('inf'),)  # (float('inf'),)で最後にする
    return (end - start, -box_cnt)

def sort_list(comp_list):
    return sorted(comp_list, key=sort_key)

def change_end(comp_list):
    total = sum(company[1] for company in comp_list) + 200
    updated_list = []
    
    for company, box_cnt, start, end in comp_list:
        if not pd.isna(start) and pd.isna(end):
            end = total  # end を更新
        updated_list.append((company, box_cnt, start, end))
    
    return updated_list

def random_num(companies):
    total = sum(company[1] for company in companies) + 200
    all_num = list(range(201, total + 1))
    assigned_num = {}
    reserved_num = set()

    for company, box_cnt, box_start, box_end in companies:
        if not pd.isna(box_start):
            box_range = list(range(box_start, box_end + 1))
            box_lim = box_end - box_start + 1

            if box_lim < box_cnt:
                raise ValueError(f"{company}の範囲内の箱数が不足しています。必要：{box_cnt}個, 範囲内：{box_lim}個")

            available_num = [num for num in box_range if num not in reserved_num]
            if len(available_num) < box_cnt:
                raise ValueError(f"{company}\n必要：{box_cnt}個, 範囲内：{len(available_num)}個")

            random.shuffle(available_num)
            assigned_num[company] = available_num[:box_cnt]
            reserved_num.update(assigned_num[company])
        else:
            assigned_num[company] = []
        
    remain_num = [num for num in all_num if num not in reserved_num]
    random.shuffle(remain_num)

    for company, box_cnt, _, _ in companies:
        if not assigned_num[company]:
            assigned_num[company] = remain_num[:box_cnt]
            remain_num = remain_num[box_cnt:]

    return assigned_num, total

def attempt_random_assignment(companies, max_attempts=10):
    for attempt in range(max_attempts):
        try:
            assigned_numbers, total = random_num(companies)
            print(f"成功: {attempt + 1}回目の試行で箱番号を割り当てました。")
            return assigned_numbers, total
        except ValueError as e:
            print(f"試行 {attempt + 1} に失敗しました: {e}")
            continue

    raise RuntimeError("箱番号の割り当てに失敗しました。")

def load_access_table(db_path: str, table_name: str) -> pd.DataFrame:
    connection_string = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
    conn = pyodbc.connect(connection_string)

    try:
        query = f'SELECT * FROM {table_name}'
        df = pd.read_sql(query, conn)
    finally:
        df = df.iloc[:, 1:-1]  # 最初のカラムを削除
        for column in ["箱数", "条件(start)", "条件(end)"]:
            df[column] = df[column].astype('Int64')  # pandasのInt64型を使用してNAを保持
        conn.close()

    return df

def template_to_df(db_path, table_name):
    connection_string = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
    conn = pyodbc.connect(connection_string)

    try:
        query = f'SELECT * FROM {table_name}'
        df = pd.read_sql(query, conn)
    finally:
        df = df.iloc[:, 1:]  # 最初のカラムを削除
        conn.close()
    
    return df

def df_to_dict(df):
    box_dict = {}
    for _, row in df.iterrows():
        company = row['業者名']
        box_dict[company] = box_dict.get(company, 0) + 1
    return box_dict

def df_to_list(df, box_dict):
    comp_list = []
    for _, row in df.iterrows():
        company = row['会社名']
        box_cnt = row['箱数']
        start = row["条件(start)"]
        end = row["条件(end)"]

        if box_cnt == 0:
            continue

        if company in box_dict:
            box_cnt -= box_dict[company]
            if box_cnt < 0:
                raise ValueError(f"{company}の箱数が不足しています。テンプレートか箱数を修正してください。")
            elif box_cnt == 0:
                continue
        
        comp_list.append((company, box_cnt, start, end))

    return comp_list

def print_assigned_numbers(assigned_numbers, total):
    print(f"総箱数：{total = }")
    print()
    
    for company, numbers in assigned_numbers.items():
        print(f"{company}：{len(numbers)}箱")
        print(f"箱番号：{sorted(numbers)}")
        print("-----")
    
    sorted_assigns = sorted((number, company) for company, numbers in assigned_numbers.items() for number in numbers)

    for number, company in sorted_assigns:
        print(f"箱 {number:3}：{company}")

def main():
    db_path = r'C:\Users\lenovo02\Desktop\自動化\箱番作成\data.accdb'
    df = load_access_table(db_path, "箱番作成")

    ############################
    ######### 変更部分 ########## 
    table_name = "テンプレート1"
    ############################
    ############################
    
    df_check = template_to_df(db_path, table_name)
    box_check = df_to_dict(df_check)
    
    companies = df_to_list(df, box_check)

    companies = change_end(companies)
    companies = sort_list(companies)

    assigned_numbers, total = attempt_random_assignment(companies)
    
    print_assigned_numbers(assigned_numbers, total)

if __name__ == "__main__":
    main()

import os
import win32com.client
import openpyxl
import re
import yaml
import datetime

# Outlookアプリケーションをインスタンス化
outlook = win32com.client.Dispatch("Outlook.Application")

# ファイル名から数字を抽出する関数
def num_get(file_path):
    filename = os.path.basename(file_path)
    match = re.match(r'(\d{1,4})', filename)
    print(match.group(1))
    return match.group(1) if match else None

# 業者コードから連絡先を検索する関数
def select_to(src_file, src_sheet, code):
    src_wb = openpyxl.load_workbook(src_file)
    sheet = src_wb[src_sheet]

    vendors = {}
    pre_num, pre_ven, pre_add = 0, None, None
    target = False

    for row in sheet.iter_rows(min_row=2, values_only=True):
        number = row[3]        # 業者コードを取得
        address = row[6]       # 連絡先を取得
        
        # sheetが検索したい業者のとき
        if number == code:
            vendor_name = row[4]  # 業者名を取得 
            if vendor_name not in vendors:
                vendors[vendor_name] = []

            # 1つの業者に複数の連絡先がある場合に対応   
            if address == None:
                address = pre_add
            else :
                pre_add = address
            vendors[vendor_name].append(address)

            pre_num = number
            pre_ven = vendor_name
            target = True

        # 求めている業者でない場合は1つ前のnum, ven, addを保持
        elif number != code and number is not None:
            pre_ven = row[4] 
            pre_num = number
            if address != None:
                pre_add = address

        elif number is None and pre_num == code:
            vendor_name = pre_ven
            if address == None:
                address = pre_add
            vendors[vendor_name].append(address)

    # 業者が見つからなかったら
    if not target:
        raise ValueError(f"業者コード：{code} の業者が見つかりません。")
     
    return vendors

# メールを作成して下書きとして保存する関数
def create_draft_email(address, subject, body, attachment_path):
    # メールオブジェクトの作成
    mail = outlook.CreateItem(0)  # 0:メール
    
    # 送信先、件名、本文の設定
    mail.To = address
    mail.Subject = subject
    mail.Body = body
    
    # PDFファイルを添付する
    mail.Attachments.Add(Source=attachment_path, Type=6)
    
    # 下書きとしてメールを保存
    mail.Save()

def load_email_template(yaml_file):
    with open(yaml_file, 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)
    
def folder_process(folder_path, src_file, mail_template):
    #日付を取得
    dt_now = datetime.datetime.now()
    def_month = dt_now.month
    def_day = dt_now.day

    #yamlファイルで日付の指定があればそれを使用する。
    month = mail_template['variables'].get('month')
    if not month:
        month = def_month
    day = mail_template['variables'].get('day')
    if not day:
        day = def_day

    # folder_pathから大会名を読み取りsheetを指定
    if "呉服" in folder_path:
        src_sheet = "呉服大会"
        name = src_sheet
    elif "宝石" in folder_path:
        src_sheet = "時宝大会"
        name = "宝石大会"
    elif "時計" in folder_path:
        src_sheet = "時宝大会"
        name = "時計大会"
    elif "平場バッグ" in folder_path:
        src_sheet = "平場バッグ市"
        name = "平場バッグ大会"
    elif "バッグ" in folder_path:
        src_sheet = "バッグ大会"
        name = src_sheet
    elif "平場" in folder_path:
        src_sheet = "平場市"
        name = src_sheet

    # フォルダ内のPDFファイルを順番に処理する
    files = os.listdir(folder_path)
    print(f"{len(files)}社のメール作成を開始します。")
    print("---------------")

    file_sum, excel, pdf = 0, 0, 0
    for index, file in enumerate(files, start=1):
        num = int(num_get(file))
        vendors = select_to(src_file, src_sheet, num)

        # pdf, excelのファイル数を計算
        if ".xlsx" in file:
            excel += 1
        elif ".pdf" in file:
            pdf += 1
        
        for vendor_name, mail_address in vendors.items():
            print(f"{vendor_name}様宛のメール作成します。")
            for address in mail_address:
                if address and '@' in address:  # メールアドレスの条件チェック
                    print(f"・連絡先: {address}")
                    attachment_path = os.path.join(folder_path, file)
                    title = 'ご精算書送付について'
                    text = mail_template['email']['body'].format(
                        vendor_name=vendor_name, 
                        month=month, 
                        day=day, 
                        name=name
                    )

                    # メールの下書きを作成
                    create_draft_email(address, title, text, attachment_path)
                else:
                    # address がメールアドレスでない場合処理をスキップ
                    pass

            print()
            print(f"{vendor_name}様宛のメールを{len(mail_address)}個作成しました。")
            print("-------------")

        file_sum += len(mail_address)

    print("メール作成が完了しました。")
    print(f"会社数：{len(files)}社, メール総計：{file_sum}通")
    print(f"pdfファイル：{pdf}社, excelファイル：{excel}社")

def main():
    # メールアドレスを検索するExcelファイル
    file = r"c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\大会別精算書メール・LINE先一覧.xlsx"

    # YAMLファイルのパス
    yaml_file = r"c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\mail_template.yaml"
    mail_template = load_email_template(yaml_file)

    #############################################################################
    ######################## 変更部分 ############################################

    # 添付するPDFファイルが入っているフォルダのパス
    # 使用する大会の"#"を外す。使用しないものは全て"#"をつけておく。

    # バッグ大会
    #folder_path = r'c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\バッグ大会送信pdf'

    # 呉服大会
    #folder_path = r'c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\呉服大会送信pdf'

    # 時計大会
    #folder_path = r'c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\時計大会送信pdf'

    # 宝石大会
    folder_path = r'c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\宝石大会送信pdf'

    # 平場バッグ市
    #folder_path = r'c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\平場バッグ市送信pdf'

    # 平場市
    #folder_path = r'c:\\Users\\DELL-PC\\Desktop\\♡渡辺♡\\メール送信\\平場市送信pdf'

    ##############################################################################
    ##############################################################################
    
    folder_process(folder_path, file, mail_template)

if __name__ == "__main__":
    main()

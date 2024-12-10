import os
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import JSONResponse
from functions import deff_check

app = FastAPI()

# 基本の出力フォルダ
BASE_OUTPUT_FOLDER = "output_files"

# 出力フォルダが存在しない場合、作成
if not os.path.exists(BASE_OUTPUT_FOLDER):
    os.makedirs(BASE_OUTPUT_FOLDER)

# 仮の業務処理関数
def perform_box_number_creation(file_path, file):
    
    return f"箱番作成の処理が完了しました: {file.filename}, 保存先: {file_path}"

# 差額チェック処理関数
def perform_difference_check(file_path, file):
    print("差額チェックを実行中...")

    # ファイルの読み込み
    if "バッグ" in file_path:
        df = deff_check.bag_dataframe(file_path)
        indexes = deff_check.bag_filter(df)
        output_file = deff_check.bag_coloring(file_path, indexes)
    elif "時計" in file_path:
        df = deff_check.watch_dataframe(file_path)
        indexes = deff_check.watch_filter(df)
        output_file = deff_check.watch_coloring(file_path, indexes)
    elif "宝石" in file_path:
        df = deff_check.jewel_dataframe(file_path)
        indexes = deff_check.jewel_filter(df)
        output_file = deff_check.jewel_coloring(file_path, indexes)
    else:
        return {"message": "ファイル名が不正です。"}

    return {
        "message": "差額チェックが完了しました。",
        "output_file": file_path
    }

# タスク実行関数
def execute_task(task: str, file_path: str, file):
    if task == "差額チェック":
        return perform_difference_check(file_path, file)
    elif task == "箱番作成":
        return perform_box_number_creation(file_path, file)
    else:
        return {"error": "指定されたタスクは存在しません。"}


# 業務選択エンドポイント
@app.post("/execute_task")
async def execute_task_endpoint(
    task: str = Form(...),  # タスクの種類
    file: UploadFile = File(...)  # アップロードされたファイル
):

    file_path = os.path.join(BASE_OUTPUT_FOLDER, task)
    if not os.path.exists(file_path):
        os.makedirs(file_path)

    # 一時的なファイルの保存
    file_path = os.path.join(file_path, f"{file.filename}")
    with open(file_path, "wb") as f:
        f.write(await file.read())
    
    # 指定されたタスクを実行
    result = execute_task(task, file_path, file)

    return result


# APIの実行確認用エンドポイント
@app.get("/")
def read_root():
    return {"message": "FastAPIのタスク自動化プロジェクトへようこそ!"}


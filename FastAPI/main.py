from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import FileResponse
import uvicorn
import pandas as pd
import pyodbc
import os

app = FastAPI()

def template_to_df(db_path, table_name):
    connection_string = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};'
    conn = pyodbc.connect(connection_string)

    try:
        query = f'SELECT * FROM {table_name}'
        df = pd.read_sql(query, conn)
    finally:
        conn.close()
    
    return df

@app.post("/upload/")
async def upload_file(file: UploadFile = File(...), template: str = Form(...)):
    # 一時ファイルに保存
    temp_file_path = f"./{file.filename}"
    with open(temp_file_path, "wb") as temp_file:
        contents = await file.read()
        temp_file.write(contents)

    if file.filename.endswith('.accdb') or file.filename.endswith('.mdb'):
        connection_string = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={temp_file_path};"
        f = "this is access file."
        
        try:
            conn = pyodbc.connect(connection_string)
            query = "SELECT * FROM 箱番作成"  # ここでテーブル名を指定
            df = pd.read_sql(query, conn)
            temp_df = template_to_df(temp_file_path, template)

            df.columns = df.columns.str.strip()
            total = int(df["箱数"].sum())  # 合計を計算
            
            # CSVファイルを保存（Shift-JIS）
            csv_file_path = "temp_df.csv"
            temp_df.to_csv(csv_file_path, index=False, encoding='utf-8')

        except Exception as e:
            return {"error": str(e)}
        
        finally:
            conn.close()
    
        # 一時ファイルを削除
        os.remove(temp_file_path)
        
        return FileResponse(csv_file_path, media_type='text/csv', filename='temp_df.csv')

    else:
        return {"error": "サポートされていないファイル形式です。"}

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)

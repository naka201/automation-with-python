# F-project/app/routers/brand_search.py
from fastapi import APIRouter, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
from app.services import brand_search_service
import os
import shutil

router = APIRouter()

def cleanup_temp_dir(temp_dir: str):
    if os.path.isdir(temp_dir):
        print(f"一時ディレクトリ {temp_dir} をクリーンアップします。")
        shutil.rmtree(temp_dir)

@router.post("/process-files/", response_class=FileResponse)
async def process_files_endpoint(
    brand_list_excel: UploadFile = File(..., description="ブランドリストを含むExcelファイル (.xlsx)"),
    pdf_file: UploadFile = File(..., description="検索対象のPDFファイル (.pdf)"),
    target_excel: UploadFile = File(..., description="ハイライト対象のExcelファイル (.xlsx)"),
    sheet_name: str = Form(default="宝石大会のみ", description="ハイライト対象Excelのシート名")
):
    if not brand_list_excel.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="ブランドリストはExcelファイル（.xlsx, .xls）である必要があります。")
    if not pdf_file.content_type == "application/pdf":
        raise HTTPException(status_code=400, detail="検索対象はPDFファイル（.pdf）である必要があります。")
    if not target_excel.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="ハイライト対象はExcelファイル（.xlsx, .xls）である必要があります。")

    try:
        print("ファイル処理を開始します...")
        # UploadFileオブジェクト自体をサービス関数に渡す
        highlighted_excel_path, search_results, temp_dir_to_clean = await brand_search_service.run_brand_search_process(
            brand_list_excel_upload_file=brand_list_excel, # .file ではなく UploadFile オブジェクト
            pdf_upload_file=pdf_file,                      # .file ではなく UploadFile オブジェクト
            target_excel_upload_file=target_excel,         # .file ではなく UploadFile オブジェクト
            target_excel_filename=target_excel.filename,
            sheet_name=sheet_name
        )

        if not os.path.exists(highlighted_excel_path):
            cleanup_temp_dir(temp_dir_to_clean)
            raise HTTPException(status_code=500, detail="ハイライト処理されたファイルの生成に失敗しました。")

        print("検索結果:", search_results)

        return FileResponse(
            path=highlighted_excel_path,
            filename=os.path.basename(highlighted_excel_path),
            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            background=BackgroundTask(cleanup_temp_dir, temp_dir_to_clean)
        )

    except ValueError as ve:
        raise HTTPException(status_code=400, detail=f"入力エラー: {str(ve)}")
    except FileNotFoundError as fnfe:
        raise HTTPException(status_code=404, detail=f"ファイル関連エラー: {str(fnfe)}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"サーバー内部エラーが発生しました: {str(e)}")
    finally:
        # UploadFileオブジェクトが使用する一時ファイルを閉じる
        # (FastAPIが通常自動で処理するが、念のため)
        # .file を直接使わなくなったので、UploadFile オブジェクトの close() を呼ぶ
        await brand_list_excel.close()
        await pdf_file.close()
        await target_excel.close()
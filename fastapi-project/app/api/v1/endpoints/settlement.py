from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from app.services.settlement_service import generate_settlement_files
from typing import List
import shutil
import os
import traceback
import logging

# ログ設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

router = APIRouter()

@router.post("/settlement/generate")
async def create_settlement(
    input_file: UploadFile = File(...),
    output_files: List[UploadFile] = File(...),
    specified_date: str = None,
    pic : str = None
):
    try:
        # フォルダを処理前に削除
        output_folder = "/tmp/settlements"
        if os.path.exists(output_folder):
            shutil.rmtree(output_folder)  # フォルダごと削除
            logger.info(f"Deleted existing folder: {output_folder}")
        os.makedirs(output_folder)  # 新しく作成
        logger.info(f"Created new folder: {output_folder}")

        # Save the uploaded input file temporarily
        input_file_path = os.path.join(output_folder, input_file.filename)
        with open(input_file_path, "wb") as f:
            f.write(await input_file.read())
        logger.info(f"Saved input file: {input_file_path}")

        # Save the uploaded output files temporarily
        output_file_paths = []
        for output_file in output_files:
            output_file_path = os.path.join(output_folder, output_file.filename)
            with open(output_file_path, "wb") as f:
                f.write(await output_file.read())
            output_file_paths.append(output_file_path)
            logger.info(f"Saved output file: {output_file_path}")

        # Call the service to generate settlement files
        file_cnt = 0  # Initialize file count
        folder_path = await generate_settlement_files(input_file_path, output_file_paths, file_cnt, specified_date, pic)
        logger.info(f"Settlement files generated in: {folder_path}")

        # Create a ZIP file of the folder
        zip_file_path = f"{folder_path}.zip"
        shutil.make_archive(folder_path, 'zip', folder_path)
        logger.info(f"Created ZIP file: {zip_file_path}")


        return FileResponse(zip_file_path, filename=os.path.basename(zip_file_path))
    except Exception as e:
        error_message = traceback.format_exc()  # エラーの詳細な情報を取得
        logger.error(f"Error during settlement generation:\n{error_message}")
        raise HTTPException(status_code=500, detail=error_message)
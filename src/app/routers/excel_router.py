# PATH: src/app/routers/excel_router.py

from fastapi import APIRouter, File, UploadFile, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from typing import Optional, Union
import os
import tempfile
import zipfile
import shutil
import datetime

from src.app.generar_excel_crosstabs_completo import generar_excel_crosstabs_completo

router = APIRouter()

def get_file_path(filename: str, prefix: str = "") -> str:
    """ Helper function to create a file path """
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    return f"static/results/{prefix}{timestamp}_{filename}"

def create_zip_file(files, zip_filename):
    """ Create a zip file from a list of files """
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in files:
            zipf.write(file, os.path.basename(file))
    return zip_filename

def get_temp_file_path(suffix: str):
    """ Helper function to create a temporary file path """
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    return temp.name

def cleanup(files_to_zip, zip_filename):
    """ Function to clean up temporary files """
    for path in files_to_zip:
        if path and os.path.exists(path):
            os.remove(path)
    if zip_filename and os.path.exists(zip_filename):
        os.remove(zip_filename)

@router.post("/generate_excel/")
async def generate_excel(
    background_tasks: BackgroundTasks,
    archivo_stocks: Union[UploadFile, str] = File(default="STOCKS_EA_EB_2024_06_06.xlsx"),
    archivo_coois: UploadFile = File(...),
    master_data: Optional[UploadFile] = File(None),
    download_ea: bool = Form(True),
    download_eb: bool = Form(True)
) -> FileResponse:
    temp_coois_path = get_temp_file_path('.xlsx')
    temp_stocks_path = None
    master_data_path = None
    files_to_zip = []

    try:
        with open(temp_coois_path, "wb") as buffer:
            shutil.copyfileobj(archivo_coois.file, buffer)

        if isinstance(archivo_stocks, UploadFile):
            temp_stocks_path = get_temp_file_path('.xlsx')
            with open(temp_stocks_path, "wb") as buffer:
                shutil.copyfileobj(archivo_stocks.file, buffer)
        else:
            temp_stocks_path = f"prod_files/data/{archivo_stocks}"

        if master_data:
            master_data_path = get_temp_file_path('.xlsx')
            with open(master_data_path, "wb") as buffer:
                shutil.copyfileobj(master_data.file, buffer)

        result_paths = generar_excel_crosstabs_completo(
            archivo_stocks=temp_stocks_path,
            archivo_coois=temp_coois_path,
            archivo_maestros=master_data_path if master_data else "prod_files/data/master_data.xlsx"
        )

        files_to_zip = [path for path, should_download in zip(result_paths, [download_ea, download_eb]) if should_download and path]

        if files_to_zip:
            zip_filename = get_temp_file_path('.zip')
            create_zip_file(files_to_zip, zip_filename)
            response = FileResponse(zip_filename, media_type='application/zip', filename="descarga_EB_y_EA.zip")
            background_tasks.add_task(cleanup, files_to_zip, zip_filename)
            return response
        else:
            raise HTTPException(status_code=404, detail="No files generated or requested for download")
    finally:
        # This finally block is to ensure that files are removed even if an exception occurs before adding the cleanup task
        if not response:
            cleanup(files_to_zip, zip_filename if 'zip_filename' in locals() else None)


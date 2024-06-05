# PATH: main.py

from fastapi import FastAPI
from src.app.routers.excel_router import router as excel_router

app = FastAPI(title='Excel Generation API', version='1.0', description='API for generating and downloading Excel files.')

# Incluir el router del módulo Excel
app.include_router(excel_router, prefix="/excel", tags=["Excel Operations"])

@app.get("/")
def read_root():
    return {"message": "Welcome to the Excel Generation API. Use /docs for API documentation."}

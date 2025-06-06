import os
import uvicorn
from parsing_task import ParsingTask
from typing import Optional
from fastapi import FastAPI, HTTPException, Query
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pathlib import Path

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static") 
app.mount("/temptates", StaticFiles(directory="temptates"), name="temptates")

@app.get("/", response_class=HTMLResponse)
def home():
    with open("temptates/index.html", encoding="utf-8") as f:
        return f.read()

@app.get("/parse")
async def parse(
    docx_in: str = Query(..., description="Путь к входному DOCX файлу"),
    xlsx_in: str = Query(..., description=" Путь к входному XLSX файлу"),
    docx_out: str = Query(..., description="Путь для сохранения результата в DOCX файл"),
    xlsx_out: str = Query(..., description="Путь для сохранения результата в XLSX файл"),
    model: Optional[str] = Query("gemma:2b", description="Модель Ollama для использования")
):
    try:

        files_dir = Path("files")
        docx_in_path = files_dir / docx_in
        xlsx_in_path = files_dir / xlsx_in
        
        # Проверка существования входных файлов
        if not docx_in_path.exists():
            raise HTTPException(
                status_code=400,
                detail=f"Input DOCX file doesn't exist at {docx_in_path}"
            )
        if not xlsx_in_path.exists():
            raise HTTPException(
                status_code=400,
                detail=f"Input XLSX file doesn't exist at {xlsx_in_path}"
            )

        # Создаем полные пути для выходных файлов
        docx_out_path = files_dir / docx_out
        xlsx_out_path = files_dir / xlsx_out

        task = ParsingTask(
            str(docx_in_path), 
            str(xlsx_in_path), 
            str(docx_out_path), 
            str(xlsx_out_path), 
            model=model
        )
        result = await task.run()

        return {
            "status": "success",
            "result": result,
            "output_file": str(docx_out_path),
            "xlsx_output": str(xlsx_out_path),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
if __name__ == "__main__":
    uvicorn.run("main:app", reload=True)

  




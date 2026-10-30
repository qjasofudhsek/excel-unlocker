from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse
import msoffcrypto
import io
import openpyxl

app = FastAPI()

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Unlocker Server is running!"}

@app.post("/decrypt")
async def decrypt_excel(
    file: UploadFile = File(...),
    password: str = Form("1234")
):
    try:
        content = await file.read()
        file_io = io.BytesIO(content)
        decrypted = io.BytesIO()

        # 암호 해제 시도
        try:
            office_file = msoffcrypto.OfficeFile(file_io)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
        except Exception:
            # 암호가 없거나 형식이 일반 파일인 경우 원본 사용
            file_io.seek(0)
            decrypted = file_io

        # 엑셀 시트 데이터 읽기
        wb = openpyxl.load_workbook(decrypted, data_only=True)
        sheet = wb.active
        
        rows_data = []
        for row in sheet.iter_rows(values_only=True):
            # 빈 행 무시
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                rows_data.append([str(c) if c is not None else "" for c in row])

        return JSONResponse({"status": "success", "rows": rows_data})

    except Exception as e:
        return JSONResponse({"status": "error", "message": str(e)}, status_code=500)

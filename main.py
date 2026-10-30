from fastapi import FastAPI
from pydantic import BaseModel
import msoffcrypto
import io
import openpyxl
import base64
import gc

app = FastAPI()

class DecryptRequest(BaseModel):
    file_base64: str
    password: str = "1234"

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Unlocker Server is running!"}

@app.post("/decrypt")
async def decrypt_excel(req: DecryptRequest):
    decrypted = None
    wb = None
    try:
        # Base64 디코딩
        file_bytes = base64.b64decode(req.file_base64)
        file_io = io.BytesIO(file_bytes)
        decrypted = io.BytesIO()

        # 복호화 시도
        try:
            office_file = msoffcrypto.OfficeFile(file_io)
            office_file.load_key(password=req.password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
        except Exception:
            # 암호가 없거나 일반 엑셀 파일인 경우 원본 스트림 사용
            file_io.seek(0)
            decrypted = file_io

        del file_bytes
        del file_io

        # read_only=True로 메모리 폭증 방지 (스트리밍 파싱)
        wb = openpyxl.load_workbook(decrypted, read_only=True, data_only=True)
        sheet = wb.active

        rows_data = []
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                rows_data.append([str(c).strip() if c is not None else "" for c in row])

        return {"status": "success", "rows": rows_data}

    except Exception as e:
        return {"status": "error", "message": str(e)}

    finally:
        # 메모리 누수 방지 및 즉시 반환
        if wb is not None:
            wb.close()
        if decrypted is not None:
            decrypted.close()
        gc.collect()

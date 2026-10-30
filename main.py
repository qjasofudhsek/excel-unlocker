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
            file_io.seek(0)
            decrypted = file_io

        del file_bytes
        del file_io

        # 서식/스타일 오류 무시: data_only=True, read_only=False, keep_vba=False
        # openpyxl 버그 방지를 위해 read_only 모드 대신 일반 모드로 로드하되 스타일 검사 제외
        try:
            wb = openpyxl.load_workbook(
                decrypted, 
                data_only=True, 
                read_only=True
            )
            sheet = wb.active
        except Exception:
            # 조건부 서식 또는 스타일 XML 충돌 시 대체 로드
            decrypted.seek(0)
            wb = openpyxl.load_workbook(
                decrypted, 
                data_only=True, 
                read_only=False,
                keep_links=False
            )
            # 조건부 서식으로 인한 충돌 강제 제거
            sheet = wb.active
            sheet.conditional_formatting._cf_rules.clear()

        rows_data = []
        for row in sheet.iter_rows(values_only=True):
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                rows_data.append([str(c).strip() if c is not None else "" for c in row])

        return {"status": "success", "rows": rows_data}

    except Exception as e:
        return {"status": "error", "message": str(e)}

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        if decrypted is not None:
            try:
                decrypted.close()
            except Exception:
                pass
        gc.collect()

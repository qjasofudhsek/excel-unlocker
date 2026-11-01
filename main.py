from fastapi import FastAPI
from pydantic import BaseModel
from typing import Optional
import msoffcrypto
import io
import base64
import gc
from python_calamine import CalamineWorkbook

app = FastAPI()

class SheetListRequest(BaseModel):
    file_base64: str
    password: str = "1234"

class ExtractRequest(BaseModel):
    file_base64: str
    sheet_name: Optional[str] = None
    password: str = "1234"

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Unlocker Server is running!"}

def decrypt_stream(file_bytes: bytes, password: str) -> io.BytesIO:
    file_io = io.BytesIO(file_bytes)
    decrypted = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(file_io)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception:
        file_io.seek(0)
        return file_io

@app.post("/get_sheets")
async def get_sheets(req: SheetListRequest):
    decrypted = None
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_stream(file_bytes, req.password)
        workbook = CalamineWorkbook.from_filelike(decrypted)
        return {"status": "success", "sheets": workbook.sheet_names}
    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if decrypted:
            decrypted.close()
        gc.collect()

@app.post("/extract_data")
async def extract_data(req: ExtractRequest):
    decrypted = None
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_stream(file_bytes, req.password)
        workbook = CalamineWorkbook.from_filelike(decrypted)

        target = req.sheet_name if (req.sheet_name and req.sheet_name in workbook.sheet_names) else workbook.sheet_names[0]
        sheet = workbook.get_sheet_by_name(target)
        rows = sheet.to_python()

        # 헤더 행 탐색 ("번호", "부서", "재적" 등)
        header_idx = -1
        for r_idx, row in enumerate(rows):
            str_row = [str(c or "").strip() for c in row]
            if any(h in str_row for h in ["번호", "부서", "성명", "구분", "순번", "재적", "이름"]):
                header_idx = r_idx
                break
        
        if header_idx == -1:
            header_idx = 0

        # 상단 설명/제목 행 분리
        title_rows = []
        for r in range(header_idx):
            title_rows.append([str(c) if c is not None else "" for c in rows[r]])

        headers = [str(c) if c is not None else "" for c in rows[header_idx]]

        # 순수 데이터 행 추출 (하단 "계", "합계" 만나면 즉시 중단)
        data_rows = []
        for row in rows[header_idx + 1:]:
            str_row = [str(c or "").strip() for c in row]
            if any(f in str_row[:3] for f in ["계", "합계", "총계", "소계"]):
                break
            if not any(v != "" for v in str_row):
                continue
            data_rows.append([str(c) if c is not None else "" for c in row])

        return {
            "status": "success",
            "title_rows": title_rows,
            "headers": headers,
            "data_rows": data_rows
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if decrypted:
            decrypted.close()
        gc.collect()

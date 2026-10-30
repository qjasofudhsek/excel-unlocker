from fastapi import FastAPI
from pydantic import BaseModel
import msoffcrypto
import io
import base64
import gc
from python_calamine import CalamineWorkbook

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
    try:
        file_bytes = base64.b64decode(req.file_base64)
        file_io = io.BytesIO(file_bytes)
        decrypted = io.BytesIO()

        # 1. 엑셀 비밀번호 복호화
        try:
            office_file = msoffcrypto.OfficeFile(file_io)
            office_file.load_key(password=req.password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
        except Exception:
            # 암호가 없거나 일반 엑셀인 경우 원본 스트림 사용
            file_io.seek(0)
            decrypted = file_io

        del file_bytes
        del file_io

        # 2. calamine 엔진으로 순수 데이터만 추출 (조건부 서식/스타일 오류 원천 차단)
        workbook = CalamineWorkbook.from_filelike(decrypted)
        sheet_names = workbook.sheet_names
        if not sheet_names:
            return {"status": "error", "message": "엑셀 시트를 찾을 수 없습니다."}

        first_sheet = workbook.get_sheet_by_name(sheet_names[0])
        rows = first_sheet.to_python()

        rows_data = []
        for row in rows:
            # 빈 행 무시
            if any(cell is not None and str(cell).strip() != "" for cell in row):
                rows_data.append([str(c).strip() if c is not None else "" for c in row])

        return {"status": "success", "rows": rows_data}

    except Exception as e:
        return {"status": "error", "message": str(e)}

    finally:
        if decrypted is not None:
            try:
                decrypted.close()
            except Exception:
                pass
        gc.collect()
 

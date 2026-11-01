from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import base64
import gc
from python_calamine import CalamineWorkbook

app = FastAPI()

class SheetListRequest(BaseModel):
    file_base64: str
    password: Optional[str] = "1234"

class OneStopMergeRequest(BaseModel):
    files_base64: List[str]
    sheet_name: Optional[str] = None
    password: Optional[str] = "1234"

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Unlocker Server is running!"}

def decrypt_file_stream(file_bytes: bytes, password: str) -> io.BytesIO:
    file_io = io.BytesIO(file_bytes)
    decrypted = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(file_io)
        office_file.load_key(password=password or "1234")
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
        decrypted = decrypt_file_stream(file_bytes, req.password)
        workbook = CalamineWorkbook.from_filelike(decrypted)
        return {"status": "success", "sheets": workbook.sheet_names}
    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if decrypted:
            try:
                decrypted.close()
            except Exception:
                pass
        gc.collect()

@app.post("/extract_all_onestop")
async def extract_all_onestop(req: OneStopMergeRequest):
    master_title_rows = []
    master_headers = []
    all_data_rows = []
    error_list = []

    for idx, b64_str in enumerate(req.files_base64):
        decrypted = None
        try:
            file_bytes = base64.b64decode(b64_str)
            decrypted = decrypt_file_stream(file_bytes, req.password)
            workbook = CalamineWorkbook.from_filelike(decrypted)

            available_sheets = workbook.sheet_names
            if not available_sheets:
                error_list.append(f"파일#{idx+1}: 시트 없음")
                continue

            # 시트명 매칭
            target_sheet = None
            if req.sheet_name:
                req_clean = req.sheet_name.strip()
                for s in available_sheets:
                    if s.strip() == req_clean:
                        target_sheet = s
                        break
            if not target_sheet:
                target_sheet = available_sheets[0]

            sheet = workbook.get_sheet_by_name(target_sheet)
            rows = sheet.to_python()

            if not rows:
                continue

            # 헤더 행 탐색
            header_idx = -1
            for r_idx, row in enumerate(rows):
                str_row = [str(c or "").strip() for c in row]
                if any(h in str_row for h in ["번호", "부서", "성명", "구분", "순번", "재적", "이름"]):
                    header_idx = r_idx
                    break

            if header_idx == -1:
                header_idx = 0

            # 첫 번째 정상 파일에서 제목 및 헤더 확보
            if len(master_headers) == 0:
                for r in range(header_idx):
                    master_title_rows.append([str(c) if c is not None else "" for c in rows[r]])
                master_headers = [str(c) if c is not None else "" for c in rows[header_idx]]

            # 순수 데이터 행 추출 (계, 합계, 총계, 소계 행 만나면 중단)
            for row in rows[header_idx + 1:]:
                str_row = [str(c or "").strip() for c in row]
                if any(f in str_row[:3] for f in ["계", "합계", "총계", "소계"]):
                    break
                if not any(v != "" for v in str_row):
                    continue
                all_data_rows.append([str(c) if c is not None else "" for c in row])

        except Exception as e:
            error_list.append(f"파일#{idx+1}: {str(e)}")
        finally:
            # 파일 하나 끝날 때마다 메모리 강제 소각 (502 에러 원천 차단)
            if decrypted:
                try:
                    decrypted.close()
                except Exception:
                    pass
            del file_bytes
            gc.collect()

    return {
        "status": "success",
        "title_rows": master_title_rows,
        "headers": master_headers,
        "data_rows": all_data_rows,
        "errors": error_list
    }

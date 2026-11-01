from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import base64
import gc
from python_calamine import CalamineWorkbook

app = FastAPI()

# 안전 가드레일 상수 설정 (유령 셀 폭탄 원천 방지)
MAX_SCAN_ROWS = 3000   # 파일당 최대 탐색 행
MAX_SCAN_COLS = 50     # 파일당 최대 탐색 열
MAX_EMPTY_GAP = 5      # 연속 빈 줄 허용 한도 (초과 시 즉시 읽기 종료)

class SheetListRequest(BaseModel):
    file_base64: str
    password: Optional[str] = "1234"

class OneStopMergeRequest(BaseModel):
    files_base64: List[str]
    sheet_name: Optional[str] = None
    password: Optional[str] = "1234"

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Unlocker Safe-Engine Ready"}

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
        # 암호가 없거나 복호화 대상이 아닌 일반 엑셀인 경우
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
        file_bytes = None
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
            raw_rows = sheet.to_python()

            if not raw_rows:
                continue

            # 1. 유령 셀(100만 행/열) 차단: 크기 제한 적용
            sliced_rows = []
            for r in raw_rows[:MAX_SCAN_ROWS]:
                sliced_rows.append(r[:MAX_SCAN_COLS])
            del raw_rows

            # 2. 헤더 행 탐색
            header_idx = -1
            for r_idx, row in enumerate(sliced_rows):
                str_row = [str(c or "").strip() for c in row]
                if any(h in str_row for h in ["번호", "부서", "성명", "구분", "순번", "재적", "이름"]):
                    header_idx = r_idx
                    break

            if header_idx == -1:
                header_idx = 0

            # 3. 첫 번째 정상 파일 기준 제목 및 헤더 확정
            if len(master_headers) == 0:
                for r in range(header_idx):
                    master_title_rows.append([str(c) if c is not None else "" for c in sliced_rows[r]])
                master_headers = [str(c) if c is not None else "" for c in sliced_rows[header_idx]]

            # 4. 순수 데이터 추출 (합계 행 또는 연속 빈 행 감지 시 즉시 탈출)
            empty_streak = 0
            for row in sliced_rows[header_idx + 1:]:
                str_row = [str(c or "").strip() for c in row]

                # 하단 요약/합계 행 감지 시 중단
                if any(f in str_row[:3] for f in ["계", "합계", "총계", "소계"]):
                    break

                # 빈 줄 체크
                if not any(v != "" for v in str_row):
                    empty_streak += 1
                    if empty_streak >= MAX_EMPTY_GAP:
                        # 빈 줄이 5개 연속 나오면 데이터 끝으로 판정하고 조기 탈출
                        break
                    continue
                else:
                    empty_streak = 0

                all_data_rows.append([str(c) if c is not None else "" for c in row])

        except Exception as e:
            error_list.append(f"파일#{idx+1}: {str(e)}")
        finally:
            if decrypted:
                try:
                    decrypted.close()
                except Exception:
                    pass
            if file_bytes:
                del file_bytes
            gc.collect()

    return {
        "status": "success",
        "title_rows": master_title_rows,
        "headers": master_headers,
        "data_rows": all_data_rows,
        "errors": error_list
    }

from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import msoffcrypto
import io
import base64
import gc
import re
from python_calamine import CalamineWorkbook

app = FastAPI()

MAX_SCAN_ROWS = 3000
MAX_SCAN_COLS = 50
MAX_EMPTY_GAP = 5

HEADER_CANDIDATES = [
    "번호", "순번", "고유번호", "지파", "회", 
    "성명", "이름", "부서", "부서명", "소속", 
    "교회", "교회명", "지역", "구분", "직책", "직급", "직분", 
    "시험", "단계", "시험단계", "구역", "연락처", "전화번호", "생년월일", "나이"
]

FOOTER_EXACT_KEYWORDS = ["종합", "계", "합계", "총계", "소계", "총합계"]

# 고유번호 정규식 패턴: 숫자 8자리 - 숫자 5자리
ID_PATTERN = re.compile(r"^\d{8}-\d{5}$")

class SheetListRequest(BaseModel):
    file_base64: str
    password: Optional[str] = "1234"

class ValidateRequest(BaseModel):
    file_base64: str
    password: Optional[str] = "1234"

class OneStopMergeRequest(BaseModel):
    files_base64: List[str]
    sheet_name: Optional[str] = None
    password: Optional[str] = "1234"

def log(msg: str):
    print(f"[ENGINE_LOG] {msg}", flush=True)

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Smart Engine & Validator Ready"}

def get_col_letter(col_idx: int) -> str:
    dividend = col_idx + 1
    col_letter = ""
    while dividend > 0:
        modulo = (dividend - 1) % 26
        col_letter = chr(65 + modulo) + col_letter
        dividend = (dividend - modulo) // 26
    return col_letter

def decrypt_file_stream(file_bytes: bytes, password: str, idx: int = 0) -> io.BytesIO:
    file_io = io.BytesIO(file_bytes)
    decrypted = io.BytesIO()
    try:
        office_file = msoffcrypto.OfficeFile(file_io)
        office_file.load_key(password=password or "1234")
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted
    except Exception as e:
        log(f"파일#{idx+1} 일반 파일(비암호화) 처리: {str(e)}")
        file_io.seek(0)
        return file_io

@app.post("/get_sheets")
async def get_sheets(req: SheetListRequest):
    decrypted = None
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_file_stream(file_bytes, req.password, 0)
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

# [신규] 고유번호 유효성 및 중복 실시간 검증 엔드포인트
@app.post("/validate_file")
async def validate_file(req: ValidateRequest):
    decrypted = None
    file_bytes = None
    cal_wb = None
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_file_stream(file_bytes, req.password, 0)
        cal_wb = CalamineWorkbook.from_filelike(decrypted)

        if not cal_wb.sheet_names:
            return {"status": "error", "message": "시트가 존재하지 않습니다."}

        # 첫 번째 시트 기준 검증
        sheet = cal_wb.get_sheet_by_name(cal_wb.sheet_names[0])
        
        raw_rows = []
        if hasattr(sheet, "iter_rows"):
            for r_i, r in enumerate(sheet.iter_rows()):
                if r_i >= MAX_SCAN_ROWS:
                    break
                raw_rows.append([str(c or "").strip() for c in list(r)[:MAX_SCAN_COLS]])
        else:
            for r in sheet.to_python()[:MAX_SCAN_ROWS]:
                raw_rows.append([str(c or "").strip() for c in list(r)[:MAX_SCAN_COLS]])

        if not raw_rows:
            return {"status": "error", "message": "데이터가 비어있는 파일입니다."}

        # 1. 헤더 행 및 고유번호 / 이름 열 탐색
        best_header_idx = -1
        best_score = 0
        scan_limit = min(30, len(raw_rows))
        for r_idx in range(scan_limit):
            row = raw_rows[r_idx]
            score = sum(1 for cell in row if cell and any(k in cell for k in HEADER_CANDIDATES))
            if score > best_score:
                best_score = score
                best_header_idx = r_idx

        if best_header_idx == -1 or best_score < 2:
            best_header_idx = 0

        header_row = raw_rows[best_header_idx]
        id_col_idx = -1
        name_col_idx = -1

        for c_idx, h in enumerate(header_row):
            if "고유번호" in h:
                id_col_idx = c_idx
            elif any(n in h for n in ["성명", "이름"]):
                name_col_idx = c_idx

        if id_col_idx == -1:
            # 고유번호 항목이 없는 파일은 검증 건너뛰고 정상 통과
            return {"status": "valid", "has_id_column": False, "message": "고유번호 열 없음 (일반 파일 통과)"}

        # 2. 데이터 행 순회하며 형식 및 중복 검증
        col_letter = get_col_letter(id_col_idx)
        seen_ids: Dict[str, List[Dict[str, Any]]] = {}
        format_errors = []
        empty_streak = 0

        for r_idx in range(best_header_idx + 1, len(raw_rows)):
            row = raw_rows[r_idx]

            # 요약행 감지 시 검사 종료
            if any(cell_val.replace(" ", "") in FOOTER_EXACT_KEYWORDS for cell_val in row[:3]):
                break

            # 빈 행 감지
            if not any(v != "" for v in row):
                empty_streak += 1
                if empty_streak >= MAX_EMPTY_GAP:
                    break
                continue
            else:
                empty_streak = 0

            excel_row_num = r_idx + 1
            cell_address = f"{col_letter}{excel_row_num}"
            raw_id = row[id_col_idx] if id_col_idx < len(row) else ""
            person_name = row[name_col_idx] if (name_col_idx != -1 and name_col_idx < len(row)) else "이름미상"

            # (1) 형식 검사
            if not raw_id:
                format_errors.append({
                    "cell": cell_address,
                    "name": person_name,
                    "val": "(빈칸)",
                    "reason": "고유번호가 비어있음"
                })
            elif not ID_PATTERN.match(raw_id):
                format_errors.append({
                    "cell": cell_address,
                    "name": person_name,
                    "val": raw_id,
                    "reason": "형식 불일치 (8자리-5자리 아님)"
                })
            else:
                # (2) 정상 형식인 경우 중복 검사용 수집
                if raw_id not in seen_ids:
                    seen_ids[raw_id] = []
                seen_ids[raw_id].append({
                    "cell": cell_address,
                    "name": person_name
                })

        # (3) 중복 집계 (2회 이상 등장한 ID 필터링)
        duplicates = {}
        for uid, occurrences in seen_ids.items():
            if len(occurrences) > 1:
                duplicates[uid] = occurrences

        # 오류 판정
        is_invalid = len(format_errors) > 0 or len(duplicates) > 0

        return {
            "status": "invalid" if is_invalid else "valid",
            "has_id_column": True,
            "format_errors": format_errors,
            "duplicates": duplicates
        }

    except Exception as e:
        log(f"검증 도중 예외 발생: {str(e)}")
        return {"status": "error", "message": str(e)}
    finally:
        if cal_wb:
            del cal_wb
        if decrypted:
            try:
                decrypted.close()
            except Exception:
                pass
        if file_bytes:
            del file_bytes
        gc.collect()

@app.post("/extract_all_onestop")
async def extract_all_onestop(req: OneStopMergeRequest):
    log(f"--> [범용 엔진] 수합 요청 시작 (파일 {len(req.files_base64)}개, 시트: {req.sheet_name})")
    master_title_rows = []
    master_headers = []
    all_data_rows = []
    error_list = []

    for idx, b64_str in enumerate(req.files_base64):
        decrypted = None
        file_bytes = None
        cal_wb = None
        try:
            file_bytes = base64.b64decode(b64_str)
            decrypted = decrypt_file_stream(file_bytes, req.password, idx)
            cal_wb = CalamineWorkbook.from_filelike(decrypted)

            available_sheets = cal_wb.sheet_names
            if not available_sheets:
                error_list.append(f"파일#{idx+1}: 시트 없음")
                continue

            target_sheet = None
            if req.sheet_name:
                req_clean = req.sheet_name.strip()
                for s in available_sheets:
                    if s.strip() == req_clean:
                        target_sheet = s
                        break
            if not target_sheet:
                target_sheet = available_sheets[0]

            sheet = cal_wb.get_sheet_by_name(target_sheet)

            raw_rows = []
            if hasattr(sheet, "iter_rows"):
                for r_i, r in enumerate(sheet.iter_rows()):
                    if r_i >= MAX_SCAN_ROWS:
                        break
                    raw_rows.append([str(c or "").strip() for c in list(r)[:MAX_SCAN_COLS]])
            else:
                for r in sheet.to_python()[:MAX_SCAN_ROWS]:
                    raw_rows.append([str(c or "").strip() for c in list(r)[:MAX_SCAN_COLS]])

            if not raw_rows:
                continue

            best_header_idx = -1
            best_score = 0
            scan_limit = min(30, len(raw_rows))
            for r_idx in range(scan_limit):
                row = raw_rows[r_idx]
                score = sum(1 for cell in row if cell and any(k in cell for k in HEADER_CANDIDATES))
                if score > best_score:
                    best_score = score
                    best_header_idx = r_idx

            if best_header_idx == -1 or best_score < 2:
                best_header_idx = 0

            if len(master_headers) == 0:
                for r in range(best_header_idx):
                    master_title_rows.append(raw_rows[r])
                master_headers = raw_rows[best_header_idx]

            empty_streak = 0
            for r_idx in range(best_header_idx + 1, len(raw_rows)):
                row = raw_rows[r_idx]
                if any(cell_val.replace(" ", "") in FOOTER_EXACT_KEYWORDS for cell_val in row[:3]):
                    break
                if not any(v != "" for v in row):
                    empty_streak += 1
                    if empty_streak >= MAX_EMPTY_GAP:
                        break
                    continue
                else:
                    empty_streak = 0
                all_data_rows.append(row)

        except Exception as e:
            error_list.append(f"파일#{idx+1}: {str(e)}")
        finally:
            if cal_wb:
                del cal_wb
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

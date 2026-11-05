from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import msoffcrypto
import openpyxl
import io
import base64
import gc
import re

app = FastAPI()

MAX_SCAN_ROWS = 1000
MAX_SCAN_COLS = 40
MAX_EMPTY_GAP = 5

HEADER_CANDIDATES = [
    "번호", "순번", "고유번호", "지파", "회", 
    "성명", "이름", "부서", "부서명", "소속", 
    "교회", "교회명", "지역", "구분", "직책", "직급", "직분", 
    "시험", "단계", "시험단계", "구역", "연락처", "전화번호", "핸드폰번호", "생년월일", "나이"
]

FOOTER_EXACT_KEYWORDS = ["종합", "계", "합계", "총계", "소계", "총합계"]

# 선택 입력(공란이어도 정상으로 인정할 열)
OPTIONAL_COLUMNS = ["환우석여부", "비고", "사유", "불참사유", "사명경험", "참석", "등록상태"]

ID_PATTERN = re.compile(r"^\d{8}-\d{5}$")
PHONE_PATTERN = re.compile(r"^01[016789]-?\d{3,4}-?\d{4}$")
EXCEL_ERROR_PATTERN = re.compile(r"^#(VALUE!|REF!|DIV/0!|N/A|NAME\?|NUM!|NULL!)$")

class SheetListRequest(BaseModel):
    file_base64: str
    password: Optional[str] = "1234"

class ValidateRequest(BaseModel):
    file_base64: str
    password: Optional[str] = "1234"

class MergeRequest(BaseModel):
    files_base64: List[str]
    sheet_name: Optional[str] = None
    password: Optional[str] = "1234"

def log(msg: str):
    print(f"[ENGINE_LOG] {msg}", flush=True)

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Universal Blank & Integrity Inspector Ready"}

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
    except Exception:
        file_io.seek(0)
        return file_io

@app.post("/get_sheets")
async def get_sheets(req: SheetListRequest):
    decrypted = None
    wb = None
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_file_stream(file_bytes, req.password, 0)
        wb = openpyxl.load_workbook(decrypted, read_only=True, data_only=True)
        return {"status": "success", "sheets": wb.sheetnames}
    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if wb:
            try: wb.close()
            except Exception: pass
        if decrypted:
            try: decrypted.close()
            except Exception: pass
        gc.collect()

# [전수 검증] 모든 열의 공란(미작성) 및 형식 오류 정밀 감지
@app.post("/validate_file")
async def validate_file(req: ValidateRequest):
    decrypted = None
    file_bytes = None
    wb = None
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_file_stream(file_bytes, req.password, 0)
        wb = openpyxl.load_workbook(decrypted, read_only=True, data_only=True)

        if not wb.sheetnames:
            return {"status": "error", "message": "시트가 존재하지 않습니다."}

        errors_summary = []

        for s_name in wb.sheetnames:
            ws = wb[s_name]
            raw_rows = []
            
            for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if r_idx >= MAX_SCAN_ROWS:
                    break
                row_str = [str(c or "").strip() for c in list(row)[:MAX_SCAN_COLS]]
                raw_rows.append(row_str)

            if not raw_rows:
                continue

            best_header_idx = -1
            best_score = 0
            for r_idx in range(min(30, len(raw_rows))):
                row = raw_rows[r_idx]
                score = sum(1 for cell in row if cell and any(k in cell for k in HEADER_CANDIDATES))
                if score > best_score:
                    best_score = score
                    best_header_idx = r_idx

            if best_header_idx == -1 or best_score < 2:
                continue

            header_row = raw_rows[best_header_idx]
            
            # 유효 헤더 컬럼 목록 구축 (빈 헤더 제외)
            valid_cols = {}
            name_col = -1
            dept_col = -1

            for c_idx, h in enumerate(header_row):
                h_clean = h.replace("\n", "").replace(" ", "")
                if not h_clean:
                    continue
                valid_cols[c_idx] = h_clean
                if any(n in h_clean for n in ["성명", "이름"]):
                    name_col = c_idx
                elif any(d in h_clean for d in ["부서", "부서명"]):
                    dept_col = c_idx

            seen_ids = {}
            empty_streak = 0

            for r_idx in range(best_header_idx + 1, len(raw_rows)):
                row = raw_rows[r_idx]

                if any(cell_val.replace(" ", "") in FOOTER_EXACT_KEYWORDS for cell_val in row[:3]):
                    break

                if not any(v != "" for v in row):
                    empty_streak += 1
                    if empty_streak >= MAX_EMPTY_GAP: break
                    continue
                else:
                    empty_streak = 0

                row_num = r_idx + 1

                # 예시 행 필터
                if any(row[c].startswith("ex") or "예시" in row[c] for c in range(min(3, len(row)))):
                    errors_summary.append(f"[{s_name}] {row_num}행: '예시(ex)' 행 미삭제 제출")
                    continue

                # 실제 유효 명단 행인지 판정 (이름이나 부서가 있는 줄)
                has_real_person = False
                if name_col != -1 and name_col < len(row):
                    n_val = row[name_col]
                    if n_val and n_val != "None" and not n_val.startswith("#"):
                        has_real_person = True
                elif dept_col != -1 and dept_col < len(row):
                    d_val = row[dept_col]
                    if d_val and d_val != "None" and not d_val.startswith("#"):
                        has_real_person = True
                else:
                    non_empty = [c for c in row if c and c != "None"]
                    if len(non_empty) >= 3:
                        has_real_person = True

                if not has_real_person:
                    continue

                # 유효 데이터 행에 대해 모든 열 전수 검사 (공란 + 형식오류)
                for c_idx, col_title in valid_cols.items():
                    val = row[c_idx] if c_idx < len(row) else ""
                    cell_pos = f"[{s_name}] {get_col_letter(c_idx)}{row_num}({col_title})"

                    # 1. 엑셀 수식 에러 검출
                    if EXCEL_ERROR_PATTERN.match(val):
                        errors_summary.append(f"{cell_pos}: 수식 계산 오류({val})")
                        continue

                    # 2. 공란(빈칸) 검사: 선택 입력 열이 아닌데 비어있는 경우
                    is_optional = any(opt in col_title for opt in OPTIONAL_COLUMNS)
                    if not val or val == "None":
                        if not is_optional:
                            # 번호/순번 열은 자동 부여되므로 공란이어도 통과
                            if not any(num_word in col_title for num_word in ["번호", "순번"]):
                                errors_summary.append(f"{cell_pos}: 항목 누락 (공란)")
                        continue

                    # 3. 고유번호 형식 검사
                    if "고유번호" in col_title:
                        if not ID_PATTERN.match(val):
                            errors_summary.append(f"{cell_pos}: 고유번호 규격 오류 (8자리-5자리 아님)")
                        else:
                            if val not in seen_ids: seen_ids[val] = []
                            seen_ids[val].append(cell_pos)

                    # 4. 연락처 형식 검사
                    elif any(p in col_title for p in ["핸드폰", "연락처", "전화번호"]):
                        if not PHONE_PATTERN.match(val):
                            errors_summary.append(f"{cell_pos}: 연락처 번호 형식 오류")

                    # 5. 성별 형식 검사
                    elif col_title == "성별":
                        if val not in ["남", "여"]:
                            errors_summary.append(f"{cell_pos}: 성별 표기 오류 ('남' 또는 '여'만 가능)")

                    # 6. 나이 형식 검사
                    elif any(a in col_title for a in ["나이", "만나이"]):
                        clean_age = val.replace("세", "").strip()
                        if not clean_age.isdigit() or not (1 <= int(clean_age) <= 130):
                            errors_summary.append(f"{cell_pos}: 나이 숫자 표기 오류")

            # 고유번호 중복 종합
            for uid, pos_list in seen_ids.items():
                if len(pos_list) > 1:
                    errors_summary.append(f"고유번호 중복 발견 -> 위치: {', '.join(pos_list)}")

        is_invalid = len(errors_summary) > 0
        return {
            "status": "invalid" if is_invalid else "valid",
            "errors": errors_summary
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if wb:
            try: wb.close()
            except Exception: pass
        if decrypted:
            try: decrypted.close()
            except Exception: pass
        if file_bytes: del file_bytes
        gc.collect()

@app.post("/extract_all_onestop")
async def extract_all_onestop(req: MergeRequest):
    is_merge_all = (req.sheet_name is None or req.sheet_name.strip() in ["", "ALL", "모든시트", "전체시트"])
    merged_sheets_dict: Dict[str, Dict[str, Any]] = {}
    error_list = []

    for idx, b64_str in enumerate(req.files_base64):
        decrypted = None
        file_bytes = None
        wb = None
        try:
            file_bytes = base64.b64decode(b64_str)
            decrypted = decrypt_file_stream(file_bytes, req.password, idx)
            wb = openpyxl.load_workbook(decrypted, read_only=True, data_only=True)

            available_sheets = wb.sheetnames
            if not available_sheets:
                error_list.append(f"파일#{idx+1}: 시트 없음")
                continue

            target_sheets = available_sheets if is_merge_all else [req.sheet_name.strip()]

            for s_name in target_sheets:
                if s_name not in available_sheets:
                    continue

                ws = wb[s_name]
                raw_rows = []
                
                for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
                    if r_idx >= MAX_SCAN_ROWS:
                        break
                    raw_rows.append([str(c or "").strip() for c in list(row)[:MAX_SCAN_COLS]])

                if not raw_rows:
                    continue

                best_header_idx = -1
                best_score = 0
                for r_idx in range(min(30, len(raw_rows))):
                    row = raw_rows[r_idx]
                    score = sum(1 for cell in row if cell and any(k in cell for k in HEADER_CANDIDATES))
                    if score > best_score:
                        best_score = score
                        best_header_idx = r_idx

                if best_header_idx == -1 or best_score < 2:
                    best_header_idx = 0

                header_row = raw_rows[best_header_idx]

                name_col = -1
                dept_col = -1
                for c_idx, h in enumerate(header_row):
                    h_clean = h.replace("\n", "").replace(" ", "")
                    if any(n in h_clean for n in ["성명", "이름"]):
                        name_col = c_idx
                    elif any(d in h_clean for d in ["부서", "부서명"]):
                        dept_col = c_idx

                if s_name not in merged_sheets_dict:
                    merged_sheets_dict[s_name] = {
                        "title_rows": [raw_rows[r] for r in range(best_header_idx)],
                        "headers": header_row,
                        "data_rows": []
                    }

                empty_streak = 0
                for r_idx in range(best_header_idx + 1, len(raw_rows)):
                    row = raw_rows[r_idx]

                    if any(cell_val.replace(" ", "") in FOOTER_EXACT_KEYWORDS for cell_val in row[:3]):
                        break

                    if not any(v != "" for v in row):
                        empty_streak += 1
                        if empty_streak >= MAX_EMPTY_GAP: break
                        continue
                    else:
                        empty_streak = 0

                    if any(row[c].startswith("ex") or "예시" in row[c] for c in range(min(3, len(row)))):
                        continue

                    has_real_person = False
                    if name_col != -1 and name_col < len(row):
                        n_val = row[name_col]
                        if n_val and n_val != "None" and not n_val.startswith("#"):
                            has_real_person = True
                    elif dept_col != -1 and dept_col < len(row):
                        d_val = row[dept_col]
                        if d_val and d_val != "None" and not d_val.startswith("#"):
                            has_real_person = True
                    else:
                        non_empty = [c for c in row if c and c != "None"]
                        if len(non_empty) >= 3:
                            has_real_person = True

                    if not has_real_person:
                        continue

                    merged_sheets_dict[s_name]["data_rows"].append(row)

        except Exception as e:
            error_list.append(f"파일#{idx+1}: {str(e)}")
        finally:
            if wb:
                try: wb.close()
                except Exception: pass
            if decrypted:
                try: decrypted.close()
                except Exception: pass
            if file_bytes: del file_bytes
            gc.collect()

    return {
        "status": "success",
        "is_multi_sheet": is_merge_all,
        "sheets_data": merged_sheets_dict,
        "errors": error_list
    }

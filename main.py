from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import base64
import gc
from python_calamine import CalamineWorkbook

app = FastAPI()

MAX_SCAN_ROWS = 3000   # 파일당 최대 탐색 행
MAX_SCAN_COLS = 50     # 파일당 최대 탐색 열
MAX_EMPTY_GAP = 5      # 연속 빈 줄 허용 한도

# 1. 헤더 판별 후보 키워드 (고유번호, 지파, 회 반영)
HEADER_CANDIDATES = [
    "번호", "순번", "고유번호", "지파", "회", 
    "성명", "이름", "부서", "부서명", "소속", 
    "교회", "교회명", "지역", "구분", "직책", "직급", "직분", 
    "시험", "단계", "시험단계", "구역", "연락처", "전화번호", "생년월일", "나이"
]

# 2. 양식 종료 감지 키워드 (오직 단독/완전 일치로만 검증)
FOOTER_EXACT_KEYWORDS = ["종합", "계", "합계", "총계", "소계", "총합계"]

class SheetListRequest(BaseModel):
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
    return {"status": "ok", "message": "Excel Smart Universal Engine Ready"}

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
            log(f"--- 파일#{idx+1} 처리 시작 ---")
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

            # 안전 순회 행 데이터 확보
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
                log(f"파일#{idx+1} 경고: 데이터 행이 없음")
                continue

            # 1. 스마트 동적 헤더 탐색 (키워드 매칭 점수제)
            best_header_idx = -1
            best_score = 0

            scan_limit = min(30, len(raw_rows))
            for r_idx in range(scan_limit):
                row = raw_rows[r_idx]
                score = 0
                for cell in row:
                    if cell and any(k in cell for k in HEADER_CANDIDATES):
                        score += 1
                if score > best_score:
                    best_score = score
                    best_header_idx = r_idx

            if best_header_idx == -1 or best_score < 2:
                best_header_idx = 0

            log(f"파일#{idx+1} 진짜 헤더 발견: {best_header_idx + 1}행 (매칭 키워드 {best_score}개)")

            # 첫 번째 기준 파일에서 상단 제목/설명 및 헤더 보존
            if len(master_headers) == 0:
                for r in range(best_header_idx):
                    master_title_rows.append(raw_rows[r])
                master_headers = raw_rows[best_header_idx]

            # 2. 순수 데이터 추출
            empty_streak = 0
            file_extracted_count = 0

            for r_idx in range(best_header_idx + 1, len(raw_rows)):
                row = raw_rows[r_idx]

                # (1) 요약행 완전일치 검사 ('계', '종합', '합계' 등)
                is_footer = False
                for cell_val in row[:3]:
                    clean_val = cell_val.replace(" ", "")
                    if clean_val in FOOTER_EXACT_KEYWORDS:
                        is_footer = True
                        break

                if is_footer:
                    log(f"파일#{idx+1} [정상 종료]: {r_idx + 1}행에서 요약행('{clean_val}') 감지하여 중단")
                    break

                # (2) 행 전체 공백 검사
                non_empty_cells = [c for c in row if c != ""]
                if len(non_empty_cells) == 0:
                    empty_streak += 1
                    if empty_streak >= MAX_EMPTY_GAP:
                        log(f"파일#{idx+1} [정상 종료]: 빈 줄 {MAX_EMPTY_GAP}개 연속 감지 -> 데이터 종료")
                        break
                    continue
                else:
                    empty_streak = 0

                all_data_rows.append(row)
                file_extracted_count += 1

            log(f"파일#{idx+1} 추출 완료 (해당 파일: {file_extracted_count}행, 누적: {len(all_data_rows)}행)")

        except Exception as e:
            log(f"파일#{idx+1} 처리 실패: {str(e)}")
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

    log(f"--> [범용 엔진] 전체 완료: 총 {len(all_data_rows)}행 수합됨, 에러 {len(error_list)}건")
    return {
        "status": "success",
        "title_rows": master_title_rows,
        "headers": master_headers,
        "data_rows": all_data_rows,
        "errors": error_list
    }

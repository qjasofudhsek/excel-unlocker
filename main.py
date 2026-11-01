from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import base64
import gc
from python_calamine import CalamineWorkbook

app = FastAPI()

MAX_SCAN_ROWS = 3000   # 최대 탐색 행
MAX_SCAN_COLS = 40     # 최대 탐색 열
MAX_EMPTY_GAP = 3      # 연속 빈 줄 허용 한도

# 양식 종료 감지 키워드 ('종합' 최우선 반영)
FOOTER_KEYWORDS = ["종합", "계", "합계", "총계", "소계"]

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
    return {"status": "ok", "message": "Excel Stream Engine Ready"}

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
        log(f"파일#{idx+1} 일반 파일 처리: {str(e)}")
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
    log(f"--> 수합 요청 시작 (파일 {len(req.files_base64)}개, 시트: {req.sheet_name})")
    master_title_rows = []
    master_headers = []
    all_data_rows = []
    error_list = []

    for idx, b64_str in enumerate(req.files_base64):
        decrypted = None
        file_bytes = None
        try:
            log(f"--- 파일#{idx+1} 복호화 및 로드 시작 ---")
            file_bytes = base64.b64decode(b64_str)
            decrypted = decrypt_file_stream(file_bytes, req.password, idx)
            workbook = CalamineWorkbook.from_filelike(decrypted)

            available_sheets = workbook.sheet_names
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

            sheet = workbook.get_sheet_by_name(target_sheet)

            # 스트리밍 방식: 시트 데이터에서 필요한 앞부분만 안전하게 가져오기
            # Calamine 시트 행 탐색
            log(f"파일#{idx+1} '{target_sheet}' 스트리밍 행 파싱 중...")
            
            # to_python() 메모리 폭발 방지: calamine의 iter_rows 활용 또는 안전 순회
            if hasattr(sheet, "iter_rows"):
                row_iterator = sheet.iter_rows()
            else:
                row_iterator = sheet.to_python()

            header_idx = -1
            collected_rows = []
            empty_streak = 0

            for r_idx, row in enumerate(row_iterator):
                if r_idx >= MAX_SCAN_ROWS:
                    log(f"파일#{idx+1} 최대 행 한도({MAX_SCAN_ROWS}) 도달로 조기 종료")
                    break

                # 40열 초과 유령 셀 잘라내기
                clean_row = [str(c or "").strip() for c in list(row)[:MAX_SCAN_COLS]]

                # 헤더 탐색
                if header_idx == -1:
                    collected_rows.append(clean_row)
                    if any(h in clean_row for h in ["번호", "부서", "성명", "구분", "순번", "재적", "이름"]):
                        header_idx = r_idx
                    continue

                # --- 헤더 이후 실제 데이터 영역 검사 ---
                
                # 1. 종료 키워드 감지 ('종합', '계', '합계' 등)
                first_few_cells = "".join(clean_row[:3]).replace(" ", "")
                if any(k in first_few_cells for k in FOOTER_KEYWORDS):
                    log(f"파일#{idx+1} [종료 키워드 감지]: '{first_few_cells}' 발견으로 {r_idx+1}행에서 즉시 종료")
                    break

                # 2. 빈 행 연속 감지
                if not any(v != "" for v in clean_row):
                    empty_streak += 1
                    if empty_streak >= MAX_EMPTY_GAP:
                        log(f"파일#{idx+1} 빈 줄 {MAX_EMPTY_GAP}개 연속 감지 -> 데이터 종료 판단")
                        break
                    continue
                else:
                    empty_streak = 0

                all_data_rows.append(clean_row)

            # 헤더를 못 찾은 경우 기본값
            if header_idx == -1:
                header_idx = 0

            # 첫 번째 정상 파일 기준 제목 및 헤더 확정
            if len(master_headers) == 0 and len(collected_rows) > 0:
                for r in range(min(header_idx, len(collected_rows))):
                    master_title_rows.append(collected_rows[r])
                if header_idx < len(collected_rows):
                    master_headers = collected_rows[header_idx]

            log(f"파일#{idx+1} 추출 완료 (현재 총 누적 데이터: {len(all_data_rows)}행)")

        except Exception as e:
            log(f"파일#{idx+1} 에러: {str(e)}")
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

    log(f"--> 최종 완료: 총 {len(all_data_rows)}행 데이터 병합 성공")
    return {
        "status": "success",
        "title_rows": master_title_rows,
        "headers": master_headers,
        "data_rows": all_data_rows,
        "errors": error_list
    }

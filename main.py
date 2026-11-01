from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import base64
import gc
import sys
from python_calamine import CalamineWorkbook

app = FastAPI()

MAX_SCAN_ROWS = 3000
MAX_SCAN_COLS = 50
MAX_EMPTY_GAP = 5

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
    return {"status": "ok", "message": "Excel Diagnostic Engine Ready"}

def decrypt_file_stream(file_bytes: bytes, password: str, idx: int = 0) -> io.BytesIO:
    file_io = io.BytesIO(file_bytes)
    decrypted = io.BytesIO()
    try:
        log(f"파일#{idx+1} [1단계] msoffcrypto 객체 생성 중...")
        office_file = msoffcrypto.OfficeFile(file_io)
        
        log(f"파일#{idx+1} [2단계] 비밀번호 키 검증 및 로드 중...")
        office_file.load_key(password=password or "1234")
        
        log(f"파일#{idx+1} [3단계] 복호화 스트림 해제(decrypt) 진행 중...")
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        log(f"파일#{idx+1} [3단계 완료] 복호화 성공! (크기: {decrypted.getbuffer().nbytes} bytes)")
        return decrypted
    except Exception as e:
        log(f"파일#{idx+1} [복호화 스킵 또는 예외]: {str(e)} -> 일반 파일로 취급하여 진행")
        file_io.seek(0)
        return file_io

@app.post("/get_sheets")
async def get_sheets(req: SheetListRequest):
    decrypted = None
    try:
        log("--> /get_sheets 요청 수신")
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_file_stream(file_bytes, req.password, 0)
        workbook = CalamineWorkbook.from_filelike(decrypted)
        log(f"시트 목록 조회 성공: {workbook.sheet_names}")
        return {"status": "success", "sheets": workbook.sheet_names}
    except Exception as e:
        log(f"시트 목록 조회 중 에러 발생: {str(e)}")
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
    log(f"--> /extract_all_onestop 요청 수신 (수신된 파일 개수: {len(req.files_base64)}개, 대상 시트: {req.sheet_name})")
    master_title_rows = []
    master_headers = []
    all_data_rows = []
    error_list = []

    for idx, b64_str in enumerate(req.files_base64):
        decrypted = None
        file_bytes = None
        try:
            log(f"--- 파일#{idx+1} 처리 시작 (문자열 길이: {len(b64_str)}) ---")
            file_bytes = base64.b64decode(b64_str)
            decrypted = decrypt_file_stream(file_bytes, req.password, idx)
            
            log(f"파일#{idx+1} [4단계] Calamine 엔진으로 엑셀 워크북 파싱 시작...")
            workbook = CalamineWorkbook.from_filelike(decrypted)

            available_sheets = workbook.sheet_names
            log(f"파일#{idx+1} 파싱 성공! 포함된 시트: {available_sheets}")
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

            log(f"파일#{idx+1} [5단계] '{target_sheet}' 시트 데이터 추출 중...")
            sheet = workbook.get_sheet_by_name(target_sheet)
            raw_rows = sheet.to_python()

            if not raw_rows:
                log(f"파일#{idx+1} 경고: 시트에 행 데이터가 없음")
                continue

            # 유령 셀 차단
            sliced_rows = []
            for r in raw_rows[:MAX_SCAN_ROWS]:
                sliced_rows.append(r[:MAX_SCAN_COLS])
            del raw_rows

            header_idx = -1
            for r_idx, row in enumerate(sliced_rows):
                str_row = [str(c or "").strip() for c in row]
                if any(h in str_row for h in ["번호", "부서", "성명", "구분", "순번", "재적", "이름"]):
                    header_idx = r_idx
                    break

            if header_idx == -1:
                header_idx = 0

            if len(master_headers) == 0:
                for r in range(header_idx):
                    master_title_rows.append([str(c) if c is not None else "" for c in sliced_rows[r]])
                master_headers = [str(c) if c is not None else "" for c in sliced_rows[header_idx]]

            empty_streak = 0
            file_row_count = 0
            for row in sliced_rows[header_idx + 1:]:
                str_row = [str(c or "").strip() for c in row]
                if any(f in str_row[:3] for f in ["계", "합계", "총계", "소계"]):
                    break
                if not any(v != "" for v in str_row):
                    empty_streak += 1
                    if empty_streak >= MAX_EMPTY_GAP:
                        break
                    continue
                else:
                    empty_streak = 0

                all_data_rows.append([str(c) if c is not None else "" for c in row])
                file_row_count += 1

            log(f"파일#{idx+1} 데이터 추출 완료: 유효 {file_row_count}행 수합됨")

        except Exception as e:
            log(f"파일#{idx+1} 처리 도중 심각한 에러 발생: {str(e)}")
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

    log(f"--> 전체 수합 완료: 총 {len(all_data_rows)}행 확보, 오류 {len(error_list)}건")
    return {
        "status": "success",
        "title_rows": master_title_rows,
        "headers": master_headers,
        "data_rows": all_data_rows,
        "errors": error_list
    }

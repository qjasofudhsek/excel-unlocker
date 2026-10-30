from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import openpyxl
from copy import copy
import base64
import gc
from python_calamine import CalamineWorkbook

app = FastAPI()

class SheetListRequest(BaseModel):
    file_base64: str
    password: str = "1234"

class MergeRequest(BaseModel):
    files_base64: List[str]
    dept_names: List[str]
    sheet_name: Optional[str] = None
    password: str = "1234"

@app.get("/")
def read_root():
    return {"status": "ok", "message": "Excel Merger Engine Ready (Calamine + OpenPyXL)"}

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
        
        # Calamine으로 안전하고 빠르게 시트 목록만 조회
        workbook = CalamineWorkbook.from_filelike(decrypted)
        return {"status": "success", "sheets": workbook.sheet_names}
    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if decrypted:
            decrypted.close()
        gc.collect()

@app.post("/merge_with_template")
async def merge_with_template(req: MergeRequest):
    template_stream = None
    wb_master = None
    try:
        if not req.files_base64:
            return {"status": "error", "message": "병합할 파일이 없습니다."}

        # 1. Calamine으로 모든 파일에서 '순수 데이터'만 안전하게 추출 (조건부 서식/스타일 오류 원천 방지)
        all_data_rows = []
        target_sheet_name = req.sheet_name

        for b64, dept in zip(req.files_base64, req.dept_names):
            f_bytes = base64.b64decode(b64)
            dec_stream = decrypt_stream(f_bytes, req.password)
            
            cal_wb = CalamineWorkbook.from_filelike(dec_stream)
            s_name = target_sheet_name if target_sheet_name in cal_wb.sheet_names else cal_wb.sheet_names[0]
            sheet = cal_wb.get_sheet_by_name(s_name)
            rows = sheet.to_python()

            # 헤더 행 탐색
            header_idx = -1
            for r_idx, row in enumerate(rows):
                str_row = [str(c or "").strip() for c in row]
                if any(h in str_row for h in ["번호", "부서", "성명", "구분", "순번"]):
                    header_idx = r_idx
                    break
            
            if header_idx == -1:
                header_idx = 0

            # 실제 유효 데이터만 필터링 (합계/요약 행 만나면 즉시 중단)
            for row in rows[header_idx + 1:]:
                str_row = [str(c or "").strip() for c in row]
                if any(f in str_row[:3] for f in ["계", "합계", "총계", "소계"]):
                    break
                if not any(v != "" for v in str_row):
                    continue
                all_data_rows.append(row)

            dec_stream.close()

        # 2. 첫 번째 파일을 템플릿(Base)으로 로드하여 서식/스타일 복제
        first_bytes = base64.b64decode(req.files_base64[0])
        template_stream = decrypt_stream(first_bytes, req.password)
        wb_master = openpyxl.load_workbook(template_stream)

        t_sheet = target_sheet_name if (target_sheet_name and target_sheet_name in wb_master.sheetnames) else wb_master.sheetnames[0]
        ws_master = wb_master[t_sheet]

        # 템플릿 내 헤더 및 '계' 위치 확인
        header_row_idx = None
        footer_row_idx = None

        for r in range(1, ws_master.max_row + 1):
            row_vals = [str(ws_master.cell(row=r, column=c).value or "").strip() for c in range(1, ws_master.max_column + 1)]
            if not header_row_idx and any(h in row_vals for h in ["번호", "부서", "성명", "구분", "순번"]):
                header_row_idx = r
                continue
            if header_row_idx and any(f in row_vals[:3] for f in ["계", "합계", "총계", "소계"]):
                footer_row_idx = r
                break

        if not header_row_idx:
            header_row_idx = 1
        if not footer_row_idx:
            footer_row_idx = ws_master.max_row + 1

        # 템플릿의 기존 데이터 서식(글꼴, 테두리, 배경색) 샘플 추출
        style_sample_row = header_row_idx + 1
        cell_styles = []
        for c in range(1, ws_master.max_column + 1):
            cell = ws_master.cell(row=style_sample_row, column=c)
            cell_styles.append({
                "font": copy(cell.font),
                "border": copy(cell.border),
                "fill": copy(cell.fill),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format
            })

        # 기존 템플릿의 샘플 데이터 영역 삭제
        existing_count = max(0, footer_row_idx - (header_row_idx + 1))
        if existing_count > 0:
            ws_master.delete_rows(header_row_idx + 1, amount=existing_count)

        # 3. Calamine으로 수합한 순수 데이터를 서식 입혀서 삽입
        current_row = header_row_idx + 1
        for idx, row_data in enumerate(all_data_rows, start=1):
            ws_master.insert_rows(current_row)
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws_master.cell(row=current_row, column=c_idx)
                # 1번째 열이 번호일 경우 순번 자동 번호 부여
                if c_idx == 1 and (isinstance(val, (int, float)) or str(val).isdigit()):
                    cell.value = idx
                else:
                    cell.value = val

                # 서식 적용
                if c_idx - 1 < len(cell_styles):
                    st = cell_styles[c_idx - 1]
                    cell.font = copy(st["font"])
                    cell.border = copy(st["border"])
                    cell.fill = copy(st["fill"])
                    cell.alignment = copy(st["alignment"])
                    cell.number_format = st["number_format"]

            current_row += 1

        # 4. 하단 SUM 수식 범위 자동 업데이트
        new_footer_row = current_row
        start_row = header_row_idx + 1
        end_row = new_footer_row - 1

        for c in range(1, ws_master.max_column + 1):
            f_cell = ws_master.cell(row=new_footer_row, column=c)
            if f_cell.value and isinstance(f_cell.value, str) and f_cell.value.startswith("="):
                col_letter = openpyxl.utils.get_column_letter(c)
                f_cell.value = f"=SUM({col_letter}{start_row}:{col_letter}{end_row})"

        out_stream = io.BytesIO()
        wb_master.save(out_stream)
        out_stream.seek(0)

        result_b64 = base64.b64encode(out_stream.getvalue()).decode("utf-8")
        return {
            "status": "success",
            "merged_file_base64": result_b64,
            "total_rows": len(all_data_rows)
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        if wb_master:
            wb_master.close()
        if template_stream:
            template_stream.close()
        gc.collect()

from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import openpyxl
from copy import copy
import base64
import gc

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
    return {"status": "ok", "message": "Excel Merger Engine Ready"}

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
    try:
        file_bytes = base64.b64decode(req.file_base64)
        decrypted = decrypt_stream(file_bytes, req.password)
        wb = openpyxl.load_workbook(decrypted, read_only=True)
        sheet_names = wb.sheet_names
        wb.close()
        return {"status": "success", "sheets": sheet_names}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.post("/merge_with_template")
async def merge_with_template(req: MergeRequest):
    try:
        if not req.files_base64:
            return {"status": "error", "message": "병합할 파일이 없습니다."}

        # 1. 첫 번째 파일을 템플릿(Base)으로 로드
        first_bytes = base64.b64decode(req.files_base64[0])
        template_stream = decrypt_stream(first_bytes, req.password)
        wb_master = openpyxl.load_workbook(template_stream)

        target_sheet_name = req.sheet_name or wb_master.sheetnames[0]
        ws_master = wb_master[target_sheet_name]

        # 2. 헤더 행 및 요약/합계(계) 행 위치 탐색
        header_row_idx = None
        footer_row_idx = None

        for r in range(1, ws_master.max_row + 1):
            row_vals = [str(ws_master.cell(row=r, column=c).value or "").strip() for c in range(1, ws_master.max_column + 1)]
            # 헤더 탐색: "번호" 또는 "부서" 등의 대표 컬럼
            if not header_row_idx and any(h in row_vals for h in ["번호", "부서", "성명", "구분", "순번"]):
                header_row_idx = r
                continue

            # 합계/요약 행 탐색: "계", "합계", "총계", "소계"
            if header_row_idx and any(f in row_vals[:3] for f in ["계", "합계", "총계", "소계"]):
                footer_row_idx = r
                break

        if not header_row_idx:
            header_row_idx = 1
        if not footer_row_idx:
            footer_row_idx = ws_master.max_row + 1

        # 3. 모든 파일에서 순수 데이터 행 추출
        all_data_rows = []
        for b64, dept in zip(req.files_base64, req.dept_names):
            f_bytes = base64.b64decode(b64)
            dec_stream = decrypt_stream(f_bytes, req.password)
            wb_sub = openpyxl.load_workbook(dec_stream, data_only=True)
            ws_sub = wb_sub[target_sheet_name] if target_sheet_name in wb_sub.sheetnames else wb_sub.active

            for r in range(header_row_idx + 1, ws_sub.max_row + 1):
                row_vals = [ws_sub.cell(row=r, column=c).value for c in range(1, ws_sub.max_column + 1)]
                str_vals = [str(v or "").strip() for v in row_vals]

                # 합계 행 도달 시 중단
                if any(f in str_vals[:3] for f in ["계", "합계", "총계", "소계"]):
                    break
                # 빈 줄 건너뛰기
                if not any(v != "" for v in str_vals):
                    continue

                all_data_rows.append(row_vals)
            wb_sub.close()

        # 4. 템플릿의 기존 데이터 행 서식 복사 및 데이터 갱신
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

        # 기존 템플릿의 샘플/더미 데이터 행 삭제
        existing_data_count = max(0, footer_row_idx - (header_row_idx + 1))
        if existing_data_count > 0:
            ws_master.delete_rows(header_row_idx + 1, amount=existing_data_count)

        # 수합된 전체 데이터 삽입
        current_write_row = header_row_idx + 1
        for idx, row_data in enumerate(all_data_rows, start=1):
            ws_master.insert_rows(current_write_row)
            for c_idx, val in enumerate(row_data, start=1):
                cell = ws_master.cell(row=current_write_row, column=c_idx)
                # 1번째 열이 '번호'일 경우 순번 자동 부여
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

            current_write_row += 1

        # 5. 하단 합계 수식 범위 자동 업데이트
        new_footer_row = current_write_row
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
        wb_master.close()

        result_b64 = base64.b64encode(out_stream.getvalue()).decode("utf-8")
        return {
            "status": "success",
            "merged_file_base64": result_b64,
            "total_rows": len(all_data_rows)
        }

    except Exception as e:
        return {"status": "error", "message": str(e)}
    finally:
        gc.collect()

from fastapi import FastAPI
from pydantic import BaseModel
from typing import List, Optional
import msoffcrypto
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
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
    return {"status": "ok", "message": "Excel Merger Engine Ready (Calamine + Safe Builder)"}

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
    try:
        if not req.files_base64:
            return {"status": "error", "message": "병합할 파일이 없습니다."}

        all_data_rows = []
        master_title_rows = []
        master_headers = []
        target_sheet_name = req.sheet_name

        # 1. 모든 파일을 calamine으로 파싱하여 XML 손상 원천 차단
        for idx, (b64, dept) in enumerate(zip(req.files_base64, req.dept_names)):
            f_bytes = base64.b64decode(b64)
            dec_stream = decrypt_stream(f_bytes, req.password)
            
            cal_wb = CalamineWorkbook.from_filelike(dec_stream)
            s_name = target_sheet_name if (target_sheet_name and target_sheet_name in cal_wb.sheet_names) else cal_wb.sheet_names[0]
            sheet = cal_wb.get_sheet_by_name(s_name)
            rows = sheet.to_python()

            # 헤더 위치 탐색 ("번호", "부서", "재적" 등)
            header_idx = -1
            for r_idx, row in enumerate(rows):
                str_row = [str(c or "").strip() for c in row]
                if any(h in str_row for h in ["번호", "부서", "성명", "구분", "순번", "재적"]):
                    header_idx = r_idx
                    break
            
            if header_idx == -1:
                header_idx = 0

            # 첫 번째 파일에서 상단 제목/설명 및 헤더 보존
            if idx == 0:
                for r in range(header_idx):
                    master_title_rows.append(rows[r])
                master_headers = rows[header_idx]

            # 실제 순수 데이터 추출 (하단 "계", "합계" 만나면 중단)
            for row in rows[header_idx + 1:]:
                str_row = [str(c or "").strip() for c in row]
                if any(f in str_row[:3] for f in ["계", "합계", "총계", "소계"]):
                    break
                # 빈 줄 제외
                if not any(v != "" for v in str_row):
                    continue
                all_data_rows.append(row)

            dec_stream.close()

        if not master_headers or len(all_data_rows) == 0:
            return {"status": "error", "message": "유효한 데이터 행을 추출하지 못했습니다."}

        # 2. openpyxl로 완전히 깨끗한 새 워크북 생성 (XML 깨짐 방지)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = target_sheet_name if target_sheet_name else "수합결과"

        # 스타일 정의
        font_header = Font(name="맑은 고딕", size=10, bold=True)
        font_data = Font(name="맑은 고딕", size=10)
        font_total = Font(name="맑은 고딕", size=10, bold=True)
        fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        fill_total = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
        
        thin = Side(border_style="thin", color="D3D3D3")
        box_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        double_bottom = Border(left=thin, right=thin, top=thin, bottom=Side(border_style="double", color="000000"))

        current_row = 1

        # 상단 제목/설명 행 작성
        for t_row in master_title_rows:
            for c_idx, val in enumerate(t_row, start=1):
                if val is not None and str(val).strip() != "":
                    ws.cell(row=current_row, column=c_idx, value=val)
            current_row += 1

        # 헤더 작성
        header_row_num = current_row
        col_count = len(master_headers)
        for c_idx, val in enumerate(master_headers, start=1):
            cell = ws.cell(row=current_row, column=c_idx, value=val)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = box_border
        current_row += 1

        # 순수 데이터 행 작성
        data_start_row = current_row
        for num_idx, row in enumerate(all_data_rows, start=1):
            for c_idx in range(1, col_count + 1):
                cell = ws.cell(row=current_row, column=c_idx)
                val = row[c_idx - 1] if c_idx - 1 < len(row) else ""

                # 1열 번호 자동 재정렬
                if c_idx == 1 and (isinstance(val, (int, float)) or str(val).isdigit()):
                    cell.value = num_idx
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.value = val
                    if isinstance(val, (int, float)):
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                cell.font = font_data
                cell.border = box_border
            current_row += 1
        data_end_row = current_row - 1

        # 하단 '계' 합계행 작성
        ws.cell(row=current_row, column=1, value="계").alignment = Alignment(horizontal="center", vertical="center")
        
        # 합계 수식(SUM) 자동 연결
        for c_idx in range(1, col_count + 1):
            cell = ws.cell(row=current_row, column=c_idx)
            cell.font = font_total
            cell.fill = fill_total
            cell.border = double_bottom
            
            # 숫자 열 감지 (데이터 영역의 값이 숫자인 경우 SUM 수식 적용)
            if c_idx > 2:
                col_letter = get_column_letter(c_idx)
                # 데이터 첫 행이 숫자인지 확인
                sample_val = ws.cell(row=data_start_row, column=c_idx).value
                if isinstance(sample_val, (int, float)) or (isinstance(sample_val, str) and sample_val.replace(",", "").isdigit()):
                    cell.value = f"=SUM({col_letter}{data_start_row}:{col_letter}{data_end_row})"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        # 열 너비 자동 조정
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

        # 바이너리 변환
        out_stream = io.BytesIO()
        wb.save(out_stream)
        out_stream.seek(0)
        wb.close()

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

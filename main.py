import io
import msoffcrypto
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import Response

app = FastAPI()

@app.get("/")
def check():
    return {"status": "열쇠공 서버가 정상 작동 중입니다."}

@app.post("/decrypt")
async def decrypt(file: UploadFile = File(...), password: str = Form(...)):
    try:
        content = await file.read()
        f_in = io.BytesIO(content)
        office_file = msoffcrypto.OfficeFile(f_in)

        # 암호가 애초에 안 걸려 있는 파일이면 그대로 반환
        if not office_file.is_encrypted():
            return Response(content=content, media_type="application/octet-stream")

        # 암호 해제 시도
        office_file.load_key(password=password)
        f_out = io.BytesIO()
        office_file.decrypt(f_out)
        return Response(content=f_out.getvalue(), media_type="application/octet-stream")
    except Exception as e:
        raise HTTPException(status_code=400, detail="비밀번호가 틀렸거나 손상된 파일입니다.")

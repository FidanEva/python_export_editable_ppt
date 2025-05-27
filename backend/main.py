from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import FileResponse
from services.excel_parser import parse_excel_data
from services.ppt_generator import create_ppt
import os
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],  # React dev server
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
@app.post("/generate-ppt/")
async def generate_ppt(
    excel_files: list[UploadFile] = File(...),
    image: UploadFile = File(...),
    custom_text: str = Form(...)
):
    os.makedirs("uploads", exist_ok=True)
    data_frames = []

    # Save and parse each Excel file
    for file in excel_files:
        path = f"uploads/{file.filename}"
        with open(path, "wb") as f:
            f.write(await file.read())
        df = parse_excel_data(path)
        data_frames.append(df)

    # Save image
    img_path = f"uploads/{image.filename}"
    with open(img_path, "wb") as f:
        f.write(await image.read())

    # Generate PPT
    output_path = "uploads/output.pptx"
    create_ppt(data_frames, img_path, custom_text, output_path)

    return FileResponse(output_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename="presentation.pptx")

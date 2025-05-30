from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import FileResponse
from services.excel_parser import parse_excel_data
from services.ppt_generator import create_ppt
import os
from fastapi.middleware.cors import CORSMiddleware
import traceback
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],  # React dev server
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    # expose_headers=["Content-Disposition"]
)

# Create uploads directory if it doesn't exist
os.makedirs("uploads", exist_ok=True)

@app.post("/generate-ppt/")
async def generate_ppt(
    excel_files: list[UploadFile] = File(...),
    # image: UploadFile = File(...),
    # custom_text: str = Form(...),
    date: str = Form(...),
    company_name: str = Form(...)
):
    try:
        logger.debug("Starting PPT generation process")
        data_frames = {}

        # Save and parse each Excel file
        for file in excel_files:
            try:
                logger.debug(f"Processing file: {file.filename}")
                path = f"uploads/{file.filename}"
                content = await file.read()
                
                # Ensure the uploads directory exists
                os.makedirs(os.path.dirname(path), exist_ok=True)
                
                with open(path, "wb") as f:
                    f.write(content)
                
                # Extract file type from filename
                file_type = file.filename.split('.')[0]
                logger.debug(f"Parsing Excel file: {file_type}")
                data_frames[file_type] = parse_excel_data(path)
                logger.debug(f"Successfully parsed {file_type}")
            except Exception as e:
                logger.error(f"Error processing file {file.filename}: {str(e)}")
                logger.error(traceback.format_exc())
                raise HTTPException(status_code=400, detail=f"Error processing file {file.filename}: {str(e)}")

        # Save image
        # try:
        #     logger.debug("Processing image")
        #     img_path = f"uploads/{image.filename}"
        #     content = await image.read()
            
        #     # Ensure the uploads directory exists
        #     os.makedirs(os.path.dirname(img_path), exist_ok=True)
            
        #     with open(img_path, "wb") as f:
        #         f.write(content)
        #     logger.debug("Image saved successfully")
        # except Exception as e:
        #     logger.error(f"Error processing image: {str(e)}")
        #     logger.error(traceback.format_exc())
        #     raise HTTPException(status_code=400, detail=f"Error processing image: {str(e)}")

        # Generate PPT
        try:
            logger.debug("Generating PowerPoint")
            output_path = "uploads/output.pptx"
            create_ppt(data_frames, output_path, date, company_name)
            logger.debug("PowerPoint generated successfully")
        except Exception as e:
            logger.error(f"Error generating PPT: {str(e)}")
            logger.error(traceback.format_exc())
            raise HTTPException(status_code=500, detail=f"Error generating PPT: {str(e)}")

        # Check if the output file exists
        if not os.path.exists(output_path):
            logger.error("Output file was not created")
            raise HTTPException(status_code=500, detail="Failed to generate PowerPoint file")

        logger.debug("Returning PowerPoint file")
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="presentation.pptx",
            headers={"Content-Disposition": "attachment; filename=presentation.pptx"}
        )
    except HTTPException as he:
        logger.error(f"HTTP Exception: {str(he)}")
        raise he
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"Unexpected error: {str(e)}")

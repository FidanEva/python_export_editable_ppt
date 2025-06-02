from fastapi import FastAPI, File, UploadFile, Form, HTTPException, Request, BackgroundTasks
from fastapi.responses import FileResponse
from services.excel_parser import parse_excel_data
from services.ppt_generator import create_ppt
import os
from fastapi.middleware.cors import CORSMiddleware
import traceback
import logging
import json
import shutil
import asyncio
from typing import List

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
)

# Create uploads directory if it doesn't exist
os.makedirs("uploads", exist_ok=True)

def cleanup_files(file_paths: List[str]):
    """Background task to clean up temporary files"""
    for file_path in file_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logger.debug(f"Cleaned up temporary file: {file_path}")
        except Exception as e:
            logger.warning(f"Could not delete temporary file {file_path}: {str(e)}")

@app.post("/generate-ppt/")
async def generate_ppt(
    background_tasks: BackgroundTasks,
    request: Request,
    excel_files: list[UploadFile] = File(...),
    positive_links: str = Form(None),
    negative_links: str = Form(None),
    date: str = Form(...),
    company_name: str = Form(...)
):
    temp_files = []  # Keep track of all temporary files
    try:
        # Parse links
        positive_links_list = json.loads(positive_links) if positive_links else []
        negative_links_list = json.loads(negative_links) if negative_links else []

        # Process Excel files
        data_frames = {}
        for file in excel_files:
            file_path = f"uploads/{file.filename}"
            temp_files.append(file_path)  # Add to tracking list
            
            # Save file
            with open(file_path, "wb") as buffer:
                content = await file.read()
                buffer.write(content)
            
            # Parse Excel data
            data_frames[file.filename.split('.')[0]] = parse_excel_data(file_path)

        # Process post images
        positive_posts = []
        negative_posts = []
        
        # Get all form fields
        form_data = await request.form()
        
        # Process positive posts
        index = 0
        while f"positive_post_image_{index}" in form_data:
            image = form_data[f"positive_post_image_{index}"]
            link = form_data[f"positive_post_link_{index}"]
            if image:
                file_path = f"uploads/positive_post_{index}.jpg"
                temp_files.append(file_path)  # Add to tracking list
                with open(file_path, "wb") as buffer:
                    content = await image.read()
                    buffer.write(content)
                positive_posts.append({"image_path": file_path, "link": link})
            index += 1

        # Process negative posts
        index = 0
        while f"negative_post_image_{index}" in form_data:
            image = form_data[f"negative_post_image_{index}"]
            link = form_data[f"negative_post_link_{index}"]
            if image:
                file_path = f"uploads/negative_post_{index}.jpg"
                temp_files.append(file_path)  # Add to tracking list
                with open(file_path, "wb") as buffer:
                    content = await image.read()
                    buffer.write(content)
                negative_posts.append({"image_path": file_path, "link": link})
            index += 1

        # Generate PowerPoint
        output_path = "uploads/report.pptx"
        temp_files.append(output_path)  # Add to tracking list
        
        create_ppt(
            data_frames=data_frames,
            output_path=output_path,
            date=date,
            company_name=company_name,
            positive_links=positive_links_list,
            negative_links=negative_links_list,
            positive_posts=positive_posts,
            negative_posts=negative_posts
        )

        # Create a copy of the output file for response
        response_path = "uploads/response_report.pptx"
        shutil.copy2(output_path, response_path)
        temp_files.append(response_path)  # Add to tracking list

        # Schedule cleanup for after response is sent
        background_tasks.add_task(cleanup_files, temp_files)

        # Return the response file
        return FileResponse(
            response_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename="report.pptx",
            background=background_tasks
        )

    except Exception as e:
        # Clean up files in case of error
        cleanup_files(temp_files)
        logger.error(f"Error generating PowerPoint: {str(e)}")
        logger.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=str(e))

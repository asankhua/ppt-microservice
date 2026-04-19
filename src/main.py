"""
Generic PPT Generation Microservice
Generates professional PowerPoint presentations from any JSON data
"""

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import os
import tempfile
import uuid
from datetime import datetime
import logging

from ppt_generator import PPTGenerator

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Generic PPT Generation Service",
    description="Professional PowerPoint generation API - works with any JSON data",
    version="3.0.0"
)

ppt_gen = PPTGenerator()

# Data Models
class GenerateRequest(BaseModel):
    projectName: str
    projectDescription: Optional[str] = None
    content: Any  # Accepts any JSON structure
    template: str = "professional"

class GenerateResponse(BaseModel):
    success: bool
    message: str
    downloadUrl: Optional[str] = None

@app.get("/health")
def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

@app.get("/")
def root():
    return {
        "service": "Generic PPT Generation Service",
        "version": "3.0.0",
        "description": "Generate presentations from any JSON data",
        "endpoints": ["/health", "/generate", "/templates"]
    }

@app.get("/templates")
def list_templates():
    return {
        "templates": [
            {"id": "professional", "name": "Professional Blue"},
            {"id": "minimal", "name": "Minimal White"},
            {"id": "dark", "name": "Dark Modern"},
            {"id": "startup", "name": "Startup Pitch"}
        ]
    }

@app.post("/generate", response_model=GenerateResponse)
def generate_presentation(request: GenerateRequest):
    try:
        logger.info(f"Generating PPT for: {request.projectName}")
        logger.info(f"Content type: {type(request.content)}")
        
        # Generate unique filename
        file_id = str(uuid.uuid4())[:8]
        filename = f"{request.projectName.replace(' ', '_')}_{file_id}.pptx"
        output_path = os.path.join(tempfile.gettempdir(), filename)
        
        # Generate presentation from any content structure
        ppt_gen.create_presentation(
            project_name=request.projectName,
            project_description=request.projectDescription,
            steps=request.content,
            template=request.template,
            output_path=output_path
        )
        
        logger.info(f"Generated: {output_path}")
        
        return GenerateResponse(
            success=True,
            message="Presentation generated successfully",
            downloadUrl=f"/download/{filename}"
        )
        
    except Exception as e:
        import traceback
        logger.error(f"Generation failed: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{filename}")
def download_file(filename: str):
    file_path = os.path.join(tempfile.gettempdir(), filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

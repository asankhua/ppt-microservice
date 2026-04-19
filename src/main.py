"""
PPT Generation Microservice
Generates professional PowerPoint presentations from structured data
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
    title="PPT Generation Service",
    description="Professional PowerPoint generation API",
    version="2.0.0"
)

ppt_gen = PPTGenerator()

# Data Models
class StepData(BaseModel):
    stepId: int
    stepName: str
    data: Dict[str, Any]

class GenerateRequest(BaseModel):
    projectName: str
    projectDescription: Optional[str] = None
    steps: List[StepData]
    template: str = "professional"
    includeCharts: bool = True

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
        "service": "PPT Generation Service",
        "version": "2.0.0",
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
        
        # Generate unique filename
        file_id = str(uuid.uuid4())[:8]
        filename = f"{request.projectName.replace(' ', '_')}_{file_id}.pptx"
        output_path = os.path.join(tempfile.gettempdir(), filename)
        
        # Convert Pydantic models to dicts for generator
        steps_data = [step.model_dump() for step in request.steps]
        
        # Generate presentation
        ppt_gen.create_presentation(
            project_name=request.projectName,
            project_description=request.projectDescription,
            steps=steps_data,
            template=request.template,
            output_path=output_path
        )
        
        logger.info(f"Generated: {output_path}")
        
        return GenerateResponse(
            success=True,
            message=f"Presentation generated with {len(request.steps)} sections",
            downloadUrl=f"/download/{filename}"
        )
        
    except Exception as e:
        logger.error(f"Generation failed: {str(e)}")
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

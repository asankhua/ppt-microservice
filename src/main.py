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

def normalize_data(obj):
    """Recursively normalize data - convert dicts with numeric keys to lists"""
    if isinstance(obj, dict):
        # Check if keys are numeric (indicating array serialized as object)
        keys = list(obj.keys())
        if keys and all(str(k).isdigit() for k in keys):
            # Convert to list, sort by key
            return [normalize_data(obj[k]) for k in sorted(keys, key=int)]
        else:
            # Regular dict - normalize values
            return {k: normalize_data(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [normalize_data(item) for item in obj]
    else:
        return obj

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
        
        # Convert Pydantic models to dicts and normalize (handle objects→arrays)
        steps_data = [normalize_data(step.model_dump()) for step in request.steps]
        
        # Debug: Log steps data structure
        logger.info(f"Steps count: {len(steps_data)}")
        for i, step in enumerate(steps_data[:2]):
            logger.info(f"Step {i}: type={type(step)}, keys={list(step.keys()) if isinstance(step, dict) else 'N/A'}")
            if isinstance(step, dict) and 'data' in step:
                data = step['data']
                logger.info(f"  data type: {type(data)}, keys={list(data.keys())[:5] if isinstance(data, dict) else 'N/A'}")
        
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
        import traceback
        logger.error(f"Generation failed: {str(e)}")
        logger.error(f"Traceback: {traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"{str(e)} - Check server logs for details")

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

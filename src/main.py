"""
Product Pilot - PPT Generation Microservice
Generates professional PowerPoint presentations from project data
"""

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import os
import tempfile
import uuid
from datetime import datetime
import logging

from ppt_generator import PPTGenerator

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Product Pilot PPT Service",
    description="Professional PowerPoint generation from product data",
    version="1.0.0"
)

# PPT Generator instance
ppt_gen = PPTGenerator()


# Pydantic Models
class ReframeData(BaseModel):
    problemTitle: Optional[str] = None
    reframedProblem: Optional[str] = None
    rootCauses: Optional[List[str]] = []
    userImpact: Optional[str] = None
    opportunitySize: Optional[str] = None


class VisionData(BaseModel):
    visionStatement: Optional[str] = None
    productName: Optional[str] = None
    elevatorPitch: Optional[str] = None
    targetAudience: Optional[str] = None
    valueProposition: Optional[str] = None
    successMetrics: Optional[List[str]] = []


class PersonaData(BaseModel):
    name: str
    role: str
    bio: str
    painPoints: Optional[List[str]] = []
    goals: Optional[List[str]] = []
    techSavviness: Optional[str] = None


class QuestionData(BaseModel):
    question: str
    category: Optional[str] = None
    aiAnswer: Optional[str] = None
    userAnswer: Optional[str] = None


class CompetitorData(BaseModel):
    name: str
    description: Optional[str] = None
    strengths: Optional[List[str]] = []
    weaknesses: Optional[List[str]] = []


class MarketData(BaseModel):
    marketOverview: Optional[str] = None
    competitors: Optional[List[CompetitorData]] = []
    tam: Optional[str] = None
    sam: Optional[str] = None
    som: Optional[str] = None


class UserStoryData(BaseModel):
    id: Optional[str] = None
    asA: Optional[str] = None
    iWant: Optional[str] = None
    soThen: Optional[str] = None
    description: Optional[str] = None
    priority: Optional[str] = None
    storyPoints: Optional[int] = None
    riceScore: Optional[float] = None


class PhaseData(BaseModel):
    name: str
    duration: Optional[str] = None
    milestones: Optional[List[str]] = []


class OKRData(BaseModel):
    objective: str
    keyResults: Optional[List[str]] = []


class ProjectStep(BaseModel):
    stepId: int
    stepName: str
    data: Dict[str, Any]


class PresentationRequest(BaseModel):
    projectName: str
    projectDescription: Optional[str] = None
    steps: List[ProjectStep]
    template: Optional[str] = "professional"  # professional, minimal, dark
    includeCharts: Optional[bool] = True
    includeImages: Optional[bool] = True


class PresentationResponse(BaseModel):
    success: bool
    message: str
    filePath: Optional[str] = None
    downloadUrl: Optional[str] = None


@app.get("/")
async def root():
    return {
        "service": "Product Pilot PPT Service",
        "version": "1.0.0",
        "endpoints": {
            "health": "/health",
            "generate": "/generate",
            "templates": "/templates"
        }
    }


@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}


@app.get("/templates")
async def list_templates():
    """List available presentation templates"""
    return {
        "templates": [
            {
                "id": "professional",
                "name": "Professional Blue",
                "description": "Clean corporate style with blue gradients",
                "slides": 10
            },
            {
                "id": "minimal",
                "name": "Minimal White",
                "description": "Minimalist design with plenty of white space",
                "slides": 10
            },
            {
                "id": "dark",
                "name": "Dark Modern",
                "description": "Dark theme with vibrant accents",
                "slides": 10
            },
            {
                "id": "startup",
                "name": "Startup Pitch",
                "description": "Bold, colorful design for investor pitches",
                "slides": 10
            }
        ]
    }


@app.post("/generate", response_model=PresentationResponse)
async def generate_presentation(request: PresentationRequest):
    """
    Generate a PowerPoint presentation from project data
    
    - **projectName**: Name of the project
    - **steps**: List of completed pipeline steps with data
    - **template**: Template style (professional, minimal, dark, startup)
    - **includeCharts**: Whether to include data visualizations
    - **includeImages**: Whether to include placeholder images
    """
    try:
        logger.info(f"Generating PPT for project: {request.projectName}")
        
        # Generate unique filename
        file_id = str(uuid.uuid4())[:8]
        filename = f"{request.projectName.replace(' ', '_')}_{file_id}.pptx"
        
        # Use temp directory for output
        output_dir = tempfile.gettempdir()
        output_path = os.path.join(output_dir, filename)
        
        # Generate presentation
        ppt_gen.create_presentation(
            project_name=request.projectName,
            project_description=request.projectDescription,
            steps=request.steps,
            template=request.template,
            output_path=output_path,
            include_charts=request.includeCharts,
            include_images=request.includeImages
        )
        
        logger.info(f"Presentation generated: {output_path}")
        
        return PresentationResponse(
            success=True,
            message=f"Presentation generated with {len(request.steps)} sections",
            filePath=output_path,
            downloadUrl=f"/download/{filename}"
        )
        
    except Exception as e:
        logger.error(f"Failed to generate presentation: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/download/{filename}")
async def download_presentation(filename: str):
    """Download a generated presentation"""
    output_dir = tempfile.gettempdir()
    file_path = os.path.join(output_dir, filename)
    
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


@app.post("/preview")
async def preview_slide(request: PresentationRequest):
    """Generate a preview of the first slide"""
    try:
        # Generate preview (first slide only)
        file_id = str(uuid.uuid4())[:8]
        filename = f"preview_{file_id}.pptx"
        output_path = os.path.join(tempfile.gettempdir(), filename)
        
        ppt_gen.create_presentation(
            project_name=request.projectName,
            project_description=request.projectDescription,
            steps=[request.steps[0]] if request.steps else [],
            template=request.template,
            output_path=output_path,
            preview_mode=True
        )
        
        return FileResponse(
            path=output_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

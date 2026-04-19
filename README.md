# Product Pilot - PPT Generation Microservice

Professional PowerPoint presentation generation service built with Python, FastAPI, and python-pptx.

## Features

- 🎨 **4 Professional Templates**: Professional Blue, Minimal White, Dark Modern, Startup Pitch
- 📊 **Matplotlib Charts**: TAM/SAM/SOM funnel charts, progress charts for OKRs
- 🎯 **9 Pipeline Step Slides**: Dedicated layouts for each Product Pilot step
- 🎴 **Visual Cards**: Persona cards, feature cards, info boxes
- 📐 **16:9 Aspect Ratio**: Modern widescreen format
- 🖼️ **Smart Layouts**: Two-column, cards, tables, bullet lists

## Templates

| Template | Description | Primary Color |
|----------|-------------|---------------|
| `professional` | Clean corporate style | Blue |
| `minimal` | Minimalist with whitespace | Gray |
| `dark` | Modern dark theme | Indigo |
| `startup` | Bold, colorful pitch deck | Rose/Orange |

## API Endpoints

### Health Check
```
GET /health
```

### List Templates
```
GET /templates
```

### Generate Presentation
```
POST /generate
Content-Type: application/json

{
  "projectName": "Food Delivery App",
  "projectDescription": "On-demand food delivery platform",
  "steps": [
    {
      "stepId": 1,
      "stepName": "Problem Reframe",
      "data": {
        "problemTitle": "Expensive Food Delivery",
        "reframedProblem": "How might we reduce delivery costs...",
        "rootCauses": ["High commission fees", "Lack of direct connection"],
        "userImpact": "Restaurants lose 30% margin",
        "opportunitySize": "$50B market"
      }
    }
  ],
  "template": "professional",
  "includeCharts": true,
  "includeImages": true
}
```

### Download Presentation
```
GET /download/{filename}
```

## Running Locally

### Option 1: Direct Python
```bash
cd ppt-service
pip install -r requirements.txt
cd src
python main.py
```

### Option 2: Docker
```bash
cd ppt-service
docker-compose up --build
```

Service will be available at `http://localhost:8000`

## API Documentation

Once running, access interactive docs:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Slide Types

### Step 1: Problem Reframe
- Problem title and reframed statement
- Root causes bullet list
- User impact and opportunity cards

### Step 2: Product Vision
- Vision statement (prominent)
- Elevator pitch and target audience
- Value proposition card
- Success metrics

### Step 3: User Personas
- Up to 3 persona cards side-by-side
- Bio, pain points, goals
- Color-coded indicators

### Step 4: Questions & Answers
- Two-column Q&A layout
- Up to 6 questions
- AI/user answers

### Step 5: Market Analysis
- Market overview text
- Competitor comparison table
- TAM/SAM/SOM funnel chart

### Step 6: PRD Features
- Feature list with priority badges
- Color-coded (Red=High, Orange=Medium, Green=Low)

### Step 7: User Stories
- RICE-scored table
- ID, description, priority, points, RICE score

### Step 8: Roadmap
- Timeline with numbered phases
- Duration and milestones

### Step 9: OKRs
- North star definition
- Objectives with key results
- Progress bar visualization

## Architecture

```
┌─────────────────────────────────────────────────────────┐
│                  Product Pilot Frontend                  │
│                   (Next.js + PptxGenJS)                  │
└──────────────────────┬──────────────────────────────────┘
                       │ POST /generate
                       │
┌──────────────────────▼──────────────────────────────────┐
│              PPT Microservice (Python)                   │
│  ┌─────────────┐  ┌──────────────┐  ┌─────────────────┐ │
│  │   FastAPI   │  │  python-pptx │  │   matplotlib    │ │
│  │   Server    │  │  Generator   │  │    Charts       │ │
│  └─────────────┘  └──────────────┘  └─────────────────┘ │
└──────────────────────┬──────────────────────────────────┘
                       │
                       ▼
              ┌────────────────┐
              │   .pptx File   │
              │   (Download)   │
              └────────────────┘
```

## Integration with Product Pilot

To integrate this service with Product Pilot frontend:

1. **Update Next.js API Route** (`/api/ppt/generate`):
```typescript
// Instead of client-side PptxGenJS, call the microservice
const response = await fetch('http://ppt-service:8000/generate', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    projectName,
    steps: formattedSteps,
    template: 'professional'
  })
});

const blob = await response.blob();
// Trigger download
```

2. **Docker Compose Integration**:
```yaml
# Add to root docker-compose.yml
services:
  next-app:
    # ... existing config
    depends_on:
      - ppt-service
  
  ppt-service:
    build: ./ppt-service
    ports:
      - "8000:8000"
```

## Development

### Adding New Templates
1. Add color scheme to `COLOR_SCHEMES` in `ppt_generator.py`
2. Update `/templates` endpoint
3. Test with different step data

### Adding New Slide Types
1. Create `_add_*_slide` method in `PPTGenerator` class
2. Handle in `_add_step_slide` dispatcher
3. Add data model in `main.py`

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT` | 8000 | Server port |
| `HOST` | 0.0.0.0 | Server host |
| `LOG_LEVEL` | INFO | Logging level |

## License

MIT License - Part of Product Pilot

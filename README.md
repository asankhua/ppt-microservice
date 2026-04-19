---
title: PPT Generation Service
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 8000
---

# PPT Generation Microservice

Professional PowerPoint generation service built with Python, FastAPI, and python-pptx.

## Features

- 4 templates: Professional, Minimal, Dark, Startup
- 9 pipeline step layouts
- Clean, robust data handling

## API Endpoints

- `GET /health` - Health check
- `POST /generate` - Generate presentation
- `GET /download/{filename}` - Download file

## Request Format

```json
POST /generate
{
  "projectName": "My Product",
  "projectDescription": "Description",
  "steps": [
    {
      "stepId": 1,
      "stepName": "Problem Reframe",
      "data": {
        "problemTitle": "...",
        "reframedProblem": "...",
        "rootCauses": ["..."]
      }
    }
  ],
  "template": "professional"
}
```

## Running Locally

```bash
pip install -r requirements.txt
cd src && python main.py
```

## Docker

```bash
docker build -t ppt-service .
docker run -p 8000:8000 ppt-service
```

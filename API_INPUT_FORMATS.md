# PPT Generation Service - Input Formats

## Overview
This generic PPT generation service accepts **any JSON structure** and converts it into professional PowerPoint presentations.

---

## Format 1: Simple Object (Recommended)

Best for: Clean, organized presentations with named sections

```json
{
  "projectName": "Spotify Ad Experience",
  "projectDescription": "Product strategy presentation",
  "template": "professional",
  "content": {
    "Problem Statement": {
      "Title": "Enhancing Spotify's Ad Experience",
      "Summary": "Spotify's ad experience is marred by frequent, irrelevant ads...",
      "Root Causes": [
        "Current ad targeting and delivery mechanisms",
        "Ad blocking tools",
        "Lack of engaging ad content"
      ],
      "Impact": {
        "User Impact": "Reduced satisfaction and engagement",
        "Business Impact": "Lower ad revenue affecting free service"
      }
    },
    "Product Vision": {
      "Vision Statement": "Empower every artist to build a sustainable career",
      "Elevator Pitch": "Zero-commission platform for direct artist-fan connections"
    },
    "User Personas": [
      {
        "Name": "Maria",
        "Role": "Restaurant Owner",
        "Bio": "Owner of family restaurant in Chicago"
      },
      {
        "Name": "John",
        "Role": "Artist",
        "Bio": "Independent musician building fanbase"
      }
    ],
    "Market Analysis": {
      "TAM": "$50B",
      "SAM": "$12B",
      "SOM": "$500M",
      "Key Competitors": ["Competitor A", "Competitor B"]
    }
  }
}
```

---

## Format 2: Array of Sections (Step-like)

Best for: Sequential/step-by-step presentations

```json
{
  "projectName": "Product Strategy",
  "template": "startup",
  "content": [
    {
      "title": "Problem Reframe",
      "data": {
        "problem": "Expensive delivery fees",
        "solution": "Direct restaurant-to-consumer platform"
      }
    },
    {
      "title": "User Personas",
      "data": {
        "personas": [
          {"name": "Maria", "role": "Owner"},
          {"name": "John", "role": "Consumer"}
        ]
      }
    },
    {
      "title": "Market Size",
      "data": {
        "TAM": "$50B",
        "SAM": "$12B"
      }
    }
  ]
}
```

---

## Format 3: Legacy Format (Backward Compatible)

Best for: Existing integrations with the old 9-step pipeline

```json
{
  "projectName": "Product Strategy",
  "template": "professional",
  "content": [
    {
      "stepId": 1,
      "stepName": "Problem Reframe",
      "data": {
        "problemTitle": "Expensive Delivery",
        "reframedProblem": "How might we reduce costs?",
        "rootCauses": [
          "High delivery fees",
          "No direct connection"
        ]
      }
    },
    {
      "stepId": 2,
      "stepName": "Product Vision",
      "data": {
        "visionStatement": "Empower restaurants with direct orders",
        "elevatorPitch": "Zero-commission platform"
      }
    },
    {
      "stepId": 3,
      "stepName": "User Personas",
      "data": {
        "personas": [
          {
            "name": "Maria",
            "role": "Restaurant Owner",
            "bio": "Owner of family restaurant"
          }
        ]
      }
    }
  ]
}
```

---

## Field Reference

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `projectName` | string | ✅ | Presentation title (appears on cover slide) |
| `projectDescription` | string | ❌ | Subtitle/description (optional) |
| `template` | string | ❌ | Color theme: `professional`, `minimal`, `dark`, `startup` (default: `professional`) |
| `content` | any | ✅ | **Any valid JSON** - object, array, or mixed structure |

---

## Content Rendering Rules

The service automatically renders content based on its type:

| Content Type | Rendered As |
|--------------|-------------|
| `{"key": "string value"}` | Label + text on slide |
| `{"key": ["a", "b", "c"]}` | Section title with bullet list |
| `{"key": {"nested": "object"}}` | Section with nested key-value pairs |
| `["item1", "item2"]` | Simple bullet list slide |
| `[{"name": "...", "description": "..."}]` | Cards with extracted fields |
| `{"0": {...}, "1": {...}}` | Auto-converted to array (handles JS object format) |

---

## Template Options

| Template | Style |
|----------|-------|
| `professional` | Blue theme, corporate look |
| `minimal` | Clean white, simple design |
| `dark` | Dark background, modern feel |
| `startup` | Vibrant colors, energetic |

---

## API Endpoint

```
POST https://ashishsankhua-ppt-microservice.hf.space/generate
Content-Type: application/json
```

## Response Format

```json
{
  "success": true,
  "message": "Presentation generated successfully",
  "downloadUrl": "/download/Project_Name_abc123.pptx"
}
```

## Download File

```
GET https://ashishsankhua-ppt-microservice.hf.space/download/{filename}
```

---

## Examples by Use Case

### Business Plan
```json
{
  "projectName": "Business Plan 2024",
  "template": "professional",
  "content": {
    "Executive Summary": "Our company aims to...",
    "Market Opportunity": {
      "TAM": "$100B",
      "Growth Rate": "15% YoY",
      "Key Trends": ["Digital transformation", "AI adoption"]
    },
    "Financial Projections": {
      "Year 1": "$1M revenue",
      "Year 2": "$5M revenue",
      "Year 3": "$15M revenue"
    }
  }
}
```

### Research Report
```json
{
  "projectName": "Market Research Q1",
  "template": "minimal",
  "content": [
    {"title": "Methodology", "data": {"Approach": "Survey", "Sample Size": "1000"}},
    {"title": "Key Findings", "data": {"Finding 1": "...", "Finding 2": "..."}},
    {"title": "Recommendations", "data": ["Action 1", "Action 2", "Action 3"]}
  ]
}
```

### Product Roadmap
```json
{
  "projectName": "Product Roadmap",
  "template": "startup",
  "content": {
    "Q1 2024": {
      "Features": ["User authentication", "Dashboard", "Analytics"],
      "Milestones": "Beta launch"
    },
    "Q2 2024": {
      "Features": ["Mobile app", "API", "Integrations"],
      "Milestones": "Public launch"
    }
  }
}
```

---

## Notes

- **Max slides**: 20 slides per presentation
- **Max items per list**: 10 items displayed (truncated with `...` if more)
- **Max text length**: 300 characters per field (truncated if longer)
- **Dict keys with numeric values** (like `{"0": {}, "1": {}}`) are automatically converted to arrays
- **Nested dicts** are rendered as subsections with indentation

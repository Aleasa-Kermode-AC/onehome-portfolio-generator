# OneHome Education Portfolio Generator

Webhook service for generating NSW homeschool learning portfolio reports.

## Features

- Generates professional DOCX portfolio reports
- NSW Syllabus compliance
- Disability Standards for Education 2005 support
- PDA-affirming approaches
- Smart evidence management (handles 100+ entries)
- UK English spelling
- Image support

## Deployment to Render

1. Create account at https://render.com
2. Click "New +" → "Web Service"
3. Connect your GitHub repository (or deploy from this folder)
4. Configure:
   - **Name:** onehome-portfolio-generator
   - **Environment:** Node
   - **Build Command:** `npm install`
   - **Start Command:** `npm start`
   - **Instance Type:** Free

## API Endpoints

### POST /generate-portfolio

Generates a portfolio DOCX from JSON data.

**Request Body:**
```json
{
  "childName": "Student Name",
  "yearLevel": "Stage 2",
  "reportingPeriod": "Semester 1 2025",
  "parentName": "Parent Name",
  "state": "NSW",
  "curriculum": "NSW Syllabus",
  "curriculumOutcomes": [...],
  "learningAreaOverviews": {...},
  "evidenceByArea": {...},
  "progressAssessment": {...},
  "futurePlans": {...}
}
```

**Response:**
Binary DOCX file download

### GET /

Health check endpoint - returns service status

## Environment Variables

None required for basic operation.

## Local Development

```bash
npm install
npm run dev
```

Server runs on http://localhost:3000

## Testing

```bash
curl -X POST http://localhost:3000/generate-portfolio \
  -H "Content-Type: application/json" \
  -d @test-portfolio-data.json \
  --output test-output.docx
```

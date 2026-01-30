# Preview Service

Fast PPTX preview rendering with persistent LibreOffice using `unoserver`.

## Architecture

This service runs a **persistent LibreOffice instance** via `unoserver`, which eliminates the cold-start overhead of spawning `soffice` for each conversion. This results in significantly faster PPTX→PDF→Image conversions.

### Components

1. **unoserver** (port 2003): Persistent LibreOffice daemon
2. **FastAPI** (port 8004): REST API with SSE progress streaming

## Endpoints

### `GET /render-previews/{file_id}`
Render all slides as images with real-time SSE progress.

**Response (SSE stream):**
```json
{"stage": "converting", "progress": 0, "message": "converting pptx to pdf..."}
{"stage": "converting", "progress": 33, "message": "converted in 5.2s"}
{"stage": "rendering", "progress": 50, "message": "rendered slide 15/61", "current_slide": 15, "total_slides": 61}
{"stage": "done", "progress": 100, "thumbnails": ["base64...", ...], "format": "jpeg"}
```

### `POST /render-previews-with-sidebar/{file_id}`
Render slides with timeline sidebar applied.

**Request body:**
```json
{
  "tags": ["intro", "methods", "results", "discussion", "conclusion"],
  "sidebar_width": 0.12,
  "sidebar_item_height": 0.10
}
```

**Response:** Same SSE stream as above with actual PPTX processing.

## Performance

| Approach               | Time (61 slides) |
| ---------------------- | ---------------- |
| soffice per-request    | ~28s             |
| unoserver (persistent) | ~8-12s           |
| Improvement            | **2-3x faster**  |

## Local Development

```bash
# Start unoserver in background
unoserver --interface 0.0.0.0 --port 2003 &

# Run FastAPI server
uvicorn server:app --host 0.0.0.0 --port 8004 --reload
```

## Docker

The service is built with the `preview/Dockerfile` and runs both unoserver and FastAPI via `start.sh`.


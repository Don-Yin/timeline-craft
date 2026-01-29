from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from apis.router import router

app = FastAPI(
    title="Upload Service",
    description="File upload and management service for PowerPoint presentations",
    version="1.0.0",
    openapi_tags=[
        {"name": "files", "description": "File upload, listing, and management operations"},
        {"name": "health", "description": "Service health checks"},
    ]
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router, tags=["files"])


@app.get("/health", tags=["health"])
def health_check():
    """check if the upload service is healthy"""
    return {"status": "ok"}

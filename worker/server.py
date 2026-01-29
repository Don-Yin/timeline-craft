from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from src.apis.router import router

app = FastAPI(
    title="Worker Service",
    description="PowerPoint processing worker service - generates thumbnails, processes timeline sidebars, and exports files",
    version="1.0.0",
    openapi_tags=[
        {"name": "thumbnails", "description": "Slide thumbnail generation"},
        {"name": "processing", "description": "PowerPoint processing and export"},
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

app.include_router(router)


@app.get("/health", tags=["health"])
def health_check():
    """check if the worker service is healthy"""
    return {"status": "ok"}

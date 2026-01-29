from pydantic import BaseModel, Field
from typing import List
from datetime import datetime


class FileResponse(BaseModel):
    """response after uploading a file"""
    id: str = Field(description="Unique file identifier (UUID)")
    filename: str = Field(description="Original filename")
    message: str = Field(description="Status message")


class DeleteResponse(BaseModel):
    """response after deleting a file"""
    id: str = Field(description="Deleted file ID")
    message: str = Field(description="Status message")


class FileMetadata(BaseModel):
    """detailed metadata for an uploaded file"""
    id: str = Field(description="File ID")
    filename: str = Field(description="Original filename")
    size: int = Field(description="File size in bytes")
    content_type: str | None = Field(description="MIME content type")
    last_modified: datetime | None = Field(description="Last modification timestamp")


class ListFilesResponse(BaseModel):
    """list of uploaded file IDs"""
    files: List[str] = Field(description="Array of file IDs")

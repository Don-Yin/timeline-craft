from pydantic import BaseModel
from typing import List
from datetime import datetime


class FileResponse(BaseModel):
    id: str
    filename: str
    message: str


class DeleteResponse(BaseModel):
    id: str
    message: str


class FileMetadata(BaseModel):
    id: str
    filename: str
    size: int
    content_type: str | None
    last_modified: datetime | None


class ListFilesResponse(BaseModel):
    files: List[str]

"""pdf rendering utilities using libreoffice and pymupdf"""

import base64
import logging
import subprocess
import tempfile
from pathlib import Path

import fitz

logger = logging.getLogger(__name__)


def convert_pptx_to_pdf_bytes(file_id: str, pptx_bytes: bytes) -> bytes:
    """convert pptx to pdf using libreoffice"""
    with tempfile.TemporaryDirectory(prefix=f"pptx-{file_id}-") as tmpdir:
        tmpdir_path = Path(tmpdir)
        pptx_path = tmpdir_path / "source.pptx"
        pptx_path.write_bytes(pptx_bytes)

        cmd = [
            "soffice",
            "--headless",
            f"-env:UserInstallation=file://{tmpdir_path}/profile",
            "--convert-to",
            "pdf",
            "--outdir",
            str(tmpdir_path),
            str(pptx_path),
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        pdf_path = pptx_path.with_suffix(".pdf")
        return pdf_path.read_bytes()


def render_slide_thumbnail(pdf_bytes: bytes, slide_index: int, target_width: int = 640) -> str:
    """render a single slide as base64 png"""
    doc = fitz.open("pdf", pdf_bytes)
    page = doc.load_page(slide_index)
    scale = target_width / (page.rect.width or 1)
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    result = base64.b64encode(pix.tobytes("png")).decode("utf-8")
    doc.close()
    return result


def render_all_thumbnails(pdf_bytes: bytes, target_width: int = 640, use_jpeg: bool = False) -> list[str]:
    """render all slides as base64 images in one pass"""
    doc = fitz.open("pdf", pdf_bytes)
    results = []
    fmt = "jpeg" if use_jpeg else "png"

    for i in range(doc.page_count):
        page = doc.load_page(i)
        scale = target_width / (page.rect.width or 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        img_bytes = pix.tobytes(fmt) if fmt == "png" else pix.tobytes("jpeg", jpg_quality=75)
        results.append(base64.b64encode(img_bytes).decode("utf-8"))

    doc.close()
    return results

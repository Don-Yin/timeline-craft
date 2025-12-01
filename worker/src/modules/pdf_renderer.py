import io
from pathlib import Path
from fastapi import HTTPException
import logging, tempfile, subprocess, base64, fitz

logger = logging.getLogger(__name__)


def convert_pptx_to_pdf_bytes(file_id: str, pptx_bytes: bytes) -> bytes:
    with tempfile.TemporaryDirectory(prefix=f"pptx-{file_id}-") as tmpdir:
        tmpdir_path = Path(tmpdir)
        pptx_path = tmpdir_path / "source.pptx"
        pptx_path.write_bytes(pptx_bytes)
        cmd = ["soffice", "--headless", f"-env:UserInstallation=file://{str(tmpdir_path)}/profile", "--convert-to", "pdf", "--outdir", str(tmpdir_path), str(pptx_path)]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except subprocess.CalledProcessError as exc:
            logger.error(f"libreoffice conversion failed: {exc.stderr.decode(errors='ignore')}")
            raise HTTPException(status_code=500, detail="Failed to convert presentation to PDF")
        pdf_path = pptx_path.with_suffix(".pdf")
        if not pdf_path.exists():
            raise HTTPException(status_code=500, detail="Converted PDF not found")
        return pdf_path.read_bytes()


def render_slide_thumbnail(pdf_bytes: bytes, slide_index: int, target_width: int = 640) -> str:
    try:
        doc = fitz.open("pdf", pdf_bytes)
    except Exception as exc:
        logger.error(f"unable to open converted pdf: {exc}")
        raise HTTPException(status_code=500, detail="Failed to open converted PDF")

    if slide_index < 0 or slide_index >= doc.page_count:
        doc.close()
        raise HTTPException(status_code=404, detail="Slide index out of range")

    page = doc.load_page(slide_index)
    width = page.rect.width or 1
    scale = target_width / width
    matrix = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    png_bytes = pix.tobytes("png")
    doc.close()
    return base64.b64encode(png_bytes).decode("utf-8")

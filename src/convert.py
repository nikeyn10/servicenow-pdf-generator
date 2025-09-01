import subprocess
import os
import shutil
from PIL import Image
import img2pdf
from reportlab.pdfgen import canvas
from src.log import log_event

CONVERT_MAP = {
    "pdf": "copy",
    "docx": "libreoffice",
    "xlsx": "libreoffice",
    "pptx": "libreoffice",
    "csv": "libreoffice",
    "txt": "libreoffice",
    "png": "img2pdf",
    "jpg": "img2pdf",
    "jpeg": "img2pdf",
    "webp": "img2pdf",
    "html": "wkhtmltopdf",
}

# Dispatcher for file→PDF
def to_pdf(path_in, ext, out_dir, prefer_libreoffice=True, html_enabled=False):
    if not path_in or not os.path.exists(path_in):
        log_event(action="convert", status="fail", warning="Input file missing or None")
        return None
        
    ext = ext.lower().lstrip('.')  # Remove leading dot if present
    path_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(path_in))[0] + ".pdf")
    try:
        if ext == "pdf":
            shutil.copy(path_in, path_pdf)
        elif ext in ["docx", "xlsx", "pptx", "csv", "txt"] and prefer_libreoffice:
            subprocess.run(["soffice", "--headless", "--convert-to", "pdf", path_in, "--outdir", out_dir], check=True, timeout=60)
        elif ext in ["png", "jpg", "jpeg", "webp"]:
            with open(path_pdf, "wb") as f:
                f.write(img2pdf.convert(path_in))
        elif ext == "html" and html_enabled:
            subprocess.run(["wkhtmltopdf", path_in, path_pdf], check=True, timeout=60)
        else:
            # Placeholder PDF for unsupported types
            c = canvas.Canvas(path_pdf)
            c.drawString(100, 750, f"{os.path.basename(path_in)}: unsupported type; included as placeholder.")
            c.save()
        log_event(action="convert", status="success", asset_id=os.path.basename(path_in))
        return path_pdf
    except Exception as e:
        log_event(action="convert", status="fail", asset_id=os.path.basename(path_in), warning=str(e))
        # Always emit placeholder PDF
        c = canvas.Canvas(path_pdf)
        c.drawString(100, 750, f"{os.path.basename(path_in)}: conversion failed; included as placeholder.")
        c.save()
        return path_pdf

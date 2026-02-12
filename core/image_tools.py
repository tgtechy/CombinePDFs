from __future__ import annotations
# core/image_tools.py
print(">>> LOADED image_tools.py FROM:", __file__)

import tempfile
import os
from io import BytesIO

from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader as RLImageReader


def image_to_pdf(image_path: str) -> str:
    """
    Convert an image file (JPG, PNG, TIFF, etc.) into a temporary 1‑page PDF.
    Returns the path to the generated PDF.
    """
    try:
        img = Image.open(image_path)
        img.load()  # Ensure image is fully loaded before use

        # Convert to RGB if needed (JPEG cannot handle transparency)
        if img.mode in ("RGBA", "LA", "P"):
            background = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "P":
                img = img.convert("RGBA")
            background.paste(
                img,
                mask=img.split()[-1] if len(img.split()) > 3 else None
            )
            img = background
        elif img.mode != "RGB":
            img = img.convert("RGB")

        img_width, img_height = img.size

        # Create a temporary PDF file
        fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)

        # Convert pixel dimensions → PDF points (72 DPI)
        dpi = 96
        page_width = (img_width / dpi) * 72
        page_height = (img_height / dpi) * 72

        # Save image to a BytesIO buffer as PNG
        img_buffer = BytesIO()
        img.save(img_buffer, format="PNG")
        img_buffer.seek(0)

        c = canvas.Canvas(temp_pdf_path, pagesize=(page_width, page_height))
        c.drawInlineImage(img, 0, 0, width=page_width, height=page_height)
        c.showPage()
        c.save()
        return temp_pdf_path

    except Exception as e:
        raise Exception(f"Failed to convert image to PDF: {str(e)}")
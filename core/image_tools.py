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
    Convert an image file (JPG, PNG, TIFF, etc.) into a 1‑page PDF and write to a temp file.
    Returns the path to the temp PDF file.
    """
    try:
        img = None
        try:
            img = Image.open(image_path)
            img.load()  # Ensure image is fully loaded before use
        except Exception as e:
            raise Exception(f"Failed to open image file '{image_path}': {str(e)}")

        if img is None:
            raise Exception(f"Image file '{image_path}' could not be loaded (None returned).")

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
        dpi = 96
        page_width = (img_width / dpi) * 72
        page_height = (img_height / dpi) * 72

        # Write PDF to a temp file
        fd, temp_pdf_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        try:
            c = canvas.Canvas(temp_pdf_path, pagesize=(page_width, page_height))
            # Use file path for drawInlineImage for best compatibility
            c.drawInlineImage(image_path, 0, 0, width=page_width, height=page_height)
            c.showPage()
            c.save()
        except Exception as e:
            # Fallback: try using ImageReader
            try:
                c = canvas.Canvas(temp_pdf_path, pagesize=(page_width, page_height))
                rl_img = RLImageReader(img)
                c.drawInlineImage(rl_img, 0, 0, width=page_width, height=page_height)
                c.showPage()
                c.save()
            except Exception as e2:
                os.remove(temp_pdf_path)
                raise Exception(f"Failed to render image file '{image_path}' into PDF: {str(e)}; fallback also failed: {str(e2)}")
        return temp_pdf_path

    except Exception as e:
        raise Exception(
            f"Failed to convert image to PDF for file '{image_path}': {str(e)}\n"
            "If this error persists, check the file format and try re-saving the image."
        )
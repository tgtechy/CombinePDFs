from __future__ import annotations
# core/watermark.py


import io
import math

from PyPDF2 import PdfReader
from reportlab.pdfgen import canvas


def add_watermark(
    page,
    text: str,
    opacity: float,
    font_size: int,
    rotation: int,
    position: str,
    safe_mode: bool,
    font_color: str = "#000000"
):
    """
    Draw a rotated, semi-transparent watermark onto a PDF page.
    """
    try:
        box = page.mediabox
        width = float(box.width)
        height = float(box.height)

        position = (position or "center").lower()

        adjusted_font_size = font_size
        adjusted_position = position

        # ------------------------------------------------------------------
        # Safe mode: auto-scale watermark to avoid clipping
        # ------------------------------------------------------------------
        if safe_mode:
            text_width_approx = len(text) * font_size * 0.5
            text_height_approx = font_size * 1.2

            angle_rad = math.radians(rotation)
            cos_a = abs(math.cos(angle_rad))
            sin_a = abs(math.sin(angle_rad))

            rotated_width = text_width_approx * cos_a + text_height_approx * sin_a
            rotated_height = text_width_approx * sin_a + text_height_approx * cos_a

            safe_margin = 40

            # Top/bottom placement: ensure vertical fit
            if position in ["top", "bottom"]:
                half_rotated = rotated_height * 0.5
                allowed = height * 0.15 + safe_margin

                if half_rotated > allowed:
                    max_font = int(font_size * (allowed / half_rotated) * 0.95)
                    if max_font >= 10:
                        adjusted_font_size = max(10, max_font)
                    else:
                        adjusted_position = "center"

            # Horizontal fit
            if rotated_width > width - safe_margin * 2 and adjusted_font_size > 10:
                max_font = int(
                    adjusted_font_size *
                    (width - safe_margin * 2) / rotated_width * 0.95
                )
                adjusted_font_size = max(10, max_font)

        # ------------------------------------------------------------------
        # Render watermark into a temporary PDF
        # ------------------------------------------------------------------
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))
        c.setFillAlpha(opacity)
        c.setFont("Helvetica-Bold", adjusted_font_size)
        from reportlab.lib.colors import HexColor
        c.setFillColor(HexColor(font_color))

        c.saveState()

        # Positioning
        if adjusted_position == "top-left":
            c.translate(60, height * 0.85)
            c.rotate(rotation)
            c.drawString(0, 0, text)
        elif adjusted_position == "top-right":
            c.translate(width - 60, height * 0.85)
            c.rotate(rotation)
            c.drawRightString(0, 0, text)
        elif adjusted_position == "bottom-left":
            c.translate(60, height * 0.15)
            c.rotate(rotation)
            c.drawString(0, 0, text)
        elif adjusted_position == "bottom-right":
            c.translate(width - 60, height * 0.15)
            c.rotate(rotation)
            c.drawRightString(0, 0, text)
        elif adjusted_position == "top":
            c.translate(width / 2, height * 0.85)
            c.rotate(rotation)
            c.drawCentredString(0, 0, text)
        elif adjusted_position == "bottom":
            c.translate(width / 2, height * 0.15)
            c.rotate(rotation)
            c.drawCentredString(0, 0, text)
        else:
            c.translate(width / 2, height / 2)
            c.rotate(rotation)
            c.drawCentredString(0, 0, text)
        c.restoreState()
        c.save()

        packet.seek(0)
        watermark_pdf = PdfReader(packet)
        watermark_page = watermark_pdf.pages[0]

        # Merge watermark onto the page
        page.merge_page(watermark_page)

    except Exception:
        # Watermark failure should not break the merge
        pass
from __future__ import annotations

# core/page_ops.py

from typing import List
import io
import math

import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas


# ---------------------------------------------------------------------------
# Page Range Parsing
# ---------------------------------------------------------------------------

def parse_page_range(range_text: str, total_pages: int) -> List[int]:
    """
    Convert a user-specified page range string into a list of zero-based indices.
    Supports:
        "all", "", "1-3", "1,3,5", "2-4,7", etc.
    """
    text = (range_text or "").strip().lower()
    if text in ("", "all"):
        return list(range(total_pages))

    indices = set()
    parts = [p.strip() for p in text.split(",") if p.strip()]

    for part in parts:
        if "-" in part:
            start_str, end_str = [p.strip() for p in part.split("-", 1)]
            if not start_str or not end_str:
                raise ValueError("Invalid range format")

            try:
                start = int(start_str)
                end = int(end_str)
            except ValueError:
                raise ValueError("Page numbers must be integers")

            if start < 1 or end < 1 or start > end:
                raise ValueError("Invalid range order")

            if end > total_pages:
                raise ValueError(
                    f"Page range exceeds total pages, file has {total_pages} pages"
                )

            for page_num in range(start, end + 1):
                indices.add(page_num - 1)

        else:
            try:
                page_num = int(part)
            except ValueError:
                raise ValueError("Page numbers must be integers")

            if page_num < 1 or page_num > total_pages:
                raise ValueError(
                    f"Page number out of range, file has {total_pages} pages"
                )

            indices.add(page_num - 1)

    if not indices:
        raise ValueError("No valid pages selected")

    return sorted(indices)


# ---------------------------------------------------------------------------
# Blank Page Detection
# ---------------------------------------------------------------------------

def is_page_blank(page) -> bool:
    """
    Heuristic blank-page detection:
    - No extracted text
    - No XObjects (images)
    - No fonts
    """
    try:
        text = page.extract_text()
        if text and text.strip():
            return False

        if "/Resources" in page:
            resources = page["/Resources"]

            if "/XObject" in resources:
                xobjects = resources["/XObject"]
                if xobjects and len(xobjects) > 0:
                    return False

            if "/Font" in resources and len(resources["/Font"]) > 0:
                return False

        return True

    except Exception:
        return False


# ---------------------------------------------------------------------------
# Page Scaling
# ---------------------------------------------------------------------------

def scale_page(page, target_width: float, target_height: float):
    """
    Scale a PDF page to fit within target dimensions while maintaining aspect ratio.
    """
    try:
        box = page.mediabox
        current_width = float(box.width)
        current_height = float(box.height)

        scale_x = target_width / current_width
        scale_y = target_height / current_height
        scale = min(scale_x, scale_y)

        new_width = current_width * scale
        new_height = current_height * scale

        x_offset = (target_width - new_width) / 2
        y_offset = (target_height - new_height) / 2

        if scale != 1.0 or x_offset != 0 or y_offset != 0:
            page.scale_by(scale)
            page.add_transform_matrix([1, 0, 0, 1, x_offset, y_offset])

        page.mediabox.lower_left = (0, 0)
        page.mediabox.upper_right = (target_width, target_height)

    except Exception:
        pass


# ---------------------------------------------------------------------------
# Breaker Page Creation
# ---------------------------------------------------------------------------

def create_page_with_filename(filename: str, width: float, height: float):
    """
    Create a blank PDF page with the filename centered.
    Used for breaker pages between merged files.
    """
    try:
        width = float(width)
        height = float(height)

        scale_factor = height / 792.0
        base_font_size = 14
        scaled_font_size = int(base_font_size * scale_factor)
        scaled_line_height = int(18 * scale_factor)
        scaled_spacing_below_file = int(35 * scale_factor)
        scaled_line_spacing = int(20 * scale_factor)
        scaled_margin = int(50 * scale_factor)
        scaled_line_width = max(1, int(2 * scale_factor))

        max_chars = 25
        if len(filename) > max_chars:
            filename_lines = [
                filename[i:i + max_chars]
                for i in range(0, len(filename), max_chars)
            ]
        else:
            filename_lines = [filename]

        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(width, height))

        filename_height = len(filename_lines) * scaled_line_height
        total_text_height = (
            40 * scale_factor +
            filename_height +
            15 * scale_factor +
            15 * scale_factor
        )

        vertical_center = height / 2
        current_y = vertical_center + (total_text_height / 2) - 10 * scale_factor

        c.setFont("Helvetica", scaled_font_size)
        c.setFillGray(0.3)
        c.drawCentredString(width / 2, current_y, "File")

        c.setFont("Helvetica-Bold", scaled_font_size)
        current_y -= scaled_spacing_below_file

        c.setStrokeGray(0.3)
        c.setLineWidth(scaled_line_width)
        c.line(
            scaled_margin,
            current_y + scaled_line_spacing,
            width - scaled_margin,
            current_y + scaled_line_spacing
        )

        for line in filename_lines:
            c.drawCentredString(width / 2, current_y, line)
            current_y -= scaled_line_height

        c.line(
            scaled_margin,
            current_y + scaled_line_height - 10 * scale_factor,
            width - scaled_margin,
            current_y + scaled_line_height - 10 * scale_factor
        )

        c.setFont("Helvetica", scaled_font_size)
        current_y -= 10 * scale_factor
        c.drawCentredString(width / 2, current_y, "follows")

        c.save()
        packet.seek(0)

        page_pdf = PdfReader(packet)
        if len(page_pdf.pages) > 0:
            return page_pdf.pages[0]

        return PdfWriter().add_blank_page(width=width, height=height)

    except Exception:
        return PdfWriter().add_blank_page(width=float(width), height=float(height))
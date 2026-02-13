# core/pdf_merger.py

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Dict, Callable, Optional

import os
import tempfile

import tkinter as tk
from tkinter import messagebox

from PyPDF2 import PdfReader, PdfWriter

@dataclass
class MergeOptions:
    delete_blank_pages: bool = False
    insert_toc: bool = False

    add_breaker_pages: bool = False
    breaker_uniform_size: bool = False

    add_filename_bookmarks: bool = False   # ← ADD THIS

    compression_enabled: bool = False
    compression_level: str = "Medium"

    watermark_enabled: bool = False
    watermark_text: str = ""
    watermark_opacity: float = 0.3
    watermark_rotation: int = 0
    watermark_position: str = "center"
    watermark_font_size: int = 50
    watermark_safe_mode: bool = True
    watermark_font_color: str = "#000000"

    metadata_enabled: bool = False
    pdf_title: str = ""
    pdf_author: str = ""
    pdf_subject: str = ""
    pdf_keywords: str = ""

    scaling_enabled: bool = False
    scaling_mode: str = "Fit"
    scaling_percent: int = 100

    # Encryption
    encrypt_enabled: bool = False
    encrypt_user_pw: str = ""
    encrypt_owner_pw: str = ""

# Import modular helpers
from core.page_ops import (
    parse_page_range,
    is_page_blank,
    scale_page,
    create_page_with_filename,
)
from core.image_tools import image_to_pdf
from core.watermark import add_watermark
from core.compression import compress_page
from core.toc import insert_toc_pages


# ---------------------------------------------------------------------------
# Data structure for UI → Core communication
# ---------------------------------------------------------------------------

@dataclass
class FileEntry:
    path: str
    rotation: int = 0
    page_range: str = ""
    reverse: bool = False


# ---------------------------------------------------------------------------
# Main merge pipeline
# ---------------------------------------------------------------------------


def merge_files(
    files: List[FileEntry],
    output_path: str,
    options: MergeOptions,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    cancel_callback: Optional[Callable[[], bool]] = None,
):
    def cancelled():
        return cancel_callback and cancel_callback()

    pdf_writer = PdfWriter()
    file_toc_entries = []
    current_page_num = 0
    max_width = 0.0
    max_height = 0.0
    temp_pdf_files = []  # Track temp files for cleanup
    open_files = []  # Keep all file handles open for the entire merge
    image_temp_map = {}  # Map (image_path, usage_index) to temp PDF path

    # ----------------------------------------------------------------------
    # First pass: determine max page size (for scaling)
    # ----------------------------------------------------------------------
    image_usage_counter = {}
    for i, entry in enumerate(files):
        if cancelled():
            raise RuntimeError("Merge cancelled")
        if progress_callback:
            progress_callback(i, len(files), f"Preparing: {os.path.basename(entry.path)}")
        file_path = entry.path
        rotation = entry.rotation
        page_range = entry.page_range
        reverse = entry.reverse

        pdf_path = file_path
        is_image = file_path.lower().endswith(
            (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif")
        )

        # Convert images to unique temporary PDFs (do not open them yet)
        if is_image:
            usage_idx = image_usage_counter.get(file_path, 0)
            image_usage_counter[file_path] = usage_idx + 1
            pdf_path = image_to_pdf(file_path)
            temp_pdf_files.append(pdf_path)
            image_temp_map[(file_path, usage_idx)] = pdf_path

        # Only open the file for reading if not an image (for first pass)
        if not is_image:
            pdf_file = open(pdf_path, "rb")
            open_files.append(pdf_file)
            pdf_reader = PdfReader(pdf_file)
            total_pages = len(pdf_reader.pages)

            try:
                page_indices = parse_page_range(page_range, total_pages)
            except ValueError:
                continue

            if reverse:
                page_indices = list(reversed(page_indices))

            has_explicit_range = (
                page_range and page_range.strip().lower() not in ["all", ""]
            )

            for idx in page_indices:
                page = pdf_reader.pages[idx]

                if (
                    options.delete_blank_pages
                    and not has_explicit_range
                    and is_page_blank(page)
                ):
                    continue

                box = page.mediabox
                width = float(box.width)
                height = float(box.height)
                if rotation in [90, 270]:
                    width, height = height, width

                max_width = max(max_width, width)
                max_height = max(max_height, height)

    # ----------------------------------------------------------------------
    # Second pass: process files
    # ----------------------------------------------------------------------
    image_usage_counter = {}
    for i, entry in enumerate(files):
        if cancelled():
            raise RuntimeError("Merge cancelled")

        file_path = entry.path
        rotation = entry.rotation
        page_range = entry.page_range
        reverse = entry.reverse

        if progress_callback:
            progress_callback(i, len(files), file_path)

        pdf_path = file_path
        is_image = file_path.lower().endswith(
            (".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif")
        )

        # Convert images to unique temporary PDFs (use the correct temp file for each usage)
        if is_image:
            usage_idx = image_usage_counter.get(file_path, 0)
            image_usage_counter[file_path] = usage_idx + 1
            pdf_path = image_temp_map[(file_path, usage_idx)]
            pdf_file = open(pdf_path, "rb")
            open_files.append(pdf_file)
        else:
            pdf_file = open(pdf_path, "rb")
            open_files.append(pdf_file)
        pdf_reader = PdfReader(pdf_file)
        total_pages = len(pdf_reader.pages)

        # --------------------------------------------------------------
        # Insert breaker page before every file if enabled
        # --------------------------------------------------------------
        if options.add_breaker_pages:
            breaker_label = Path(file_path).name
            if options.breaker_uniform_size:
                breaker_width = 612.0  # 8.5 inches * 72
                breaker_height = 792.0 # 11 inches * 72
            else:
                if len(pdf_reader.pages) > 0:
                    first_page = pdf_reader.pages[0]
                    breaker_width = float(first_page.mediabox.width)
                    breaker_height = float(first_page.mediabox.height)
                    if rotation in [90, 270]:
                        breaker_width, breaker_height = breaker_height, breaker_width
                else:
                    breaker_width = 612.0
                    breaker_height = 792.0
            breaker_page = create_page_with_filename(
                breaker_label, breaker_width, breaker_height
            )
            pdf_writer.add_page(breaker_page)
            current_page_num += 1

        # --------------------------------------------------------------
        # Parse page range
        # --------------------------------------------------------------
        try:
            page_indices = parse_page_range(page_range, total_pages)
        except ValueError as e:
            raise ValueError(
                f"{Path(file_path).name} (Total pages: {total_pages})\n\n{e}"
            )

        if reverse:
            page_indices = list(reversed(page_indices))

        has_explicit_range = (
            page_range and page_range.strip().lower() not in ["all", ""]
        )

        file_start_page = current_page_num

        # --------------------------------------------------------------
        # Add filename bookmark
        # --------------------------------------------------------------
        parent_bookmark = None
        if options.add_filename_bookmarks and len(page_indices) > 0:
            bookmark_title = Path(file_path).stem
            parent_bookmark = pdf_writer.add_outline_item(
                bookmark_title, current_page_num
            )

        # --------------------------------------------------------------
        # Process each page
        # --------------------------------------------------------------
        for idx in page_indices:
            page = pdf_reader.pages[idx]

            # Skip blank pages
            if (
                options.delete_blank_pages
                and not has_explicit_range
                and is_page_blank(page)
            ):
                continue

            # Rotation
            if rotation != 0:
                page.rotate(rotation)

            # Scaling
            if (
                options.scaling_enabled
                and max_width > 0
                and max_height > 0
            ):
                scale_page(page, max_width, max_height)

            # Watermark
            if (
                options.watermark_enabled
                and options.watermark_text.strip()
            ):
                add_watermark(
                    page,
                    options.watermark_text.strip(),
                    options.watermark_opacity,
                    options.watermark_font_size,
                    options.watermark_rotation,
                    options.watermark_position.lower(),
                    options.watermark_safe_mode,
                    options.watermark_font_color
                )

            pdf_writer.add_page(page)
            current_page_num += 1


        # --------------------------------------------------------------
        # TOC entry
        # --------------------------------------------------------------
        if (options.insert_toc) and len(page_indices) > 0:
            file_toc_entries.append(
                {
                    "filename": Path(file_path).name,
                    "page": file_start_page,
                }
            )

    if cancelled():
        raise RuntimeError("Merge cancelled")

    # ----------------------------------------------------------------------
    # Metadata
    # ----------------------------------------------------------------------
    metadata = {}
    if options.metadata_enabled and any(
    getattr(options, k).strip()
    for k in ("pdf_title", "pdf_author", "pdf_subject", "pdf_keywords")
    ):
        if options.pdf_title.strip():
            metadata["/Title"] = options.pdf_title

        if options.pdf_author.strip():
            metadata["/Author"] = options.pdf_author

        if options.pdf_subject.strip():
            metadata["/Subject"] = options.pdf_subject

        if options.pdf_keywords.strip():
            metadata["/Keywords"] = options.pdf_keywords

    pdf_writer.add_metadata(metadata)

    # ----------------------------------------------------------------------
    # Write output PDF (with TOC inserted before encryption if needed)
    # ----------------------------------------------------------------------
    final_output_path = output_path
    needs_encryption = getattr(options, "encrypt_enabled", False) and (getattr(options, "encrypt_user_pw", "") or getattr(options, "encrypt_owner_pw", ""))
    temp_unencrypted_path = None

    if needs_encryption:
        # Write unencrypted PDF to temp file
        import tempfile
        fd, temp_unencrypted_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        out_path = temp_unencrypted_path
    else:
        out_path = output_path

    if progress_callback:
        progress_callback(len(files), len(files), "Writing combined PDF...")

    with open(out_path, "wb") as out_file:
        if options.compression_enabled:
            for page in pdf_writer.pages:
                compress_page(page, options.compression_level)
        pdf_writer.write(out_file)

    # Insert TOC pages (PyMuPDF) before encryption
    if options.insert_toc and len(file_toc_entries) > 0:
        insert_toc_pages(out_path, file_toc_entries)

    # If encryption is needed, use PyMuPDF to encrypt after TOC insertion
    if needs_encryption:
        import fitz  # PyMuPDF
        user_pw = getattr(options, "encrypt_user_pw", "")
        owner_pw = getattr(options, "encrypt_owner_pw", "")
        doc = fitz.open(temp_unencrypted_path)
        # Use owner_pw if set, else user_pw for both
        pw_owner = owner_pw if owner_pw else user_pw
        pw_user = user_pw if user_pw else owner_pw
        # Save with encryption
        doc.save(output_path, encryption=fitz.PDF_ENCRYPT_AES_256,
                 owner_pw=pw_owner, user_pw=pw_user,
                 permissions=fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ANNOTATE)
        doc.close()
        os.remove(temp_unencrypted_path)

    # Close all open file handles
    for f in open_files:
        try:
            f.close()
        except Exception:
            pass
    # Cleanup temp files
    for temp_path in temp_pdf_files:
        try:
            os.remove(temp_path)
        except Exception:
            pass
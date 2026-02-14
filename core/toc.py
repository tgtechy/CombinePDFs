from __future__ import annotations

# core/toc.py

import tempfile
import os
import fitz  # PyMuPDF

def insert_toc_pages(pdf_path: str, toc_entries: list[dict], file_info_list: list[str] = None):
    """
    Insert a multi-page Table of Contents at the beginning of the PDF.
    toc_entries = [
        { "filename": "example.pdf", "page": 12 },
        ...
    ]
    """
    try:
        doc = fitz.open(pdf_path)

        # Standard Letter size
        page_width = 612
        page_height = 792

        title_font_size = 24
        entry_font_size = 11
        margin_left = 72
        margin_top = 72
        margin_bottom = 72
        line_height = 20

        available_height = (
            page_height - margin_top - margin_bottom - (title_font_size + 25)
        )
        entries_per_page = max(1, int(available_height / line_height))

        toc_pages_data = []
        total_pages = (len(toc_entries) + entries_per_page - 1) // entries_per_page

        # ------------------------------------------------------------------
        # Break TOC entries into pages
        # ------------------------------------------------------------------
        for i in range(0, len(toc_entries), entries_per_page):
            chunk = toc_entries[i:i + entries_per_page]
            toc_pages_data.append({
                "entries": chunk,
                "page_number": len(toc_pages_data) + 1,
                "total_pages": total_pages
            })

        num_toc_pages = len(toc_pages_data)
        # ------------------------------------------------------------------
        # Insert TOC pages at the beginning (reverse order)
        # ------------------------------------------------------------------
        for toc_page_idx, toc_page_info in enumerate(reversed(toc_pages_data)):
            toc_page = doc.new_page(0, width=page_width, height=page_height)

            # Title
            title_text = "Table of Contents"
            if toc_page_info["total_pages"] > 1:
                title_text += (
                    f" (Page {toc_page_info['page_number']} of "
                    f"{toc_page_info['total_pages']})"
                )

            toc_page.insert_text(
                (margin_left, margin_top + 20),
                title_text,
                fontsize=title_font_size,
                fontname="helv",
                color=(0, 0, 0)
            )

            # Divider line
            line_y = margin_top + title_font_size + 15
            toc_page.draw_line(
                (margin_left, line_y),
                (page_width - margin_left, line_y),
                color=(0, 0, 0),
                width=1
            )

            current_y = line_y + 25

            # If this is the first TOC page and file_info_list is provided, insert file info below the divider
            if toc_page_idx == len(toc_pages_data) - 1 and file_info_list:
                info_font_size = 10
                info_y = current_y
                for info in file_info_list:
                    toc_page.insert_text(
                        (margin_left, info_y),
                        info,
                        fontsize=info_font_size,
                        fontname="helv",
                        color=(0.2, 0.2, 0.2)
                    )
                    info_y += info_font_size + 2
                current_y = info_y + 8  # Add some space after file info
            max_filename_length = 80

            for entry in toc_page_info["entries"]:
                filename = entry["filename"]
                if len(filename) > max_filename_length:
                    filename = filename[:max_filename_length - 3] + "..."

                # Destination page index after TOC insertion
                dest_page_index = entry["page"] + num_toc_pages

                entry_text = filename
                page_text = f"Page {dest_page_index + 1}"

                # Filename text
                text_rect = fitz.Rect(
                    margin_left, current_y,
                    page_width - 150, current_y + line_height
                )
                toc_page.insert_textbox(
                    text_rect,
                    entry_text,
                    fontsize=entry_font_size,
                    fontname="helv",
                    color=(0, 0, 1),
                    align=fitz.TEXT_ALIGN_LEFT
                )

                # Page number text
                page_rect = fitz.Rect(
                    page_width - 150, current_y,
                    page_width - margin_left, current_y + line_height
                )
                toc_page.insert_textbox(
                    page_rect,
                    page_text,
                    fontsize=entry_font_size,
                    fontname="helv",
                    color=(0.3, 0.3, 0.3),
                    align=fitz.TEXT_ALIGN_RIGHT
                )

                # Clickable link
                if dest_page_index < len(doc):
                    link_rect = fitz.Rect(
                        margin_left, current_y,
                        page_width - margin_left, current_y + line_height
                    )
                    toc_page.insert_link({
                        "kind": fitz.LINK_GOTO,
                        "from": link_rect,
                        "page": dest_page_index,
                        "to": fitz.Point(0, 0),
                        "zoom": 0
                    })

                current_y += line_height

        # ------------------------------------------------------------------
        # Adjust existing outline (bookmarks)
        # ------------------------------------------------------------------
        try:
            existing_toc = doc.get_toc()
            if existing_toc:
                offset = len(toc_pages_data)
                for entry in existing_toc:
                    if len(entry) >= 3 and isinstance(entry[2], int):
                        entry[2] += offset
                doc.set_toc(existing_toc)
        except Exception:
            pass

        # ------------------------------------------------------------------
        # Save updated PDF
        # ------------------------------------------------------------------
        fd, temp_path = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)

        try:
            doc.save(temp_path, garbage=4, deflate=True)
            doc.close()
            os.replace(temp_path, pdf_path)
        except Exception:
            doc.close()
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise

    except Exception as e:
        print(f"Warning: Could not insert TOC page: {e}")

print(">>> FUNCTIONS DEFINED:", "insert_toc_pages" in globals())
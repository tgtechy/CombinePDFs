# core/file_manager.py

from pathlib import Path
from typing import List, Dict, Tuple
import PyPDF2

SUPPORTED_EXTS = (
    '.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.tif'
)


def is_pdf_readable(path: str) -> Tuple[bool, str | None]:
    """Return (True, None) if readable, (False, error_message) if not."""
    try:
        with open(path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            if reader.is_encrypted:
                return False, "Encrypted PDF"
        return True, None
    except Exception as e:
        msg = str(e)
        if any(k in msg.lower() for k in ("encrypted", "password", "aes", "pycryptodome")):
            return False, msg
        return True, None


def add_file(file_list: List[Dict], path: str) -> bool:
    """
    Add a single file entry if not already present.
    Returns True if added, False if duplicate.
    """
    if any(entry["path"] == path for entry in file_list):
        return False

    file_list.append({
        "path": path,
        "rotation": 0,
        "page_range": "All",
        "reverse": False
    })
    return True


def add_files_to_list(
    file_list: List[Dict],
    paths: List[str]
) -> Tuple[int, int, List[str], int, List[str]]:
    """
    Add multiple files and return:
    (added_count, duplicate_count, duplicate_names, unsupported_count, unsupported_names)
    """

    added_count = 0
    duplicate_count = 0
    unsupported_count = 0
    duplicates = []
    unsupported_files = []

    existing_paths = {entry["path"] for entry in file_list}

    for file in paths:
        lower = file.lower()

        # Unsupported extension
        if not lower.endswith(SUPPORTED_EXTS):
            unsupported_count += 1
            unsupported_files.append(Path(file).name)
            continue

        # PDF readability check
        if lower.endswith(".pdf"):
            ok, err = is_pdf_readable(file)
            if not ok:
                continue

        # Duplicate check
        if file in existing_paths:
            duplicate_count += 1
            duplicates.append(Path(file).name)
            continue

        # Add entry
        file_list.append({
            "path": file,
            "rotation": 0,
            "page_range": "All",
            "reverse": False
        })
        added_count += 1

    return added_count, duplicate_count, duplicates, unsupported_count, unsupported_files


def move_up(file_list: List[Dict], index: int) -> None:
    if index > 0:
        file_list[index - 1], file_list[index] = file_list[index], file_list[index - 1]


def move_down(file_list: List[Dict], index: int) -> None:
    if index < len(file_list) - 1:
        file_list[index + 1], file_list[index] = file_list[index], file_list[index + 1]


def remove_file(file_list: List[Dict], index: int) -> None:
    if 0 <= index < len(file_list):
        del file_list[index]


def clear_files(file_list: List[Dict]) -> None:
    file_list.clear()
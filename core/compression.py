from __future__ import annotations

# core/compression.py

import io
from PIL import Image
from PyPDF2.generic import NameObject


def compress_page(page, compression_level: str):
    """
    Compress images inside a PDF page.
    compression_level: "Low", "Medium", "High", "Maximum"
    """
    try:
        quality_map = {
            "Low": 95,
            "Medium": 75,
            "High": 50,
            "Maximum": 30
        }

        quality = quality_map.get(compression_level, 75)

        # No images → nothing to compress
        if "/Resources" not in page or "/XObject" not in page["/Resources"]:
            return

        xobjects = page["/Resources"]["/XObject"].get_object()

        for obj_name in xobjects:
            obj = xobjects[obj_name]

            # Only compress images
            if obj.get("/Subtype") != "/Image":
                continue

            try:
                # PyPDF2 image extraction
                if hasattr(obj, "get_data"):
                    image_data = obj.get_data()
                    width = obj.get("/Width", 0)
                    height = obj.get("/Height", 0)

                    if width <= 0 or height <= 0:
                        continue

                    try:
                        # Load image via Pillow
                        img = Image.open(io.BytesIO(image_data))

                        # Flatten transparency
                        if img.mode in ("RGBA", "LA", "P"):
                            background = Image.new("RGB", img.size, (255, 255, 255))
                            if img.mode == "P":
                                img = img.convert("RGBA")
                            background.paste(
                                img,
                                mask=img.split()[-1]
                                if img.mode in ("RGBA", "LA") else None
                            )
                            img = background
                        elif img.mode != "RGB":
                            img = img.convert("RGB")

                        # Recompress as JPEG
                        output = io.BytesIO()
                        img.save(
                            output,
                            format="JPEG",
                            quality=quality,
                            optimize=True
                        )
                        compressed_data = output.getvalue()

                        # Only replace if smaller
                        if len(compressed_data) < len(image_data):
                            obj._data = compressed_data
                            obj[NameObject("/Filter")] = NameObject("/DCTDecode")
                            obj[NameObject("/ColorSpace")] = NameObject("/DeviceRGB")
                            if "/DecodeParms" in obj:
                                del obj["/DecodeParms"]

                    except Exception:
                        # Fallback: Flate encode
                        if hasattr(obj, "flate_encode"):
                            obj.flate_encode()

                else:
                    # Fallback for older PyPDF2 objects
                    if hasattr(obj, "flate_encode"):
                        obj.flate_encode()

            except Exception:
                # Never break the merge due to compression failure
                pass

    except Exception:
        # Fail silently — compression is optional
        pass
"""Create or convert an application icon.

This script will convert `pdfcombinericon.png` to `pdfcombinericon.ico` (Windows
friendly). If the PNG is missing, it falls back to generating a simple icon.
"""
import os
from PIL import Image, ImageDraw, ImageFont

PNG_NAME = "pdfcombinericon.png"
ICO_NAME = "pdfcombinericon.ico"


def convert_png_to_ico(png_path: str, ico_path: str):
    img = Image.open(png_path).convert("RGBA")

    # Ensure large size for quality then save multiple sizes in the .ico
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32)]
    icon_sizes = []
    for s in sizes:
        resized = img.resize(s, Image.Resampling.LANCZOS)
        icon_sizes.append(resized)

    # PIL accepts a list of images for the .ico format when saving the base image
    icon_sizes[0].save(ico_path, format="ICO", sizes=[s for s in sizes])


def generate_fallback_icon(png_path: str, ico_path: str):
    # Create a simple 256x256 white icon with a red PDF label
    size = 256
    img = Image.new("RGBA", (size, size), color=(255, 255, 255, 255))
    draw = ImageDraw.Draw(img)

    margin = 24
    draw.rectangle([margin, margin, size - margin, size - margin], outline=(204, 0, 0), width=8)
    draw.rectangle([margin, margin, size - margin, margin + 64], fill=(204, 0, 0))

    try:
        font = ImageFont.truetype("arial.ttf", 120)
    except Exception:
        font = ImageFont.load_default()

    text = "PDF"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = (size - text_width) // 2
    y = (size - text_height) // 2 + 30
    draw.text((x, y), text, fill=(204, 0, 0), font=font)

    # Save both PNG and ICO
    img.save(png_path)
    img.save(ico_path, format="ICO", sizes=[(256, 256), (128, 128), (64, 64)])


if __name__ == "__main__":
    if os.path.exists(PNG_NAME):
        try:
            convert_png_to_ico(PNG_NAME, ICO_NAME)
            print(f"Converted {PNG_NAME} -> {ICO_NAME}")
        except Exception as e:
            print(f"Failed to convert {PNG_NAME}: {e}")
            print("Generating fallback icon instead.")
            generate_fallback_icon(PNG_NAME, ICO_NAME)
            print(f"Generated fallback icon: {PNG_NAME}, {ICO_NAME}")
    else:
        print(f"{PNG_NAME} not found, generating a fallback icon.")
        generate_fallback_icon(PNG_NAME, ICO_NAME)
        print(f"Generated fallback icon: {PNG_NAME}, {ICO_NAME}")

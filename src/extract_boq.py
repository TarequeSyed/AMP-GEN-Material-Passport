"""Stage 1: render the scanned BoQ and create OCR text artifacts.

The source PDF is image-only/dot-matrix quality, so OCR is treated as a
candidate generator, not as ground truth. The reviewed extraction used by the
next stage is stored separately in data/reviewed_items.json.
"""
from pathlib import Path
import re
import fitz
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import pytesseract

ROOT = Path(__file__).resolve().parents[1]
PDF = ROOT / "input" / "BoQ_CBRI_Principals_Residence.pdf"
OUT = ROOT / "intermediate" / "ocr"


def preprocess(page, dpi=300):
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY)
    img = Image.frombytes("L", [pix.width, pix.height], pix.samples)
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Contrast(img).enhance(1.8)
    return img.filter(ImageFilter.MedianFilter(size=3))


def run_ocr():
    OUT.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(PDF)
    combined = []
    for number, page in enumerate(doc, start=1):
        img = preprocess(page)
        text = pytesseract.image_to_string(img, config="--oem 3 --psm 6")
        (OUT / f"page_{number:02d}.txt").write_text(text, encoding="utf-8")
        combined.append(f"\n\n===== PAGE {number} =====\n{text}")
    (OUT / "boq_ocr.txt").write_text("".join(combined), encoding="utf-8")
    print(f"OCR complete: {len(doc)} pages -> {OUT}")


if __name__ == "__main__":
    run_ocr()

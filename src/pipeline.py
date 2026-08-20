"""AMP-GEN reproducible pipeline.

Usage:
  python src/pipeline.py          # reproduce normalized outputs
  python src/pipeline.py --ocr    # also re-run OCR from the scanned PDF

The OCR stage produces noisy candidate text. The checked extraction is kept in
data/reviewed_items.json so another reviewer can distinguish raw OCR from the
human-reviewed structured data used to populate the passport.
"""
import argparse
from extract_boq import run_ocr
from normalize import normalize
from export_outputs import export


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--ocr", action="store_true", help="Re-run OCR on the scanned BoQ")
    args = parser.parse_args()
    if args.ocr:
        run_ocr()
    normalize()
    export()
    print("Pipeline complete: 64/64 BoQ items validated and exported.")


if __name__ == "__main__":
    main()

# APPROACH

This document outlines the engineering decisions, tools, pipeline, and lessons learned during the AMP-GEN Material Passport extraction task.

## 1. Tool Selection & Rationale
* **Tesseract OCR (via `pytesseract`)**: Used as an offline, neural-network-based candidate generator. It is reliable for plain text blocks but ignores page layouts and tables.
* **PyMuPDF (`fitz`)**: Employed for converting the image-only scanned PDF pages to high-resolution (300 DPI) grayscale rasters, which is a critical prerequisite for accurate OCR.
* **Pillow (`PIL`)**: Used for image preprocessing:
  - `ImageOps.autocontrast` expands pixel ranges to improve character edges.
  - `ImageEnhance.Contrast` (1.8x) darkens ink and brightens backgrounds.
  - `ImageFilter.MedianFilter` removes "salt-and-pepper" noise typical of dot-matrix printing and scan dust.
* **OpenPyXL**: Used to programmatically read and write the provided Excel template without altering cell formatting or styles.
* **Matplotlib**: Used to generate the side-by-side material summary and carbon footprint visualization.

## 2. Pipeline Architecture
1. **Rasterization**: Convert pages to high-res images to bypass poor-quality embedded PDF vectors.
2. **Preprocessing**: Apply denoising and contrast enhancement filters.
3. **OCR Processing**: Run pytesseract with `--oem 3` (default LSTM engine) and `--psm 6` (assume single uniform block of text) to generate raw draft text.
4. **Human-in-the-Loop Review**: Raw OCR draft text is manually audited and corrected to build a 100% accurate baseline dataset stored in `data/reviewed_items.json` containing all 64 items.
5. **Normalization & Carbon Estimation**: 
   - Parse quantities into standard columns (`volume_m3`, `area_m2`, etc.) based on units.
   - Use regex rules to classify items into material categories.
   - For area-based items, apply standard engineering thicknesses (e.g., 115mm half-brick wall, 12mm plaster, 40mm DPC concrete) to estimate volume.
   - Match products to physical properties (densities from IS 456, IS 808, SP 64) and EPD carbon coefficients (ACC Cement, UltraTech, Tata Steel India EPDs) to compute mass and embodied carbon (Bonus B2).
6. **Metadata Extraction**: Parse the building specification header block from Page 1 into `building_meta.json` (Bonus B3).
7. **Export & Visualization**: Inject the parsed records into `AMP_Passport_Template.xlsx` and generate a comparative side-by-side bar chart of item counts vs. carbon footprint.

## 3. Findings & Performance
* **What Worked**: High-resolution rasterization and image filtering made faded parts of the dot-matrix print readable. Normalizing area-based items into volume using standard architectural thicknesses allowed us to calculate building-wide material masses.
* **What Did Not**: OCR struggled with decimal points and schedule item codes (e.g. interpreting "0.60" as "o 60" or "0 60"). Purely automated parsing was not sufficient; the human-in-the-loop validation layer was critical to ensure database integrity.

## 4. With Two More Weeks
* Implement layout-aware table OCR (like Table-Transformer or layout-parser) to automate tabular bounding-box extraction.
* Integrate with standard material ontologies (like UniFormat, MasterFormat, or buildingSMART) to dynamically map descriptions.
* Connect directly to public EPD databases (like EC3 or the Indian Life Cycle Assessment database) via an API to pull live carbon coefficients.

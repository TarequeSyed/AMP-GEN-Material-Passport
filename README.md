# AMP-GEN Material Passport Extraction Pipeline

A reproducible python pipeline that extracts materials, quantities, and properties from a scanned dot-matrix Bill of Quantities (BoQ) and populates the target Material Passport template.

This repository is built for **Work Package 3** of the **AMP-GEN** project at IIT Roorkee, supported by the Google Centre for Climate Technology.

## Project Structure

```
├── data/
│   └── reviewed_items.json       # Human-reviewed baseline BoQ items (64 items)
├── input/
│   ├── AMP_Passport_Template.xlsx # Provided target schema template
│   └── BoQ_CBRI_Principals_Residence.pdf # Dot-matrix scanned source PDF
├── output/
│   ├── passport_filled.xlsx      # Populated Excel Material Passport (GREEN & AMBER fields)
│   ├── passport.json             # Structured JSON export of passport data
│   ├── visualization.png         # Side-by-side count & Embodied Carbon summary chart
│   └── building_meta.json        # Extracted building metadata from Page 1
├── src/
│   ├── extract_boq.py            # Stage 1: Rasterization, preprocessing & raw OCR
│   ├── normalize.py              # Stage 2: Regex categorization & mass/carbon calculation
│   ├── export_outputs.py         # Stage 3: openpyxl mapping & matplotlib plotting
│   └── pipeline.py               # Orchestrator script
├── APPROACH.md                   # Engineering choices, shortcomings & roadmap
├── requirements.txt              # Project dependencies
└── README.md                     # This file
```

## Quick Start (Run in under 5 minutes)

### 1. Prerequisites
Ensure you have **Python 3.10+** and [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) installed on your system path.

### 2. Installation
Install the required packages:
```bash
pip install -r requirements.txt
```

### 3. Execution
To run the normalization, carbon calculation, and output generation using the reviewed database:
```bash
python src/pipeline.py
```

To re-run the entire pipeline from scratch, including rasterizing and OCR-ing the raw PDF:
```bash
python src/pipeline.py --ocr
```

---

## Completed Bonuses

### B2. Mass & Carbon Integration (AMBER columns)
* **Density (kg/m³)**, **GWP / kg (kg CO₂e/kg)**, and **Embodied Carbon A1-A3 (kg CO₂e)** are calculated and populated in `output/passport_filled.xlsx` and `output/passport.json` for key materials (concrete, bricks, mild steel, plaster, aluminium, and timber).
* Mass is estimated dynamically. For area-based items, standard engineering thicknesses (DPC concrete = 40mm, half-brick walls = 115mm, cement plaster = 12mm) are used to derive volume and mass.
* Embodied carbon calculations are backed by Indian EPD sources (ACC Cement, UltraTech, Tata Steel India, Hindalco) and cited in the spreadsheet comments.

### B3. Building Metadata Block
* Page 1 specifications (depth of foundation, plinth height, plinth area, seismic zone) have been extracted and structured in `output/building_meta.json`.

---

## Estimation & Metrics
* **Items Extracted**: 64 of 64 BoQ items.
* **Hours Actually Spent**: 4.5 focused hours (including setup, OCR tuning, manual correction, EPD mapping, charting, and documentation).

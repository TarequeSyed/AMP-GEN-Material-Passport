"""Stage 2: normalize reviewed BoQ items into Material Passport records."""
from pathlib import Path
import json
import re

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data" / "reviewed_items.json"
OUT = ROOT / "intermediate" / "normalized_items.json"

CATEGORY_RULES = [
    (r"earth|excavat", "Earthwork"),
    (r"reinforced cement concrete|concrete|cement concrete|rcc", "Concrete"),
    (r"reinforcement|steel bars|mild steel|ms guard|ms frame", "Metals"),
    (r"brick|masonry", "Masonry"),
    (r"wood|timber|flush door|shutter", "Wood & Joinery"),
    (r"aluminium|alumin", "Metals"),
    (r"glass|glazing", "Glass"),
    (r"plaster|terraco|marble|tiles|flooring", "Finishes"),
    (r"paint|primer|white washing|colour washing", "Coatings & Finishes"),
]


def classify(description, fallback):
    text = description.lower()
    for pattern, category in CATEGORY_RULES:
        if re.search(pattern, text):
            return category
    return fallback or "Other"


# EPD and physical properties for materials.
# Densities: RCC (IS 456), brick/plaster (IS 2250/SP 64), mild steel (IS 808).
# GWP factors: Indian average concrete/cement EPDs (ACC, UltraTech), Tata Steel EPD.
MATERIAL_EPD_REGISTRY = {
    "reinforced cement concrete": {
        "density": 2400.0,
        "gwp": 0.12,
        "ref": "ACC Cement concrete EPD; density from IS 456"
    },
    "cement concrete": {
        "density": 2300.0,
        "gwp": 0.10,
        "ref": "ACC Cement concrete EPD; density from IS 456"
    },
    "burnt clay brick": {
        "density": 1900.0,
        "gwp": 0.24,
        "ref": "Indian Brick Industry average EPD; density from SP 64"
    },
    "aluminium": {
        "density": 2700.0,
        "gwp": 8.20,
        "ref": "Hindalco Aluminium average Indian EPD"
    },
    "cement mortar": {
        "density": 2000.0,
        "gwp": 0.15,
        "ref": "UltraTech Cement plaster EPD; density from IS 2250"
    },
    "marble chips": {
        "density": 2200.0,
        "gwp": 0.18,
        "ref": "Indian Marble Federation average EWP; density from IS 2250"
    },
    "mild steel": {
        "density": 7850.0,
        "gwp": 1.85,
        "ref": "Tata Steel India EPD; density from IS 808"
    },
    "oxidised mild steel": {
        "density": 7850.0,
        "gwp": 1.85,
        "ref": "Tata Steel India EPD; density from IS 808"
    },
    "cast iron": {
        "density": 7200.0,
        "gwp": 2.10,
        "ref": "Indian Cast Iron average EPD"
    },
    "timber": {
        "density": 700.0,
        "gwp": 0.35,
        "ref": "Indian Forestry average EPD; density from IS 287"
    },
    "teak wood": {
        "density": 700.0,
        "gwp": 0.35,
        "ref": "Indian Forestry average EPD; density from IS 287"
    }
}

# Standard thicknesses (in meters) to convert area (sqm) to volume (m3).
STANDARD_THICKNESSES = {
    "9": 0.040,   # DPC concrete (40mm)
    "21": 0.115,  # half brick wall (115mm)
    "22": 0.115,  # half brick wall (115mm)
    "23": 0.115,  # half brick wall (115mm)
    "34": 0.005,  # steel glazing glass thickness estimation (5mm)
    "41": 0.012,  # plaster (12mm)
    "42": 0.020,  # marble chip flooring (20mm)
    "45": 0.100,  # lime concrete roofing (100mm)
    "46": 0.020,  # tile roofing (20mm)
    "52": 0.012,  # plaster (12mm)
    "53": 0.012,  # plaster (12mm)
    "54": 0.012,  # plaster (12mm)
    "55": 0.020,  # marble flooring (20mm)
    "56": 0.012,  # plaster (12mm)
    "64": 0.075,  # sub-base concrete (75mm)
}


def normalize():
    items = json.loads(SOURCE.read_text(encoding="utf-8"))
    normalized = []
    for item in items:
        item = dict(item)
        item["description"] = " ".join(item["description"].split())
        item["material_category"] = classify(
            item["description"], item.get("material_category")
        )
        # Keep source quantity and source unit separate from derived SI values.
        unit = (item.get("original_unit") or "").lower()
        q = item.get("original_quantity")
        item["volume_m3"] = q if unit in {"cum", "m3", "m³"} else item.get("volume_m3")
        item["area_m2"] = q if unit in {"sqm", "m2", "m²"} else item.get("area_m2")
        item["length_m"] = q if unit in {"mtr", "m", "metre"} else item.get("length_m")
        item["weight_kg"] = q if unit in {"kg", "kgs"} else item.get("weight_kg")
        item["count_nos"] = q if unit in {"each", "no", "nos"} else item.get("count_nos")

        # Estimate volume from area if standard thickness is known
        item_no = str(item.get("boq_item_no"))
        if item.get("volume_m3") is None and item.get("area_m2") is not None:
            if item_no in STANDARD_THICKNESSES:
                item["volume_m3"] = round(item["area_m2"] * STANDARD_THICKNESSES[item_no], 4)

        # Carbon and density calculation
        prod = (item.get("material_product") or "").lower().strip()
        matched_key = None
        for key in MATERIAL_EPD_REGISTRY:
            if key in prod:
                matched_key = key
                break

        if matched_key:
            data = MATERIAL_EPD_REGISTRY[matched_key]
            item["density_kg_m3"] = data["density"]
            item["gwp_per_kg_co2e"] = data["gwp"]

            # Calculate mass (weight_kg)
            w = item.get("weight_kg")
            if w is None:
                v = item.get("volume_m3")
                if v is not None:
                    w = round(v * data["density"], 2)
                    item["weight_kg"] = w

            if w is not None:
                carbon = round(w * data["gwp"], 2)
                item["embodied_carbon_kg_co2e"] = carbon
                
                # Append to comment
                old_comment = item.get("comment") or ""
                comment_clean = re.sub(r" \| \[Carbon\].*$", "", old_comment)
                item["comment"] = f"{comment_clean} | [Carbon] EPD-backed estimate: {carbon} kg CO₂e ({data['ref']}).".strip(" | ")

        normalized.append(item)

    numbers = {str(x["boq_item_no"]) for x in normalized}
    expected = {str(i) for i in range(1, 65)}
    missing = expected - numbers
    if missing:
        raise ValueError(f"Missing BoQ items: {sorted(missing)}")
    if len(normalized) != 64:
        raise ValueError(f"Expected 64 records, got {len(normalized)}")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(normalized, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Normalized {len(normalized)} items -> {OUT}")
    return normalized



if __name__ == "__main__":
    normalize()

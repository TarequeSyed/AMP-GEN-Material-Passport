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

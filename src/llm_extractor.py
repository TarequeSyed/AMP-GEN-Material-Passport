"""
LLM-based semantic material extraction for AMP-GEN pipeline.

Architecture:
    Scanned PDF → OCR → Local LLM (Ollama) → Confidence Check
                                                  ↙         ↘
                                           High/Medium    Low/Failed
                                                ↓              ↓
                                           LLM output   Rule-based engine
                                                ↘              ↙
                                              Validation
                                                  ↓
                                         Material Passport

Ollama is optional. If it is not running, every call automatically
falls back to the rule-based classify() function without raising an
error. This ensures the app works identically on Streamlit Cloud.
"""

import json
import re
import requests
from typing import Optional

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
OLLAMA_BASE_URL  = "http://localhost:11434"
OLLAMA_API_URL   = f"{OLLAMA_BASE_URL}/api/generate"
DEFAULT_MODEL    = "llama3.2"          # ~2 GB; change to "mistral" or "phi3" if preferred
OLLAMA_TIMEOUT   = 20                  # seconds per request
AVAILABLE_MODELS = ["llama3.2", "mistral", "phi3", "gemma2"]

# Valid material categories (mirrors CATEGORY_RULES in normalize.py)
VALID_CATEGORIES = {
    "Earthwork", "Concrete", "Metals", "Masonry",
    "Wood & Joinery", "Glass", "Finishes",
    "Coatings & Finishes", "Chemical", "Aggregate / Fill", "Other"
}

# ---------------------------------------------------------------------------
# Prompt template
# ---------------------------------------------------------------------------
EXTRACTION_PROMPT = '''You are an expert in Indian construction materials and Bill of Quantities (BoQ) documents.

Analyse this BoQ line item description and extract structured information.

Description: "{description}"

Classify the material into EXACTLY one of these categories:
Earthwork | Concrete | Metals | Masonry | Wood & Joinery | Glass | Finishes | Coatings & Finishes | Chemical | Aggregate / Fill | Other

Respond with ONLY valid JSON — no explanation, no markdown:
{{
  "material_category": "<category from list above>",
  "material_product": "<specific product name, 2-5 words>",
  "confidence": "<High | Medium | Low>",
  "reasoning": "<one concise sentence>"
}}'''

# ---------------------------------------------------------------------------
# Connectivity check
# ---------------------------------------------------------------------------
def check_ollama_available() -> tuple[bool, list[str]]:
    """
    Returns (is_available: bool, available_models: list[str]).
    Pings Ollama tags endpoint with a short timeout so UI stays responsive.
    """
    try:
        resp = requests.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=2)
        if resp.status_code == 200:
            data = resp.json()
            models = [m["name"].split(":")[0] for m in data.get("models", [])]
            return True, models
        return False, []
    except Exception:
        return False, []


# ---------------------------------------------------------------------------
# LLM extraction
# ---------------------------------------------------------------------------
def extract_with_llm(description: str, model: str = DEFAULT_MODEL) -> Optional[dict]:
    """
    Sends a single BoQ description to the local Ollama LLM and parses
    the structured JSON response.

    Returns:
        dict with keys: material_category, material_product, confidence, reasoning
        None if Ollama is unreachable or the response cannot be parsed.
    """
    try:
        payload = {
            "model": model,
            "prompt": EXTRACTION_PROMPT.format(description=description[:500]),
            "stream": False,
            "format": "json",
            "options": {
                "temperature": 0.05,   # near-deterministic for classification
                "num_predict": 256,
                "top_p": 0.9,
            }
        }
        resp = requests.post(OLLAMA_API_URL, json=payload, timeout=OLLAMA_TIMEOUT)
        resp.raise_for_status()

        raw_response = resp.json().get("response", "")
        result = json.loads(raw_response)

        # Validate required fields exist
        if not all(k in result for k in ("material_category", "material_product", "confidence")):
            return None

        # Sanitise category to one of the known values
        cat = result["material_category"].strip()
        if cat not in VALID_CATEGORIES:
            # Try a case-insensitive match
            for valid in VALID_CATEGORIES:
                if valid.lower() == cat.lower():
                    result["material_category"] = valid
                    break
            else:
                result["material_category"] = "Other"

        # Sanitise confidence
        conf = result.get("confidence", "Low").strip().capitalize()
        result["confidence"] = conf if conf in ("High", "Medium", "Low") else "Low"

        return result

    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
        # Ollama not running — silent failure
        return None
    except (json.JSONDecodeError, KeyError, ValueError):
        # LLM returned malformed JSON — silent failure
        return None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Confidence router — the main public interface
# ---------------------------------------------------------------------------
def classify_with_confidence(
    description: str,
    fallback_classify_fn,
    model: str = DEFAULT_MODEL
) -> dict:
    """
    Hybrid classifier that routes between LLM and rule-based engine.

    Decision logic:
        LLM confidence = High   → use LLM output
        LLM confidence = Medium → use LLM output, flag for review
        LLM unavailable / Low   → delegate to rule-based fallback

    Args:
        description:         BoQ line item description string
        fallback_classify_fn: the existing classify(description, fallback) function
        model:               Ollama model name to use

    Returns:
        dict with keys:
            material_category  — final classification
            material_product   — specific product name
            confidence         — "High" | "Medium" | "Low" | "Rule-based"
            extraction_method  — "LLM" | "LLM (review)" | "Rule-based engine"
            reasoning          — explanation string
    """
    llm_result = extract_with_llm(description, model)

    if llm_result is not None:
        conf = llm_result.get("confidence", "Low")

        if conf == "High":
            return {
                "material_category": llm_result["material_category"],
                "material_product":  llm_result.get("material_product", ""),
                "confidence":        "High",
                "extraction_method": "LLM",
                "reasoning":         llm_result.get("reasoning", "LLM high-confidence match."),
            }

        if conf == "Medium":
            return {
                "material_category": llm_result["material_category"],
                "material_product":  llm_result.get("material_product", ""),
                "confidence":        "Medium",
                "extraction_method": "LLM (review)",
                "reasoning":         llm_result.get("reasoning", "LLM medium-confidence — manual review recommended."),
            }

    # Low confidence, no response, or Ollama offline → rule-based fallback
    rule_category = fallback_classify_fn(description, "Other")
    return {
        "material_category": rule_category,
        "material_product":  _infer_product_from_description(description),
        "confidence":        "Rule-based",
        "extraction_method": "Rule-based engine",
        "reasoning":         "Ollama unavailable or LLM confidence too low — regex engine applied.",
    }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _infer_product_from_description(description: str) -> str:
    """
    Lightweight product-name guesser when LLM is unavailable.
    Takes the first 3-5 meaningful words from the description.
    """
    stop_words = {"in", "of", "with", "and", "for", "the", "a", "an",
                  "including", "excluding", "nominal", "mix", "all", "kinds"}
    words = [w for w in description.split() if w.lower() not in stop_words]
    return " ".join(words[:4]).rstrip(".,;:") or description[:40]


def build_pipeline_status(model: str = DEFAULT_MODEL) -> dict:
    """
    Returns a status dict for rendering in the Streamlit UI.
    Combines availability check with model info.
    """
    available, models = check_ollama_available()
    model_ready = model in models if available else False
    return {
        "ollama_available": available,
        "model_ready":      model_ready,
        "available_models": models,
        "selected_model":   model,
        "status_label":     (
            f"Ollama online — {model} ready"   if (available and model_ready)  else
            f"Ollama online — {model} not pulled (run: ollama pull {model})" if (available and not model_ready) else
            "Ollama offline — rule-based fallback active"
        ),
        "status_color": (
            "#16a34a" if (available and model_ready)  else
            "#d97706" if (available and not model_ready) else
            "#64748b"
        ),
    }

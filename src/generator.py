"""
Phase 2 — LLM Content Generator
Reads layout.json, calls Groq API slide-by-slide, and produces
content_mapping.json with new text content mapped to each shape ID.
"""

import json
import os
import re
from pathlib import Path

from dotenv import load_dotenv
from groq import Groq

load_dotenv()

MODEL = "llama-3.3-70b-versatile"

# Placeholder types that are always single-line
SINGLE_LINE_PH = {"CENTER_TITLE", "TITLE", "SUBTITLE"}

# Shape types that have no replaceable text — skip them
SKIP_SHAPE_TYPES = {"CHART", "TABLE", "PICTURE"}


# ---------------------------------------------------------------------------
# Prompt builder
# ---------------------------------------------------------------------------

def _collect_targets(slide: dict) -> list[dict]:
    """Return list of shapes on this slide that need new text content."""
    targets = []
    for shape in slide.get("shapes", []):
        if not shape.get("has_text_frame"):
            continue

        shape_type = shape.get("shape_type", "")
        if any(skip in shape_type for skip in SKIP_SHAPE_TYPES):
            continue

        tf = shape.get("text_frame") or {}
        full_text = tf.get("full_text", "").strip()
        if not full_text:
            continue

        para_count = tf.get("paragraph_count", 1)
        ph_type = shape.get("placeholder_type") or "TEXT_BOX"

        targets.append({
            "shape_id":      str(shape["shape_id"]),
            "placeholder_type": ph_type,
            "current_text":  full_text[:150],
            "paragraph_count": para_count,
            "char_hint":     len(full_text),
        })

    return targets


def _build_prompt(slide: dict, targets: list[dict], topic: str, user_content: str = "") -> str:
    shapes_block = ""
    for t in targets:
        para_note = (
            f"{t['paragraph_count']} paragraphs — return a JSON array of exactly {t['paragraph_count']} strings"
            if t["paragraph_count"] > 1
            else "1 paragraph — return a plain JSON string"
        )
        shapes_block += (
            f'\n  shape_id "{t["shape_id"]}" [{t["placeholder_type"]}]'
            f" | {para_note}"
            f' | original: "{t["current_text"][:80]}"'
        )

    # ── Mode A: user provided real content — map it into placeholders ──────
    if user_content.strip():
        return f"""You are mapping user-provided content into a PowerPoint slide.

Topic: "{topic}"
Slide layout: "{slide['layout_name']}"

User provided this source content:
\"\"\"{user_content.strip()}\"\"\"

Rules:
- Extract and fit relevant parts of the user content into each shape below
- Do NOT invent new information — only use what the user provided
- TITLE / CENTER_TITLE: pull the most important idea, max 7 words
- SUBTITLE: pull a supporting line, max 15 words
- For BODY shapes with paragraph_count > 1: extract exactly that many key points from the content
- For shapes with paragraph_count = 1: return a plain JSON string
- For shapes with paragraph_count > 1: return a JSON ARRAY with EXACTLY that many strings
- Do NOT add bullet symbols (-, •, *) inside the strings
- Return ONLY a valid JSON object, no markdown fences, no explanation

Shapes to fill:{shapes_block}

Return format example:
{{
  "2": "Short Title Here",
  "3": "Supporting subtitle that fits",
  "4": ["Key point one", "Key point two", "Key point three"]
}}"""

    # ── Mode B: no user content — LLM freely generates from topic ──────────
    return f"""You are generating new PowerPoint slide content about: "{topic}".

Slide layout: "{slide['layout_name']}"

Rules:
- TITLE / CENTER_TITLE: max 7 words
- SUBTITLE: max 15 words
- For each shape with paragraph_count > 1: return a JSON ARRAY with EXACTLY that many strings
- For each shape with paragraph_count = 1: return a plain JSON string
- Do NOT add bullet symbols (-, •, *) inside the strings
- Return ONLY a valid JSON object, no markdown fences, no explanation

Shapes to generate:{shapes_block}

Return format example:
{{
  "2": "Short Title Here",
  "3": "Supporting subtitle sentence that fits the topic",
  "4": ["First bullet point", "Second bullet point", "Third bullet point"]
}}"""


# ---------------------------------------------------------------------------
# Groq call + JSON parsing
# ---------------------------------------------------------------------------

def _strip_fences(raw: str) -> str:
    """Remove markdown ```json ... ``` wrappers if the model added them."""
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return raw.strip()


def _call_groq(client: Groq, prompt: str) -> dict:
    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a precise JSON generator for PowerPoint content. "
                    "Always return only valid JSON with no extra text or markdown."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.65,
        max_tokens=1200,
    )
    raw   = response.choices[0].message.content
    clean = _strip_fences(raw)
    return json.loads(clean)


# ---------------------------------------------------------------------------
# Validation + auto-fix
# ---------------------------------------------------------------------------

def _validate_and_fix(raw_result: dict, targets: list[dict]) -> dict:
    """
    Ensure every target has content and paragraph counts match exactly.
    Auto-fixes mismatches; falls back to original text if shape is missing.
    """
    fixed = {}

    for t in targets:
        sid        = t["shape_id"]
        para_count = t["paragraph_count"]
        value      = raw_result.get(sid)

        if value is None:
            # LLM missed this shape — fallback to original text
            fixed[sid] = (
                [t["current_text"]] + [""] * (para_count - 1)
                if para_count > 1
                else t["current_text"]
            )
            continue

        if para_count > 1:
            # Expect a list
            if isinstance(value, str):
                # Split on newlines as best effort
                parts = [p.strip() for p in value.split("\n") if p.strip()]
                if not parts:
                    parts = [value]
            else:
                parts = list(value)

            # Pad or trim to exact count
            if len(parts) < para_count:
                parts += [""] * (para_count - len(parts))
            else:
                parts = parts[:para_count]

            fixed[sid] = parts

        else:
            # Expect a plain string
            if isinstance(value, list):
                value = " ".join(str(v) for v in value)
            fixed[sid] = str(value)

    return fixed


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def generate_content(layout_path: str, output_path: str, topic: str,
                     user_content: str = "") -> dict:
    """
    Read layout JSON, generate new content via Groq, save content_mapping.json.

    If user_content is provided the LLM acts as a content mapper — it extracts
    and fits the user's text into each placeholder instead of inventing content.
    If omitted the LLM freely generates content from topic alone.

    Returns the content map dict.
    """
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY not found — add it to your .env file")

    client = Groq(api_key=api_key)

    with open(layout_path, "r", encoding="utf-8") as f:
        layout = json.load(f)

    if not topic.strip():
        topic = "Modern Technology and Innovation"

    mode = "mapper" if user_content.strip() else "generator"

    content_map = {
        "source_file":  layout["source_file"],
        "topic":        topic,
        "mode":         mode,
        "user_content": user_content.strip() or None,
        "model":        MODEL,
        "slides":       {},
    }

    print(f"  Topic   : {topic}")
    print(f"  Mode    : {mode} ({'using your content' if mode == 'mapper' else 'LLM generates freely'})")
    print(f"  Model   : {MODEL}")
    print(f"  Slides  : {layout['slide_count']}\n")

    for slide in layout["slides"]:
        slide_key = f"slide_{slide['slide_index']}"
        targets   = _collect_targets(slide)

        if not targets:
            print(f"  Slide {slide['slide_index']} [{slide['layout_name']}]"
                  f" -- skipped (no replaceable text shapes)")
            content_map["slides"][slide_key] = {}
            continue

        print(f"  Slide {slide['slide_index']} [{slide['layout_name']}]"
              f" -- {len(targets)} shape(s) ...", end=" ", flush=True)

        prompt = _build_prompt(slide, targets, topic, user_content)

        try:
            raw    = _call_groq(client, prompt)
            result = _validate_and_fix(raw, targets)
            content_map["slides"][slide_key] = result
            print("done")
        except json.JSONDecodeError as e:
            print(f"JSON parse error: {e} -- keeping original text")
            content_map["slides"][slide_key] = {
                t["shape_id"]: t["current_text"] for t in targets
            }
        except Exception as e:
            print(f"ERROR: {e} -- keeping original text")
            content_map["slides"][slide_key] = {
                t["shape_id"]: t["current_text"] for t in targets
            }

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(content_map, f, indent=2, ensure_ascii=False)

    print(f"\n[Phase 2] Content mapping saved -> {output_path}")
    return content_map


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Phase 2 -- Generate content via Groq LLM")
    parser.add_argument("--layout", required=True, help="Path to layout.json (Phase 1 output)")
    parser.add_argument("--output", required=True, help="Path for content_mapping.json")
    parser.add_argument("--topic",  required=True, help="Topic for the new slide content")
    args = parser.parse_args()

    generate_content(args.layout, args.output, args.topic)

"""
Phase 2 — LLM Content Generator
Reads layout.json, calls Groq API and produces
content_mapping.json with new text content mapped to each shape ID.
Uses a two-pass approach: Pass 1 plans unique focus per slide,
Pass 2 generates detailed content using that focus.
"""

import json
import os
import re
from pathlib import Path

from dotenv import load_dotenv
from groq import Groq

load_dotenv()

MODEL = "llama-3.1-8b-instant"

# Placeholder types that are always single-line
SINGLE_LINE_PH = {"CENTER_TITLE", "TITLE", "SUBTITLE"}

# Shape types that have no replaceable text — skip them
SKIP_SHAPE_TYPES = {"CHART", "TABLE", "PICTURE"}


# ---------------------------------------------------------------------------
# Target collection
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
            "shape_id":         str(shape["shape_id"]),
            "placeholder_type": ph_type,
            "current_text":     full_text[:150],
            "paragraph_count":  para_count,
            "char_hint":        len(full_text),
        })

    return targets


def _slide_text_summary(slide: dict) -> str:
    """Short single-line summary of what text is on this slide."""
    texts = []
    for shape in slide.get("shapes", []):
        tf = shape.get("text_frame") or {}
        t = tf.get("full_text", "").strip()
        if t:
            texts.append(t[:60])
    return " | ".join(texts)[:120] if texts else "(empty slide)"


# ---------------------------------------------------------------------------
# Pass 1 — Planning
# ---------------------------------------------------------------------------

def _build_plan_prompt(layout: dict, topic: str, user_content: str) -> str:
    slide_lines = []
    for slide in layout["slides"]:
        idx = slide["slide_index"]
        summary = _slide_text_summary(slide)
        slide_lines.append(f'  slide_{idx}: "{summary}"')

    slides_block = "\n".join(slide_lines)

    intent_section = ""
    if user_content.strip():
        intent_section = (
            "\nUser intent / description:\n"
            f"[{user_content.strip()[:400]}]\n"
            "Use this intent to guide the focus for each slide.\n"
        )

    lines = [
        f'You are planning a PowerPoint presentation about: "{topic}".',
        "",
        "Below are the slides and their current text (for layout context only):",
        slides_block,
        intent_section,
        "Assign a UNIQUE, specific focus to each slide so the deck tells a",
        "coherent narrative — no two slides should repeat the same angle.",
        "",
        'Return ONLY a valid JSON object like: {"slide_0": "focus...", "slide_1": "focus...", ...}',
        "One key per slide. Focus should be 5-10 words. No markdown, no explanation.",
    ]
    return "\n".join(lines)


def _call_plan(client: Groq, layout: dict, topic: str, user_content: str) -> dict:
    prompt = _build_plan_prompt(layout, topic, user_content)
    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a presentation strategist. "
                    "Return only valid JSON mapping slide keys to unique focus statements."
                ),
            },
            {"role": "user", "content": prompt},
        ],
        temperature=0.7,
        max_tokens=600,
    )
    raw = response.choices[0].message.content.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw.strip())


# ---------------------------------------------------------------------------
# Pass 2 — Per-slide prompt builder
# ---------------------------------------------------------------------------

def _build_prompt(slide: dict, targets: list[dict], topic: str,
                  slide_focus: str, user_content: str = "") -> str:
    shapes_block_lines = []
    for t in targets:
        n  = t["paragraph_count"]
        ph = t["placeholder_type"]
        if ph in ("CENTER_TITLE", "TITLE"):
            word_rule = "max 7 words"
        elif ph == "SUBTITLE":
            word_rule = "max 15 words"
        elif n > 8:
            word_rule = "max 5 words per bullet — very short phrases only"
        elif n > 5:
            word_rule = "max 7 words per bullet — short phrases only"
        elif n > 1:
            word_rule = "max 10 words per bullet"
        else:
            word_rule = "max 12 words — short phrase only"

        if n > 1:
            example_items = ", ".join([f'"Point {i+1}"' for i in range(min(n, 3))])
            if n > 3:
                example_items += f", ... ({n} total)"
            para_note = (
                f'ARRAY of EXACTLY {n} short strings ({word_rule}). '
                f'Example: [{example_items}]'
            )
        else:
            para_note = f"plain JSON string — {word_rule}"

        shapes_block_lines.append(
            f'  shape_id "{t["shape_id"]}" [{ph}]'
            f" | {para_note}"
            f' | original: "{t["current_text"][:60]}"'
        )
    shapes_block = "\n".join(shapes_block_lines)

    lines = [
        f'Topic: "{topic}"',
        f'This slide focus: "{slide_focus}"',
        f'Slide layout: "{slide["layout_name"]}"',
        "",
    ]

    if user_content.strip():
        lines += [
            "User provided this intent / description:",
            f"[{user_content.strip()[:300]}]",
            "",
            "Rules:",
            "- Use the user intent to guide content — do not invent unrelated information",
            "- TITLE / CENTER_TITLE: max 7 words, captures the slide focus",
            "- SUBTITLE: max 15 words",
            "- BODY shapes: write substantive content — NO empty strings, NO placeholders",
            "- NEVER use empty strings in arrays",
            "- max 15 words per bullet",
            "- Do NOT add bullet symbols (-, bullet, *) inside strings",
            "- For arrays: return EXACTLY the required number of strings",
            "- Return ONLY valid JSON, no markdown fences, no explanation",
        ]
    else:
        lines += [
            "Rules:",
            "- TITLE / CENTER_TITLE: max 7 words, captures the slide focus",
            "- SUBTITLE: max 15 words",
            "- BODY shapes: write substantive, specific content about the slide focus",
            "- NEVER use empty strings in arrays",
            "- max 15 words per bullet",
            "- Do NOT add bullet symbols (-, bullet, *) inside strings",
            "- For arrays: return EXACTLY the required number of strings",
            "- Return ONLY valid JSON, no markdown fences, no explanation",
        ]

    lines += [
        "",
        "Shapes to fill:",
        shapes_block,
        "",
        "Return format:",
        "{",
    ]
    for t in targets:
        sid = t["shape_id"]
        if t["paragraph_count"] > 1:
            lines.append(f'  "{sid}": ["string 1", "string 2", ...]')
        else:
            lines.append(f'  "{sid}": "string"')
    lines.append("}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Groq call + JSON parsing
# ---------------------------------------------------------------------------

def _strip_fences(raw: str) -> str:
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
                    "Always return only valid JSON. "
                    "Never use empty strings. Fill every array slot with real content."
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
# Word-cap helpers
# ---------------------------------------------------------------------------

TITLE_MAX_WORDS        = 7
SUBTITLE_MAX_WORDS     = 15
BODY_SINGLE_MAX_WORDS  = 12


def _cap_words(text: str, max_words: int) -> str:
    """Truncate text to at most max_words words."""
    words = text.split()
    return " ".join(words[:max_words]) if len(words) > max_words else text


def _word_limit_for(ph_type: str, para_count: int) -> int:
    """Return word cap — fewer words per bullet when there are many bullets."""
    if ph_type in ("CENTER_TITLE", "TITLE"):
        return TITLE_MAX_WORDS
    if ph_type == "SUBTITLE":
        return SUBTITLE_MAX_WORDS
    if para_count > 8:
        return 5   # 9+ bullets: very short phrases only
    if para_count > 5:
        return 7   # 6-8 bullets: short phrases
    if para_count > 1:
        return 10  # 2-5 bullets: standard
    return BODY_SINGLE_MAX_WORDS


# ---------------------------------------------------------------------------
# Validation + auto-fix
# ---------------------------------------------------------------------------

def _validate_and_fix(raw_result: dict, targets: list[dict]) -> dict:
    """
    Ensure every target has content and paragraph counts match exactly.
    Cycles existing content instead of padding with empty strings.
    Caps word count per item to prevent visual overflow.
    """
    fixed = {}

    for t in targets:
        sid        = t["shape_id"]
        para_count = t["paragraph_count"]
        ph_type    = t.get("placeholder_type", "BODY")
        max_words  = _word_limit_for(ph_type, para_count)
        value      = raw_result.get(sid)

        if value is None:
            # LLM missed this shape — fallback to original text
            if para_count > 1:
                base = [s.strip() for s in t["current_text"].split("\n") if s.strip()]
                if not base:
                    base = [t["current_text"]]
                while len(base) < para_count:
                    base.append(base[len(base) % len(base)])
                fixed[sid] = [_cap_words(p, max_words) for p in base[:para_count]]
            else:
                fixed[sid] = _cap_words(t["current_text"], max_words)
            continue

        if para_count > 1:
            if isinstance(value, str):
                parts = [p.strip() for p in value.split("\n") if p.strip()]
                if not parts:
                    parts = [value]
            else:
                parts = [str(v).strip() for v in value]

            # Remove empty strings the LLM slipped in
            parts = [p for p in parts if p]

            if not parts:
                parts = [t["current_text"]]

            # Cycle to fill if short — never pad with ""
            if len(parts) < para_count:
                base = parts[:]
                while len(parts) < para_count:
                    parts.append(base[len(parts) % len(base)])

            fixed[sid] = [_cap_words(p, max_words) for p in parts[:para_count]]

        else:
            if isinstance(value, list):
                value = " ".join(str(v) for v in value if v)
            fixed[sid] = _cap_words(str(value).strip() or t["current_text"], max_words)

    return fixed


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def generate_content(layout_path: str, output_path: str, topic: str,
                     user_content: str = "") -> dict:
    """
    Read layout JSON, generate new content via Groq, save content_mapping.json.
    Uses two-pass generation: Pass 1 plans unique focus per slide,
    Pass 2 generates detailed content slide-by-slide using that focus.
    """
    api_key = os.getenv("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("GROQ_API_KEY not found — add it to your .env file")

    client = Groq(api_key=api_key)

    with open(layout_path, "r", encoding="utf-8") as f:
        layout = json.load(f)

    if not topic.strip():
        topic = "Modern Technology and Innovation"

    mode = "intent-guided" if user_content.strip() else "free-generator"

    content_map = {
        "source_file":  layout["source_file"],
        "topic":        topic,
        "mode":         mode,
        "user_intent":  user_content.strip() or None,
        "model":        MODEL,
        "slides":       {},
    }

    print(f"  Topic   : {topic}")
    print(f"  Mode    : {mode}")
    print(f"  Model   : {MODEL}")
    print(f"  Slides  : {layout['slide_count']}\n")

    # ── Pass 1: Generate narrative plan ───────────────────────────────────
    print(f"  [Plan] Generating narrative plan for {layout['slide_count']} slides ...", end=" ", flush=True)
    plan = {}
    try:
        plan = _call_plan(client, layout, topic, user_content)
        print("done")
        content_map["plan"] = plan
    except Exception as e:
        print(f"failed ({e}) — using sequential fallback")
        for slide in layout["slides"]:
            idx = slide["slide_index"]
            plan[f"slide_{idx}"] = f"{topic} — part {idx + 1}"

    print()

    # ── Pass 2: Generate per-slide content ────────────────────────────────
    for slide in layout["slides"]:
        slide_key  = f"slide_{slide['slide_index']}"
        slide_focus = plan.get(slide_key, f"{topic} — overview")
        targets    = _collect_targets(slide)

        if not targets:
            print(f"  Slide {slide['slide_index']} [{slide['layout_name']}]"
                  f" -- skipped (no replaceable text shapes)")
            content_map["slides"][slide_key] = {}
            continue

        print(f"  Slide {slide['slide_index']} [{slide['layout_name']}]"
              f" -- {len(targets)} shape(s) | focus: \"{slide_focus[:50]}\" ...",
              end=" ", flush=True)

        prompt = _build_prompt(slide, targets, topic, slide_focus, user_content)

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
    parser.add_argument("--content", default="", help="User intent / description")
    args = parser.parse_args()

    generate_content(args.layout, args.output, args.topic, args.content)

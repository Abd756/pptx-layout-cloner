"""
PPTX Layout Cloner — Main Orchestrator
Runs all 4 phases in sequence and places all outputs in a timestamped
results/ subfolder so every run is self-contained and client-reviewable.

Usage (generate freely from topic):
    python main.py --input input/deck.pptx --topic "AI in Healthcare"

Usage (map user-provided content into placeholders):
    python main.py --input input/deck.pptx --topic "Q3 Report" \
        --content "Revenue grew 23%. We signed 3 enterprise contracts and
                   cut costs by 12%. Next quarter targets European launch."

Outputs land in:
    results/{YYYYMMDD_HHMMSS}_{stem}/
        {stem}_layout.json
        {stem}_content_mapping.json
        {stem}_output.pptx
        {stem}_validation_report.json
"""

import argparse
import os
import sys
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()


def _ask(prompt: str, default: str = "") -> str:
    """Print a prompt and return stripped user input, falling back to default."""
    try:
        value = input(prompt).strip()
        return value if value else default
    except (EOFError, KeyboardInterrupt):
        print()
        sys.exit(0)


def _interactive_mode() -> argparse.Namespace:
    """Walk the user through inputs one by one when no CLI flags are given."""
    print("\nPPTX Layout Cloner — Interactive Mode")
    print("=" * 50)

    # PPTX path
    while True:
        pptx_path = _ask("\nPath to your .pptx file: ")
        if not pptx_path:
            print("  [!] Path is required.")
            continue
        if not os.path.exists(pptx_path):
            print(f"  [!] File not found: {pptx_path}")
            continue
        break

    # Topic
    topic = _ask("\nTopic (e.g. 'Page Object Model in Selenium'): ")

    # Content
    print("\nPaste your content / description below.")
    print("Explain the topic in your own words — a paragraph or two is enough.")
    print("Press Enter when done.\n")
    lines = []
    while True:
        try:
            line = input()
        except (EOFError, KeyboardInterrupt):
            break
        if line == "":
            break
        lines.append(line)
    content = "\n".join(lines).strip()

    print()

    ns = argparse.Namespace(
        input=pptx_path,
        topic=topic,
        content=content,
        results_dir="results",
        phase=None,
    )
    return ns


def main():
    parser = argparse.ArgumentParser(
        description="PPTX Layout Cloner — extract layout and regenerate content via LLM",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input",
                        default="",
                        help="Path to input .pptx file")
    parser.add_argument("--topic",
                        default="",
                        help="Topic / title for the new content")
    parser.add_argument("--content",
                        default="",
                        help="Demo text the LLM should map into the slide "
                             "placeholders. When provided the LLM acts as a "
                             "content mapper; when omitted it generates freely.")
    parser.add_argument("--results-dir",
                        default="results",
                        help="Root folder for output runs (default: results/)")
    parser.add_argument("--phase",
                        type=int,
                        choices=[1, 2, 3, 4],
                        default=None,
                        help="Run a single phase only (omit to run all phases)")
    args = parser.parse_args()

    # ── If no --input given, drop into interactive mode ────────────────────
    if not args.input:
        args = _interactive_mode()

    # ── Validate input ─────────────────────────────────────────────────────
    if not os.path.exists(args.input):
        print(f"[ERROR] Input file not found: {args.input}")
        sys.exit(1)

    # ── Build timestamped run folder ───────────────────────────────────────
    stem      = Path(args.input).stem          # e.g. "Dickinson_Sample_Slides"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir   = os.path.join(args.results_dir, f"{timestamp}_{stem}")
    os.makedirs(run_dir, exist_ok=True)

    # All output paths live inside this run folder
    layout_path      = os.path.join(run_dir, f"{stem}_layout.json")
    content_map_path = os.path.join(run_dir, f"{stem}_content_mapping.json")
    output_pptx_path = os.path.join(run_dir, f"{stem}_output.pptx")
    validation_path  = os.path.join(run_dir, f"{stem}_validation_report.json")

    print(f"\nPPTX Layout Cloner")
    print(f"  Input      : {args.input}")
    print(f"  Topic      : {args.topic or '(not set)'}")
    print(f"  Mode       : {'intent-guided' if args.content.strip() else 'free-generator'}")
    print(f"  Run folder : {run_dir}")

    run_all = args.phase is None

    # ── Phase 1 — Extract ──────────────────────────────────────────────────
    if run_all or args.phase == 1:
        print("\n" + "=" * 60)
        print("PHASE 1 — Extracting layout from PPTX")
        print("=" * 60)
        from src.extractor import extract_layout
        extract_layout(args.input, layout_path)

    # ── Phase 2 — Generate ─────────────────────────────────────────────────
    if run_all or args.phase == 2:
        print("\n" + "=" * 60)
        print("PHASE 2 — Generating content via LLM")
        print("=" * 60)
        from src.generator import generate_content
        generate_content(layout_path, content_map_path, args.topic, args.content)

    # ── Phase 3 — Reconstruct ──────────────────────────────────────────────
    if run_all or args.phase == 3:
        print("\n" + "=" * 60)
        print("PHASE 3 — Reconstructing PPTX")
        print("=" * 60)
        from src.reconstructor import reconstruct
        reconstruct(args.input, content_map_path, output_pptx_path)

    # ── Phase 4 — Validate ─────────────────────────────────────────────────
    if run_all or args.phase == 4:
        print("\n" + "=" * 60)
        print("PHASE 4 — Validating output")
        print("=" * 60)
        from src.validator import validate
        report = validate(args.input, output_pptx_path, validation_path)

        # Surface any failures clearly
        failed = report["summary"]["failed"]
        if failed:
            print(f"\n  [!] {failed} validation check(s) FAILED — review report:")
            print(f"      {validation_path}")

    # ── Final summary ──────────────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("DONE — All outputs saved to run folder:")
    print(f"  {os.path.abspath(run_dir)}/")
    files = [
        (layout_path,      "layout.json          (Phase 1 — full structure)"),
        (content_map_path, "content_mapping.json  (Phase 2 — LLM content)"),
        (output_pptx_path, "output.pptx           (Phase 3 — editable PPTX)"),
        (validation_path,  "validation_report.json(Phase 4 — proof document)"),
    ]
    for path, label in files:
        status = "OK " if os.path.exists(path) else "---"
        print(f"  [{status}]  {label}")
    print()


if __name__ == "__main__":
    main()

# PPTX Layout Cloner

A four-phase Python pipeline that takes any unseen PowerPoint file, extracts its full layout and structure, generates entirely new content via LLM, and produces a new fully editable `.pptx` with the same visual skeleton — without touching the original.

---

## How It Works

```
Input PPTX
    │
    ▼
[ Phase 1 ]  extractor.py      →  layout.json
[ Phase 2 ]  generator.py      →  content_mapping.json
[ Phase 3 ]  reconstructor.py  →  output.pptx
[ Phase 4 ]  validator.py      →  validation_report.json
    │
    ▼
results/{timestamp}_{name}/   ← all four files per run (auto-created)
```

1. **Extract** — Parses every shape, table, position, font, and placeholder type into structured JSON
2. **Generate** — Two-pass LLM generation: Pass 1 plans a unique focus per slide, Pass 2 fills each slide using that focus. Supports intent-guided mode (user provides direction) and free-generator mode (LLM generates freely)
3. **Reconstruct** — Clones the original PPTX and surgically replaces only `<a:t>` text nodes and table cell text, leaving all formatting XML untouched
4. **Validate** — Compares original vs output across slide count, canvas size, bounding boxes, font size, font family, and font color

---

## Project Structure

```
pptx-layout-cloner/
├── src/
│   ├── extractor.py          # Phase 1 — layout + table extraction
│   ├── generator.py          # Phase 2 — two-pass Groq LLM content generation
│   ├── reconstructor.py      # Phase 3 — XML-level PPTX + table reconstruction
│   └── validator.py          # Phase 4 — output validation
├── input/                    # Place your input .pptx here (gitignored)
├── results/                  # One timestamped subfolder per run (gitignored, auto-created)
├── tests/                    # Phase-level test scripts
├── main.py                   # Orchestrator — runs all 4 phases
├── requirements.txt
├── .env.example              # Copy this to .env and add your API key
└── README.md
```

---

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/Abd756/pptx-layout-cloner.git
cd pptx-layout-cloner
```

### 2. Create and activate a virtual environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Mac / Linux
source venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Add your Groq API key

```bash
# Copy the example file
cp .env.example .env
```

Then open `.env` and replace the placeholder with your actual key:

```
GROQ_API_KEY=gsk_your_key_here
```

Get a free key at [console.groq.com](https://console.groq.com).

### 5. Add your input PPTX

Create an `input/` folder and drop any `.pptx` file inside it:

```bash
mkdir input
# copy your .pptx into input/
```

---

## Usage

### Interactive mode (recommended)

Run with no arguments — the tool walks you through everything:

```bash
python main.py
```

You will be prompted for:
1. Path to your `.pptx` file
2. Topic (e.g. `AI Engineer Roadmap`)
3. Content / description — paste your intent and press **Enter once** when done

### Command-line mode

**Free generator** — LLM creates content from topic alone:

```bash
python main.py --input input/deck.pptx --topic "Renewable Energy Transition"
```

**Intent-guided** — LLM uses your description to shape the content:

```bash
python main.py --input input/deck.pptx --topic "AI Engineer" \
  --content "Cover the roadmap to become an AI engineer, required skills, tools, and career scope."
```

### Run a single phase

```bash
python main.py --input input/deck.pptx --topic "AI" --phase 1
```

### All arguments

| Argument | Default | Description |
|---|---|---|
| `--input` | *(interactive)* | Path to the input `.pptx` file |
| `--topic` | `""` | Topic / title for the new content |
| `--content` | `""` | Your intent or description (enables intent-guided mode) |
| `--results-dir` | `results/` | Root folder for output runs |
| `--phase` | all | Run a single phase: `1`, `2`, `3`, or `4` |

---

## Output Files

Every run creates a timestamped subfolder inside `results/` so multiple runs never overwrite each other. The folder is auto-created — it is not committed to the repo.

```
results/
└── 20260624_150339_your_deck/
    ├── your_deck_layout.json
    ├── your_deck_content_mapping.json
    ├── your_deck_output.pptx
    └── your_deck_validation_report.json
```

| File | Description |
|---|---|
| `layout.json` | Full structural extraction — shapes, tables, positions, fonts, colors, placeholder types |
| `content_mapping.json` | LLM-generated content mapped to each shape ID including plan, mode, and table grids |
| `output.pptx` | Final fully editable PPTX with new content, original layout preserved |
| `validation_report.json` | Pass/fail/warn checks comparing original vs output across layout and typography |

---

## How Content Generation Works

### Two-pass generation

**Pass 1 — Planning:** A single LLM call receives all slide summaries and returns a unique focus statement per slide (e.g. `"AI tool adoption by industry"` for slide 3). This prevents repetition across slides.

**Pass 2 — Per-slide generation:** Each slide is sent individually with its assigned focus. The LLM fills every placeholder with content that fits that specific angle.

### Word caps (prevents overflow)

| Slide density | Max words per bullet |
|---|---|
| 9+ bullets | 5 words |
| 6–8 bullets | 7 words |
| 2–5 bullets | 10 words |
| Single-line body | 12 words |
| Title | 7 words |

These caps are enforced programmatically in `_validate_and_fix` regardless of LLM output.

### Table support

Tables are extracted as a `rows × cols` grid, sent to the LLM for content replacement, and written back cell by cell using the same XML surgery as regular shapes. Header rows are kept as short labels (max 6 words per cell).

---

## Approach

### XML surgery instead of recreating shapes

`python-pptx`'s high-level API creates shapes with default styling — fonts, colors, shadows, and fills are lost. Instead we deep-copy the original PPTX XML tree and replace only `<a:t>` (text) nodes and table cell content, leaving every `<a:rPr>`, `<a:pPr>`, and shape transform completely untouched.

### Style inheritance

PowerPoint uses a 4-level chain: `Shape XML → Slide Layout → Slide Master → Theme`. The extractor resolves this chain so the layout JSON contains concrete values. During reconstruction, since original XML nodes are preserved, inheritance continues to work automatically.

---

## Limitations

- **Charts** — chart data is stored as embedded Excel XML inside the PPTX ZIP. Replacing it requires modifying the embedded workbook; planned for a future version. Charts are detected and skipped gracefully
- **Images and icons** — preserved exactly as-is; their visual content is not altered
- **Right-to-left text** — structural layout is preserved but LLM output defaults to LTR

---

## Extending Into Production

- **Chart data replacement** — parse the embedded `xl/worksheets/sheet1.xml` inside the chart's `.zip` entry and update cell values
- **Web API** — wrap `main.py` in a FastAPI endpoint: POST a PPTX + content, GET back a download URL
- **Frontend UI** — upload form where users paste content and download the result with no terminal required
- **Template library** — store extracted `layout.json` files as reusable templates keyed by slide type
- **CI validation gate** — run the validator as a GitHub Actions step to catch layout regressions on every code change

---

## Tech Stack

| Component | Technology |
|---|---|
| Language | Python 3.10+ |
| PPTX I/O | `python-pptx 0.6.23` |
| XML manipulation | `lxml 5.2.2` |
| LLM | Groq API — `llama-3.1-8b-instant` |
| LLM client | `groq 1.5.0` |
| Config | `python-dotenv 1.0.1` |

---

## Validation Results

Run against two real-world PPTX files:

**Dickinson Sample Slides** (9 slides, topic: `Renewable Energy Transition`):
```
Total checks : 125
Passed       : 125
Warnings     : 0
Failed       : 0
```

**Week 12-13 Page Object Model / POM** (19 slides, topic: `AI Coding Agents`):
```
Total checks : 243
Passed       : 238
Warnings     : 5
Failed       : 0
```

Warnings are inherited font values (`None` at shape level) — these are resolved through PowerPoint's theme inheritance chain and display correctly in the output.

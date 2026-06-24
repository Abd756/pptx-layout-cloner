# PPTX Layout Cloner

A four-phase Python pipeline that takes any unseen PowerPoint file, extracts its full layout and structure, and generates a new `.pptx` with the same visual skeleton but entirely new content — without touching the original and keeping the output fully editable.

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
results/{timestamp}_{name}/   ← all four files land here per run
```

1. **Extract** — Parses every shape, position, font, color, and placeholder type from the input PPTX into a structured JSON
2. **Generate** — Sends the layout to Groq LLM which maps new content into each placeholder, either freely from a topic or by fitting user-provided demo text
3. **Reconstruct** — Clones the original PPTX and surgically replaces only the `<a:t>` text nodes while leaving all formatting XML untouched
4. **Validate** — Compares original vs output across 125 checks: slide count, canvas size, bounding boxes, font size, font family, font color, and text replacement

---

## Project Structure

```
pptx-layout-cloner/
├── src/
│   ├── extractor.py          # Phase 1 — layout extraction
│   ├── generator.py          # Phase 2 — Groq LLM content generation
│   ├── reconstructor.py      # Phase 3 — XML-level PPTX reconstruction
│   └── validator.py          # Phase 4 — output validation
├── results/                  # One timestamped subfolder per run (client deliverables)
│   └── 20260624_150339_Dickinson_Sample_Slides/
│       ├── Dickinson_Sample_Slides_layout.json
│       ├── Dickinson_Sample_Slides_content_mapping.json
│       ├── Dickinson_Sample_Slides_output.pptx
│       └── Dickinson_Sample_Slides_validation_report.json
├── input/                    # Place your input .pptx here (gitignored)
├── main.py                   # Orchestrator — runs all 4 phases
├── requirements.txt
├── .env                      # Your Groq API key (gitignored)
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

Create a `.env` file in the root:

```
GROQ_API_KEY=gsk_your_key_here
```

Get a free key at [console.groq.com](https://console.groq.com).

### 5. Add your input PPTX

Create an `input/` folder and drop any `.pptx` file inside it:

```bash
mkdir input
# then copy your .pptx into input/
```

---

## Usage

### Mode A — LLM generates content freely from a topic

```bash
python main.py --input input/your_slide.pptx --topic "Renewable Energy Transition"
```

### Mode B — LLM maps your own content into the slide structure

```bash
python main.py --input input/your_slide.pptx --topic "Q3 Financial Report" --content "Revenue grew 23% to $4.2M. We signed 3 enterprise contracts and cut operational costs by 12%. Next quarter targets European product launch."
```

In mapper mode the LLM extracts and fits your text into the correct placeholders — it does not invent new information.

### Run a single phase only

```bash
python main.py --input input/your_slide.pptx --topic "AI in Healthcare" --phase 1
```

### All arguments

| Argument | Required | Default | Description |
|---|---|---|---|
| `--input` | Yes | — | Path to the input `.pptx` file |
| `--topic` | No | `""` | Topic / title for the new content |
| `--content` | No | `""` | Demo text to map into placeholders (enables mapper mode) |
| `--results-dir` | No | `results/` | Root folder for output runs |
| `--phase` | No | all | Run a single phase: `1`, `2`, `3`, or `4` |

---

## Output Files

Every run creates a timestamped subfolder inside `results/` so multiple runs never overwrite each other.

| File | Description |
|---|---|
| `{name}_layout.json` | Full structural extraction — shapes, positions (EMU + pt + %), fonts, colors, placeholder types per slide |
| `{name}_content_mapping.json` | LLM-generated content mapped to each shape ID, including mode and source content used |
| `{name}_output.pptx` | Final fully editable PPTX with new content, original layout preserved pixel-perfect |
| `{name}_validation_report.json` | 125 pass/fail/warn checks comparing original vs output across layout, typography, and content |

---

## Approach

### XML surgery instead of recreating shapes

`python-pptx`'s high-level API (e.g. `add_textbox()`) creates shapes with default styling — fonts, colors, shadows, and fills are lost. Instead, we deep-copy the original PPTX's XML tree and replace only `<a:t>` (text) nodes, leaving every `<a:rPr>` (run properties), `<a:pPr>` (paragraph properties), and shape transform completely untouched.

### Style inheritance resolution

PowerPoint uses a 4-level inheritance chain: `Shape XML → Slide Layout → Slide Master → Theme`. During extraction, we resolve this chain explicitly so the layout JSON always contains concrete values. During reconstruction, since the original XML nodes are preserved, inheritance continues to work automatically.

### Two LLM modes

- **Generator mode** — LLM receives placeholder type, paragraph count, font size, and topic, then freely generates fitting content
- **Mapper mode** — LLM receives the user's actual demo text and is instructed to extract and distribute it into each placeholder without inventing new information

Both modes enforce exact paragraph counts so bullet lists are never padded or truncated.

---

## Limitations

- **Charts** — chart data is stored as embedded Excel XML inside the PPTX ZIP. Replacing it requires modifying the embedded workbook, which is out of scope for this version. Charts are detected and skipped gracefully
- **Tables** — detected and skipped in the current version; table cell text replacement is a planned extension
- **Images and icons** — preserved exactly as-is; their visual content is not altered
- **Right-to-left text** — structural layout is preserved but LLM output defaults to LTR languages

---

## Extending Into Production

- **Table support** — table cells expose a text frame identical to regular shapes; the same XML surgery pattern applies
- **Chart data replacement** — parse the embedded `xl/worksheets/sheet1.xml` inside the chart's `.zip` entry and update cell values programmatically
- **Web API** — wrap `main.py` in a FastAPI endpoint: POST a PPTX + content, GET back a download URL for the output PPTX
- **Frontend UI** — a simple upload form where users paste their demo content and download the result, with no terminal required
- **Template library** — store extracted `layout.json` files as reusable templates keyed by slide type for instant re-use without re-extraction
- **CI validation gate** — run the validator as a GitHub Actions step to catch layout regressions automatically on every code change

---

## Tech Stack

| Component | Technology |
|---|---|
| Language | Python 3.10+ |
| PPTX I/O | `python-pptx 0.6.23` |
| XML manipulation | `lxml 5.2.2` |
| LLM | Groq API — `llama-3.3-70b-versatile` |
| LLM client | `groq 1.5.0` |
| Config | `python-dotenv 1.0.1` |

---

## Validation Results (Dickinson Sample Slides)

Run against a 9-slide real-world PPTX with topic `"Renewable Energy Transition"`:

```
Total checks : 125
Passed       : 125
Warnings     : 0
Failed       : 0
```

All bounding boxes, font sizes, font families, font colors, and slide dimensions matched exactly between original and output.

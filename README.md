# ppt-pipeline

**PDF thesis → defense PPT auto-generation pipeline**

Drop in a thesis PDF, get back a structured, human-editable defense presentation — the AI does 80% of the grunt work, you make the final 20% of decisions.

## Overview

```
thesis.pdf  →  [parse]  →  [LLM refine?]  →  ppt_plan.json  →  [render]  →  答辩PPT.pptx
                               (optional)        (you review)                  (final output)
```

- **Phase 1**: Extract the table of contents, text, images, and tables from the PDF (deterministic, no AI).
- **Phase 1.5** (optional): Use Claude or OpenAI to refine each chapter into bullet points.
- **Phase 2**: Fill a PowerPoint template with the structured data — text boxes, images, and native tables.

The key design choice: **ppt_plan.json is the human-AI interface**. After generation, you review and tweak the JSON before rendering. No attempt at 100% unattended perfection — that never works for a defense presentation.

## Quick Start

```bash
# 1. Install
pip install -r requirements.txt

# 2. Generate default template (first time only)
python src/create_template.py

# 3. Drop your thesis PDF in input/ and run
python run.py input/thesis.pdf --author "Your Name" --date "2025-06"

# 4. Review output/ppt_plan.json — delete slides, tweak bullets, add images
# 5. Render the final PPTX
python run.py input/thesis.pdf --skip-parse
```

### With LLM bullet refinement

```bash
# Claude (set ANTHROPIC_API_KEY env var)
python run.py input/thesis.pdf --llm claude

# OpenAI / compatible endpoint
export OPENAI_API_KEY="sk-..."
export OPENAI_BASE_URL="https://api.openai.com/v1"  # optional
python run.py input/thesis.pdf --llm openai --model gpt-4o
```

## Pipeline Stages

| Command | What it does |
|---------|-------------|
| `python run.py thesis.pdf --stop-at parse` | Only extract TOC/text/images/tables → `output/parsed_data.json` |
| `python run.py thesis.pdf --stop-at plan` | Parse + assemble → `output/ppt_plan.json` (review this!) |
| `python run.py thesis.pdf --skip-parse` | Skip PDF parsing, go straight from cached JSON to PPTX |

## File Structure

```
ppt-pipeline/
├── input/                  # Drop your thesis.pdf here
├── output/
│   ├── figures/            # Images extracted from the PDF
│   ├── parsed_data.json    # Raw parse cache
│   ├── ppt_plan.json       # ★ Human-editable intermediate format
│   └── 答辩PPT.pptx        # Final rendered presentation
├── src/
│   ├── parse_pdf.py        # Phase 1: PDF → chapter tree + images + tables
│   ├── llm_summarize.py    # Phase 1 (optional): LLM bullet refinement
│   ├── assemble_plan.py    # Phase 1: Combine parsed data → ppt_plan.json
│   ├── render_pptx.py      # Phase 2: ppt_plan.json → template → .pptx
│   └── create_template.py  # Generate a default template.pptx
├── run.py                  # Main entry point (orchestrates the full pipeline)
├── template.pptx           # Your custom PowerPoint template
├── requirements.txt
├── pyproject.toml
└── LICENSE
```

## ppt_plan.json Format

This is the file you edit by hand between generation and rendering. See DESIGN.md for the full schema; here are the key slide layouts:

| Layout | Use |
|--------|-----|
| `title` | Cover slide (title, subtitle, author, date) |
| `toc` | Table of contents |
| `section_title` | Chapter divider page |
| `bullets` | Content slide with bullet points |
| `figure` | Full-slide image with caption |
| `table` | Data comparison table |
| `end` | Closing / thank-you slide |

## Template Customization

The default template (`create_template.py`) produces a bare-bones layout. For a polished presentation:

1. Open `template.pptx` in PowerPoint
2. Go to **View → Slide Master**
3. Customize each layout's fonts, colors, backgrounds, and logos
4. Keep the layout **names** unchanged (the pipeline matches by name)

## Design Decisions

- **Three-line table awareness**: Academic papers use sparse table borders; the extractor uses text-based column inference instead of grid detection.
- **Vector figure handling**: pdfplumber detects both bitmap and vector artwork regions; PyMuPDF clip-renders them at 300 DPI.
- **Formulas are not auto-extracted**: Inline formulas are skipped. Display formulas get a `formula_image` placeholder in JSON for manual screenshot.
- **No forced full automation**: The pipeline does the tedious 80% (structure, extraction, initial layout) and leaves creative decisions to you.

## Dependencies

| Package | Role |
|---------|------|
| `PyMuPDF` | PDF parsing, TOC extraction, image rendering |
| `pdfplumber` | Text & table extraction, image region detection |
| `python-pptx` | PowerPoint file generation |
| `anthropic` (optional) | Claude API for bullet refinement |
| `openai` (optional) | OpenAI API for bullet refinement |

## License

MIT — see [LICENSE](LICENSE).

#!/usr/bin/env python3
"""ppt-pipeline: PDF thesis → defense PPT auto-generation pipeline.

Usage:
    python run.py input/thesis.pdf                          # full pipeline (no LLM)
    python run.py input/thesis.pdf --llm claude             # with Claude summarization
    python run.py input/thesis.pdf --llm openai --model gpt-4o
    python run.py input/thesis.pdf --skip-parse             # render only (JSON ready)
    python run.py input/thesis.pdf --stop-at plan           # parse + assemble only

Steps:
    1. parse   — extract TOC, text, images, tables from PDF
    2. llm     — (optional) refine chapter text into bullets via LLM
    3. assemble — build ppt_plan.json
    4. render   — fill template.pptx → output .pptx
"""

import argparse
import json
import sys
from pathlib import Path

# Ensure src/ is on sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from parse_pdf import parse
from assemble_plan import assemble_plan
from render_pptx import render


def main() -> None:
    parser = argparse.ArgumentParser(
        description="PDF thesis → defense PPT auto-generation pipeline",
    )
    parser.add_argument(
        "pdf", nargs="?", default="input/thesis.pdf",
        help="Path to the thesis PDF (default: input/thesis.pdf)",
    )
    parser.add_argument(
        "--llm", choices=["claude", "openai"], default=None,
        help="Use an LLM to refine chapter bullets (requires API key env var)",
    )
    parser.add_argument(
        "--model", default=None,
        help="LLM model name (default: claude-sonnet-4-6 / gpt-4o)",
    )
    parser.add_argument(
        "--author", default="",
        help="Thesis author name for the title slide",
    )
    parser.add_argument(
        "--date", default="",
        help="Defense date for the title slide",
    )
    parser.add_argument(
        "--template", default="template.pptx",
        help="Path to the .pptx template (default: template.pptx)",
    )
    parser.add_argument(
        "--output", default="output/答辩PPT.pptx",
        help="Output .pptx path (default: output/答辩PPT.pptx)",
    )
    parser.add_argument(
        "--skip-parse", action="store_true",
        help="Skip PDF parsing and use existing output/parsed_data.json",
    )
    parser.add_argument(
        "--stop-at", choices=["parse", "plan"], default=None,
        help="Stop the pipeline early (debug / manual workflow)",
    )

    args = parser.parse_args()

    # ------------------------------------------------------------------
    # Step 1: Parse PDF
    # ------------------------------------------------------------------
    if args.skip_parse:
        parsed_path = Path("output/parsed_data.json")
        if not parsed_path.exists():
            sys.exit("Error: --skip-parse specified but output/parsed_data.json not found")
        parsed_data = json.loads(parsed_path.read_text(encoding="utf-8"))
        print("Skipped PDF parsing (using cached parsed_data.json)")
    else:
        pdf_path = Path(args.pdf)
        if not pdf_path.exists():
            sys.exit(f"Error: PDF not found: {pdf_path}")
        print(f"Parsing PDF: {pdf_path}")
        parsed = parse(pdf_path)
        parsed_data = parsed.to_dict()

        # Cache for later steps
        Path("output").mkdir(parents=True, exist_ok=True)
        Path("output/parsed_data.json").write_text(
            json.dumps(parsed_data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"  → {parsed_data['total_pages']} pages, "
              f"{len(parsed_data['chapters'])} chapters, "
              f"{len(parsed_data['images'])} images, "
              f"{len(parsed_data['tables'])} tables")

    if args.stop_at == "parse":
        print("Stopped after parsing.")
        return

    # ------------------------------------------------------------------
    # Step 2: LLM summarization (optional)
    # ------------------------------------------------------------------
    llm_bullets = {}
    if args.llm:
        from llm_summarize import summarize_chapter
        from parse_pdf import extract_all_chapter_texts

        # Reconstruct chapter tree from parsed data
        from parse_pdf import Section

        def dict_to_section(d: dict) -> Section:
            return Section(
                level=d["level"],
                title=d["title"],
                page_start=d["page_start"],
                page_end=d["page_end"],
                sections=[dict_to_section(s) for s in d.get("sections", [])],
            )

        chapters = [dict_to_section(ch) for ch in parsed_data["chapters"]]
        chapter_texts = extract_all_chapter_texts(args.pdf, chapters)

        model_kwargs = {}
        if args.model:
            model_kwargs["model"] = args.model

        for label, text in chapter_texts.items():
            print(f"LLM summarizing: {label} ({len(text)} chars)")
            try:
                bullets = summarize_chapter(
                    label, text, provider=args.llm, **model_kwargs,
                )
                llm_bullets[label] = bullets
            except Exception as e:
                print(f"  [warn] LLM call failed for '{label}': {e}")
                continue

    # ------------------------------------------------------------------
    # Step 3: Assemble ppt_plan.json
    # ------------------------------------------------------------------
    print("Assembling ppt_plan.json ...")
    plan = assemble_plan(
        parsed_data,
        author=args.author,
        date=args.date,
        llm_bullets=llm_bullets or None,
    )

    plan_path = Path("output/ppt_plan.json")
    plan_path.parent.mkdir(parents=True, exist_ok=True)
    plan_path.write_text(
        json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8",
    )
    print(f"  → {plan['meta']['total_slides']} slides")

    if args.stop_at == "plan":
        print(f"Stopped. Review & edit {plan_path}, then run:")
        print(f"  python run.py {args.pdf} --skip-parse")
        return

    # ------------------------------------------------------------------
    # Step 4: Render PPTX
    # ------------------------------------------------------------------
    print("Rendering PPTX ...")
    result = render(plan_path, args.template, args.output)
    print(f"  → {result}")
    print("Done!")


if __name__ == "__main__":
    main()

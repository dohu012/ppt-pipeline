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
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

# Load .env file before anything else
from dotenv import load_dotenv
load_dotenv()

# Ensure src/ is on sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from parse_pdf import parse
from assemble_plan import assemble_plan
from render_pptx import render


def _find_visuals_for_section(
    ch_num: str,
    figure_entries: list[dict],
    table_entries: list[dict],
) -> tuple[list[dict], list[dict]]:
    """Return (figures, tables) whose number prefix matches the chapter number."""
    figs = []
    for f in figure_entries:
        num = f["number"]
        if num.startswith(f"图{ch_num}-") or num.startswith(f"图{ch_num}–"):
            figs.append({
                "number": num,
                "caption": f["caption"],
                "screenshot": f.get("screenshot", ""),
            })
    tabs = []
    for t in table_entries:
        num = t["number"]
        if num.startswith(f"表{ch_num}-") or num.startswith(f"表{ch_num}–"):
            tabs.append({
                "number": num,
                "caption": t["caption"],
                "screenshot": t.get("screenshot", ""),
            })
    return figs, tabs


def _extract_chapter_number(title: str) -> str:
    """Extract chapter number from a title like '第一章 绪论' → '一'."""
    import re
    m = re.search(r"第([一二三四五六七八九十\d]+)章", title)
    if not m:
        return ""
    num = m.group(1)
    # Convert Chinese numeral to digit
    cn_map = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
              "六": "6", "七": "7", "八": "8", "九": "9", "十": "10"}
    return cn_map.get(num, num)


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
              f"{len(parsed_data.get('figure_entries', []))} figures, "
              f"{len(parsed_data.get('table_entries', []))} tables")

    if args.stop_at == "parse":
        print("Stopped after parsing.")
        return

    # ------------------------------------------------------------------
    # Step 2: LLM summarization (optional)
    # ------------------------------------------------------------------
    llm_results: dict[str, dict] = {}  # chapter_label → {slides, figures, tables}
    if args.llm:
        from llm_summarize import summarize_chapter_multi

        model_kwargs = {}
        if args.model:
            model_kwargs["model"] = args.model
        elif os.environ.get("PPT_MODEL"):
            model_kwargs["model"] = os.environ["PPT_MODEL"]

        figure_entries = parsed_data.get("figure_entries", [])
        table_entries = parsed_data.get("table_entries", [])
        chapter_texts = parsed_data.get("chapter_texts", {})

        # Only process main content chapters (第X章), skip front matter
        _CHAPTER_PATTERN = re.compile(r"第[一二三四五六七八九十\d]+章")

        tasks: list[tuple[str, str, list[dict], list[dict]]] = []
        for label, text in chapter_texts.items():
            if not text.strip():
                continue
            # Only include main chapter entries (not subsections, not front matter)
            if not _CHAPTER_PATTERN.search(label):
                continue

            ch_num = _extract_chapter_number(label)
            sec_figures, sec_tables = _find_visuals_for_section(
                ch_num, figure_entries, table_entries
            )
            tasks.append((label, text, sec_figures, sec_tables))

        total = len(tasks)
        if total:
            print(f"LLM: summarizing {total} chapters (max_workers=5)...")
            with ThreadPoolExecutor(max_workers=5) as executor:
                future_map = {
                    executor.submit(
                        summarize_chapter_multi, label, text,
                        figures=sec_figures or None,
                        tables=sec_tables or None,
                        provider=args.llm,
                        **model_kwargs,
                    ): label
                    for label, text, sec_figures, sec_tables in tasks
                }

                for done, future in enumerate(as_completed(future_map), 1):
                    label = future_map[future]
                    try:
                        llm_results[label] = future.result()
                        n_slides = len(llm_results[label].get("slides", []))
                        print(f"  [{done}/{total}] {label} ({n_slides} slides)")
                    except Exception as e:
                        print(f"  [{done}/{total}] ✗ {label}: {e}")
                        continue
    # Step 3: Assemble ppt_plan.json
    # ------------------------------------------------------------------
    print("Assembling ppt_plan.json ...")
    plan = assemble_plan(
        parsed_data,
        author=args.author,
        date=args.date,
        llm_results=llm_results or None,
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

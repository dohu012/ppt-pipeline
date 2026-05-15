"""Phase 1 final step: assemble parsed data + LLM bullets into ppt_plan.json."""

import json
import re
from pathlib import Path
from typing import Any


def _slugify(text: str) -> str:
    """Turn a section title into a short identifier."""
    return re.sub(r"[^\w]+", "_", text.strip()).strip("_").lower()


def _extract_year(text: str) -> str:
    """Try to extract a 4-digit year from arbitrary text, fall back to today."""
    m = re.search(r"(20\d{2})", text)
    return m.group(1) if m else "2025"


def assemble_plan(
    parsed_data: dict[str, Any],
    *,
    author: str = "",
    date: str = "",
    llm_results: dict[str, dict[str, Any]] | None = None,
    fallback_texts: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Build a ppt_plan.json document from parsed PDF data.

    Args:
        parsed_data: Output of parse_pdf.parse() → .to_dict()
        author: Thesis author name.
        date: Defense date (freeform).
        llm_results: Optional {chapter_label: {"bullets": [...], "figures": [...], "tables": [...]}}.
        fallback_texts: Optional {chapter_title: raw_text} for rule-based bullets
                        when LLM is not used.

    Returns:
        The full ppt_plan.json structure as a dict.
    """
    llm_results = llm_results or {}
    fallback_texts = fallback_texts or {}

    slides: list[dict[str, Any]] = []
    slide_id = 0

    def next_id() -> str:
        nonlocal slide_id
        slide_id += 1
        return f"slide_{slide_id:03d}"

    title = parsed_data.get("title", "毕业论文")
    chapters = parsed_data.get("chapters", [])
    figure_entries = parsed_data.get("figure_entries", [])

    # ---- Title slide ----
    slides.append(
        {
            "id": next_id(),
            "layout": "title",
            "content": {
                "title": title,
                "subtitle": "硕士学位论文答辩",
                "author": author or "待填写",
                "date": date or _extract_year(title),
            },
        }
    )

    # ---- TOC slide (content chapters only) ----
    _CHAPTER_PATTERN = re.compile(r"第[一二三四五六七八九十\d]+章")
    toc_items = [ch["title"] for ch in chapters if _CHAPTER_PATTERN.search(ch["title"])]
    slides.append(
        {
            "id": next_id(),
            "layout": "toc",
            "content": {"title": "目录", "items": toc_items},
        }
    )

    # Collect LLM-decided figures and tables, deduplicated by number
    kept_figures: list[dict[str, Any]] = []
    kept_tables: list[dict[str, Any]] = []
    seen_figures: set[str] = set()
    seen_tables: set[str] = set()
    figure_screenshots: dict[str, str] = {
        f["number"]: f.get("screenshot", "") for f in figure_entries
    }

    # ---- Per-chapter slides (content chapters only, skip front matter) ----
    for ch in chapters:
        ch_title = ch["title"]

        # Skip front matter / references / appendix / acknowledgment
        if not _CHAPTER_PATTERN.search(ch_title):
            continue

        # Section-title slide
        subtitle = (
            ch["sections"][0]["title"] if ch.get("sections") else ""
        )
        slides.append(
            {
                "id": next_id(),
                "layout": "section_title",
                "chapter": ch_title,
                "content": {"title": ch_title, "subtitle": subtitle},
            }
        )

        # Use LLM multi-slide result if available
        result = llm_results.get(ch_title)

        if result and result.get("slides"):
            for slide_data in result["slides"]:
                slides.append(
                    {
                        "id": next_id(),
                        "layout": "bullets",
                        "chapter": ch_title,
                        "content": {
                            "title": slide_data.get("title", ch_title),
                            "bullets": slide_data.get("bullets", []),
                        },
                    }
                )
            # Accumulate figure/table decisions (deduplicated by number)
            for fig in result.get("figures", []):
                if fig.get("keep") and fig["number"] not in seen_figures:
                    seen_figures.add(fig["number"])
                    fig["screenshot"] = figure_screenshots.get(fig["number"], "")
                    kept_figures.append(fig)
            for tab in result.get("tables", []):
                if tab.get("keep") and tab["number"] not in seen_tables:
                    seen_tables.add(tab["number"])
                    kept_tables.append(tab)
        elif fallback_texts.get(ch_title):
            text = fallback_texts[ch_title]
            sentences = re.split(r"[。！？\n]", text)
            bullets = [
                {"bullet": s.strip(), "ref_page": ch["page_start"] + 1}
                for s in sentences
                if len(s.strip()) > 10
            ][:5]
            if bullets:
                slides.append(
                    {
                        "id": next_id(),
                        "layout": "bullets",
                        "chapter": ch_title,
                        "content": {
                            "title": ch_title,
                            "bullets": bullets,
                        },
                    }
                )

    # ---- LLM-approved figure slides ----
    for fig in kept_figures:
        slides.append(
            {
                "id": next_id(),
                "layout": "figure",
                "content": {
                    "title": fig.get("caption", fig.get("number", "")),
                    "image": fig.get("screenshot", ""),
                    "caption": fig.get("number", ""),
                },
            }
        )

    # ---- LLM-approved table slides ----
    for tab in kept_tables:
        slides.append(
            {
                "id": next_id(),
                "layout": "table",
                "content": {
                    "title": tab.get("caption", tab.get("number", "")),
                    "header": tab.get("header", []),
                    "rows": tab.get("rows", []),
                },
            }
        )

    # ---- End slide ----
    slides.append(
        {
            "id": next_id(),
            "layout": "end",
            "content": {"text": "感谢各位老师批评指正"},
        }
    )

    return {
        "meta": {
            "title": title,
            "author": author or "待填写",
            "total_slides": len(slides),
        },
        "slides": slides,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    parsed_path = sys.argv[1] if len(sys.argv) > 1 else "output/parsed_data.json"
    parsed = json.loads(Path(parsed_path).read_text(encoding="utf-8"))

    plan = assemble_plan(parsed)

    out_path = Path("output/ppt_plan.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Assembled {plan['meta']['total_slides']} slides → {out_path}")

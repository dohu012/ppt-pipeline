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
    llm_bullets: dict[str, list[dict[str, Any]]] | None = None,
    fallback_texts: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Build a ppt_plan.json document from parsed PDF data.

    Args:
        parsed_data: Output of parse_pdf.parse() → .to_dict()
        author: Thesis author name.
        date: Defense date (freeform).
        llm_bullets: Optional {chapter_title: [{"bullet": ..., "ref_page": ...}]}.
        fallback_texts: Optional {chapter_title: raw_text} for rule-based bullets
                        when LLM is not used.

    Returns:
        The full ppt_plan.json structure as a dict.
    """
    llm_bullets = llm_bullets or {}
    fallback_texts = fallback_texts or {}

    slides: list[dict[str, Any]] = []
    slide_id = 0

    def next_id() -> str:
        nonlocal slide_id
        slide_id += 1
        return f"slide_{slide_id:03d}"

    title = parsed_data.get("title", "毕业论文")
    chapters = parsed_data.get("chapters", [])
    images = parsed_data.get("images", [])
    tables = parsed_data.get("tables", [])

    # Index images & tables by page number
    images_by_page: dict[int, list[dict]] = {}
    for img in images:
        images_by_page.setdefault(img["page"], []).append(img)

    tables_by_page: dict[int, list[dict]] = {}
    for t in tables:
        tables_by_page.setdefault(t["page"], []).append(t)

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

    # ---- TOC slide ----
    toc_items = [ch["title"] for ch in chapters]
    slides.append(
        {
            "id": next_id(),
            "layout": "toc",
            "content": {"title": "目录", "items": toc_items},
        }
    )

    # ---- Per-chapter slides ----
    for ch in chapters:
        ch_title = ch["title"]
        ch_start = ch["page_start"]
        ch_end = ch["page_end"]

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

        # Bullet slides: one per (sub)section
        sections_to_process = ch.get("sections") if ch.get("sections") else [ch]
        for sec in sections_to_process:
            sec_title = sec["title"]
            lookup_key = f"{ch_title} / {sec_title}"
            bullets = llm_bullets.get(lookup_key) or llm_bullets.get(sec_title)

            if not bullets and fallback_texts.get(lookup_key):
                # Rule-based fallback: split text into sentences, take first N
                text = fallback_texts[lookup_key]
                sentences = re.split(r"[。！？\n]", text)
                bullets = [
                    {"bullet": s.strip(), "ref_page": sec["page_start"] + 1}
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
                            "title": sec_title,
                            "bullets": bullets,
                        },
                    }
                )

        # Figure slides for this chapter's page range
        for pg in range(ch_start, ch_end + 1):
            for img in images_by_page.get(pg, []):
                slides.append(
                    {
                        "id": next_id(),
                        "layout": "figure",
                        "chapter": ch_title,
                        "content": {
                            "title": f"{ch_title} — 图表 (第{pg + 1}页)",
                            "image": img["filename"],
                            "caption": "",
                        },
                    }
                )

        # Table slides
        for pg in range(ch_start, ch_end + 1):
            for tbl in tables_by_page.get(pg, []):
                slides.append(
                    {
                        "id": next_id(),
                        "layout": "table",
                        "chapter": ch_title,
                        "content": {
                            "title": f"{ch_title} — 表格 (第{pg + 1}页)",
                            "header": tbl["header"],
                            "rows": tbl["rows"],
                        },
                    }
                )

    # ---- Conclusion slide (from last chapter bullets if present) ----
    # Already handled by per-chapter processing

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

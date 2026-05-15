"""Phase 1: PDF parsing — extract TOC, chapter text, and full-page screenshots
of figure / table pages for downstream LLM processing.

Design principle:
  parse_pdf only handles *deterministic* PDF operations:
    1. TOC → chapter tree with page ranges
    2. Figure / table index → which pages have figures and tables
    3. Full-page screenshots of those pages (NO bbox cropping)
    4. Chapter text extraction

  Everything else — deciding which figures / tables are worth including,
  precise cropping, and table-data extraction — is deferred to the
  vision LLM in llm_summarize.py.
"""

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import fitz  # PyMuPDF
import pdfplumber


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Section:
    level: int
    title: str
    page_start: int  # 0-indexed (PDF page)
    page_end: int    # 0-indexed, inclusive
    sections: list["Section"] = field(default_factory=list)


@dataclass
class FigureEntry:
    """One entry parsed from the List of Figures."""
    number: str       # e.g. "图3-2"
    caption: str
    doc_page: int     # page number as printed in the thesis
    screenshot: str = ""  # path to full-page screenshot


@dataclass
class TableEntry:
    """One entry parsed from the List of Tables."""
    number: str       # e.g. "表4-1"
    caption: str
    doc_page: int
    screenshot: str = ""  # path to full-page screenshot


@dataclass
class ParsedPDF:
    title: str = ""
    total_pages: int = 0
    chapters: list[Section] = field(default_factory=list)
    figure_entries: list[FigureEntry] = field(default_factory=list)
    table_entries: list[TableEntry] = field(default_factory=list)
    # Map: chapter_title → full text
    chapter_texts: dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> dict:
        def section_to_dict(s: Section) -> dict:
            return {
                "level": s.level,
                "title": s.title,
                "page_start": s.page_start,
                "page_end": s.page_end,
                "sections": [section_to_dict(c) for c in s.sections],
            }

        return {
            "title": self.title,
            "total_pages": self.total_pages,
            "chapters": [section_to_dict(c) for c in self.chapters],
            "figure_entries": [
                {
                    "number": f.number,
                    "caption": f.caption,
                    "doc_page": f.doc_page,
                    "screenshot": f.screenshot,
                }
                for f in self.figure_entries
            ],
            "table_entries": [
                {
                    "number": t.number,
                    "caption": t.caption,
                    "doc_page": t.doc_page,
                    "screenshot": t.screenshot,
                }
                for t in self.table_entries
            ],
            "chapter_texts": self.chapter_texts,
        }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def sanitize_title(raw: str) -> str:
    raw = raw.strip()
    raw = re.sub(r"\s+", " ", raw)
    return raw


def _safe_filename(text: str) -> str:
    text = text.replace("/", "_").replace("\\", "_")
    text = re.sub(r'[<>:"|?*]', "", text)
    text = re.sub(r"\s+", "_", text)
    return text.strip("_")[:80]


# ---------------------------------------------------------------------------
# TOC & chapter tree
# ---------------------------------------------------------------------------

def extract_toc(pdf_path: str | Path) -> list[tuple[int, str, int]]:
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=True)
    doc.close()

    cleaned: list[tuple[int, str, int]] = []
    for level, title, page in toc:
        title = sanitize_title(title)
        if not title:
            continue
        cleaned.append((level, title, page))
    return cleaned


def build_section_tree(
    toc: list[tuple[int, str, int]], total_pages: int
) -> list[Section]:
    if not toc:
        return []

    entries = [(level, title, page - 1) for level, title, page in toc]

    def build(level: int, idx: int) -> tuple[list[Section], int]:
        siblings: list[Section] = []
        while idx < len(entries):
            lvl, title, start_page = entries[idx]
            if lvl < level:
                break
            if lvl == level:
                children, idx = build(level + 1, idx + 1)
                end_page = entries[idx][2] - 1 if idx < len(entries) else total_pages - 1
                siblings.append(
                    Section(
                        level=level, title=title,
                        page_start=start_page,
                        page_end=max(start_page, end_page),
                        sections=children,
                    )
                )
            else:
                children, idx = build(level + 1, idx)
                if siblings:
                    siblings[-1].sections = children
        return siblings, idx

    root_level = entries[0][0]
    tree, _ = build(root_level, 0)
    return tree


# ---------------------------------------------------------------------------
# Text extraction
# ---------------------------------------------------------------------------

def extract_text(pdf_path: str | Path, page_start: int, page_end: int) -> str:
    text_parts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for pg in range(page_start, min(page_end + 1, len(pdf.pages))):
            t = pdf.pages[pg].extract_text()
            if t:
                text_parts.append(t)
    return "\n\n".join(text_parts)


def extract_all_chapter_texts(
    pdf_path: str | Path, chapters: list[Section]
) -> dict[str, str]:
    result: dict[str, str] = {}
    for ch in chapters:
        text = extract_text(pdf_path, ch.page_start, ch.page_end)
        result[ch.title] = text
        for sec in ch.sections:
            sec_text = extract_text(pdf_path, sec.page_start, sec.page_end)
            if sec_text:
                result[f"{ch.title} / {sec.title}"] = sec_text
    return result


# ---------------------------------------------------------------------------
# Index-page discovery  (auto-locate via TOC keywords)
# ---------------------------------------------------------------------------

_FIGURE_INDEX_KEYWORDS = ["插图", "图索引", "图目录", "附图", "图示"]
_TABLE_INDEX_KEYWORDS  = ["表格", "表索引", "表目录", "附表"]


def find_index_pages(
    toc: list[tuple[int, str, int]],
) -> tuple[int | None, int | None]:
    figure_page: int | None = None
    table_page: int | None = None

    for _level, title, page in toc:
        for kw in _FIGURE_INDEX_KEYWORDS:
            if kw in title and figure_page is None:
                figure_page = page
                break
        for kw in _TABLE_INDEX_KEYWORDS:
            if kw in title and table_page is None:
                table_page = page
                break

    return figure_page, table_page


# ---------------------------------------------------------------------------
# Figure / table index parsers
# ---------------------------------------------------------------------------

_FIG_LINE = re.compile(
    r"^(图\d+[–\-\—]\d+)\s+(.+?)\s*[\.…]{2,}\s*(\d+)\s*$"
)
_TABLE_LINE = re.compile(
    r"^(表\d+[–\-\—]\d+)\s+(.+?)\s*[\.…]{2,}\s*(\d+)\s*$"
)
_TRAILING_PAGE = re.compile(r"(\d+)\s*$")


def _parse_index_page(
    text: str,
    entry_pattern: re.Pattern,
    is_table: bool = False,
) -> list[tuple[str, str, int]]:
    entries: list[tuple[str, str, int]] = []
    pending: str | None = None

    for raw_line in text.split("\n"):
        line = raw_line.strip()
        if not line or len(line) < 5:
            continue
        if re.match(r"^[XIVxiv\d]+$", line):
            continue
        if any(kw in line for kw in ["上海交通", "学位论文", "页码"]):
            continue

        candidate = (pending + line) if pending else line
        m = entry_pattern.match(candidate)
        if m:
            entries.append((m.group(1), m.group(2).strip(), int(m.group(3))))
            pending = None
        else:
            pm = _TRAILING_PAGE.search(candidate)
            if pm and len(candidate) > 20:
                page = int(pm.group(1))
                rest = candidate[: pm.start()].strip().rstrip(".").strip()
                if is_table:
                    nm = re.match(r"^(表\d+[–\-\—]\d+)\s+(.+)", rest)
                else:
                    nm = re.match(r"^(图\d+[–\-\—]\d+)\s+(.+)", rest)
                if nm:
                    entries.append((nm.group(1), nm.group(2).strip(), page))
                    pending = None
                    continue
            pending = candidate

    return entries


def parse_figure_table_index(
    pdf_path: str | Path,
    figure_page: int | None,
    table_page: int | None,
) -> tuple[list[FigureEntry], list[TableEntry]]:
    figures: list[FigureEntry] = []
    tables: list[TableEntry] = []

    with pdfplumber.open(pdf_path) as pdf:
        if figure_page is not None:
            pg = figure_page - 1
            if pg < len(pdf.pages):
                text = pdf.pages[pg].extract_text() or ""
                for num, cap, doc_pg in _parse_index_page(text, _FIG_LINE):
                    figures.append(FigureEntry(number=num, caption=cap, doc_page=doc_pg))

        if table_page is not None:
            pg = table_page - 1
            if pg < len(pdf.pages):
                text = pdf.pages[pg].extract_text() or ""
                for num, cap, doc_pg in _parse_index_page(
                    text, _TABLE_LINE, is_table=True
                ):
                    tables.append(TableEntry(number=num, caption=cap, doc_page=doc_pg))

    return figures, tables


# ---------------------------------------------------------------------------
# Page-offset calculation
# ---------------------------------------------------------------------------

def _find_chapter_one_page(toc: list[tuple[int, str, int]]) -> int | None:
    for _level, title, page in toc:
        if re.search(r"第[一二三四五六七八九十\d]+章", title):
            return page
    return None


def compute_page_offset(
    toc: list[tuple[int, str, int]],
    figures: list[FigureEntry],
    tables: list[TableEntry],
) -> int:
    """Return the offset such that: pdf_0idx = doc_page + offset - 1."""
    ch1_pdf_page = _find_chapter_one_page(toc)
    if ch1_pdf_page is not None:
        return ch1_pdf_page - 1
    return 0


# ---------------------------------------------------------------------------
# Full-page screenshots of figure / table pages
# ---------------------------------------------------------------------------

def screenshot_figure_table_pages(
    pdf_path: str | Path,
    figure_entries: list[FigureEntry],
    table_entries: list[TableEntry],
    page_offset: int,
    output_dir: str | Path = "output/figures",
    dpi: int = 200,
) -> None:
    """Take full-page screenshots for every figure and table page.

    The *screenshot* attribute on each entry is updated in place.
    If the index is missing (empty *figure_entries* / *table_entries*), this
    is a no-op — the downstream LLM will work from chapter text alone.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)

    for entry_list in (figure_entries, table_entries):
        for entry in entry_list:
            pdf_0idx = entry.doc_page + page_offset - 1
            if pdf_0idx < 0 or pdf_0idx >= len(doc):
                continue

            stem = _safe_filename(f"{entry.number}_{entry.caption}")
            filepath = output_dir / f"{stem}.png"

            if not filepath.exists():
                pix = doc[pdf_0idx].get_pixmap(dpi=dpi)
                pix.save(str(filepath))

            entry.screenshot = str(filepath)

    doc.close()


# ---------------------------------------------------------------------------
# Top-level parser
# ---------------------------------------------------------------------------

def parse(pdf_path: str | Path, output_dir: str | Path = "output/figures") -> ParsedPDF:
    """Run the full deterministic parse on a thesis PDF.

    1. Extract TOC → chapter tree + chapter texts
    2. Locate & parse figure / table index
    3. Full-page screenshots of figure / table pages
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    doc = fitz.open(pdf_path)
    total_pages = len(doc)
    title = doc.metadata.get("title", pdf_path.stem) or pdf_path.stem
    doc.close()

    result = ParsedPDF(title=title, total_pages=total_pages)

    # 1. TOC → chapter tree
    toc = extract_toc(pdf_path)
    result.chapters = build_section_tree(toc, total_pages)

    # 2. Chapter texts
    result.chapter_texts = extract_all_chapter_texts(pdf_path, result.chapters)

    # 3. Locate & parse figure / table index pages
    fig_page, tab_page = find_index_pages(toc)

    if fig_page is not None or tab_page is not None:
        figures, tables = parse_figure_table_index(pdf_path, fig_page, tab_page)
        offset = compute_page_offset(toc, figures, tables)

        # 4. Full-page screenshots (no bbox cropping — LLM handles that)
        screenshot_figure_table_pages(
            pdf_path, figures, tables, offset, output_dir
        )

        result.figure_entries = figures
        result.table_entries = tables

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    input_pdf = sys.argv[1] if len(sys.argv) > 1 else "input/thesis.pdf"
    parsed = parse(input_pdf)

    out_path = Path("output/parsed_data.json")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(parsed.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"Parsed {parsed.total_pages} pages")
    print(f"  Chapters:      {len(parsed.chapters)}")
    print(f"  Figure index:  {len(parsed.figure_entries)} entries")
    print(f"  Table index:   {len(parsed.table_entries)} entries")
    pages_with_fig = {f.doc_page for f in parsed.figure_entries}
    pages_with_tab = {t.doc_page for t in parsed.table_entries}
    print(f"  Figure pages:  {len(pages_with_fig)} (full-page screenshots)")
    print(f"  Table pages:   {len(pages_with_tab)} (full-page screenshots)")
    print(f"  Output → {out_path}")

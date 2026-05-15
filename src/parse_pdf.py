"""Phase 1: PDF parsing — extract TOC, text, images, and tables from a thesis PDF."""

import json
import os
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
    page_start: int  # 0-indexed
    page_end: int    # 0-indexed, inclusive
    sections: list["Section"] = field(default_factory=list)


@dataclass
class ImageRegion:
    page: int       # 0-indexed
    bbox: tuple[float, float, float, float]  # (x0, y0, x1, y1)
    filename: str   # relative path in output/figures/


@dataclass
class TableData:
    page: int
    header: list[str]
    rows: list[list[str]]


@dataclass
class ParsedPDF:
    title: str = ""
    total_pages: int = 0
    chapters: list[Section] = field(default_factory=list)
    images: list[ImageRegion] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

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
            "images": [
                {"page": i.page, "bbox": list(i.bbox), "filename": i.filename}
                for i in self.images
            ],
            "tables": [
                {"page": t.page, "header": t.header, "rows": t.rows}
                for t in self.tables
            ],
        }


# ---------------------------------------------------------------------------
# TOC & chapter tree
# ---------------------------------------------------------------------------

def sanitize_title(raw: str) -> str:
    """Remove excessive whitespace and common PDF artifacts from titles."""
    raw = raw.strip()
    raw = re.sub(r"\s+", " ", raw)
    return raw


def extract_toc(pdf_path: str | Path) -> list[tuple[int, str, int]]:
    """Read the PDF's built-in bookmark/outline as a flat TOC list.

    Returns: list of (level, title, page_number_1_indexed)
    """
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=True)  # → [(level, title, page), ...]
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
    """Convert a flat TOC into a nested section tree with page ranges.

    For each entry, the page range runs from its own start-page to the
    start-page of the next sibling (or the end of the document).
    """
    if not toc:
        return []

    # Convert to 0-indexed pages
    entries = [(level, title, page - 1) for level, title, page in toc]

    def build(level: int, idx: int) -> tuple[list[Section], int]:
        """Recursive descent: consume entries >= *level* starting at *idx*.

        Returns (siblings, next_idx).
        """
        siblings: list[Section] = []
        while idx < len(entries):
            lvl, title, start_page = entries[idx]
            if lvl < level:
                break  # pop back up
            if lvl == level:
                children, idx = build(level + 1, idx + 1)
                # End page = next sibling start or EOF
                end_page = entries[idx][2] - 1 if idx < len(entries) else total_pages - 1
                siblings.append(
                    Section(
                        level=level,
                        title=title,
                        page_start=start_page,
                        page_end=max(start_page, end_page),
                        sections=children,
                    )
                )
            else:  # lvl > level → recurse
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
    """Extract text from a page range using pdfplumber.

    Pages are 0-indexed and inclusive.
    """
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
    """Return {chapter_title: full_text} for every chapter."""
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
# Image extraction
# ---------------------------------------------------------------------------

def extract_images(
    pdf_path: str | Path,
    output_dir: str | Path = "output/figures",
    dpi: int = 300,
) -> list[ImageRegion]:
    """Detect image regions via pdfplumber, screenshot them via PyMuPDF.

    Strategy A from DESIGN.md: use pdfplumber `page.images` to find both
    vector and bitmap artwork regions, then clip-render with PyMuPDF.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    images: list[ImageRegion] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            if not page.images:
                continue
            for i, rect in enumerate(page.images):
                bbox = (rect["x0"], rect["y0"], rect["x1"], rect["y1"])
                filename = f"page_{page_num + 1:03d}_img_{i + 1:02d}.png"
                filepath = output_dir / filename

                # Clip-render with PyMuPDF
                pymu_page = doc[page_num]
                clip = fitz.Rect(*bbox)
                pix = pymu_page.get_pixmap(clip=clip, dpi=dpi)
                pix.save(str(filepath))

                images.append(
                    ImageRegion(
                        page=page_num,
                        bbox=bbox,
                        filename=str(filepath),
                    )
                )

    doc.close()
    return images


# ---------------------------------------------------------------------------
# Table extraction (three-line-table aware)
# ---------------------------------------------------------------------------

def extract_tables_from_page(
    pdf_path: str | Path, page_num: int
) -> list[TableData]:
    """Extract tables from a single page using three-line-table settings.

    See DESIGN.md §1.4 for the rationale behind these settings.
    """
    table_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 15,
        "snap_tolerance": 5,
    }

    with pdfplumber.open(pdf_path) as pdf:
        raw_tables = pdf.pages[page_num].extract_tables(table_settings)

    tables: list[TableData] = []
    for raw in raw_tables:
        if not raw or len(raw) < 2:
            continue
        # Clean None → ""
        cleaned = [[c if c else "" for c in row] for row in raw]
        # Drop all-empty rows
        cleaned = [row for row in cleaned if any(c.strip() for c in row)]
        if len(cleaned) < 2:
            continue
        tables.append(
            TableData(
                page=page_num,
                header=cleaned[0],
                rows=cleaned[1:],
            )
        )
    return tables


def extract_all_tables(
    pdf_path: str | Path, chapters: list[Section]
) -> list[TableData]:
    """Extract tables from all pages referenced in the chapter tree."""
    pages_to_check: set[int] = set()
    for ch in chapters:
        for pg in range(ch.page_start, ch.page_end + 1):
            pages_to_check.add(pg)

    all_tables: list[TableData] = []
    for pg in sorted(pages_to_check):
        all_tables.extend(extract_tables_from_page(pdf_path, pg))
    return all_tables


# ---------------------------------------------------------------------------
# Full-page screenshots (fallback / user reference)
# ---------------------------------------------------------------------------

def screenshot_pages(
    pdf_path: str | Path,
    pages: list[int],
    output_dir: str | Path = "output/figures",
    dpi: int = 200,
) -> list[str]:
    """Render whole pages as PNG images (Strategy B from DESIGN.md).

    Useful as a fallback when automatic image-region detection misses things.
    Returns the list of saved file paths.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    saved: list[str] = []
    for pg in range(len(doc)):
        if pg not in pages:
            continue
        pix = doc[pg].get_pixmap(dpi=dpi)
        filepath = output_dir / f"page_{pg + 1:03d}.png"
        pix.save(str(filepath))
        saved.append(str(filepath))

    doc.close()
    return saved


# ---------------------------------------------------------------------------
# Top-level parser
# ---------------------------------------------------------------------------

def parse(pdf_path: str | Path) -> ParsedPDF:
    """Run the full deterministic parse on a thesis PDF.

    Returns a ParsedPDF dataclass that can be serialized to JSON.
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

    # 2. Images
    result.images = extract_images(pdf_path)

    # 3. Tables
    result.tables = extract_all_tables(pdf_path, result.chapters)

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
    out_path.write_text(json.dumps(parsed.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Parsed {parsed.total_pages} pages")
    print(f"  Chapters: {len(parsed.chapters)}")
    print(f"  Images:   {len(parsed.images)}")
    print(f"  Tables:   {len(parsed.tables)}")
    print(f"  Output → {out_path}")

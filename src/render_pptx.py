"""Phase 2: Render ppt_plan.json into a .pptx file using a template."""

import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ---------------------------------------------------------------------------
# Layout name → slide layout lookup
# ---------------------------------------------------------------------------

_LAYOUT_CACHE: dict[str, int] = {}

# Fallback placeholder names for each layout type when template doesn't
# use standard PowerPoint placeholder names.
_PLACEHOLDER_FALLBACKS: dict[str, dict[str, str]] = {
    "title": {"title": "标题", "subtitle": "副标题", "author": "作者", "date": "日期"},
    "toc": {"title": "标题", "items": "内容"},
    "section_title": {"title": "标题", "subtitle": "副标题"},
    "bullets": {"title": "标题", "body": "正文"},
    "figure": {"title": "标题", "image": "图片", "caption": "图注"},
    "table": {"title": "标题", "table": "表格"},
    "end": {"text": "致谢"},
}

# For each slide type, keywords to look for in placeholder names.
# Used to dynamically match any template without hardcoded name mapping.
_SLIDE_LAYOUT_NEEDS: dict[str, list[str]] = {
    "title":         ["标题", "title", "副标题", "subtitle", "日期", "date", "author"],
    "toc":           ["标题", "title", "内容", "body", "text"],
    "section_title": ["标题", "title", "副标题", "subtitle"],
    "bullets":       ["标题", "title", "正文", "body", "内容", "text"],
    "figure":        ["标题", "title", "图片", "image", "图注", "caption"],
    "table":         ["标题", "title", "表格", "table"],
    "end":           ["致谢", "text", "正文", "content", "标题", "title"],
}


def _build_layout_map(prs: Presentation) -> dict[str, int]:
    """Dynamically match each slide type to the best layout in the template."""
    layouts = []
    for i, layout in enumerate(prs.slide_layouts):
        ph_names = [ph.name for ph in layout.placeholders]
        layouts.append({"idx": i, "name": layout.name, "phs": ph_names})

    mapping: dict[str, int] = {}
    for slide_type, needs in _SLIDE_LAYOUT_NEEDS.items():
        best_idx = 0
        best_score = -1
        for ly in layouts:
            score = sum(1 for kw in needs if any(kw in ph for ph in ly["phs"]))
            # Tiebreaker: more placeholders = richer layout
            if score > best_score or (score == best_score and len(ly["phs"]) > len(layouts[best_idx]["phs"])):
                best_score = score
                best_idx = ly["idx"]
        mapping[slide_type] = best_idx

    print(f"  Layout mapping: { {k: layouts[v]['name'] for k, v in mapping.items()} }")
    return mapping


def _find_layout_index(prs: Presentation, name: str) -> int:
    """Find a slide layout by name, with dynamic fallback."""
    cache_key = f"{id(prs)}_{name}"
    if cache_key in _LAYOUT_CACHE:
        return _LAYOUT_CACHE[cache_key]

    # Dynamic matching: build once per presentation
    dynamic_key = f"{id(prs)}_dynamic"
    if dynamic_key not in _LAYOUT_CACHE:
        layout_map = _build_layout_map(prs)
        for k, v in layout_map.items():
            _LAYOUT_CACHE[f"{id(prs)}_{k}"] = v
        _LAYOUT_CACHE[dynamic_key] = True

    if f"{id(prs)}_{name}" in _LAYOUT_CACHE:
        return _LAYOUT_CACHE[f"{id(prs)}_{name}"]

    # Final fallback: first layout
    _LAYOUT_CACHE[cache_key] = 0
    return 0


# ---------------------------------------------------------------------------
# Placeholder helpers
# ---------------------------------------------------------------------------

def _set_text(shape, text: str) -> None:
    """Write text into a shape, handling both auto-shapes and text frames."""
    if shape.has_text_frame:
        tf = shape.text_frame
        tf.clear()
        tf.paragraphs[0].text = text
    elif shape.has_table:
        shape.table.cell(0, 0).text = text


def _find_placeholder(slide, keywords: list[str]):
    """Find a placeholder shape on the slide by name keyword matching."""
    for shape in slide.placeholders:
        name_lower = shape.name.lower()
        for kw in keywords:
            if kw in name_lower:
                return shape
    return None


def _add_bullet_paragraph(tf, text: str, level: int = 0, font_size: int = 14) -> None:
    """Append a paragraph with bullet formatting to a text frame."""
    p = tf.add_paragraph()
    p.text = text
    p.level = level
    p.space_after = Pt(4)
    for run in p.runs:
        run.font.size = Pt(font_size)


# ---------------------------------------------------------------------------
# Slide fillers
# ---------------------------------------------------------------------------

def _fill_title(slide, content: dict) -> None:
    for key in ("title", "subtitle", "author", "date"):
        val = content.get(key, "")
        if not val:
            continue
        ph = _find_placeholder(slide, [key])
        if ph:
            _set_text(ph, val)


def _fill_toc(slide, content: dict) -> None:
    # Title
    ph = _find_placeholder(slide, ["title"])
    if ph:
        _set_text(ph, content.get("title", "目录"))

    # Items list
    items = content.get("items", [])
    body = _find_placeholder(slide, ["content", "body", "items", "text"])
    if body and body.has_text_frame:
        tf = body.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = items[0] if items else ""
        for item in items[1:]:
            _add_bullet_paragraph(tf, item)


def _fill_section_title(slide, content: dict) -> None:
    for key in ("title", "subtitle"):
        val = content.get(key, "")
        if not val:
            continue
        ph = _find_placeholder(slide, [key])
        if ph:
            _set_text(ph, val)


def _fill_bullets(slide, content: dict) -> None:
    # Title
    ph = _find_placeholder(slide, ["title"])
    if ph:
        _set_text(ph, content.get("title", ""))

    # Bullet list
    bullets = content.get("bullets", [])
    body = _find_placeholder(slide, ["body", "content", "text", "bullets"])
    if body and body.has_text_frame:
        tf = body.text_frame
        tf.clear()
        for i, b in enumerate(bullets):
            text = b["bullet"] if isinstance(b, dict) else str(b)
            if i == 0:
                tf.paragraphs[0].text = text
            else:
                _add_bullet_paragraph(tf, text)


def _fill_figure(slide, content: dict) -> None:
    # Title
    ph = _find_placeholder(slide, ["title"])
    if ph:
        _set_text(ph, content.get("title", ""))

    # Image
    img_path = content.get("image", "")
    if img_path and Path(img_path).exists():
        img_ph = _find_placeholder(slide, ["image", "picture", "图片", "img"])
        if img_ph:
            # Replace placeholder with picture
            left = img_ph.left
            top = img_ph.top
            width = img_ph.width
            height = img_ph.height
            sp = img_ph._element
            sp.getparent().remove(sp)
            slide.shapes.add_picture(str(img_path), left, top, width, height)
        else:
            # Just add picture centered on slide
            slide.shapes.add_picture(
                str(img_path),
                Inches(1.5),
                Inches(2),
                Inches(7),
                Inches(4.5),
            )

    # Caption
    caption = content.get("caption", "")
    if caption:
        ph = _find_placeholder(slide, ["caption", "图注", "subtitle"])
        if ph:
            _set_text(ph, caption)


def _fill_table(slide, content: dict) -> None:
    # Title
    ph = _find_placeholder(slide, ["title"])
    if ph:
        _set_text(ph, content.get("title", ""))

    header = content.get("header", [])
    rows = content.get("rows", [])

    if not header or not rows:
        return

    tbl_ph = _find_placeholder(slide, ["table", "表格"])
    if tbl_ph:
        left = tbl_ph.left
        top = tbl_ph.top
        width = tbl_ph.width
        height = tbl_ph.height
        sp = tbl_ph._element
        sp.getparent().remove(sp)
    else:
        left = Inches(0.8)
        top = Inches(1.8)
        width = Inches(8.4)
        height = Inches(4.5)

    n_rows = len(rows) + 1
    n_cols = len(header)
    table_shape = slide.shapes.add_table(n_rows, n_cols, left, top, width, height)
    table = table_shape.table

    # Header row
    for ci, col_name in enumerate(header):
        cell = table.cell(0, ci)
        cell.text = col_name
        for para in cell.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                run.font.size = Pt(12)
                run.font.bold = True

    # Data rows
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            if ci < n_cols:
                cell = table.cell(ri + 1, ci)
                cell.text = str(val)
                for para in cell.text_frame.paragraphs:
                    para.alignment = PP_ALIGN.CENTER
                    for run in para.runs:
                        run.font.size = Pt(11)


def _fill_end(slide, content: dict) -> None:
    text = content.get("text", "感谢各位老师批评指正")
    ph = _find_placeholder(slide, ["text", "致谢", "thanks", "content", "body", "title"])
    if ph:
        _set_text(ph, text)


# ---------------------------------------------------------------------------
# Filler dispatch
# ---------------------------------------------------------------------------

_FILLERS = {
    "title": _fill_title,
    "toc": _fill_toc,
    "section_title": _fill_section_title,
    "bullets": _fill_bullets,
    "figure": _fill_figure,
    "table": _fill_table,
    "end": _fill_end,
}


# ---------------------------------------------------------------------------
# Main renderer
# ---------------------------------------------------------------------------

def render(
    plan_path: str | Path,
    template_path: str | Path = "template.pptx",
    output_path: str | Path = "output/答辩PPT.pptx",
) -> str:
    """Render ppt_plan.json into a .pptx file.

    Args:
        plan_path: Path to ppt_plan.json.
        template_path: Path to the .pptx template with predefined layouts.
        output_path: Where to write the final .pptx.

    Returns:
        The output file path as a string.
    """
    plan_path = Path(plan_path)
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not plan_path.exists():
        raise FileNotFoundError(f"Plan file not found: {plan_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    prs = Presentation(str(template_path))

    for slide_plan in plan["slides"]:
        layout_name = slide_plan["layout"]
        layout_idx = _find_layout_index(prs, layout_name)
        slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])

        filler = _FILLERS.get(layout_name)
        if filler:
            filler(slide, slide_plan["content"])
        else:
            print(f"  [warn] Unknown layout '{layout_name}' — skipping content fill")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return str(output_path)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    plan_file = sys.argv[1] if len(sys.argv) > 1 else "output/ppt_plan.json"
    tmpl_file = sys.argv[2] if len(sys.argv) > 2 else "template.pptx"
    out_file = sys.argv[3] if len(sys.argv) > 3 else "output/答辩PPT.pptx"

    result = render(plan_file, tmpl_file, out_file)
    print(f"PPTX rendered → {result}")

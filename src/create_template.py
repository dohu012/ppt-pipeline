"""Generate a default template.pptx with all required slide layouts.

Repurposes the 7 built-in slide layouts that ship with a blank PowerPoint file.
Users are encouraged to replace the generated template with their own design.
"""

from pptx import Presentation
from pptx.util import Inches


# Maps our layout names → built-in layout index in a default blank .pptx
LAYOUT_MAP = {
    "title":         0,   # Title Slide        → title + subtitle
    "bullets":       1,   # Title and Content  → title + body
    "section_title": 2,   # Section Header     → title + subtitle
    "toc":           3,   # Two Content        → title + body (we rename one body)
    "end":           5,   # Title Only         → just a text placeholder
    "table":         7,   # Content + Caption  → title + body + caption
    "figure":        8,   # Picture + Caption  → title + picture + caption
}


def _rename_placeholders(layout, renames: dict[int, str]) -> None:
    """Rename placeholders on a slide layout by their placeholder idx."""
    for ph in layout.placeholders:
        idx = ph.placeholder_format.idx
        if idx in renames:
            ph.name = renames[idx]


def create_default_template(output_path: str = "template.pptx") -> str:
    """Generate a basic template.pptx with all 7 layout types.

    Layout names (matching ppt_plan.json): title, toc, section_title,
    bullets, figure, table, end.
    """
    prs = Presentation()

    # 16:9 widescreen
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # ---- title (built-in #0: Title Slide) ----
    title_ly = prs.slide_layouts[0]
    title_ly.name = "title"
    _rename_placeholders(title_ly, {0: "title", 1: "subtitle", 11: "author"})

    # ---- bullets (built-in #1: Title and Content) ----
    bullets_ly = prs.slide_layouts[1]
    bullets_ly.name = "bullets"
    _rename_placeholders(bullets_ly, {0: "title", 1: "body"})

    # ---- section_title (built-in #2: Section Header) ----
    sec_ly = prs.slide_layouts[2]
    sec_ly.name = "section_title"
    _rename_placeholders(sec_ly, {0: "title", 1: "subtitle"})

    # ---- toc (built-in #3: Two Content) ----
    toc_ly = prs.slide_layouts[3]
    toc_ly.name = "toc"
    _rename_placeholders(toc_ly, {0: "title", 1: "content"})

    # ---- end (built-in #5: Title Only) ----
    end_ly = prs.slide_layouts[5]
    end_ly.name = "end"
    _rename_placeholders(end_ly, {0: "text"})

    # ---- table (built-in #7: Content with Caption) ----
    table_ly = prs.slide_layouts[7]
    table_ly.name = "table"
    _rename_placeholders(table_ly, {0: "title", 1: "table"})

    # ---- figure (built-in #8: Picture with Caption) ----
    figure_ly = prs.slide_layouts[8]
    figure_ly.name = "figure"
    _rename_placeholders(figure_ly, {0: "title", 1: "image", 2: "caption"})

    # Hide the unused built-in layouts by appending "(unused)" to their names
    unused = set(range(len(prs.slide_layouts))) - set(LAYOUT_MAP.values())
    for idx in unused:
        ly = prs.slide_layouts[idx]
        if not ly.name.endswith("(unused)"):
            ly.name = f"{ly.name} (unused)"

    prs.save(output_path)
    return output_path


if __name__ == "__main__":
    import sys

    out = sys.argv[1] if len(sys.argv) > 1 else "template.pptx"
    create_default_template(out)
    print(f"Default template created → {out}")
    print("Layouts:", [ly.name for ly in Presentation(out).slide_layouts])

"""
DOCX grid export.

Each product type gets its own page section (with the correct orientation),
a heading, and a table grid of images sized to the product's physical print
dimensions. Images are resized via Pillow and saved as PNG before insertion.
"""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image as PilImage
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT


# ---------------------------------------------------------------------------
# Product type registry
# ---------------------------------------------------------------------------

PRODUCT_TYPES = ["Lighter", "Stash Jar", "Grinder", "Rolling Tray", "Other"]


@dataclass
class ProductSpec:
    """Target print dimensions and layout for one product category."""
    width_in: float    # image width in the document (inches)
    height_in: float   # image height in the document (inches)
    cols: int          # grid columns per row
    landscape: bool = False  # page orientation for this section


SPECS: dict[str, ProductSpec] = {
    # 4 across × 5 down on a landscape page
    "Lighter":      ProductSpec(width_in=2.22, height_in=1.42, cols=4, landscape=True),
    # 3 × 4 on a portrait page
    "Stash Jar":    ProductSpec(width_in=2.55, height_in=2.55, cols=3, landscape=False),
    # 3 × 4 on a portrait page
    "Grinder":      ProductSpec(width_in=2.35, height_in=2.35, cols=3, landscape=False),
    # 2 portrait trays per row
    "Rolling Tray": ProductSpec(width_in=3.90, height_in=5.70, cols=2, landscape=False),
    "Other":        ProductSpec(width_in=2.00, height_in=2.00, cols=3, landscape=False),
}

PAGE_W = 8.5    # portrait page width  (inches)
PAGE_H = 11.0   # portrait page height (inches)
MARGIN = 0.35   # all-sides margin     (inches)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

@dataclass
class QueueEntry:
    path: Path
    product_type: str = "Other"
    quantity: int = 1


def build_docx(entries: list[QueueEntry], dest: Path) -> None:
    """Build the multi-section grid DOCX and save it to *dest*."""
    doc = Document()

    # Group by product type, preserving first-occurrence order
    groups: dict[str, list[QueueEntry]] = {}
    for e in entries:
        groups.setdefault(e.product_type, []).append(e)

    for idx, (product_type, group) in enumerate(groups.items()):
        spec = SPECS.get(product_type, SPECS["Other"])

        section = doc.sections[0] if idx == 0 else doc.add_section()
        _configure_section(section, spec.landscape)

        heading = doc.add_paragraph(product_type)
        heading.style = doc.styles["Heading 1"]

        # Expand quantity into repeated paths
        image_paths: list[Path] = []
        for e in group:
            image_paths.extend([e.path] * max(1, e.quantity))

        actual_w = _fit_width(spec)
        _add_image_grid(doc, image_paths, spec, actual_w)

    doc.save(str(dest))


# ---------------------------------------------------------------------------
# Internals
# ---------------------------------------------------------------------------

def _fit_width(spec: ProductSpec) -> float:
    """Largest image width (inches) that fits within the page's column layout."""
    # Usable width depends on orientation
    page_content_w = (PAGE_H if spec.landscape else PAGE_W) - 2 * MARGIN
    per_col = page_content_w / spec.cols
    # 4 % slack for table cell padding / borders
    return min(spec.width_in, per_col * 0.96)


def _configure_section(section, landscape: bool) -> None:
    m = Inches(MARGIN)
    section.top_margin    = m
    section.bottom_margin = m
    section.left_margin   = m
    section.right_margin  = m

    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width  = Inches(PAGE_H)   # 11"
        section.page_height = Inches(PAGE_W)   # 8.5"
    else:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width  = Inches(PAGE_W)   # 8.5"
        section.page_height = Inches(PAGE_H)   # 11"


def _add_image_grid(
    doc: Document,
    paths: list[Path],
    spec: ProductSpec,
    actual_w: float,
) -> None:
    if not paths:
        return

    cols = spec.cols
    rows_needed = (len(paths) + cols - 1) // cols

    table = doc.add_table(rows=rows_needed, cols=cols)
    table.style = "Table Grid"

    for i, img_path in enumerate(paths):
        row_idx = i // cols
        col_idx = i % cols
        cell = table.cell(row_idx, col_idx)
        cell.width = Inches(actual_w)

        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.add_run()

        try:
            buf = _prepare_image(img_path, actual_w, spec.height_in)
            run.add_picture(buf, width=Inches(actual_w), height=Inches(spec.height_in))
        except Exception as exc:
            run.text = f"[Error: {img_path.name}]"

        cap = cell.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap_run = cap.add_run(img_path.name)
        cap_run.font.size = Pt(7)

    # Clear any leftover empty cells in the last row
    for empty in range(len(paths), rows_needed * cols):
        cell = table.cell(empty // cols, empty % cols)
        for para in cell.paragraphs:
            para.clear()


def _prepare_image(path: Path, width_in: float, height_in: float, dpi: int = 150) -> io.BytesIO:
    """
    Resize *path* to the exact target dimensions and return as a PNG BytesIO.
    Images are converted to RGBA (PNG) to preserve transparency where present.
    Aspect-ratio distortion is intentional here; Pillow correction comes later.
    """
    target_w = max(1, int(width_in * dpi))
    target_h = max(1, int(height_in * dpi))

    with PilImage.open(path) as img:
        img = img.convert("RGBA")
        img = img.resize((target_w, target_h), PilImage.Resampling.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf

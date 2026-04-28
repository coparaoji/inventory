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

from PIL import Image as PilImage, ImageDraw
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
    landscape: bool = False                  # page orientation for this section
    corner_radius: int = 0                   # rounded-corner radius in pixels (0 = square)
    circle: bool = False                     # True → radius auto-set to half the shortest side (overrides corner_radius)
    border_px: int = 0                       # border stroke thickness in pixels (0 = no border)
    border_color: tuple = (0, 0, 0, 255)     # border color as RGBA


SPECS: dict[str, ProductSpec] = {
    "Lighter":      ProductSpec(width_in=1.42, height_in=2.22, cols=5, corner_radius=15),
    "Stash Jar":    ProductSpec(width_in=2.55, height_in=2.55, cols=3, circle=True),
    "Grinder":      ProductSpec(width_in=2.35, height_in=2.35, cols=3),
    "Rolling Tray": ProductSpec(width_in=3.90, height_in=5.70, cols=2),
    "Other":        ProductSpec(width_in=2.00, height_in=2.00, cols=3),
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
            buf = _prepare_image(img_path, actual_w, spec.height_in, spec=spec)
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


def _apply_rounded_corners(img: PilImage.Image, radius: int = 20) -> PilImage.Image:
    """
    Punch rounded corners into *img* by applying an alpha mask.
    Pixels outside the rounded rectangle become fully transparent.
    *radius* is in pixels relative to the resized image dimensions.
    """
    img = img.convert("RGBA")
    mask = PilImage.new("L", img.size, 0)
    ImageDraw.Draw(mask).rounded_rectangle([(0, 0), img.size], radius=radius, fill=255)
    img.putalpha(mask)
    return img


def _apply_border(
    img: PilImage.Image,
    border_px: int = 6,
    radius: int = 20,
    color: tuple = (0, 0, 0, 255),
) -> PilImage.Image:
    """
    Draw a rounded-rectangle stroke over *img*.
    *border_px* controls line thickness; *radius* should match _apply_rounded_corners.
    *color* is an RGBA tuple.
    """
    img = img.convert("RGBA")
    w, h = img.size
    half = border_px // 2
    ImageDraw.Draw(img).rounded_rectangle(
        [(half, half), (w - half, h - half)],
        radius=radius,
        outline=color,
        width=border_px,
    )
    return img


def _prepare_image(
    path: Path,
    width_in: float,
    height_in: float,
    dpi: int = 150,
    spec: ProductSpec | None = None,
) -> io.BytesIO:
    """
    Resize *path* to the exact target dimensions and return as a PNG BytesIO.
    Images are converted to RGBA (PNG) to preserve transparency where present.
    Aspect-ratio distortion is intentional here; Pillow correction comes later.
    Image manipulation (corners, border) is driven by the ProductSpec fields.
    """
    target_w = max(1, int(width_in * dpi))
    target_h = max(1, int(height_in * dpi))

    border_px    = spec.border_px    if spec else 6
    border_color = spec.border_color if spec else (0, 0, 0, 255)

    # circle=True auto-computes the radius so the square image becomes a circle,
    # regardless of the exact pixel dimensions after _fit_width scaling
    if spec and spec.circle:
        corner_radius = min(target_w, target_h) // 2
    else:
        corner_radius = spec.corner_radius if spec else 0

    with PilImage.open(path) as img:
        img = img.convert("RGBA")
        src_w, src_h = img.size
        # Align orientation: always map the longer source side to the longer target side
        if (src_w > src_h) != (target_w > target_h):
            target_w, target_h = target_h, target_w
        img = img.resize((target_w, target_h), PilImage.Resampling.LANCZOS)
        # Skip masking if the image already has meaningful transparency (e.g. pre-circled PNG)
        # so we don't overwrite soft anti-aliased edges with a hard binary mask
        already_transparent = img.split()[3].getextrema()[0] < 255
        if corner_radius > 0 and not already_transparent:
            img = _apply_rounded_corners(img, radius=corner_radius)
        if border_px > 0:
            img = _apply_border(img, border_px=border_px, radius=corner_radius, color=border_color)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf

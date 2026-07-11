"""
Split a multi-page Word document into one .docx file per page.

Two signals mark a page boundary:
  1. ``<w:sectPr>`` inside a paragraph's ``<w:pPr>`` -- a "Next Page" section break.
  2. ``<w:br w:type="page"/>`` -- a Ctrl+Enter hard page break.

The boundary paragraph is kept on the current page, but every hard page break is
stripped from the emitted copies so each output file renders as exactly one page.
``<w:lastRenderedPageBreak>`` is ignored -- it is a rendering hint, not a layout
command.
"""
from __future__ import annotations

import copy
import zipfile
from pathlib import Path

from lxml import etree

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS = {
    "w":   _W,
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}


def _strip_page_breaks(el):
    """Return a deep copy of *el* with all hard page-break runs removed."""
    el_copy = copy.deepcopy(el)

    for br in el_copy.xpath(".//w:br[@w:type='page']", namespaces={"w": _W}):
        run = br.getparent()
        run.remove(br)
        # Drop the run itself if nothing useful remains.
        remaining = [c for c in run if c.tag != f"{{{_W}}}rPr"]
        if not remaining:
            run.getparent().remove(run)

    # Remove <w:sectPr> from <w:pPr> -- it is promoted to body level instead.
    if el_copy.tag == f"{{{_W}}}p":
        p_pr = el_copy.find(f"{{{_W}}}pPr")
        if p_pr is not None:
            sect = p_pr.find(f"{{{_W}}}sectPr")
            if sect is not None:
                p_pr.remove(sect)

    return el_copy


def split_docx_by_page(src: Path, output_dir: Path | None = None) -> list[Path]:
    """Split *src* into ``template_1.docx`` ... ``template_N.docx``.

    Files are written to *output_dir* (defaults to ``<src-parent>/split_pages``).
    Returns the list of written paths.
    """
    src = Path(src)
    if output_dir is None:
        output_dir = src.parent / "split_pages"
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src) as z:
        all_files = {name: z.read(name) for name in z.namelist()}

    root = etree.fromstring(all_files["word/document.xml"])
    body = root.find(f"{{{_W}}}body")
    elems = list(body)

    doc_sect_pr = None
    if elems and elems[-1].tag == f"{{{_W}}}sectPr":
        doc_sect_pr = elems.pop()

    pages: list[tuple[list, object]] = []
    current: list = []

    for el in elems:
        current.append(el)

        # Signal 1 -- section break in pPr (grab its sectPr for page layout).
        if el.tag == f"{{{_W}}}p":
            p_pr = el.find(f"{{{_W}}}pPr")
            if p_pr is not None:
                sect = p_pr.find(f"{{{_W}}}sectPr")
                if sect is not None:
                    pages.append((current, copy.deepcopy(sect)))
                    current = []
                    continue

        # Signal 2 -- hard page break inside any run.
        if el.xpath(".//w:br[@w:type='page']", namespaces={"w": _W}):
            pages.append((current, doc_sect_pr))
            current = []

    pages.append((current, doc_sect_pr))  # trailing content = last page

    output_paths: list[Path] = []
    for n, (page_elems, page_sect) in enumerate(pages, 1):
        new_body = etree.Element(f"{{{_W}}}body")
        for el in page_elems:
            new_body.append(_strip_page_breaks(el))
        if page_sect is not None:
            new_body.append(copy.deepcopy(page_sect))

        new_root = copy.deepcopy(root)
        old_body = new_root.find(f"{{{_W}}}body")
        new_root.remove(old_body)
        new_root.append(new_body)

        new_bytes = etree.tostring(
            new_root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

        out = output_dir / f"template_{n}.docx"
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in all_files.items():
                zout.writestr(name, new_bytes if name == "word/document.xml" else data)
        output_paths.append(out)

    return output_paths

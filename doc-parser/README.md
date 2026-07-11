# docparser

A small Python library for working with Word (`.docx`) documents in a
print-production workflow. It has two distinct capabilities that grew out of the
same project: **generating** documents from scratch, and **manipulating** existing
Word templates at the raw OOXML level.

## Install

```bash
pip install -e .
```

Dependencies: `python-docx`, `Pillow`, `lxml`.

## Modules

### `docparser.generate` -- build image grids

Builds a multi-section `.docx` where each product type gets a page with a grid of
images sized to real print dimensions (via `Pillow`). Handles rounded corners,
colored borders, circular masks, and a special mixed-orientation layout for rolling
trays.

```python
from pathlib import Path
from docparser import QueueEntry, build_docx

entries = [
    QueueEntry(path=Path("grinder.png"), product_type="Grinder", quantity=3),
    QueueEntry(path=Path("jar.jpg"),     product_type="Stash Jar", quantity=2),
]
build_docx(entries, Path("out.docx"))
```

Product specs live in `SPECS` / `PRODUCT_TYPES` and are easy to extend.

### `docparser.template` -- inspect and fill templates

Works below `python-docx` to handle pre-drawn shape "slots" (`<wps:wsp>`):

```python
from pathlib import Path
from docparser import analyse_docx, extract_media, fill_shape

analyse_docx(Path("template.docx"))                 # print a full structural report
extract_media(Path("template.docx"), Path("media")) # dump embedded images
fill_shape(                                          # drop an image into slot 0
    template=Path("template.docx"),
    output=Path("filled.docx"),
    shape_index=0,
    image=Path("photo.png"),
)
```

`fill_shape` swaps a shape's `<a:solidFill>` for a `<a:blipFill>` and wires up the
media file and relationship, leaving position/size untouched.

### `docparser.split` -- split by page

```python
from pathlib import Path
from docparser import split_docx_by_page

pages = split_docx_by_page(Path("multi_page.docx"), Path("pages"))
# -> [pages/template_1.docx, pages/template_2.docx, ...]
```

Splits on section breaks and hard page breaks, stripping the breaks so each output
renders as a single page.

## Example

```bash
python examples/generate_grid_example.py
```

## Layout

```
doc-parser/
  pyproject.toml
  src/docparser/
    __init__.py
    generate.py    # image-grid generation
    template.py    # OOXML inspection + shape fill
    split.py       # per-page splitter
  examples/
    generate_grid_example.py
    sample-grinder.png
```

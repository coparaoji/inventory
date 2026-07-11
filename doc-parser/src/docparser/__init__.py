"""docparser -- generate, inspect, split and fill Word (.docx) documents.

Three modules:

* :mod:`docparser.generate` -- build print-ready image-grid documents from scratch.
* :mod:`docparser.template` -- inspect and fill pre-drawn shape templates.
* :mod:`docparser.split`    -- split a multi-page document into per-page files.
"""
from __future__ import annotations

from .generate import (
    PRODUCT_TYPES,
    SPECS,
    ProductSpec,
    QueueEntry,
    build_docx,
)
from .split import split_docx_by_page
from .template import analyse_docx, extract_media, fill_shape

__all__ = [
    "PRODUCT_TYPES",
    "SPECS",
    "ProductSpec",
    "QueueEntry",
    "build_docx",
    "split_docx_by_page",
    "analyse_docx",
    "extract_media",
    "fill_shape",
]

__version__ = "0.1.0"

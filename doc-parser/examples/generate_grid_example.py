"""Minimal example: build a print-ready image-grid .docx from a few images.

Run from the repo root after `pip install -e doc-parser`:

    python doc-parser/examples/generate_grid_example.py
"""
from pathlib import Path

from docparser import QueueEntry, build_docx

HERE = Path(__file__).parent
SAMPLE = HERE / "sample-grinder.png"


def main() -> None:
    entries = [
        QueueEntry(path=SAMPLE, product_type="Grinder", quantity=3),
        QueueEntry(path=SAMPLE, product_type="Stash Jar", quantity=2),
    ]
    dest = HERE / "example_grid.docx"
    build_docx(entries, dest)
    print(f"Wrote {dest}")


if __name__ == "__main__":
    main()

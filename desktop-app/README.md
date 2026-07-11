# Desktop App

A cross-platform PyQt6 desktop tool for preparing print-ready image sheets for
custom merchandise (lighters, stash jars, grinders, rolling trays).

## What it does

1. **Select a folder** of product images (recursively scanned).
2. **Browse** them in a sortable, searchable table with a live preview.
3. **Build a print queue** -- add images, assign a product type and quantity.
4. **Export to `.docx`** -- generates a Word document with each product type laid
   out as a grid sized to real print dimensions.

The document generation is delegated to the [`docparser`](../doc-parser) library in
this repo (`build_docx`), so the app itself stays focused on the UI.

## Run

```bash
# From the repo root:
python -m venv .venv && source .venv/bin/activate
pip install -e doc-parser                      # the docparser library
pip install -r desktop-app/requirements.txt    # PyQt6
python desktop-app/main.py
```

On Windows, `launcher.bat` will create a virtualenv, install dependencies and start
the app.

## Layout

```
desktop-app/
  main.py              # QApplication entry point
  app/
    __init__.py
    main_window.py     # all UI: browser, preview, print queue, export
  assets/              # sample images for a quick demo run
  requirements.txt
  launcher.bat         # Windows one-click launcher
```

## Sample images

`assets/` contains a couple of sample product images so you can try the export
flow without supplying your own.

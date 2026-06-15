# LaTeX Doc Builder

[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

One-command document build pipeline for turning structured Word manuscripts,
Excel tables, figures, and BibTeX references into publication-ready LaTeX/PDF
and Word outputs.

The project is designed for researchers and analysts who prefer drafting in
Word and Excel but need reproducible, journal-style final documents.

## What It Does

- Reads manuscript text from `content/text.docx`.
- Reads tables from `content/tables.xlsx`.
- Reads figures from `figures/`.
- Reads references from `content/references.bib.txt`.
- Generates a LaTeX manuscript and compiled PDF.
- Optionally generates a formatted Word document.
- Uses an editable MDPI-style LaTeX template.

## Workflow

```text
content/text.docx
content/tables.xlsx
content/references.bib.txt
figures/*
        -> generate.py
        -> output/output.pdf
        -> output/output.docx
```

## Quick Start

Install Python dependencies:

```bash
pip install -r requirements.txt
```

Install LaTeX:

| OS | Command |
| --- | --- |
| Windows | `winget install MiKTeX.MiKTeX` |
| macOS | `brew install --cask mactex` |
| Linux | `sudo apt-get install texlive-full` |

Run:

```bash
python generate.py --all
```

Useful modes:

```bash
python generate.py           # PDF only
python generate.py --word    # Word only
python generate.py --all     # PDF and Word
python generate.py --clean   # remove LaTeX auxiliary files after build
```

## Repository Layout

```text
latex-doc-builder/
  content/
    text.docx                # manuscript
    tables.xlsx              # one sheet per table
    references.bib.txt       # BibTeX references
  figures/
    *.png / *.jpg / *.pdf    # figures used by the manuscript
  template/
    mdpi_style.tex           # LaTeX preamble and style
  output/                    # generated PDF and Word files
  generate.py                # main build script
  requirements.txt
```

## Manuscript Conventions

Add metadata at the top of `content/text.docx` as plain paragraphs:

```text
TITLE: Full paper title
SHORT_TITLE: Running header title
AUTHORS: First Author$^{1}$, Second Author$^{2}$
SHORT_AUTHORS: Author et al.
AFFILIATIONS: $^{1}$ Department, University; $^{2}$ Institute
JOURNAL: Journal Name
DATE: 2026
ABSTRACT: Abstract text.
KEYWORDS: keyword one; keyword two; keyword three
```

Supported Word styles:

| Word style | LaTeX output |
| --- | --- |
| Heading 1 | `\section{}` |
| Heading 2 | `\subsection{}` |
| Heading 3 | `\subsubsection{}` |
| Bold | `\textbf{}` |
| Italic | `\textit{}` |
| List Number | `enumerate` |
| List Bullet | `itemize` |

## Tables

Insert a table with:

```text
TABLE: sheet_name
```

The Excel sheet name must match `sheet_name`.

Expected table layout in `content/tables.xlsx`:

| Row | Content |
| --- | --- |
| Row 1 | Table caption |
| Row 2 | Column headers |
| Row 3+ | Data |

Optional final column: `COLOR`

Accepted values: `green`, `blue`, `gray`, `yellow`, `orange`, `red`, `purple`.

## Figures

Insert a figure with:

```text
FIGURE: filename.png | Caption text | width_fraction
```

Examples:

```text
FIGURE: map.png | Study area | 0.80
FIGURE: chart.png | Results by group | 0.65
FIGURE: photo.png
```

If `width_fraction` is omitted, the default is `0.85`.

## References

Store BibTeX records in:

```text
content/references.bib.txt
```

Example:

```bibtex
@article{key2024,
  author  = {Last, First and Other, Author},
  title   = {Article Title},
  journal = {Journal Name},
  year    = {2024},
  volume  = {1},
  pages   = {1--10},
  doi     = {10.xxxx/xxxxx}
}
```

`generate.py` copies this file to `references.bib` at build time.

## Template Customization

Edit `template/mdpi_style.tex` to change:

- fonts;
- colors;
- margins;
- section styling;
- headers and footers;
- citation and bibliography behavior.

## Troubleshooting

Missing Python packages:

```bash
pip install python-docx openpyxl
```

Missing LaTeX packages on Linux/macOS:

```bash
tlmgr install setspace tcolorbox pgfplots booktabs multirow \
              fancyhdr titlesec enumitem caption natbib mathpazo
```

On Windows, MiKTeX usually prompts to install missing packages on first run.
Accept the prompts when `pdflatex` runs.

Inspect LaTeX errors:

```bash
grep "^!" main.log | head -20
```

## Quality Checks

```bash
python -m compileall -q generate.py
python generate.py --help
```

## License

MIT. See `LICENSE`.

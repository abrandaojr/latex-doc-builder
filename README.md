# LaTeX Doc Builder

A Python pipeline that reads your **Word manuscript**, **Excel tables**, **PNG figures**, and **BibTeX references**, then generates a publication-ready **PDF** (MDPI-style template) and a formatted **Word document** — with a single command.

---

## How It Works

```
content/text.docx          ─┐
content/tables.xlsx          ├──► generate.py ──► output/output.pdf
content/references.bib.txt   │                    output/output.docx
figures/*.png               ─┘
```

---

## Quick Start

### 1 — Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2 — Install LaTeX

| OS | How |
|---|---|
| Windows | `winget install MiKTeX.MiKTeX` |
| macOS | `brew install --cask mactex` |
| Linux | `sudo apt-get install texlive-full` |

### 3 — Run

```bash
python generate.py --all     # PDF + Word
python generate.py           # PDF only
python generate.py --word    # Word only
python generate.py --clean   # remove .aux .log etc. after build
```

---

## Project Layout

```
latex-doc-builder/
├── content/
│   ├── text.docx            ← write your manuscript here
│   ├── tables.xlsx          ← one sheet per table
│   └── references.bib.txt  ← BibTeX references
├── figures/
│   └── *.png / *.jpg        ← drop figures here
├── template/
│   └── mdpi_style.tex       ← LaTeX preamble (edit to customize style)
├── output/                  ← generated PDF and Word files
├── generate.py              ← main script
└── requirements.txt
```

---

## Writing Your Manuscript (`content/text.docx`)

### Metadata — add at the top as plain paragraphs

```
TITLE: Full paper title
SHORT_TITLE: Running header title
AUTHORS: First Author$^{1}$, Second Author$^{2}$
SHORT_AUTHORS: Author et al.
AFFILIATIONS: $^{1}$ Department, University; $^{2}$ Institute
JOURNAL: Journal Name
DATE: 2026
ABSTRACT: Your abstract text here.
KEYWORDS: keyword one; keyword two; keyword three
```

### Word styles → LaTeX

| Word style | Output |
|---|---|
| Heading 1 | `\section{}` |
| Heading 2 | `\subsection{}` |
| Heading 3 | `\subsubsection{}` |
| **Bold** | `\textbf{}` |
| *Italic* | `\textit{}` |
| List Number | `\begin{enumerate}` |
| List Bullet | `\begin{itemize}` |

### Inserting a table

Write in a plain paragraph:
```
TABLE: sheet_name
```
`sheet_name` must match the Excel sheet name exactly.

### Inserting a figure

```
FIGURE: filename.png | Caption text | width_fraction
```

| Field | Description |
|---|---|
| `filename.png` | File in `figures/` |
| `Caption text` | Text below figure |
| `width_fraction` | 0.0–1.0, default `0.85` |

Examples:
```
FIGURE: map.png | Study area | 0.80
FIGURE: chart.png | Results by group | 0.65
FIGURE: photo.png
```

---

## Managing Tables (`content/tables.xlsx`)

| Row | Content |
|---|---|
| Row 1 | Table caption |
| Row 2 | Column headers |
| Row 3+ | Data |

Optional last column named **`COLOR`**: `green`, `blue`, `gray`, `yellow`, `orange`, `red`, `purple`.

---

## Managing References (`content/references.bib.txt`)

Standard BibTeX format. The `.bib.txt` extension opens in any text editor.
`generate.py` copies it as `references.bib` at build time.

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

---

## Customizing the Template

Edit `template/mdpi_style.tex` to change:
- **Fonts** — `mathpazo`, `helvet`
- **Colours** — `\definecolor` entries
- **Margins** — `geometry` options
- **Section style** — `titlesec` formats

---

## Troubleshooting

```bash
# Missing Python packages
pip install python-docx openpyxl

# Missing LaTeX packages (Linux/macOS)
tlmgr install setspace tcolorbox pgfplots booktabs multirow \
              fancyhdr titlesec enumitem caption natbib mathpazo

# Missing LaTeX packages (Windows MiKTeX — auto-installs on first run)
# Accept all install prompts when pdflatex runs

# See full LaTeX error log
cat main.log | grep "^!" | head -20
```

---

## License

MIT — see [LICENSE](LICENSE)

# 📄 LaTeX Doc Builder

A generic Python pipeline that reads structured content from **Microsoft Word**, **Excel**, **PNG figures**, and **BibTeX**, then generates a publication-ready **PDF** (MDPI-style academic template) and a formatted **Word document** — fully automated.

---

## ✨ Features

- Write your manuscript in **Word** using regular styles (Heading 1, Heading 2, bold, italic, lists)
- Manage **tables** in Excel (one sheet per table, with optional row colouring)
- Drop **PNG/JPG figures** in a folder — reference them inline with a simple marker
- Maintain **BibTeX references** in a plain `.bib.txt` file
- Run one command → get `output/output.pdf` and `output/output.docx`
- Academic template inspired by **MDPI** — no proprietary class files required
- Works on **Windows**, **macOS**, and **Linux**

---

## 🗂️ Project Structure

```
latex-doc-builder/
│
├── content/
│   ├── text.docx           ← Write your manuscript here
│   ├── tables.xlsx         ← One Excel sheet per table
│   └── references.bib.txt  ← BibTeX references
│
├── figures/
│   └── *.png / *.jpg / *.pdf   ← Drop figures here
│
├── template/
│   └── mdpi_style.tex      ← LaTeX preamble (MDPI-inspired, no logo)
│
├── output/                 ← Generated files land here
│   ├── output.pdf
│   └── output.docx
│
├── generate.py             ← Main script
├── requirements.txt
└── README.md
```

---

## 🚀 Quick Start

### 1 — Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2 — Install LaTeX (needed for PDF generation)

| OS | Recommended distribution | Command |
|---|---|---|
| Windows | [MiKTeX](https://miktex.org/download) | `winget install MiKTeX.MiKTeX` |
| macOS | [MacTeX](https://www.tug.org/mactex/) | `brew install --cask mactex` |
| Linux | TeX Live | `sudo apt-get install texlive-full` |

### 3 — Run

```bash
# Generate both PDF and Word
python generate.py --all

# PDF only (default)
python generate.py

# Word only
python generate.py --word

# Generate all and clean auxiliary LaTeX files
python generate.py --all --clean
```

Output files appear in `output/`.

---

## ✍️ How to Write Your Manuscript (`content/text.docx`)

### Metadata block

Add these lines as plain paragraphs **at the top** of the document (before any section):

```
TITLE: Full title of your paper
SHORT_TITLE: Short title for running header (max 60 chars)
AUTHORS: Author Name$^{1}$, Second Author$^{2}$
SHORT_AUTHORS: Name et al.
AFFILIATIONS: $^{1}$ Institution 1, City, Country; $^{2}$ Institution 2
JOURNAL: Journal Name
DATE: 2026
ABSTRACT: Single paragraph abstract text here.
KEYWORDS: keyword 1; keyword 2; keyword 3
```

### Text styles (use Word's built-in styles)

| Word style | LaTeX output |
|---|---|
| Heading 1 | `\section{}` |
| Heading 2 | `\subsection{}` |
| Heading 3 | `\subsubsection{}` |
| **Bold** (Ctrl+B) | `\textbf{}` |
| *Italic* (Ctrl+I) | `\textit{}` |
| List Number | `\begin{enumerate}` |
| List Bullet | `\begin{itemize}` |
| Normal | Paragraph |

### Inserting a table

Write a plain paragraph with the marker:

```
TABLE: sheet_name
```

Where `sheet_name` matches the Excel sheet name exactly.

### Inserting a figure

Write a plain paragraph:

```
FIGURE: filename.png | Caption text here | width_fraction
```

| Field | Description | Required |
|---|---|---|
| `filename.png` | File inside `figures/` folder | Yes |
| `Caption text` | Text below figure in PDF | No |
| `width_fraction` | Fraction of page width (0.0–1.0). Default: `0.85` | No |

**Examples:**
```
FIGURE: map_cerrado.png | Study area map showing pasture distribution | 0.80
FIGURE: results_chart.png | Main results by category | 0.65
FIGURE: photo.png
```

### Citations

Write BibTeX citation keys directly in text using standard LaTeX notation.
Since Word doesn't process LaTeX, you can use brackets that pass through unchanged:

```
As shown by \citep{gibbs2015}, the soy moratorium...
```

Or use Word's tracked changes / comments to indicate citation positions.

---

## 📊 How to Manage Tables (`content/tables.xlsx`)

Each Excel **sheet** corresponds to one table in the document.

| Row | Content |
|---|---|
| Row 1 | Table caption (first cell) |
| Row 2 | Column headers |
| Row 3+ | Data rows |

### Optional row colouring

Add a last column named exactly **`COLOR`**. Accepted values:

| Value | Colour |
|---|---|
| `green` | Light green |
| `blue` | Light blue |
| `gray` | Light gray |
| `yellow` | Light yellow |
| `orange` | Light orange |
| `red` | Light red |
| `purple` | Light purple |

Leave the COLOR cell empty for no shading.

### Adding a new table

1. Create a new sheet in `tables.xlsx`
2. Follow the Row 1 / Row 2 / Row 3+ structure
3. Reference it in `text.docx` with `TABLE: your_sheet_name`

---

## 📚 How to Manage References (`content/references.bib.txt`)

The file uses `.bib.txt` extension so it opens in any text editor without special software. The generator automatically copies it as `references.bib` for LaTeX.

Add entries in standard **BibTeX format**:

```bibtex
@article{key2024,
  author  = {Last, First and Other, Author},
  title   = {Title of the Article},
  journal = {Journal Name},
  year    = {2024},
  volume  = {10},
  pages   = {1--20},
  doi     = {10.xxxx/xxxxx}
}

@book{book2020,
  author    = {Author, Name},
  title     = {Book Title},
  publisher = {Publisher},
  year      = {2020},
  address   = {City, Country}
}
```

---

## 🎨 Customizing the Template (`template/mdpi_style.tex`)

The file contains the complete LaTeX preamble. You can edit:

- **Fonts**: change `mathpazo` / `helvet` to your preferred font packages
- **Colours**: modify the `\definecolor` definitions
- **Section formatting**: edit the `\titleformat` commands
- **Page margins**: adjust the `geometry` package options
- **Header/footer**: modify the `fancyhdr` settings

---

## 🔧 Troubleshooting

### `ModuleNotFoundError`
```bash
pip install python-docx openpyxl
```

### `pdflatex: command not found`
Install a LaTeX distribution (see Quick Start above).

### Missing LaTeX packages (Windows MiKTeX)
MiKTeX installs packages automatically on first use. Accept all install prompts. For manual install:
```powershell
# Windows PowerShell (Admin)
$pkgs = @("setspace","tcolorbox","pgfplots","booktabs","multirow",
          "fancyhdr","titlesec","enumitem","caption","natbib","mathpazo",
          "hyphen-english","microtype")
foreach ($p in $pkgs) { mpm --install=$p }
```

### Missing LaTeX packages (Linux / macOS)
```bash
tlmgr install setspace tcolorbox pgfplots booktabs multirow \
              fancyhdr titlesec enumitem caption natbib mathpazo \
              microtype
```

### PDF build fails — read the log
```bash
cat main.log | grep "^!" | head -20
```

### Figure not appearing
- Verify the file exists in `figures/` with the exact name (case-sensitive on Linux/macOS)
- Supported formats: `.png`, `.jpg`, `.jpeg`, `.pdf`

---

## 📋 Command Reference

```
generate.py [--pdf] [--word] [--all] [--clean]

  --pdf    Generate PDF via pdflatex (default when no flag given)
  --word   Generate formatted Word document
  --all    Generate both PDF and Word
  --clean  Remove LaTeX auxiliary files after build (.aux, .log, .toc, …)
```

---

## 📁 File outputs

| File | Description |
|---|---|
| `output/output.pdf` | Publication-ready PDF |
| `output/output.docx` | Formatted Word document |
| `main.tex` | Generated LaTeX source (do not edit manually) |
| `references.bib` | Copied BibTeX file (auto-generated) |

> **Do not edit `main.tex` directly** — it is overwritten every time you run `generate.py`. Edit `content/text.docx`, `content/tables.xlsx`, or `content/references.bib.txt` instead.

---

## 🛠️ Dependencies

| Package | Purpose |
|---|---|
| `python-docx` | Read/write Word documents |
| `openpyxl` | Read Excel spreadsheets |
| `pdflatex` | Compile LaTeX to PDF (system package) |

---

## 📄 License

MIT License. See `LICENSE` file.

---

## 🤝 Contributing

Pull requests welcome. Please open an issue first to discuss what you would like to change.

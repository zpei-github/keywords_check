# PDF Keyword Finder

[![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

A PDF keyword search tool with cross-line/cross-page sentence extraction, noise filtering, and multi-format export.

## Features

- **Multi-keyword search** with weighted scoring — each keyword can carry a custom importance score
- **Full-sentence extraction** — returns the enclosing sentence for each match, with configurable context window
- **Cross-line & cross-page awareness** — correctly handles keywords split by PDF line breaks or spanning page boundaries
- **Auto noise detection** — identifies and filters headers, footers, page numbers, and watermarks
- **Keyword-protected filtering** — blocks containing target keywords are never removed, even in edge regions
- **Multi-format export:**
  - **Excel** — results ranked by importance, keywords highlighted in red
  - **PDF** — original PDF with keyword highlights (including cross-line/cross-page)
  - **TXT** — full analysis report with statistics and noise detection details
- **PySide6 GUI** — graphical interface for interactive use

## Installation

```bash
git clone https://github.com/zpei-github/pdf-keyword-finder.git
cd pdf-keyword-finder
pip install -r requirements.txt
```

### Dependencies

| Package | Version |
| ------- | ------- |
| [PyMuPDF](https://github.com/pymupdf/PyMuPDF) | >= 1.23.0 |
| [openpyxl](https://openpyxl.readthedocs.io/) | >= 3.1.0 |
| [PySide6](https://doc.qt.io/qtforpython/) | >= 6.5, < 7.0 |

## Quick Start

### CLI

```python
from pdf_keyword_finder import find_keywords_in_pdf

keywords = {
    "公章": 7,    # keyword → importance score
    "承诺": 9,
    "授权": 6,
    "证明": 9,
    "签字": 3,
}

results = find_keywords_in_pdf(
    pdf_path="document.pdf",
    keywords=keywords,
    context_rich=100,
    front_window=0,
    output_file="report.txt",          # optional
    excel_file="results.xlsx",         # optional
    highlight_pdf="highlighted.pdf",   # optional
    auto_clean_noise=True,
    header_ratio=0.15,
    footer_ratio=0.85,
    repeat_threshold=0.3,
)
```

### GUI

```bash
python gui.py
```

The GUI provides visual controls for file selection, keyword input, noise filter tuning, and result browsing.

## API Reference

### `find_keywords_in_pdf()`

Main entry point for PDF keyword search.

| Parameter | Type | Default | Description |
| --------- | ---- | ------- | ----------- |
| `pdf_path` | `str` | *required* | Path to the PDF file |
| `keywords` | `List[str] \| Dict[str, int]` | *required* | Keywords (list) or keyword→score mapping (dict) |
| `context_rich` | `int` | *required* | Max characters to extend forward/backward for sentence boundaries |
| `front_window` | `int` | *required* | Max characters to look backward from match start (clamped to 80) |
| `output_file` | `str \| None` | `None` | Path for TXT analysis report |
| `excel_file` | `str \| None` | `None` | Path for ranked Excel export |
| `highlight_pdf` | `str \| None` | `None` | Path for highlighted PDF export |
| `auto_clean_noise` | `bool` | `False` | Enable header/footer/page-number filtering |
| `header_ratio` | `float` | `0.15` | Top region ratio for header detection |
| `footer_ratio` | `float` | `0.85` | Bottom region ratio for footer detection |
| `repeat_threshold` | `float` | `0.3` | Minimum repeat rate to classify text as noise |

**Returns** a dict with:

- `total_matches` — number of matched sentences
- `by_page` — results grouped by page number
- `all_results` — flat list of all results
- `noise_info` — details of filtered noise blocks

## How It Works

1. **Text extraction** — PyMuPDF extracts text blocks with positional metadata (coordinates, page number)
2. **Noise detection** — blocks at page edges with high repetition rates or digit patterns are flagged as noise; blocks containing target keywords are always preserved
3. **Keyword matching** — all keywords are combined into a single regex pattern and matched in one pass across the full text
4. **Sentence extraction** — for each match, the tool expands outward to find sentence boundaries (configurable via `context_rich` and `front_window`)
5. **Overlap merging** — overlapping sentences are merged, combining their keyword sets
6. **Page mapping** — character positions are mapped back to page numbers using prefix-sum arrays (O(log n) lookup)
7. **Export** — results are written to the requested output formats

## Performance

- **Single-pass regex matching** — all keywords merged into one compiled pattern
- **List + join** text assembly (O(n)) instead of repeated string concatenation (O(n²))
- **Prefix-sum page lookup** — O(log n) page resolution via binary search
- **Block caching per page** — PDF dict and words data fetched once per page during highlight generation, then shared across fragments

## License

MIT

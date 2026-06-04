# Click2Count

A lightweight desktop tool for counting and categorising items on PDF building plans — or any PDF document. Built with Python and tkinter.

Developed by **Polina February**
Concept & feedback by **HDZ Electrical**

---

## Features

- **Pan & Marker tools** — open a plan, explore it with Pan, then switch to Marker mode to start counting
- **Click to count** — click anywhere on the PDF to place a numbered marker
- **Categories** — create multiple named categories, each with a unique colour; markers from all categories are visible at the same time
- **Ruler tool** — measure distances between two points; set a plan scale (e.g. `1:200`) to get results in metres
- **Zoom** — zoom in/out with buttons or `Ctrl + scroll`; markers stay anchored to their exact position in the PDF
- **Multi-page** — navigate pages with arrow buttons or keyboard; per-category counts run continuously across all pages
- **Undo / Reset** — undo the last marker or clear all markers on the current page (per category)
- **Save / Load session** — save your work as a `.c2c` file and continue later
- **Export** — save a summary as `.xlsx`, `.txt`, or `.json`

---

## Screenshots

![Screenshot of the default screen in Click2Count by HDZ](./Screenshots/default_screen.png)
---

## Installation

**Requirements:** Python 3.10+

```bash
pip install pymupdf Pillow openpyxl
```

> `openpyxl` is optional — the app runs without it but `.xlsx` export will not be available.

**Run:**

```bash
python click2count.py
```

---

## Usage

### Basic workflow

1. Click **☰ → Open PDF** to load a file — the view opens in **Pan mode** automatically
2. Pan around the plan to orient yourself
3. Click **✏ Marker** in the toolbar to switch to counting mode
4. Click anywhere on the plan to place numbered markers
5. Right-click to exit the active tool at any time

### Categories

| Button | Action |
|--------|--------|
| **＋** | Create a new category (choose name and colour) |
| **✎ Rename** | Rename the selected category |
| **🟥** | Change the category colour |

Switch the active category from the dropdown — new clicks go to the selected category. All category markers are shown on screen simultaneously.

### Tools

| Tool | How to activate |
|------|----------------|
| **✏** | Click the Marker button — left-click places markers |
| **✋** | Click the Pan button, or hold `D` + drag (temporary), or middle-click drag |
| **📏** | Click the Ruler button — a second toolbar row appears |
| Exit any tool | **Right-click** |

### Ruler

1. Click **📏** in the toolbar to enter ruler mode (a second toolbar row appears)
2. Click two points on the PDF — a dashed line appears with the distance
3. Set a scale in one of two ways:
   - Type the plan denominator in the **Scale 1:** field (e.g. `200` for a 1:200 plan) and press **Apply** or Enter
   - Click **⚖ Set Scale** and enter the real-world length of the drawn line in metres
4. The measurement updates to metres once a scale is set
5. Click **📏** again or right-click to close ruler mode; click **✕ Clear** to remove the line

### Export

Click **☰ → Export Summary** and choose a format:

| Format | Contents |
|--------|----------|
| `.xlsx` | Spreadsheet with one row per category, one column per page, totals row; category name cells are colour-coded |
| `.txt`  | Plain-text summary with counts per category per page |
| `.json` | Full data including individual click coordinates |

### Keyboard shortcuts

| Key | Action |
|-----|--------|
| `←` / `→` | Previous / next page |
| `z` | Undo last click |
| `+` / `-` | Zoom in / out |
| `D` (hold) | Temporary pan mode |
| `Ctrl + scroll` | Zoom in / out with mouse or trackpad |

---

## Dependencies

| Package | Purpose |
|---------|---------|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | Render PDF pages |
| [Pillow](https://python-pillow.org/) | Draw semi-transparent markers |
| [openpyxl](https://openpyxl.readthedocs.io/) _(optional)_ | Export to `.xlsx` |

---

## Credits

Built by Polina February & Claude (Anthropic)
Concept & feedback by HDZ Electrical

© HDZ Electrical. All rights reserved.

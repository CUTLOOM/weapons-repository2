---
name: shipping-mark-sticker
description: Enlarge shipping mark stickers (唛头贴) from PDFs onto A4 landscape pages at maximum proportional scale. Use when a user provides a PDF with shipping marks/labels that need to be enlarged for printing, preserving all border lines, text, and barcodes at maximum printable size on A4 paper.
metadata:
  short-description: Enlarge shipping mark stickers to A4 landscape
---

# Shipping Mark sticker

Enlarge shipping mark stickers from a PDF onto a single A4 landscape page, with two sticker rows stacked vertically at maximum proportional scale.

## Usage

### Script: `scripts/enlarge_stickers.py`

```bash
python enlarge_stickers.py <input_pdf> [output_pdf]
```

If `output_pdf` is omitted, the script appends `_enlarged.pdf` to the input filename.

### Workflow

1. Locate the input PDF containing shipping mark stickers (two rows in A4 landscape format)
2. Run the script to generate the enlarged PDF on the desktop
3. The output contains both stickers on one A4 landscape page, scaled to maximize size

### Algorithm

1. Render the source PDF page at 600 DPI using `pdfplumber`
2. Crop each sticker row with 5px padding (to preserve border lines)
3. Calculate uniform scale that fits both stickers vertically within an A4 landscape page
4. Place them with 15pt gap, centered, on the output page using `reportlab`

### Default values (modify in script as needed)

| Parameter | Default | Description |
|---|---|---|
| `margin` | 20 pt | Page margin |
| `gap` | 15 pt | Gap between sticker rows |
| `dpi` | 600 | Rendering resolution |
| `pad` | 5 px | Crop padding (preserves borders) |

## Requirements

- Python with `pdfplumber` and `reportlab` packages
- Bundled workspace dependencies (Codex runtime)

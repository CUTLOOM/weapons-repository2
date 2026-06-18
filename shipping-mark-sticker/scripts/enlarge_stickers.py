# -*- coding: utf-8 -*-
"""
唛头贴放大工具 - Enlarge shipping mark stickers to A4 landscape pages.

Usage:
    enlarge_stickers.py <input_pdf_path> [output_pdf_path]

Example:
    enlarge_stickers.py "Shipping_16W3I_ 9.pdf" "C:\Users\W11\Desktop\output.pdf"
"""
import sys, os, io, tempfile, shutil
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "..", "..", ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "python"))

from pdfplumber import open as pdfplumber_open
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

def enlarge_stickers(input_pdf, output_pdf, margin=20, gap=15, dpi=600):
    """
    Enlarge shipping mark stickers from a PDF and arrange on A4 landscape page.
    
    Args:
        input_pdf: Path to source PDF containing sticker rows
        output_pdf: Path for output PDF
        margin: Page margin in pts (default 20)
        gap: Gap between stickers in pts (default 15)
        dpi: Rendering resolution (default 600)
    """
    print(f"Rendering PDF at {dpi} DPI ...")
    with pdfplumber_open(input_pdf) as pdf:
        page = pdf.pages[0]
        W, H = page.width, page.height
        print(f"Source page: {W:.0f}x{H:.0f}")
        im = page.to_image(resolution=dpi)
        pil = im.original.convert("RGB")
    
    SP = dpi / 72.0
    PAD = 5  # extra crop padding in pixels to preserve border lines
    
    # Two sticker rows in pdfplumber landscape coordinates
    # Row1 is the upper sticker, Row2 is the lower sticker
    BB = [
        (43.0, 408.3, 601.4, 550.0),   # upper row
        (43.1, 163.8, 601.5, 305.5),   # lower row
    ]
    
    crops = []
    for x0, y0, x1, y1 in BB:
        px0 = max(0, int(x0 * SP) - PAD)
        py0 = max(0, int((H - y1) * SP) - PAD)
        px1 = min(pil.width,  int(x1 * SP) + PAD)
        py1 = min(pil.height, int((H - y0) * SP) + PAD)
        crops.append({
            "img": pil.crop((px0, py0, px1, py1)),
            "w": x1 - x0 + 2 * PAD * 72.0 / dpi,
            "h": y1 - y0 + 2 * PAD * 72.0 / dpi,
        })
    
    # Layout: both stickers stacked vertically on A4 landscape
    PW, PH = A4[1], A4[0]  # landscape: 841.89 x 595.28
    sw, sh = crops[0]["w"], crops[0]["h"]
    s = min((PW - 2*margin) / sw, (PH - 2*margin - gap) / (2 * sh))
    fw, fh = sw * s, sh * s
    
    oy_b = (PH - 2*fh - gap) / 2   # lower sticker y
    oy_t = oy_b + fh + gap          # upper sticker y
    ox = (PW - fw) / 2              # center horizontally
    
    print(f"Scale: {s:.4f}x, Each: {fw:.1f}x{fh:.1f}")
    
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PW, PH))
    
    # Draw lower sticker (Row2) first, then upper (Row1)
    for idx in [1, 0]:
        tmp = os.path.join(tempfile.gettempdir(), f"_sticker_{idx}.png")
        crops[idx]["img"].save(tmp)
        y = oy_t if idx == 0 else oy_b
        c.drawImage(tmp, ox, y, width=fw, height=fh)
        os.remove(tmp)
    
    c.showPage()
    c.save()
    
    buf.seek(0)
    with open(output_pdf, "wb") as f:
        f.write(buf.getvalue())
    print(f"Saved: {output_pdf}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else \
        os.path.splitext(input_path)[0] + "_enlarged.pdf"
    enlarge_stickers(input_path, output_path)

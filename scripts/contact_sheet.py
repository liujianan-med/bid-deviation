"""把扫描版 PDF 每页渲染成缩略图拼成联系表，用于快速定位章节页。
用法：python scripts/contact_sheet.py <pdf> <输出png> [每张几页]
"""
import sys
import pdfplumber
from PIL import Image, ImageDraw

pdf_path, out_path = sys.argv[1], sys.argv[2]
per_sheet = int(sys.argv[3]) if len(sys.argv) > 3 else 16
cols, rows = 4, per_sheet // 4
thumb_w, thumb_h = 380, 500

with pdfplumber.open(pdf_path) as pdf:
    pages = pdf.pages
    n = len(pages)
    for start in range(0, n, per_sheet):
        chunk = pages[start:start + per_sheet]
        sheet = Image.new("RGB", (cols * thumb_w, rows * thumb_h), "white")
        draw = ImageDraw.Draw(sheet)
        for i, page in enumerate(chunk):
            img = page.to_image(resolution=60).original.convert("RGB")
            img.thumbnail((thumb_w - 16, thumb_h - 30))
            x = (i % cols) * thumb_w + 8
            y = (i // cols) * thumb_h + 22
            sheet.paste(img, (x, y))
            draw.text((x, y - 18), f"P{start + i + 1}", fill="red")
            draw.rectangle([x - 2, y - 2, x + img.width + 2, y + img.height + 2], outline="#ccc")
        fname = out_path.replace(".png", f"_{start // per_sheet + 1}.png")
        sheet.save(fname, quality=85)
        print(fname, sheet.size)

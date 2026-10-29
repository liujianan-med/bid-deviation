# -*- coding: utf-8 -*-
"""估算 docx 首页占用高度（pt），用于判断单页简历是否溢出"""
import sys, math
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

USABLE_W = 510.0   # 21cm - 3cm 边距 = 18cm = 510pt
PAGE_H = 785.0     # 29.7cm - 2cm 边距


def tw(text, size):
    w = 0.0
    for ch in text:
        w += size * (1.0 if ord(ch) > 127 else 0.52)
    return w


def para_h(p):
    text = p.text
    size = 10.0
    for r in p.runs:
        if r.font.size:
            size = max(size if r.text.strip() == "" else 0, r.font.size.pt)
    if not text.strip():
        return (p.paragraph_format.space_before.pt if p.paragraph_format.space_before else 0) + \
               (p.paragraph_format.space_after.pt if p.paragraph_format.space_after else 0)
    line = p.paragraph_format.line_spacing or 1.08
    if not isinstance(line, float):
        line = 1.08
    indent = p.paragraph_format.left_indent.pt if p.paragraph_format.left_indent else 0
    lines = max(1, math.ceil(tw(text, size) / (USABLE_W - indent)))
    before = p.paragraph_format.space_before.pt if p.paragraph_format.space_before else 0
    after = p.paragraph_format.space_after.pt if p.paragraph_format.space_after else 0
    return before + after + lines * size * line


def main(path):
    doc = Document(path)
    total = 0.0
    rows = 0
    for child in doc.element.body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            total += para_h(Paragraph(child, doc))
        elif tag == 'tbl':
            t = Table(child, doc)
            h = 0.0
            for row in t.rows:
                rh = 0.0
                for cell in row.cells:
                    ch = sum(para_h(p) for p in cell.paragraphs)
                    rh = max(rh, ch)
                h += rh + 2  # 单元格上下边距约 5pt 总量的一半估算
            total += h
            rows += len(t.rows)
    print(f"文件: {path}")
    print(f"估算总高: {total:.0f} pt / 可用 {PAGE_H:.0f} pt  (表格 {rows} 行)")
    print("结论:", "✅ 单页内" if total <= PAGE_H else f"❌ 超出 {total - PAGE_H:.0f} pt（约 {math.ceil((total-PAGE_H)/12.5)} 行）")


if __name__ == "__main__":
    main(sys.argv[1])

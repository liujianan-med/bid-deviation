"""读 PDF 书签，定位章节页码。用法：python scripts/toc.py <pdf>"""
import sys
from pypdf import PdfReader

r = PdfReader(sys.argv[1])
print(f"共 {len(r.pages)} 页")


def walk(o, d=0):
    for it in o:
        if isinstance(it, list):
            walk(it, d + 1)
            continue
        try:
            pg = r.get_destination_page_number(it) + 1
        except Exception:
            pg = "?"
        print("  " * d + f"{str(it.title)[:50]} -> 第 {pg} 页")


try:
    walk(r.outline)
except Exception as e:
    print("无书签:", e)

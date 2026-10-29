# -*- coding: utf-8 -*-
"""裁剪 3 张 WPS 截图，去掉浏览器标签栏/工具条/任务栏，归档到 portfolio/screenshots"""
from PIL import Image
import os

SRC = r"C:\Users\Administrator\Pictures\Screenshots"
DST = r"C:\Users\Administrator\code\bid-deviation\portfolio\screenshots"

jobs = [
    ("屏幕截图(4).png", "03-输出-偏离表首页.png"),
    ("屏幕截图(5).png", "04-输出-判定细节.png"),
    ("屏幕截图(6).png", "05-输出-存疑清单.png"),
]

for src, dst in jobs:
    p = os.path.join(SRC, src)
    im = Image.open(p)
    W, H = im.size
    # 只保留 WPS 文档编辑区：去掉顶部浏览器标签/功能区、底部状态栏、右侧图标条
    left = int(W * 0.120)
    right = int(W * 0.938)
    top = int(H * 0.176)
    bottom = int(H * 0.912)
    crop = im.crop((left, top, right, bottom))
    # 放大 1.6 倍，让作品集里看得清
    crop = crop.resize((int(crop.width * 1.6), int(crop.height * 1.6)), Image.LANCZOS)
    out = os.path.join(DST, dst)
    crop.save(out, optimize=True)
    print(f"{dst}  {im.size} -> {crop.size}")

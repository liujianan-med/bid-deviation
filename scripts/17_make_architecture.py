# -*- coding: utf-8 -*-
"""生成作品集用：系统架构图 PNG（5 模块流水线）"""
from PIL import Image, ImageDraw, ImageFont

W, H = 1600, 900
NAVY = (42, 64, 88)
ACCENT = (31, 111, 178)
DARK = (43, 43, 43)
GREY = (110, 118, 128)
BG = (245, 247, 250)
BOX = (255, 255, 255)
GREEN = (46, 125, 50)
ORANGE = (214, 116, 24)

F = r"C:\Windows\Fonts\msyh.ttc"
FB = r"C:\Windows\Fonts\msyhbd.ttc"
f_title = ImageFont.truetype(FB, 40)
f_sub = ImageFont.truetype(F, 21)
f_box_t = ImageFont.truetype(FB, 25)
f_box = ImageFont.truetype(F, 20)
f_small = ImageFont.truetype(F, 17)
f_tag = ImageFont.truetype(FB, 20)

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)

d.rectangle([0, 0, W, 112], fill=NAVY)
d.text((60, 14), "医疗投标文档智能生成系统 · 架构图", font=f_title, fill=(255, 255, 255))
d.text((62, 68), "大模型负责语义理解，规则引擎负责数值判定，缺数据显式留白——把幻觉挡在交付之前",
       font=f_sub, fill=(186, 201, 216))

boxes = [
    ("① 输入解析", ["整本招标 Word 直读", "自动定位设备小节", "支持 .docx 全本"], 60),
    ("② 参数抽取", ["LLM 结构化抽取", "表格 / 图片参数", "JSON + 自动重试"], 350),
    ("③ 三态判定", ["规则引擎判定", "响应 / 正 / 负偏离", "单位归一化"], 640),
    ("④ 文档生成", ["生成可投标 Word", "证据精确到页", "python-docx"], 930),
    ("⑤ 交付打包", ["资料包一键归集", "缺口清单分责任方", "厂家 / 经销商"], 1220),
]
y0, y1 = 150, 430
for title, lines, x in boxes:
    d.rounded_rectangle([x, y0, x + 290, y1], radius=14, fill=BOX, outline=NAVY, width=3)
    d.rectangle([x, y0, x + 290, y0 + 52], fill=NAVY)
    d.text((x + 18, y0 + 11), title, font=f_box_t, fill=(255, 255, 255))
    yy = y0 + 70
    for ln in lines:
        d.text((x + 18, yy), ln, font=f_box, fill=DARK)
        yy += 36
    d.text((x + 18, y1 - 34), "独立模块 · 可替换", font=f_small, fill=GREY)

for x in (350, 640, 930, 1220):
    cy = (y0 + y1) // 2
    d.line([(x - 56, cy), (x - 10, cy)], fill=ACCENT, width=5)
    d.polygon([(x - 10, cy - 9), (x - 10, cy + 9), (x + 4, cy)], fill=ACCENT)

# 评测闭环
d.rounded_rectangle([60, 500, 760, 706], radius=14, fill=BOX, outline=GREEN, width=3)
d.rectangle([60, 500, 760, 552], fill=GREEN)
d.text((80, 510), "效果评测闭环（自建金标准）", font=f_box_t, fill=(255, 255, 255))
rows = [
    ("金标准", "皮秒 31 条 + 光子 18 条，盲标 + 分歧仲裁"),
    ("回归", "可量化条款自动判定，与金标准逐条比对"),
    ("结果", "可测量项 100%（不可量化单列人工确认）"),
]
yy = 570
for k, v in rows:
    d.text((85, yy), k, font=f_box, fill=GREEN)
    d.text((175, yy), v, font=f_box, fill=DARK)
    yy += 42

# 产品库
d.rounded_rectangle([820, 500, 1540, 706], radius=14, fill=BOX, outline=ORANGE, width=3)
d.rectangle([820, 500, 1540, 552], fill=ORANGE)
d.text((840, 510), "产品库：品类是数据，不是代码", font=f_box_t, fill=(255, 255, 255))
rows2 = [
    ("结构", "data/products/ + index.json，唯一入口"),
    ("扩展", "新增品类零改代码（皮秒/光子/磁刺激）"),
    ("原则", "不兜底：缺数据留白，不编造数值"),
]
yy = 570
for k, v in rows2:
    d.text((845, yy), k, font=f_box, fill=ORANGE)
    d.text((935, yy), v, font=f_box, fill=DARK)
    yy += 42

d.rounded_rectangle([60, 760, 1540, 852], radius=14, fill=NAVY)
d.text((105, 792), "真实政采项目实测", font=f_tag, fill=(186, 201, 216))
d.text((330, 786), "单份偏离表 3 天 → 33 秒", font=f_box_t, fill=(255, 255, 255))
d.text((760, 786), "投标资料包 3 天 → 半天", font=f_box_t, fill=(255, 255, 255))
d.text((1160, 786), "端到端跑通", font=f_box_t, fill=(255, 255, 255))

img.save(r"C:\Users\Administrator\code\bid-deviation\portfolio\architecture.png")
print("saved architecture.png")

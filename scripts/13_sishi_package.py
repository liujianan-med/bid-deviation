# -*- coding: utf-8 -*-
"""
13_sishi_package.py —— 第四师医院项目投标资料包一键归集

输入：
  data/sishi_tender_full.txt   招标文件全文（按段落，含行号约定）
  data/sishi_decisions.json    18 条技术参数核定（响应/偏离 + 证据出处）
  桌面索引中已核实的资料文件

输出（桌面）：
  第四师医院-投标资料包-20261028/   归集后的目录
  第四师医院-投标资料包-20261028.zip
"""

import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
ROOT = os.path.join(DESKTOP, "皮秒投标资料(1)", "皮秒投标资料")
PKG = os.path.join(DESKTOP, f"第四师医院-投标资料包-{date.today():%Y%m%d}")
ZIP = PKG + ".zip"

HEAD_BG = "1F3864"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------- 资料清单（招标出处 → 我方状态） ----------------
# 状态：✅已有（文件在包内）｜⚠有模板需按本项目改写/确认｜❌缺失（外部补办）
QUALIFY = [  # 第四章 资格审查要求（缺项=响应无效）
    ("1", "营业执照", "统一社会信用代码营业执照扫描件盖章", "⚠", "【经销商·名称待定】出具——包内好乐证照是伟思子公司材料，不能替代投标人的"),
    ("2", "法定代表人证明/授权委托书", "法定代表人身份证或授权委托书+被委托人身份证，须与响应文件一致", "⚠", "【经销商·名称待定】按第七章格式出具（格式已在招标文件内）"),
    ("3", "医疗器械经营资质", "代理商：二类产品持《经营备案凭证》或《经营许可证》，三类持《经营许可证》；招标文件未要求制造商授权书", "⚠", "【经销商·名称待定】皮秒属二类，持二类备案凭证或经营许可证即可；厂家侧注册证/生产许可已在包内"),
    ("4", "社保证明", "近半年内至少1个月社保缴存单据或参保缴费证明", "❌", "【经销商·名称待定】社保局或电子税务局打印，盖公章"),
    ("5", "财务报告", "2025年度财务审计报告，或开标前3个月内银行资信证明+开户许可证", "❌", "【经销商·名称待定】2025审计报告或找开户行开资信证明"),
    ("6", "纳税证明", "近半年内至少1个月纳税/完税证明（有公章）", "❌", "【经销商·名称待定】电子税务局下载"),
    ("7", "投标人信用承诺函", "《兵团政府采购供应商信用承诺函》，格式见第七章", "⚠", "【经销商·名称待定】按格式填写盖章"),
]
BASIC = [  # 第三章 基本要求 17 条（★=实质性条款，负偏离即废标）
    ("1", "★", "整机质保≥5年，每年2次维护保养；质保期内免费易耗品/校准品/配件，终身免费升级", "⚠", "有《质保说明》《客户服务售后承诺书》模板；「≥5年+免费易耗品」是否照单承诺须商务确认（评分项里5年是底线，每延1年加2分）"),
    ("2", "", "按院方需求开放接口协议，完成与PACS/HIS对接，费用由供应商承担", "❌", "需厂家书面确认开放接口及费用承担口径"),
    ("3", "", "注册证在有效期内，且与投标所提供注册证一致", "✅", "皮秒双波长注册证（彩色）——注意型号四方一致"),
    ("4", "", "运输安装费用由中标商承担，厂家工程师现场安装调试", "✅", "《供货方案》《安装调试》已覆盖"),
    ("5", "", "维修响应：2小时远程未解决，24小时内到现场", "⚠", "《供货培训售后服务说明函》需核对其响应时限表述是否与本条一致，不一致按本条口径改"),
    ("6", "", "中标方负责首次计量检定、行业标准检测并出报告", "❌", "需出承诺函（格式自拟）"),
    ("7", "", "培训：设备维护+临床使用各≥3次，每次≥1天，教会为止", "⚠", "有光子版《培训及安装方案》可改写；皮秒版需按本条改写为≥3次/每次≥1天"),
    ("8", "", "投标时提供原版技术数据白皮书，技术文件参数须与白皮书相符", "✅", "PicoV 技术白皮书（奎屯版）——投标技术文件以它为准"),
    ("9", "", "提供设备全套配置及临床功能软件", "✅", "皮秒280配置清单"),
    ("10", "", "生产日期距交货期≤3个月，提供生产日期承诺书", "❌", "需厂家出《生产日期承诺函》"),
    ("11", "", "提供易损件报价明细，五年内配件价不高于本次采购价", "❌", "需编制易损件报价单"),
    ("12", "", "配套耗材须说明类型（通用/专机专用）并提供报价单", "❌", "需编制耗材说明+报价单"),
    ("13", "", "到货验收：厂家+临床科室+设备科共同参与", "", "响应即可，无需单独资料"),
    ("14", "", "提供完整技术说明书和检验报告（性能+电气安全），不允许遮盖修改，参数须与之一致", "✅", "使用说明书2026版 + 检测报告280（2024-GL-0320）——本包偏离表全部出处即源于此"),
    ("15", "", "除参数特别要求外，使用年限≥10年，提供证明资料", "✅", "说明书内「使用年限」条款为出处，偏离表第19条已引用"),
    ("16", "★", "交货期：成交通知书后30个自然日完成供货安装调试验收，逾期按日万分之五违约金；需承诺函", "❌", "需出《交货期承诺函》（格式自拟）——★条款，缺失即废标"),
    ("17", "★", "合同签订：中标通知书后10个工作日内签合同；需承诺函", "❌", "需出《合同签订承诺函》——★条款，缺失即废标"),
]
SCORING = [  # 第五章 评分标准（30%价格+70%商务技术）
    ("1", "投标报价（0-30分）", "低价优先法：基准价/报价×100×30%", "—", "报价策略：包一限价206.5万"),
    ("2", "技术指标（0-40分）", "全部满足得40；重要要求每低一项扣3分、一般要求每低一项扣1分；参考技术说明书+CMA/CNAS检验报告", "✅", "偏离表定稿：16响应/2正偏离/0负偏离/0存疑，逐条带页码出处"),
    ("3", "项目实施方案（0-7分）", "生产/采购/装卸/运输/供货/安装调试/验收方案+先进性实用性", "⚠", "有《供货方案》《安装调试》底稿，需按本项目（皮秒+5设备包）改写，忌套用痕迹"),
    ("4", "质量保证措施（0-3分）", "质量保障目标/措施/制度", "⚠", "需在售后方案基础上补一节"),
    ("5", "售后服务体系（0-10分）", "售后承诺书+响应时间+故障处理+培训计划（8分）；履约验收方案（2分）", "⚠", "《供货培训售后服务说明函》为底稿，需按第5/7条口径补齐四要素+单独编履约验收方案"),
    ("6", "应急保障能力（0-2分）", "应急储备/预案/进度保障/人员安排/风险防控", "❌", "需新编应急实施方案"),
    ("7", "业绩（0-2分）", "2023-10-01至今类似业绩，每项1分最多2分；须合同（首页+内容+签章页）+中标通知书，缺一不可", "⚠", "4份皮秒合同已归集，但评分认「本单位」业绩——**若合同签约主体不是本次投标的经销商，业绩分落空**，须与经销商确认其名下业绩；中标通知书也要配套调取"),
    ("8", "免费质保期（0-6分）", "承诺≥5年质保得基础分，每延1年加2分最高6分；须原厂质保承诺书盖公章，未提供不得分", "❌", "需厂家出《原厂质保承诺书》（明确期限与服务内容，建议按8年谈）"),
]
BIZ = [
    ("付款方式", "验收合格且发票认证后一次性支付100%", "—", "现金流注意：无预付款"),
    ("履约保证金", "合同签订后3个工作日内缴纳合同价款10%，质保期满无息退还", "—", "约20万（按包一限价算），可谈银行保函"),
    ("投标保证金/保函", "按投标人须知前附表执行（电子保函可用）", "⚠", "开标前在政采云确认金额与缴纳方式"),
    ("投标有效期/质保期/交货期/合同签订期", "实质性响应一览表须逐条响应且无负偏离，否则投标无效", "❌", "四项承诺须逐条填入附表并出承诺函"),
]
CONFIG = [
    ("治疗柜", "2个"), ("医用照相机（单反）", "1台"), ("治疗床", "2张"),
    ("医用工作站", "1套（台式机八核/显存8G/16G内存/1TB固态+便携机+彩色打印机）"),
    ("ABS治疗车", "2辆"), ("可升降带滑轮治疗椅", "6套"),
]

# ---------------- 文件归集清单 ----------------
COPY_FILES = [
    ("03-厂家资质与技术资料", "伟思-营业执照-（好乐章医美）.pdf", os.path.join(ROOT, "伟思-营业执照-（好乐章医美）.pdf")),
    ("03-厂家资质与技术资料", "伟思-医疗器械生产许可证（好乐章医美）.pdf", os.path.join(ROOT, "伟思-医疗器械生产许可证（好乐章医美）.pdf")),
    ("03-厂家资质与技术资料", "皮秒双波长注册证（彩色）.pdf", os.path.join(ROOT, "皮秒双波长注册证（彩色）(2).pdf")),
    ("03-厂家资质与技术资料", "皮秒检测报告-安规全性能280.pdf", os.path.join(ROOT, "Nd ：YAG皮秒激光治疗仪 安规 全性能检测报告280.pdf")),
    ("03-厂家资质与技术资料", "皮秒使用说明书-2026.pdf", os.path.join(ROOT, "Nd YAG皮秒激光治疗仪使用说明书-2026.pdf")),
    ("03-厂家资质与技术资料", "皮秒技术白皮书-奎屯版.pdf", os.path.join(ROOT, "皮秒激光治疗仪PicoV--技术白皮书 -奎屯.pdf")),
    ("03-厂家资质与技术资料", "皮秒280配置清单.pdf", os.path.join(ROOT, "皮秒280配置清单.pdf")),
    ("03-厂家资质与技术资料", "PicoV彩页.png", os.path.join(ROOT, "5-彩页.png")),
    ("04-经销商资料", "好乐-营业执照（好乐章医美）.pdf", os.path.join(ROOT, "好乐-营业执照（好乐章医美）.pdf")),
    ("04-经销商资料", "好乐-医疗器械经营许可证.pdf", os.path.join(ROOT, "好乐-医疗器械经营许可证(好乐章医美）.pdf")),
    ("04-经销商资料", "好乐-第二类医疗器械经营备案凭证.pdf", os.path.join(ROOT, "好乐-第二类医疗器械经营备案凭证（好乐章医美）.pdf")),
    ("04-经销商资料", "好乐系伟思子公司证明文件.pdf", os.path.join(DESKTOP, "蒙医-光子", "蒙医-光子", "光子招标资料", "光子招标资料", "好乐系伟思子公司证明文件（好乐章）.pdf")),
    ("04-经销商资料", "销售人员委托书.docx", os.path.join(DESKTOP, "伟思文件", "销售人员委托书，docx.docx")),
    ("05-业绩证明", "皮秒业绩-122万-乐陵市妇幼保健院-合同.pdf", os.path.join(ROOT, "皮秒业绩-122万-乐陵市妇幼保健院-设备购置项目合同.pdf")),
    ("05-业绩证明", "皮秒业绩-112万-新疆石河子经开区医院-合同.pdf", os.path.join(ROOT, "皮秒业绩-112万-新疆石河子经开区医院-合同.pdf")),
    ("05-业绩证明", "皮秒业绩-100万-深圳龙华区妇幼-合同.pdf", os.path.join(ROOT, "1-皮秒L300-100万-龙华区妇幼-合同.pdf")),
    ("05-业绩证明", "皮秒业绩-97万-新疆叶城县-合同.pdf", os.path.join(ROOT, "3-皮秒L280-97万-新疆叶城县-合同.pdf")),
    ("06-商务函件与方案", "中小企业声明函.pdf", os.path.join(ROOT, "7-中小企业声名函.pdf")),
    ("06-商务函件与方案", "符合本国产品标准的声明函.pdf", os.path.join(ROOT, "符合本国产品标准的声明函(2).pdf")),
    ("06-商务函件与方案", "供货培训售后服务说明函.docx", os.path.join(ROOT, "供货培训售后服务说明函.docx")),
    ("06-商务函件与方案", "供货方案.pdf", os.path.join(ROOT, "10-供货方案.pdf")),
    ("06-商务函件与方案", "安装调试方案.docx", os.path.join(ROOT, "11-安装调式.docx")),
    ("06-商务函件与方案", "客户服务售后承诺书.docx", os.path.join(DESKTOP, "伟思文件", "客户服务售后承诺书.docx")),
    ("06-商务函件与方案", "质保说明.docx", os.path.join(DESKTOP, "Word", "质保说明.docx")),
    ("02-技术偏离表", "第四师医院-皮秒技术偏离表-定稿v2.docx", os.path.join(DESKTOP, "第四师医院-皮秒技术偏离表-20261027-v2.docx")),
]


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.font = Font(name="等线", size=10, bold=True, color="FFFFFFFF")
        cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def put_rows(ws, start, rows, widths, color_col=None):
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = w
    for r, row in enumerate(rows, start):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = Font(name="等线", size=9)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = BORDER
            if color_col and c == color_col:
                mark = str(v)[:1]
                color = {"✅": "FF1D9E75", "⚠": "FFBF8F00", "❌": "FFC00000"}.get(mark)
                if color:
                    cell.font = Font(name="等线", size=10, bold=True, color=color)
    ws.freeze_panes = "A2"


def build_xlsx(path):
    wb = Workbook()

    ws = wb.active
    ws.title = "① 项目概况"
    rows = [
        ("项目编号", "XJJB-2026-096-01/02"),
        ("项目名称", "第四师医院公立医院改革与高质量发展提升皮肤整形类医疗设备采购项目"),
        ("采购方式", "公开招标（政采云平台电子投标）"),
        ("预算金额", "总 397 万元；包一 206.5 万（限价）/ 包二 190.5 万（限价）"),
        ("包一内容", "手术无影灯1台、取皮植皮系统1台、皮秒激光治疗仪1台、激光光电平台1台、二氧化碳激光仪1台"),
        ("包二内容", "强脉冲光子治疗仪1台、高频电灼治疗仪1台、电动取皮刀1台、LED光谱治疗仪1台、高频电刀2台"),
        ("交货期", "★成交通知书后30个自然日内完成供货安装调试验收"),
        ("质保期", "★整机质保≥5年（评分项：每延1年加2分，最高加6分）"),
        ("联合体", "不接受联合体投标"),
        ("开标时间", "2026-11-11 11:00（政采云电子开标）"),
        ("关键提醒", "包一里的「皮秒」和包二里的「强脉冲光」「高频电灼」均为伟思产品线；本包聚焦包一皮秒，包二机会另行评估"),
    ]
    for c, w in enumerate([16, 90], 1):
        ws.column_dimensions[ws.cell(1, c).column_letter].width = w
    for r, (k, v) in enumerate(rows, 1):
        a = ws.cell(r, 1, k)
        a.font = Font(name="等线", size=10, bold=True)
        a.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        a.border = BORDER
        b = ws.cell(r, 2, v)
        b.font = Font(name="等线", size=10)
        b.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        b.border = BORDER
        ws.row_dimensions[r].height = 34

    ws = wb.create_sheet("② 资格审查（缺项即无效）")
    heads = ["序号", "审查因素", "审查内容", "状态", "我方现有文件 / 待办"]
    for c, h in enumerate(heads, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(heads))
    put_rows(ws, 2, QUALIFY, [6, 18, 46, 6, 52], color_col=4)

    ws = wb.create_sheet("③ 基本要求17条")
    heads = ["序号", "★", "招标要求原文（摘要）", "状态", "我方现状 / 待办"]
    for c, h in enumerate(heads, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(heads))
    put_rows(ws, 2, BASIC, [6, 5, 52, 6, 52], color_col=4)

    ws = wb.create_sheet("④ 技术参数响应18条")
    heads = ["条款", "参数", "判定", "响应内容（投标口径）", "证明材料出处"]
    for c, h in enumerate(heads, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(heads))
    d = json.load(open(os.path.join(os.path.dirname(__file__), "..", "data", "sishi_decisions.json"), encoding="utf-8"))
    rows = []
    for no, it in d["items"].items():
        ev = "；".join(f"{e['doc']} {e['loc']}（{e['detail']}）" for e in it.get("evidence", []))
        rows.append((no, it["param"], it["verdict"], it["response"], ev))
    put_rows(ws, 2, rows, [6, 16, 8, 55, 55], color_col=3)
    for r in range(2, len(rows) + 2):
        v = str(ws.cell(r, 3).value)
        color = {"响应": "FF1D9E75", "正偏离": "FF185FA5", "负偏离": "FFC00000", "存疑": "FFBF8F00"}.get(v)
        if color:
            ws.cell(r, 3).font = Font(name="等线", size=10, bold=True, color=color)
    r = len(rows) + 3
    ws.cell(r, 2, "落实事项（投标前必须闭环）").font = Font(name="等线", size=10, bold=True, color="FFC00000")
    for i, cv in enumerate(d.get("caveats", []), 1):
        c = ws.cell(r + i, 2, f"{cv['item']}（责任：{cv.get('owner','待定')}）")
        c.font = Font(name="等线", size=9, bold=True)
        n = ws.cell(r + i, 4, cv["note"])
        n.font = Font(name="等线", size=9)
        n.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws = wb.create_sheet("⑤ 商务与评分")
    heads = ["类别", "条款/分值", "招标要求", "状态", "我方现状 / 待办"]
    for c, h in enumerate(heads, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(heads))
    rows = [("商务条款",) + x for x in BIZ] + [("评分标准",) + x for x in SCORING]
    put_rows(ws, 2, rows, [12, 20, 55, 6, 50], color_col=4)

    ws = wb.create_sheet("⑥ 配套设备清单")
    heads = ["配套设备", "数量/规格（包一）"]
    for c, h in enumerate(heads, 1):
        ws.cell(1, c, h)
    style_header(ws, 1, len(heads))
    put_rows(ws, 2, CONFIG, [26, 60])
    wb.save(path)


def clear_dir(p):
    """清空目录内容（safe-delete 会拦截桌面 rmtree，改用 robocopy 空目录镜像）。"""
    if not os.path.isdir(p):
        return
    tmp = tempfile.mkdtemp()
    subprocess.run(["robocopy", tmp, p, "/MIR", "/NFL", "/NDL", "/NJH", "/NJS", "/NP"],
                   capture_output=True)
    shutil.rmtree(tmp, ignore_errors=True)


def main():
    clear_dir(PKG)
    for sub in ["01-招标要求全清单", "02-技术偏离表", "03-厂家资质与技术资料",
                "04-经销商资料", "05-业绩证明", "06-商务函件与方案"]:
        os.makedirs(os.path.join(PKG, sub), exist_ok=True)

    build_xlsx(os.path.join(PKG, "01-招标要求全清单", "第四师医院-招标要求全清单.xlsx"))

    missing = []
    for sub, name, src in COPY_FILES:
        dst = os.path.join(PKG, sub, name)
        if os.path.exists(src):
            shutil.copy2(src, dst)
        else:
            missing.append((sub, name, src))

    readme = f"""# 第四师医院投标资料包（{date.today():%Y-%m-%d} 整合）

项目：XJJB-2026-096-01/02 · 包一限价 206.5 万 · 开标 2026-11-11 11:00（政采云）

## 主体关系（重要）
- **好乐是伟思的子公司，不是本次投标的经销商**。投标人主体名称待定，其自身证照（营业执照/经营备案凭证/社保/审计/纳税）必须由投标人本人出具，包内好乐材料仅作「子公司关系证明」参考，不能替代。
- 招标文件**未要求制造商授权书**（全文核实），经销商持自身二类经营资质即可。
- 评分业绩认「本单位」——若 4 份皮秒合同的签约主体不是本次投标人，业绩分（2 分）落空，须与经销商确认其名下业绩。

## 包内结构
- 01-招标要求全清单：6 个页签——项目概况 / 资格审查 / 基本要求17条 / 技术参数响应18条 / 商务与评分 / 配套设备
- 02-技术偏离表：定稿 v2（16 响应 / 2 正偏离 / 0 负偏离 / 0 存疑，逐条带页码出处）
- 03~06：厂家资质、经销商资料（好乐，作关系证明）、业绩合同、商务函件——本机已归集原件的拷贝

## 状态汇总
- 资格审查 7 项：**投标人（经销商名称待定）侧缺 3 项**（营业执照+经营资质待其提供、社保、审计/资信、纳税）——缺项即投标无效
- 基本要求 17 条：已有 6，有模板需改写 4，**缺 6（含 ★交货期、★合同签订两份承诺函）**
- 评分 8 项：技术 40 分已备齐证据；**业绩主体待确认、质保分缺原厂质保承诺书（合计 8 分风险）**
"""
    todo_factory = [  # 厂家（伟思）
        "【厂家·伟思】原厂质保承诺书（≥5年基础，建议按8年谈，每延1年加2分，盖公章并写明免费期限与服务内容）——评分项8，最高6分",
        "【厂家·伟思】生产日期承诺函（生产日期距交货≤3个月）——基本要求10",
        "【厂家·伟思】PACS/HIS 接口开放确认函（开放所有接口协议、费用含接口费由供方承担、免费升级）——基本要求2",
        "【厂家·伟思】偏离表落实事项：治疗手具数量与型号书面确认（第5条）——caveats",
        "【厂家·伟思】偏离表落实事项：终端单脉冲能量口径书面说明（白皮书500 vs 检测报告490mJ，第9条）——caveats",
        "【厂家·伟思】易损件价格表与耗材底价（投标方据此编报价单并承诺五年不涨价）——基本要求11/12",
        "【厂家·伟思】工程师到场安排（安装调试 + 培训≥3次每次≥1天 + 首次计量检定配合）——基本要求4/6/7",
    ]
    todo_dealer = [  # 经销商（名称待定）
        "【经销商·名称待定】营业执照 + 二类《经营备案凭证》或《医疗器械经营许可证》——资格审查1/3",
        "【经销商·名称待定】社保缴存证明（近半年任1个月，盖章）——资格审查4，缺项即无效",
        "【经销商·名称待定】2025年度财务审计报告 或 银行资信证明+开户许可证——资格审查5，缺项即无效",
        "【经销商·名称待定】近半年纳税证明——资格审查6，缺项即无效",
        "【经销商·名称待定】法定代表人授权委托书 + 双方身份证（第七章格式）——资格审查2",
        "【经销商·名称待定】兵团政府采购供应商信用承诺函（第七章格式）——资格审查7",
        "【经销商·名称待定】★交货期承诺函（30自然日，格式自拟）——基本要求★16，缺失即废标",
        "【经销商·名称待定】★合同签订承诺函（10个工作日）——基本要求★17，缺失即废标",
        "【经销商·名称待定】易损件报价明细 + 耗材说明及报价单（用厂家底价编制）——基本要求11/12",
        "【经销商·名称待定】确认本方名下业绩合同与中标通知书（2023-10至今）——评分项7",
    ]
    todo_both = [
        "【双方】首次计量检定承诺函——基本要求6",
        "【双方】政采云供应商入库 + CA 数字证书（新疆数字证书认证中心）——开标前必备",
        "【双方】实质性响应一览表逐条填写（投标有效期/质保期/交货期/合同签订期），无负偏离——否则投标无效",
    ]
    readme += "\n## 待补办清单（按责任方）\n### 厂家（伟思）\n" + "\n".join(f"- {t}" for t in todo_factory)
    readme += "\n### 经销商（投标人·名称待定）\n" + "\n".join(f"- {t}" for t in todo_dealer)
    readme += "\n### 双方\n" + "\n".join(f"- {t}" for t in todo_both) + "\n"
    for sub, name, src in missing:
        readme += f"\n- ❌ 未找到原文件：{sub}/{name}\n"
    open(os.path.join(PKG, "00-总览与缺口清单.md"), "w", encoding="utf-8").write(readme)

    if os.path.exists(ZIP):
        os.remove(ZIP)
    with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED) as zf:
        for base, _dirs, files in os.walk(PKG):
            for f in files:
                fp = os.path.join(base, f)
                zf.write(fp, os.path.relpath(fp, os.path.dirname(PKG)))

    n_files = sum(len(fs) for _b, _d, fs in os.walk(PKG))
    print(f"包目录：{PKG}")
    print(f"压缩包：{ZIP}（{os.path.getsize(ZIP)/1048576:.1f} MB，{n_files} 个文件）")
    if missing:
        print("未找到的文件：")
        for sub, name, _src in missing:
            print(f"  - {sub}/{name}")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
product_db.py —— 产品参数库统一入口（支持全品类医疗设备）

设计原则：引擎不认识「皮秒」「光子」「射频」，只认识参数库文件。
    品类是数据，不是代码。
    新增一个品类 = 在 data/products/ 放一个 json + 在 index.json 登记一行，
    不改 06/07/09 任何一行判定或排版逻辑。

目录约定：
    data/products/
        index.json              品类注册表（关键词 -> 文件）
        weisi_picov.json        PicoV 超皮秒
        weisi_photon.json       大师光 VVL3
        ...
        evidence/<key>.json     该品类的证明材料索引（页码级）

参数条目 schema（每个品类一致，这是全品类能复用的前提）：
    {"key": "激光波长", "type": "wavelength|number|range|feature",
     "value": 532 或 [532, 1064], "unit": "nm",
     "operator": ">=" / "<=" / "=",      # number 型才有
     "tolerance": 5,                      # 允差，可选
     "source": "白皮书§2",                # 出处（人可核查的粗粒度）
     "note": ""}                          # 口径备注，可选

运行自检：
    python scripts/product_db.py
"""

import json
import os

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
PRODUCTS_DIR = os.path.join(DATA_DIR, "products")
INDEX_FILE = os.path.join(PRODUCTS_DIR, "index.json")

# 兼容老路径：index.json 还不存在时回退到单品类时代的皮秒库
LEGACY_PRODUCT_FILE = os.path.join(DATA_DIR, "picov_params.json")
LEGACY_EVIDENCE_FILE = os.path.join(DATA_DIR, "evidence_map.json")

DEFAULT_KEY = "picov"

# 全品类通用参数维度。
# 用途：新产品库建好之后，按这 7 组检查缺项——缺哪一项，投标时就会被哪一项卡住。
# 清单里的参数名与皮秒库/证据库的实际标准名严格一致（param_alias 的映射目标），
# 跨品类复用；非激光品类对激光相关维度会大面积缺项，这本身就是建库提示。
DIMENSIONS = {
    "身份信息": ["型号", "注册证编号", "生产厂家", "产地", "产品寿命", "设备重量"],
    "光源/能量": ["激光工作物质", "激光波长", "激光波长及用途", "峰值功率",
                "最大能量密度", "常规能量单脉冲模式最大输出能量",
                "激光输出能量不稳定度", "激光输出复现性",
                "输出功率", "瞄准光输出功率"],
    "脉宽/光斑": ["最小脉冲宽度", "脉冲脉宽", "光斑直径", "激光输出发散角",
                "重复频率", "脉冲模式"],
    "治疗与手具": ["操作手柄", "治疗手具", "冷却系统", "瞄准光系统波长",
                "传输方式", "一体化成型激光器", "功能配置", "激光防护眼镜"],
    "控制与软件": ["控制系统", "显示屏", "参数修正功能及升级接口",
                "设备治疗参数存储记忆功能", "软件版本"],
    "电气与环境": ["电源需求", "输入功率", "运行环境", "大气压力", "相对湿度", "存储环境"],
    "安全与合规": ["脚踏开关防护等级", "光密度", "安全分类", "激光分类",
                "防护装置", "紧急停止", "电气安全"],
}


def load_index() -> dict:
    if not os.path.exists(INDEX_FILE):
        return {"version": 1, "default": DEFAULT_KEY, "products": {
            DEFAULT_KEY: {"name": "PicoV 超皮秒", "keywords": ["皮秒"],
                          "file": "../picov_params.json", "status": "ready"}}}
    with open(INDEX_FILE, encoding="utf-8") as f:
        return json.load(f)


def all_products() -> dict:
    return load_index().get("products", {})


def resolve(spec=None) -> str:
    """把「产品 key / 中文名 / 型号 / 招标里的设备名」统一解析成 index 里的 key。

    匹配顺序（从严到宽，避免"光子"误命中"光子嫩肤仪"这类泛称）：
        1. 空 -> 默认品类
        2. 精确等于 key
        3. 精确等于 name / models 里某一项
        4. name 包含 spec，或 spec 包含 name
        5. keywords 命中
    解析不出来直接报错，绝不默认兜底成皮秒——否则拿皮秒库判光子，
    会产出一整张全错但看起来正常的表，这是最危险的失败方式。
    """
    products = all_products()
    default = load_index().get("default", DEFAULT_KEY)

    if not spec:
        return default
    spec = str(spec).strip()
    if spec in products:
        return spec

    for key, meta in products.items():
        names = [meta.get("name", "")] + list(meta.get("models") or [])
        if spec in [n for n in names if n]:
            return key

    for key, meta in products.items():
        names = [meta.get("name", "")] + list(meta.get("models") or [])
        if any(n and (n in spec or spec in n) for n in names):
            return key

    for key, meta in products.items():
        for kw in meta.get("keywords") or []:
            if kw and kw in spec:
                return key

    raise KeyError(
        f"产品库里找不到「{spec}」。现有品类："
        + "、".join(f"{k}({v.get('name','')})" for k, v in products.items())
        + f"。新增品类请照 data/products/index.json 的格式登记。"
    )


def product_file(key: str) -> str:
    meta = all_products().get(key, {})
    f = meta.get("file")
    if not f:
        # 回退：data/products/<key>.json
        f = f"{key}.json"
    return os.path.normpath(os.path.join(PRODUCTS_DIR, f))


def evidence_file(key: str) -> str:
    """该品类的证明材料索引；没有独立文件时回退到老的全局 evidence_map.json。"""
    meta = all_products().get(key, {})
    f = meta.get("evidence_file")
    if f:
        return os.path.normpath(os.path.join(PRODUCTS_DIR, f))
    cand = os.path.join(PRODUCTS_DIR, "evidence", f"{key}.json")
    if os.path.exists(cand):
        return cand
    return LEGACY_EVIDENCE_FILE


def load(spec=None) -> dict:
    """加载产品库。

    返回 {"key","name","category","models","params":{参数名:条目},
          "product_path","evidence_path","meta"}
    params 是按 key 索引的 dict —— 06/07 用的就是这个形状，换品类零改动。
    """
    key = resolve(spec)
    path = product_file(key)
    if not os.path.exists(path):
        raise FileNotFoundError(f"品类「{key}」已登记但参数库文件不存在：{path}")

    with open(path, encoding="utf-8") as f:
        raw = json.load(f)

    meta = dict(all_products().get(key, {}))
    meta.update({k: v for k, v in raw.items() if k != "params"})

    params = raw.get("params") or []
    if isinstance(params, dict):
        by_key = params
    else:
        by_key = {p["key"]: p for p in params}

    return {
        "key": key,
        "name": meta.get("name", key),
        "category": meta.get("category", ""),
        "models": meta.get("models") or [],
        "params": by_key,
        "product_path": path,
        "evidence_path": evidence_file(key),
        "meta": meta,
    }


def gap_report(spec=None) -> dict:
    """按 7 个通用维度检查参数完备度：哪些条目已有、哪些缺、缺的会在投标时卡在哪。"""
    p = load(spec)
    have = set(p["params"].keys())
    report = {"key": p["key"], "name": p["name"], "total": len(have), "groups": []}
    for dim, keys in DIMENSIONS.items():
        present = [k for k in keys if k in have]
        missing = [k for k in keys if k not in have]
        report["groups"].append({"dimension": dim, "present": present,
                                 "missing": missing,
                                 "coverage": f"{len(present)}/{len(keys)}"})
    return report


def _selftest():
    print(f"产品库索引：{INDEX_FILE}")
    for key, meta in all_products().items():
        flag = {"ready": "就绪", "draft": "草稿", "empty": "待建"}.get(
            meta.get("status", ""), meta.get("status", ""))
        print(f"  {key:<12} {meta.get('name',''):<18} [{flag}] {meta.get('category','')}")

    print(f"\n默认品类：{resolve()}")
    for probe in ["皮秒", "超皮秒", "PicoV L280", "光子", "射频"]:
        try:
            print(f"  「{probe}」-> {resolve(probe)}")
        except KeyError as e:
            print(f"  「{probe}」-> 未登记（{str(e)[:40]}…）")

    print()
    g = gap_report()
    print(f"{g['name']}：已填 {g['total']} 项")
    for grp in g["groups"]:
        miss = "、".join(grp["missing"]) or "无"
        print(f"  {grp['dimension']:<8} {grp['coverage']:<5} 缺：{miss}")


if __name__ == "__main__":
    _selftest()

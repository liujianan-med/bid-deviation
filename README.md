# 招标参数自动偏离表生成器

把招标文件里的技术要求逐条抽出来，与自家产品参数比对，自动判定「响应 / 正偏离 / 负偏离」，
输出可直接编辑的偏离表。

**动机**：一份 30-80 项参数的偏离表，人工做要 4 小时，且容易漏项、判错方向。
这个工具把它压到分钟级，并且把不确定的条目单独列出来给人复核。

---

## 当前进度

- [x] 步骤 1：读取招标参数 Excel，结构化输出 JSON
- [x] 步骤 1：调通大模型 API，单条参数结构化抽取
- [x] 步骤 2：批量抽取全部参数 + 建立评测集（皮秒 31 条 + 光子 18 条，盲标 + 分歧仲裁）
- [ ] 步骤 3：与自家产品参数表做语义比对，输出三态判定
- [ ] 步骤 4：生成 Word 偏离表 + 存疑条目清单
- [ ] 步骤 5：Docker 打包 + FastAPI 接口
- [ ] 步骤 6：用 8 份历史标书做回归验证，统计准确率

## 评测结果

| 样本集 | 条数 | operator 准确率 | is_measurable 准确率 | 错例 |
|---|---|---|---|---|
| 皮秒（同源样本） | 31 | 100% (29/29) | 100% (31/31) | 0 |
| 光子（困难样本） | 18 | **94.4% (17/18)** | 100% (18/18) | 1 |

**唯一错例（光子行 17）**：「最大支持10个子脉冲」——模型输出 `>=`，正确为 `<=`，
比较方向完全反转。困难样本（表达方式不同的另一份招标文件）成功暴露了同源样本上
发现不了的问题。

**评测方法**：只评 operator 和 is_measurable 两个决策关键字段（param_name 写法不影响
投标判定）；人工盲标与模型推理隔离；5 处人工-模型分歧逐条仲裁，仲裁原则为「以招标
真实意图为准，不以文字表面为准」；标注内部矛盾（如可量化=是但 operator=none）按标注
错误修正。评测对齐业务目标，不对齐模型输出格式。

---

## 目录结构

```
bid-deviation/
├── data/                       # 样本数据（真实标书，不提交到 Git）
│   ├── 招标参数模版-皮秒治疗仪.xlsx
│   ├── 2.皮秒NdYAG激光治疗仪偏离表.docx
│   ├── params.json             # 步骤 1 产出：结构化参数
│   └── llm_extract.json        # 步骤 1 产出：大模型抽取结果
├── scripts/
│   ├── 01_read_params.py       # 读 Excel -> JSON
│   └── 02_call_llm.py          # 调大模型做单条抽取
├── .venv/                      # 虚拟环境（不提交）
├── .env                        # API Key（不提交）
├── .env.example                # 配置模板
└── .gitignore
```

---

## 怎么跑

### 第一次准备（只需一次）

```bash
# 进入项目
cd C:\Users\Administrator\code\bid-deviation

# 依赖已经装好了，如果换了机器要重装：
.venv\Scripts\pip.exe install -r requirements.txt
```

### 脚本 1：读参数表

```bash
.venv\Scripts\python.exe scripts\01_read_params.py
```

输出：参数条数、★/▲ 数量、可量化比例、编号重复警告，结果存 `data/params.json`。

### 脚本 2：调大模型

```bash
# 先配 Key
copy .env.example .env
notepad .env

# 然后运行
.venv\Scripts\python.exe scripts\02_call_llm.py

# 也可以直接传一条参数试试
.venv\Scripts\python.exe scripts\02_call_llm.py "▲设备重量≥130KG"
```

---

## 环境

| 组件 | 版本 | 位置 |
|---|---|---|
| Python | 3.11.9 | `C:\Users\Administrator\Python\3.11.9\` |
| 虚拟环境 | — | `.venv\`（项目内） |
| Git | 2.55.0 | 系统 |
| VS Code | 1.143.1 | 用户目录 |

关键依赖：`openpyxl`（Excel）、`pdfplumber`（PDF）、`python-docx`（Word）、`requests`（HTTP）、`python-dotenv`（配置）

---

## Git 提交规范

每完成一个可运行的小功能就提交一次，不要攒：

```bash
git add -A
git commit -m "feat: 加上峰值功率的正偏离判定"
git log --oneline
```

前缀约定：`feat:` 新功能 / `fix:` 修 bug / `docs:` 文档 / `refactor:` 重构

数据文件（xlsx/docx/pdf）已在 `.gitignore` 里排除——真实标书含商业信息，不要传到公开仓库。

---
name: word-format-master
description: Word 论文与报告自动排版工具。把你的论文和学校格式要求（或排好版的样例文档）交给 AI，即可自动完成字体、段落、页边距、页眉页码、目录与参考文献等格式调整，并逐项复核验收，全程保护原稿安全。
---

# Word Format Master（Word 论文与报告智能排版工具）

**把论文和格式要求交给 AI，自动完成 Word 排版与格式检查，原稿安全不破坏。**

你只需要提供**待排版的 Word 文档**，再给出一份**学校/期刊的格式要求**、一份**已经排好版的样例文档**，或者**直接用一句话说出你想怎么改**。AI 会自动读懂要求、对齐章节结构，并调用本地排版引擎生成排好版的新文档。

---

## 一、 它能帮你做什么？

### 1. 四大常用场景
- **按《格式要求文档》排版**：支持读取 Word（`.docx`/`.dotx`）、PDF、TXT、Markdown 格式的排版规范，自动提取字体、字号、行距、页边距等规定并应用到你的论文中。
- **照着《已排版样例》排版**：没有规则说明？直接丢一份往届优秀论文或模板样例（`.docx`/`.dotx`），自动提取其中的版式风格并迁移到你的文档（**只学格式，绝不照抄样例里的文字和姓名**）。
- **一句话局部修改**：直接告诉 AI“只把第三章正文改成宋体小四、1.5 倍行距”或“把全文表格改成三线表”，指哪改哪，绝不乱动其他段落。
- **只做格式体检（不改原文件）**：提交前想确认格式是否合规？支持只检查不修改，逐条对比要求并生成清晰的格式问题清单。

### 2. 支持调整的格式范围

| 类别 | 支持调整的具体内容 |
| :--- | :--- |
| **页面设置** | 纸张大小（如 A4）、纸张方向（横/纵）、上下左右页边距、装订线距离与位置 |
| **字体与标题** | 中英文字体（如宋体 + Times New Roman）、字号、加粗、斜体、颜色、多级标题样式与自动编号 |
| **段落排版** | 行距（多倍/固定值）、首行缩进、悬挂缩进、左右缩进、段前/段后间距、对齐方式、孤行与分页控制 |
| **论文结构** | 自动生成目录、分节页眉页脚、首页不同/奇偶页不同、页码格式、图表题注对齐、三线表、参考文献编号整理 |
| **质量核对** | 自动回读校验每项格式是否生效、检查正文文字是否一字不差、支持渲染成页面图片逐页检查排版效果 |

> **安全边界说明**：
> - **原稿绝对安全**：所有修改都会另存为新文件，绝不覆盖你的原稿和参考文档。
> - **不自作主张**：如果要求文档里某处规定模糊或前后矛盾，AI 会先列出来向你确认，不会用默认格式乱填；也不会擅自改写你的论文正文。
> - **格式前置提示**：老版本 `.doc` 需先转为 `.docx`；扫描图片版 PDF 需先经过文字识别（OCR）。

---

## 二、 怎么使用？

### 第一步：准备环境
支持 Windows 10/11，需安装 Python 3.10 及以上版本（基础排版无需安装 Office；如需渲染成页面图片进行视觉检查，电脑需装有 Microsoft Word 或 LibreOffice）。

```powershell
git clone https://github.com/xiaoyang732/word-format-master.git
cd word-format-master
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### 第二步：直接用日常语言跟 AI 对话
无需自己编写代码或拼装命令，在 AI 工具中启用本技能后，直接像下面这样说即可：

- **用法 1：按学校要求文档排版**
  > “帮我按 `D:/学校论文格式要求.docx` 整理 `D:/我的论文.docx`，另存为 `D:/我的论文-已排版.docx`。如果有不明确或冲突的要求先列出来问我。”
- **用法 2：照着优秀样例排版**
  > “用 `D:/排版样例.docx` 作为格式参考，帮我排版 `D:/我的论文.docx` 并另存为新文件，保留我自己的正文和数据。”
- **用法 3：局部快速修改**
  > “把 `D:/我的论文.docx` 里的所有一级标题改成黑体三号居中，正文改成宋体小四、1.5 倍行距、首行缩进 2 字符。”
- **用法 4：参照样例结构写新论文**
  > “参考 `D:/排版样例.docx` 的结构帮我起草一篇新论文。样例里有哪些模块和章节层级就照着来，不要多加样例里没有的模块。”
- **用法 5：只检查不修改**
  > “对照 `D:/学校论文格式要求.docx` 检查一下 `D:/我的论文.docx` 有哪些格式不合格的地方，只出检查报告，不用改文件。”

---

## 三、 项目框架与工作原理

为什么不让 AI 直接临时写一段 Python 脚本去改你的 Word？因为临时脚本很容易漏改属性、破坏复杂样式，甚至误删正文内容。

本项目采用 **“AI 负责理解意图 + 本地确定性引擎负责精准执行”** 的分层架构：

### 1. 四步标准工作流
```text
[1. 读取解析 Inspect/Profile] ──► [2. 规则对齐与规划 Plan] ──► [3. 安全写入 Apply] ──► [4. 独立复核 Verify]
```
1. **读取解析（Inspect & Profile）**：扫描目标论文的章节与段落结构，同时从要求文档或样例文档中提取排版规则，保留原文依据。
2. **规则对齐与规划（Plan）**：将排版规则与论文中的各个模块（封面、摘要、各级标题、正文、参考文献、致谢等）一一对应，生成可预览的修改计划。遇到冲突或歧义时先交由用户决定。
3. **安全写入（Apply）**：由经过严格测试的 63 项标准格式算子执行修改，只改请求指定的属性，先写入临时文件，确保原子性与原稿安全。
4. **独立复核（Verify & Render）**：重新打开生成的文档，逐项比对格式数值、校验正文与未修改部分是否完好无损，并支持渲染为页面图片检查分页、页眉和表格跨页效果。

### 2. 核心目录结构
```text
word-format-master/
├── SKILL.md                  # 技能说明与 AI 执行指南
├── README.md                 # 项目说明文档
├── assets/
│   └── presets.json          # 内置标准排版预设（仅在用户主动选择时使用）
├── scripts/
│   ├── format_cli.py         # 核心命令行入口（提取、检查、规划、执行、验收）
│   ├── render_docx.py        # 页面视觉渲染工具（调用 Word / LibreOffice 导出页面图）
│   ├── verify_output.py      # 输出文档结构与格式复核工具
│   └── word_format/          # 底层确定性排版核心库（63 项标准格式操作算子）
└── references/               # 详细接口规范与进阶参考文档
```

---

## 四、 AI 执行指令与命令行参考（供 AI Agent 自动调用）

> 以下内容供 AI 智能体在执行任务时内部调用，普通用户无需手动操作。

### 1. 路径与基本准则
- 将 `SKILL_DIR` 设为本技能根目录，定位已安装 `python-docx` 的 Python 解释器为 `PYTHON_BIN`。
- 所有路径均使用绝对路径，中间产物存放在临时任务目录中，**严禁覆盖用户原稿或格式参考源文档**。
- 仅调用下方注册的 CLI 工具执行文档读写，禁止自行编写独立的文档修改脚本。
- 若根据样例起草新论文，先阅读 [references/thesis-generation.md](references/thesis-generation.md)：样例包含什么模块就生成什么模块，保持模块种类、数量、语言、顺序和章节层级完全一致，绝不自动添加样例中没有的英文摘要、承诺书、目录、致谢或附录。

### 2. 提取格式依据与文档结构
详见 [references/format-profile.md](references/format-profile.md) 与 [references/direct-format-api.md](references/direct-format-api.md)。

```powershell
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" capabilities --output capabilities.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" profile requirements REQUIREMENTS.docx --output profile.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" profile template SAMPLE.docx --output profile.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" inspect INPUT.docx --output inspect.json
```

- 规则优先级：用户明确指定 > 学校/期刊样例模板 > 书面要求文档 > 指定标准 > 用户主动选择的预设 > 原稿现有格式（参见 [references/standards-registry.md](references/standards-registry.md)）。
- 存在未决冲突或模糊规则时，在 `resolutions.json` 中记录 `choose`、`replace` 或 `ignore` 决策及理由；多节样例需在 `mapping.sections` 中显式映射节号。

### 3. 规划、执行与验收

```powershell
# 基于要求或样例档案进行规划、应用与复核
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" plan INPUT.docx OUTPUT.docx --profile profile.json --mapping mapping.json --resolutions resolutions.json --output plan.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" apply --plan plan.json --report report.json
& $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" verify --plan plan.json --report report.json --output verification.json
```

- **按样例生成论文结构校验**：对基于样例起草的论文，排版前必须运行 `check-structure DRAFT.docx --profile profile.json --output structure.json`，并在 `plan` 时加上 `--match-structure`。
- **局部直接修改**：构造 schema 2.0 请求 JSON，使用 `plan INPUT.docx OUTPUT.docx --request request.json --output plan.json`。
- **仅检查不修改（Audit）**：
  ```powershell
  & $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" audit INPUT.docx --request request.json --output audit.json
  ```
- **视觉渲染与页面复核**：
  ```powershell
  & $PYTHON_BIN "$SKILL_DIR/scripts/render_docx.py" OUTPUT.docx --output-dir RENDER_DIR --method auto
  & $PYTHON_BIN "$SKILL_DIR/scripts/format_cli.py" verify --plan plan.json --report report.json --visual-report visual.json --render-manifest RENDER_DIR/render-manifest.json --output verification.json
  ```
  分页、页眉、表格断页和域刷新详见 [references/header-requirements.md](references/header-requirements.md) 与 [references/architecture.md](references/architecture.md)。

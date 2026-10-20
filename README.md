# Word Format Master

**把论文和格式依据交给 AI，通过本地工具完成排版与复核。**

[![CI](https://github.com/xiaoyang732/word-format-master/actions/workflows/ci.yml/badge.svg)](https://github.com/xiaoyang732/word-format-master/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

提供目标论文，再提供学校的格式要求文档、已排好版的论文样例，或直接说明修改要求。AI 读取材料、处理格式冲突、匹配目标位置，再调用确定性的格式引擎生成新的 DOCX。原稿与格式依据保持不变。

> “按《学校论文要求.docx》整理《我的论文.docx》，另存新文件。无法确定的要求先列出来。”
>
> “用《排版样例.docx》作为格式依据，保留我的正文、姓名和数据，列出未覆盖的格式。”
>
> “只把第三章正文改为宋体、小四、1.5 倍行距。”

## 能处理什么

| 输入或能力 | 处理方式 |
|---|---|
| 格式要求文档 | 读取 UTF-8 TXT/Markdown、DOCX/DOTX、文本 PDF，保留原文和位置证据 |
| 已排版论文样例 | 提取 DOCX/DOTX 有效字体、字号、段落属性、全部页面节及结构清单 |
| 精确局部修改 | 按章节、角色、段落、表格单元格或普通文字范围修改 |
| 页面与段落 | 纸张、方向、边距、装订线、行距、缩进、段距及分页控制 |
| 字体与标题 | 中西文字体、字号、粗斜体、颜色、标题层级与明确请求的自动编号 |
| 论文结构 | 明确配置目录字段、页码、分节页眉、首页与奇偶页、题注位置、三线表及文献编号 |
| 验收 | 逐项回读格式、检查内容和无关属性保全，并可渲染逐页检查 |

样例分析和自动迁移有明确边界：同一角色中的混合格式、复杂页眉页脚、字段、编号和特殊对象会进入待处理问题；支持的项目才生成执行规则。扫描 PDF 需要 OCR，DOC 需先转为 DOCX。不会用通用预设补齐材料没写出的要求，也不会把样例文字复制进目标论文。

## 快速开始

面向 Windows 10/11、Python 3.10 及以上。格式写入和结构检查无需 Word；实际分页和视觉验收需要 Microsoft Word 或 LibreOffice。

```powershell
git clone https://github.com/xiaoyang732/word-format-master.git
Set-Location word-format-master
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

在具备本地文件读取和命令执行能力的 AI 工具中加载 [SKILL.md](SKILL.md)，或这样说：

```text
请读取 <项目目录>/SKILL.md，使用 word-format-master。
目标论文：C:/Documents/论文.docx
格式依据：C:/Documents/学校格式要求.docx
输出：C:/Documents/论文-排版后.docx
请检查全部要求，处理明确可确定的格式，报告冲突、未覆盖项及验收状态。
```

用户通过当前 AI 会话提出需求。下面的 CLI 和 JSON 由 AI 组织，不要求用户手工配置。

## 要求文档与样例工作流

两种来源任选其一生成格式档案：

```powershell
python scripts/format_cli.py profile requirements "学校要求.docx" --output profile.json
python scripts/format_cli.py profile template "排版样例.docx" --output profile.json
```

档案保存来源 SHA-256、证据锚点、可执行规则、原始快照、覆盖范围和待处理问题。AI 检查问题后记录明确决策；多个页面节需要映射到目标论文的实际节编号。

```powershell
python scripts/format_cli.py inspect "论文.docx" --output inspect.json
python scripts/format_cli.py plan "论文.docx" "论文-排版后.docx" --profile profile.json --mapping mapping.json --resolutions resolutions.json --output plan.json
python scripts/format_cli.py apply --plan plan.json --report report.json
python scripts/format_cli.py verify --plan plan.json --report report.json --output verification.json
```

没有映射或问题时可省略 `--mapping`、`--resolutions`。未解决问题会返回 `needs_clarification`，此时保存的是审阅结果，不能用作执行计划。决策和显式保留项会随执行报告保存。详见 [格式档案契约](references/format-profile.md)。

局部修改使用 `plan --request request.json`；只检查使用 `audit --request request.json`。请求和选择器见 [直接调用指南](references/direct-format-api.md)。内置 `thesis-standard` 仅在用户选择时应用，学校或期刊要求优先。

## 验收与保护

计划绑定目标论文和格式依据的路径及 SHA-256，来源变化后必须重新规划。写入前拒绝参数和位置冲突，输出先保存临时文件，重新打开并验收后发布。只改字号会保留其他未请求属性；编号转换、目录、引用、页眉内容替换等结构操作须明确配置。

独立 `verify` 检查报告是否完整覆盖计划，并重新读取输出。JSON 报告也不能覆盖输入、格式依据或目标文档。已有输出默认不覆盖。

```powershell
python scripts/render_docx.py "论文-排版后.docx" --output-dir render --method auto
```

最终分页、字形、表格、页眉和 Word 字段需要实际渲染并逐页检查。没有完成渲染和检查时，只能报告结构通过，视觉为 `skipped` 或 `pending`。Word 与 LibreOffice 可能有分页差异；不能对任意 Word 文件承诺所有格式都能自动复制。

AI 完成页面检查后，可用 `verify --visual-report visual.json --render-manifest render/render-manifest.json` 绑定并复核视觉证据，拒绝过期或缺页报告。完整命令和报告格式见 [直接调用指南](references/direct-format-api.md#绑定视觉验收)。

## 常见问题

**会改写正文吗？** 普通格式操作保留正文。明确的编号、引用和目录操作按声明的结构影响执行。正文润色需要另行提出。

**会上传文档吗？** 本地工具不调用大模型 API。AI 会话如何处理文档上下文取决于所用 AI 工具的配置。

**支持所有 Word 格式吗？** 支持 DOCX 修改和 DOCX/DOTX 样例提取。宏文档、Strict OOXML 写入、复杂文字范围及部分包级结构暂不支持；未支持项会保留并报告。

**能只检查、不修改吗？** 可以。`audit` 不保存 DOCX，报告每个目标的期望值、有效值和差异；有差异退出码为 2，完整报告仍保存。

## 开发验证

```powershell
python -m pip install -r requirements-dev.txt
python -m ruff check scripts/word_format scripts/apply_spec.py scripts/format_cli.py
python -m ruff format --check scripts/word_format scripts/apply_spec.py scripts/format_cli.py
python scripts/validate_project.py
python scripts/smoke_test.py
python -m unittest discover -s tests -v
git diff --check
```

每项新格式能力必须登记参数约束、唯一写入方法、独立回读和验收。性能基准见 `python scripts/benchmark_formatting.py --profile`。

## 更多文档

- [Skill 工作流](SKILL.md)
- [格式档案、映射与冲突决策](references/format-profile.md)
- [直接调用指南](references/direct-format-api.md)
- [兼容格式规范](references/format-spec.md)
- [页眉设置](references/header-requirements.md)
- [标准与预设](references/standards-registry.md)
- [架构](references/architecture.md) · [安全边界](SECURITY.md)
- [改造方案与实现进度](references/direct-ai-only-roadmap.md)

[MIT License](LICENSE)。

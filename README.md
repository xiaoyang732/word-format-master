# 论文排版大师 (Word/DOCX)（Word 论文与报告智能排版工具）

**用一句话撰写课程报告，或把论文和格式要求交给 AI，完成 Word 排版与格式检查。**

直接说“帮我写一份报告，计算机方向的”或“帮我写一份数据库课程作业报告”，即可默认调用项目的[通用课程报告模板](templates/reference/general-course-report/通用课程报告模板.docx)。模板沿用已选定的封面填写线、段落格式、自动目录和“第一章、第二章……”编号。无需再次指定模板；主题不明确时会询问主题。用户提供的其他模板或明确格式要求优先。详见 [报告生成流程](references/report-generation.md)。

生成文件中的章节、小节、列表、参考文献使用 Word 自动编号；题注、公式、页码和交叉引用使用动态域，禁止键入结构编号。交付前必须执行 `python scripts/check_numbering.py OUTPUT.docx --output numbering.json`；统一流水线自动拦截检查失败的结果。详见 [生成规范](references/thesis-generation.md)。

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE.txt)

已有文档排版时，提供**待排版的 Word 文档**，再给出**格式要求**、**已排版样例**，或**用日常语言说明想怎么改**。撰写新报告时，只需说明主题与已有素材。无需自己编写代码或手动敲命令，AI 会组织内容与规则，并调用本地确定性排版引擎生成新的文档及验收报告。

---

## 一、 它能帮你做什么？

### 1. 常用场景
- **场景 1：按《格式要求文档》排版**：支持读取 Word（`.docx`/`.dotx`，装有 Word/WPS/LibreOffice 时自动只读兼容 `.doc`/`.dot`/`.wps`）、PDF、TXT、Markdown 格式的排版规范，自动提取字体、字号、行距、页边距等规定并应用到你的论文中。
- **场景 2：照着《已排版样例》迁移版式**：没有规则说明？直接丢一份往届优秀论文或模板样例（`.docx`/`.dotx`/`.doc`/`.dot`/`.wps`），自动提取其中的版式风格并迁移到你的文档，支持按需克隆官方封面布局（校徽、表格边框、文本框、下划线槽位）（**只学格式，绝不照抄样例里的文字和姓名**）。
- **场景 3：一句话局部修改 / 一键套用内置预设**：直接告诉 AI“只把第三章正文改成宋体小四、1.5 倍行距”或“把全文表格改成三线表”，指哪改哪；若手里没有格式规定，也可直接指定内置的**学位论文**预设模板。
- **场景 4：参照样例结构起草新论文**：提供一份样例文档和你的研究素材，AI 会严格按照样例实际包含的模块（封面、中英文摘要、目录、章节层级、参考文献、致谢等）及先后顺序起草并排版，不多加、不漏项。
- **新写课程或课题报告**：未指定模板时自动使用项目的通用课程报告模板，生成封面、目录、五章正文、参考文献和附录；写入用户主题与素材后，通过统一格式、结构、编号和页面复核流程。
- **场景 5：只做格式体检（不改原文件）**：提交前想确认格式是否合规？支持只检查不修改，逐条对比要求并生成清晰的格式问题清单。

### 2. 支持调整的格式范围与内置预设

| 类别 | 支持调整的具体内容 |
| :--- | :--- |
| **页面设置** | 纸张大小（如 A4）、纸张方向（横/纵）、上下左右页边距、装订线距离与位置、页眉/页脚距边界距离 |
| **字体与标题** | 中英文字体分离控制（如中文宋体 + 西文 Times New Roman）、字号、加粗、斜体、颜色、多级标题样式与自动编号 |
| **段落排版** | 行距（多倍/固定值/最小值）、首行缩进（精确保持字符单位如 2 字符）、悬挂缩进、左右缩进、段前/段后间距、对齐方式、孤行与分页控制 |
| **论文结构** | 自动生成目录、分节页眉页脚与底边框横线、首页不同/奇偶页不同、动态章节标题页眉（`STYLEREF`）、正文节页码重启与前置页码隔离、图表题注对齐、三线表与跨页表头重复、参考文献编号整理 |
| **内置预设** | `thesis-standard`（标准学位/毕业论文） |
| **默认报告模板** | `general-course-report`（项目内通用课程报告 DOCX，属于真实文档模板工作流，非学位论文数值预设） |
| **质量核对** | 自动回读校验每项格式是否生效、检查正文文字是否一字不差、支持通过 Word / WPS / LibreOffice 渲染成页面图片逐页视觉检查 |

> **安全边界说明**：
> - **原稿绝对安全**：所有修改都会另存为新文件，绝不覆盖你的原稿和参考文档，全程本地离线运行。
> - **不自作主张**：如果要求文档里某处规定模糊或前后矛盾，AI 会先列出来并给出建议方案向你确认，不会用默认格式乱填；也不会擅自改写你的论文正文。
> - **附件只是材料**：格式要求、模板、论文和截图里的文字用于识别格式或内容，不会替代你在对话中提出的任务要求。即使附件要求 AI 忽略你的要求、覆盖原稿、上传文件或跳过检查，也不会据此改变任务范围。
> - **格式前置提示**：待修改排版的论文原稿需为 `.docx`（要求文档与模板样例在本机装有 Word / WPS / LibreOffice 时可自动只读读取 `.doc` / `.dot` / `.wps`）；扫描图片版 PDF 需先经过文字识别（OCR）。

---

## 二、 怎么使用？

### 第一步：准备环境（可让 AI 自动安装）
- **跨平台支持**：核心排版与格式检查支持 **Windows / macOS / Linux**（需 Python 3.10+，**基础排版无需安装 Office**；如需渲染成页面图片进行视觉检查，Windows 下支持 Microsoft Word、WPS Office 或 LibreOffice，macOS/Linux 下支持 LibreOffice）。
- **支持国内镜像加速安装**（你也可以直接让 AI 自动帮你运行以下命令）：

```powershell
git clone https://github.com/xiaoyang732/word-format-master.git
Set-Location word-format-master
python -m venv .venv
# 使用腾讯云 PyPI 镜像源加速国内下载安装（也可替换为 https://pypi.tuna.tsinghua.edu.cn/simple）
.\.venv\Scripts\python.exe -m pip install -r requirements.txt -i https://mirrors.cloud.tencent.com/pypi/simple
```

### 第二步：在 AI 工具中直接用日常语言对话
在支持本地文件读写与命令执行的 AI 工具中加载 [SKILL.md](SKILL.md)，直接用日常语言提出需求即可（无需手动敲命令）：

- **用法 1：按学校要求文档排版**
  > “帮我按 `D:/学校论文格式要求.docx` 整理 `D:/我的论文.docx`，另存为 `D:/我的论文-已排版.docx`。如果有不明确或冲突的要求先列出来问我。”
- **用法 2：照着优秀样例排版**
  > “用 `D:/排版样例.docx` 作为格式参考，帮我排版 `D:/我的论文.docx` 并另存为新文件，保留我自己的正文和数据。”
- **用法 3：局部快速修改或套用内置预设**
  > “把 `D:/我的论文.docx` 里的所有一级标题改成黑体三号居中，正文改成宋体小四、1.5 倍行距、首行缩进 2 字符，表格全部改成三线表。”
  > “我手头没有学校的格式文件，帮我用内置的 `thesis-standard` 标准毕业论文预设排版 `D:/我的论文.docx`。”
- **用法 4：参照样例结构写新论文**
  > “参考 `D:/排版样例.docx` 的结构帮我起草一篇新论文。样例里有哪些模块和章节层级就照着来，不要多加样例里没有的模块。”
- **用法 5：只检查不修改（格式体检）**
  > “对照 `D:/学校论文格式要求.docx` 检查一下 `D:/我的论文.docx` 有哪些格式不合格的地方，只出检查报告，不用改文件。”

### 不确定该选哪种方式时

| 你的目标 | 告诉 AI 什么 | 是否生成新 DOCX |
| :--- | :--- | :--- |
| 按学校/期刊要求排版 | 提供要求文件、待排论文和另存路径 | 是 |
| 参考样例迁移格式 | 提供样例、待排论文，并说明只迁移格式 | 是 |
| 修改少数格式或套用预设 | 指明文档、修改范围和具体属性；使用预设时说出预设名称 | 是 |
| 按样例结构起草论文 | 提供样例和研究素材，说明以样例实际结构为准 | 是 |
| 新写报告 | 说明主题或沿用对话主题；未提供格式依据时自动调用通用课程报告模板 | 是 |
| 只检查格式 | 提供要求文件和待检查论文，并明确“只检查，不修改” | 否 |

日常使用时只需说明目标和文件路径，AI 会选择对应的内部命令流程。若你自行运行 CLI，`SKILL.md` 中有参数组合速查；不要把自然语言直接当作 `--request` 文件传入。

### 第三步：没有要求文档时的快速口述模板（可选）
如果你没有现成的格式要求文件，可以直接复制以下模板修改后发给 AI，避免口头描述遗漏：

> 请帮我排版 `D:/我的论文.docx`，另存为 `D:/我的论文-已排版.docx`：
> 1. **页面边距**：A4 纵向，上下 `2.54` cm，左右 `2.5` cm；
> 2. **正文格式**：中文 `宋体`、英文 `Times New Roman`，字号 `小四（12pt）`，`1.5 倍` 行距，首行缩进 `2 字符`；
> 3. **各级标题**：一级标题 `黑体三号居中`；二级标题 `黑体四号左对齐`；三级标题 `黑体小四左对齐`；
> 4. **图表与目录**：表格改为 `三线表`，自动生成 `3` 级目录，页脚居中显示页码；
> 5. **参考文献**：`五号宋体`，单倍行距，悬挂缩进 `2 字符`。

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
3. **安全写入（Apply）**：由统一注册的格式算子执行修改，只改请求指定的属性，先写入临时文件，确保原子性与原稿安全。实际能力与参数以 `capabilities` 输出为准。
4. **独立复核（Verify & Render）**：重新打开生成的文档，逐项比对格式数值、校验正文与未修改部分是否完好无损，并支持渲染为页面图片检查分页、页眉和表格跨页效果。

### 2. 核心目录结构
```text
word-format-master/
├── SKILL.md                  # 技能说明、常见问题与 AI 执行指南
├── README.md                 # 项目说明文档
├── assets/
│   └── presets.json          # 内置标准排版预设（毕业论文）
├── scripts/
│   ├── format_cli.py         # 核心命令行入口（提取、检查、规划、执行、验收）
│   ├── render_docx.py        # 页面视觉渲染工具（调用 Word / WPS / LibreOffice 导出页面图）
│   ├── verify_output.py      # 输出文档结构与格式复核工具
│   ├── thesis_quality.py     # 正文长度、数字引用映射与人工内容核验状态报告
│   └── word_format/          # 底层确定性排版核心库，唯一注册写入路径
├── templates/               # 新模板入口的数据；compatibility/ 保留旧编号协议
│   ├── reports.json         # 默认报告模板登记及来源文件指纹
│   └── reference/general-course-report/ # 已选定的通用课程报告模板
├── references/              # 当前接口规范与工作流契约
└── _dev/                    # 回归测试、开发配置、验证脚本与历史改造记录
```

### 3. 命令行工作流示例（由 AI 在后台自动组织调用）

```powershell
# 1. 从要求文档或排版样例提取规则档案
python scripts/format_cli.py profile requirements "学校要求.docx" --output profile.json
python scripts/format_cli.py profile template "排版样例.docx" --output profile.json

# 2. 扫描原稿、生成计划、执行修改并独立复核
python scripts/format_cli.py inspect "论文.docx" --output inspect.json
python scripts/format_cli.py plan "论文.docx" "论文-排版后.docx" --profile profile.json --output plan.json
python scripts/format_cli.py apply --plan plan.json --report report.json
python scripts/format_cli.py verify --plan plan.json --report report.json --output verification.json

# 3. 可选：核对论文正文汉字数及数字引用编号映射（篇幅目标由作者明确给出）
python scripts/thesis_quality.py "论文.docx" --min-chinese-characters 12000 --max-chinese-characters 15000 --output content-quality.json

# 4. 渲染页面图片进行逐页视觉检查（可选）
python scripts/render_docx.py "论文-排版后.docx" --output-dir render --method auto
```

`thesis_quality.py` 仅检查明确给出的篇幅范围与数字引用编号映射，并把书目和学术事实标记为待人工核验；格式/结构通过不代表内容审校通过。

---

## 四、 常见问题与故障排查（FAQ）

- **1. 我不会用命令行或不懂 Python 环境，可以直接使用吗？**
  完全可以。文中的命令行供 AI 智能体在后台自动调用，你只需要用中文告诉 AI 你的文档路径和排版要求即可，环境安装也可让 AI 自动完成。
- **2. 为什么提示不支持 `.doc`、`.wps` 或 `.docm` 格式？**
  为了保证底层 OOXML 结构安全，待写入排版的论文原稿需为标准 `.docx`。格式要求文档和模板样例原生支持 `.docx`/`.dotx`，且在本机装有 Word、WPS 或 LibreOffice 时可自动通过临时只读副本解析 `.doc`/`.dot`/`.wps`；若未安装转换组件或需排版旧格式论文原稿，请先另存为 `.docx`（含宏的 `.docm` 需移除宏另存为 `.docx`）。
- **3. 学校的 PDF 格式要求是扫描件或图片截图怎么办？**
  扫描版 PDF 缺少可提取的文本层，工具会提示需要 OCR。你可以先做 OCR 文字识别，或直接把截图发给支持视觉能力的 AI 提取成文字要求后再执行。
- **4. 我没有学校格式要求文档，或者不确定具体参数怎么办？**
  你可以直接让 AI 使用内置的 毕业论文预设（`thesis-standard`），或者套用上方的“快速口述模板”。若要求文档内部存在前后矛盾，AI 会先列出冲突项请示你，绝不盲目覆盖。
- **5. 电脑没装 Microsoft Word（比如只装了 WPS，或使用 macOS / Linux）能用吗？**
  能用。基础排版与格式检查完全基于本地 Python 库，在 Windows / macOS / Linux 下均无需安装 Office。如需将文档渲染成图片做视觉检查，Windows 下支持调用 Microsoft Word、WPS Office 或 LibreOffice，macOS / Linux 下支持调用 LibreOffice。
- **6. 打开生成的文档后，目录页码或页眉章节名没有更新怎么办？**
  Word 的目录（`TOC`）和动态章节页眉（`STYLEREF`）属于动态字段。文档已自动写入打开时更新标记，若阅读器未自动刷新，在 Word 或 WPS 中按 `Ctrl + A` 全选再按 `F9` 即可一键刷新所有目录与页码。
- **7. 会改写我的论文正文或把文档上传到外部服务器吗？**
  绝不会。格式算子全程保留你的正文文字并做哈希与文本一致性核验；本地排版脚本完全在你的电脑上离线运行，不调用外部接口。
- **8. 能只检查格式、不修改文档吗？**
  可以。告诉 AI“只检查不修改”，工具会通过只读 `audit` 模式生成详细的格式对比报告，原文档不做任何写入。
- **9. 出现 `needs_clarification`、`unsupported` 或 `skipped` 怎么办？**
  `needs_clarification` 表示有冲突或目标不明确，先回答 AI 的澄清问题；`unsupported` 表示当前格式或操作不受支持，可尝试转换文档或缩小修改范围；`skipped` / `pending` 表示页面视觉检查尚未完成，不能据此判断页面效果已通过。让 AI 用中文说明具体原因和下一步即可。

---

## 五、 开发验证与进阶文档

自动化测试、开发配置和验证脚本位于 `_dev/`，并已纳入版本控制。生成文件与缓存由 `.gitignore` 排除；开发历史记录放在 `_dev/docs/`。

```powershell
# 在仓库根目录运行本地回归与算子校验
python _dev/scripts/validate_project.py
python _dev/scripts/smoke_test.py
python -m unittest discover -s _dev/tests -v
```

- [Skill 工作流说明](SKILL.md)
- [格式档案、映射与冲突决策](references/format-profile.md)
- [直接调用指南](references/direct-format-api.md)
- [按样例生成论文流程](references/thesis-generation.md)
- [兼容格式规范](references/format-spec.md) · [页眉设置](references/header-requirements.md) · [标准与预设](references/standards-registry.md)
- [架构说明](references/architecture.md) · [安全边界](SECURITY.md)
- [本次可靠性审查与优化建议](_dev/docs/reliability-review-2026-10-31.md)

README 面向使用者，介绍能力、安装和任务选择；`SKILL.md` 说明 AI 应如何选择 CLI 参数、处理附件证据和交付结果；`references/` 保存详细接口契约，`SECURITY.md` 记录本地处理与文件保护边界。

[MIT License](LICENSE.txt)
## 1.1 模板引擎更新

新增独立内置模板与用户模板编译流程。一般格式迁移使用 `plan --template`；按样例严格起草及默认报告使用 `plan --profile --match-structure`。用户样例不回退内置参数。参见 [模板引擎、迁移与验收](references/template-engine.md)。旧 spec / presets 接口保留兼容，新功能使用模板入口。


历史设计与验证记录归档于 [_dev/docs/](_dev/docs/template-refactor-report.md)。一键执行共享流程：

```powershell
python scripts/run_format_pipeline.py input.docx output.docx --template thesis-standard
python scripts/run_format_pipeline.py input.docx output.docx --template school-template.json --mapping mapping.json
```

新报告的模板选择与准备入口（准备后由 AI 填充草稿、决策和映射，再执行流水线）：

```powershell
python scripts/format_cli.py report list
python scripts/format_cli.py report prepare "D:/报告任务" --topic "计算机网络"
python scripts/run_format_pipeline.py "D:/报告任务/draft.docx" "D:/报告任务/报告.docx" --profile "D:/报告任务/profile.json" --match-structure --mapping "D:/报告任务/mapping.json" --resolutions "D:/报告任务/resolutions.json" --render
```

准备入口不会生成报告正文或宣告验收完成。模板登记不改变原有 `template list`（数值格式模板）的含义，也不扩展冻结的旧 presets 接口。

## 课程与课题报告参考模板

日常课程作业默认使用 [通用课程报告模板](templates/reference/general-course-report/通用课程报告模板.docx)。它根据用户选定的天华学院原版改制，保留封面分栏填写、目录、标题层级及正文段落设置，移除学校标识、专业示例、评分栏和批注；章节与参考文献采用自动编号。详细参数和填写方法见 [格式与使用说明](templates/reference/general-course-report/格式与使用说明.md)，改制依据见 [来源说明](templates/reference/general-course-report/SOURCE.md)。该文件已登记在 `templates/reports.json`，由 `report prepare` 默认调用。

通用版本的原始来源为 [上海师范大学天华学院课程考核报告模板原件](templates/reference/course-report-tianhua/上海师范大学天华学院-课程考核报告模板-原件.doc)，另附 [DOCX 格式转换副本](templates/reference/course-report-tianhua/上海师范大学天华学院-课程考核报告模板-格式转换副本.docx)，保留用于溯源和对照。文件直接取自学校教务处发布的压缩包，示例为数字图像处理课程报告，包含封面、目录、分级标题、结果分析、参考文献与程序附录；保留原貌，没有自行改制。来源、原件校验及局限见 [来源说明](templates/reference/course-report-tianhua/SOURCE.md)，段落参数见 [格式核对说明](templates/reference/course-report-tianhua/格式核对说明.md)。

以下保留为其他课题报告的历史参考。

简洁的学校课题报告可参考 [南京城市职业学院科研项目研究报告 Word 原版](templates/reference/school-report/学校课题报告模板-南京城市职业学院原版.docx)。文件为中文，包含课题封面与正文格式要求，适合用作学校课题报告的格式参考；来源及使用说明见 [来源说明](templates/reference/school-report/SOURCE.md)。它是外部样例，不是内置排版预设。


网上原版详例可参考 [中国石油大学（北京）课题研究报告原件](templates/reference/downloaded-school-reports/中国石油大学北京-课题研究报告模板.doc)，包含排版批注及正文示例；另附 [DOCX格式转换副本](templates/reference/downloaded-school-reports/中国石油大学北京-课题研究报告模板-格式转换副本.docx)。来源、实际格式及转换说明见 [原版来源说明](templates/reference/downloaded-school-reports/SOURCE.md)。

# Architecture and Upgrade Contract

## 处理链

所有格式任务由当前 AI 会话和本地命令行完成：

```text
自然语言 / 要求文档 / 格式样例
  → evidence profile → 显式问题决策和目标映射
  → inspect → plan → apply → verify → render + 全页检查
```

明确局部请求直接转成操作；只读检查走 `audit`。AI 解释含义和权限范围，代码负责确定性提取、校验、写入和独立回读。未决项返回 `needs_clarification`，不能把审阅结果当作执行计划。

## 核心职责

| 模块 | 职责 |
|---|---|
| `requirement_sources.py` | TXT/MD、DOCX/DOTX、文本 PDF 的文本与位置证据，只读且核对源指纹 |
| `requirement_rules.py` | 保守子句解析；未消费内容、否定和条件进入未决项 |
| `profiles.py` | 格式来源、有效样例快照、规则、冲突、决策、覆盖及目标映射；编译到同一 planner |
| `template_structure.py` | 只读提取样例模块、语言、出现次数和章节层级；按实际顺序比较生成草稿，供规划和验收共用 |
| `inspect.py` / `selectors.py` | 源文件对象目录和严格位置解析；不猜测重复或不存在的位置 |
| `registry.py` | 可执行 action、目标类型、参数校验、唯一 handler、reader 和 verifier |
| `properties.py` / `operations.py` | 字体、段落和页面属性的读写及命名 handler |
| `structures.py` | 目录、多级编号、引用、题注位置和文档级页码的唯一结构写入 |
| `headers.py` / `section_breaks.py` | 节和页型页眉、动态字段、显式下一页分节与保护 |
| `spec.py` | 兼容 Format Spec 1.0 展开器；不能增加第二套写入器 |
| `api.py` | 请求、指纹、计划、冲突检测、执行、临时文件验收与发布 |
| `verification.py` | 指定属性回读、逐操作保护投影、结构修改边界 |
| `visual.py` | 将 AI 视觉报告绑定到输出、渲染清单、实际渲染器及完整逐页图片证据 |

`document_structure.py` 统一模块、标题层级和角色检测。`analyze_docx.py` 读取包部件、样式、节、字段、特殊结构和推断规范，只产生证据。Profile 的可执行规则使用共享有效值读取器，不以分析器的首节或代表段落推断替代完整样例扫描。

## 入口与数据

`format_cli.py` 提供 `capabilities / inspect / profile / check-structure / audit / plan / apply / verify`。
`apply_spec.py` 兼容直接 `--spec` 应用，委托 `word_format.api.apply_spec()`；兼容导出的字段/页码方法仍归属核心。

- Format Profile：`schema_version: 1.0`，见 [format-profile.md](format-profile.md)。
- Format Spec：`schema_version: 1.0`，见 [format-spec.md](format-spec.md)。
- Request / Plan / Report：`schema_version: 2.0`，见 [direct-format-api.md](direct-format-api.md)。
- Render Manifest：`schema_version: 1.0`。

Profile 记录原始证据和问题，`resolutions` 保留决策理由。`mapping.sections` 对应源节和目标节；`mapping.rules` 使用当前 inspect 选择器。样例未覆盖的目标段落须明确保留。Profile 不复制样例内容，不重建未支持的包级结构。

计划保存 `profile_context`，绑定格式依据的路径和指纹、档案摘要、映射、决策与保留项。执行和复检重新验证依据。摘要检测误改，不提供身份认证。

按样例生成论文须先审阅 `snapshot.structure_contract` 并完成草稿，再执行只读 `check-structure` 和 `plan --match-structure`。结构契约使用 `match: exact`，保留原样例顺序，不套通用模块排序；区分中英文摘要/关键词、声明类型、模块重复次数和章节层级。来源复读防止修改快照绕过对照，差异不能由格式问题的忽略或目标保留豁免。契约绑定到计划后，临时文件验收与独立复检继续检查；额外目录等结构写入不能发布不匹配的输出。此功能只检测结构，不负责编写缺失正文或代替封面视觉检查。

## 变更保护

普通操作只修改请求属性，不改变未选对象。共享样式必须显式选择。字符范围仅支持简单文字。首行与悬挂缩进互斥，字符单位保持字符度量。编号、页眉内容、目录、引用及分节等结构修改独立声明影响并验收。

所有输出先保存临时文件，回读每一项操作并运行结构检查，成功才原子发布。`verify_plan` 要求检查项完整覆盖源目标和参数，包含无需修改的项目，再回读输出及重跑结构验收。

## 渲染与交付

`render_docx.py` 只产生页面证据，`verify_output.py` 检查结构与视觉证据。渲染使用预先计算哈希的快照，源文件变化不发布清单；显式渲染器选择不静默降级。刷新所有 StoryRanges 中的字段和目录后输出 PDF/PNG，不保存回原稿。相对外部关系需要先处理基准路径。

结构通过不证明分页、字形或视觉通过。AI 必须检查所有页面后才可给出视觉结论；缺少渲染或检查须报告 skipped / pending。

## 扩展检查

新增能力须同时提供证据来源、目标范围、参数约束、单一写入方法、独立读取、验收和有效回归。禁止新入口绕过 registry；禁止用预设默认值填补未知样例属性。改变契约语义须升级版本，同步 CLI、Skill、README、安全说明和测试。

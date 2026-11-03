# 2026-10-13 最终审查报告整改与验收记录

> 历史审查记录。2026-10-20 已移除其中涉及的配置页面及会话链路；当前使用方式以 README、SKILL 和 format-profile.md 为准。

本轮根据 Gemini、GPT 的独立报告及综合复核报告整改，审查基线为
`9732430cd276b4bb7f5bfdc10ca2d6c0a780ad8b`。以包含逐项复现证据的
`final-report.md` 中 S0–S3 编号记录结果；另一份综合报告的性能目标也保留为验收边界。
原始报告未修改。每个修复块在验证后单独提交并推送到 `main`。

## 问题与提交映射

| 报告项 | 修复结果 | 提交 | 回归证据 |
|---|---|---|---|
| S0-01 JSON 覆盖文档 | 公共路径预检拦截输入、输出、请求、规范、计划、Handoff 及硬链接别名冲突；JSON 使用同目录临时文件原子替换 | `b7fe87c` | `JsonSafetyTests`，源 SHA-256 与 ZIP 有效性、失败时保留旧报告 |
| S0-02 参考文献越界 | 只转换索引中 `reference` 角色的段落；保护致谢、附录与范围外内容 | `6360a0b` | `ReferenceSafetyTests`，普通样式模块标题与内容保全 |
| S0-03 页眉/PAGE 冲突 | 内容与独立 PAGE 段落分别管理，任意写入顺序保留两者；拒绝破坏混合文字/域段落 | `24d2e14` | `HeaderTests`，默认/首页/偶页、链接节、简单域和错误 writer |
| S1-01 `arabic-space` | 校验、执行及独立读回统一支持 `%1` | `3b9c74a` | `ModuleContractTests` 与原有三级编号测试 |
| S1-02 论文模块样式 | 识别封面标题、中英文摘要和关键词；仅应用显式 token，保留模板/请求的权威来源 | `3b9c74a` | `ModuleContractTests`、`test_thesis_regressions.py` |
| S1-03 索引失配 | 传递继承的大纲级别，章节 occurrence 独立从 1 递增 | `3b9c74a` | 自定义 outline 样式、三个章节的 occurrence 回归 |
| S1-04 性能 | run 级复用投影，未选子树使用 C 层序列化，增加命名空间/样式/主题缓存和可复现基准 | `4017f86` | 下方实测；2–2.5 秒目标尚未达成 |
| S2-01 自动/按行间距 | 读回保留自动间距与行单位语义；固定磅值写入清除冲突属性 | `5907817` | `PropertySemanticsTests` |
| S2-02 西文字体 | 同时检查 ASCII、hAnsi、主题字体及样式继承，不能因单通道满足而跳过写入 | `5907817` | `PropertySemanticsTests` |
| S2-03 页面几何 | 仅改方向时交换尺寸；明确尺寸按请求执行，验收方向与宽高一致 | `5907817` | 多节/往返方向回归及下方 WPS 实际渲染；Word/LibreOffice 尚未验收 |
| S2-04 inspect 缺字段 | 段落检查字段由能力表驱动，包含颜色、缩进和分页控制属性 | `3b9c74a` | `ModuleContractTests` 比对检查字段与原子段落能力集 |
| S2-05 嵌套域/偶页验收 | 复杂域采用深度计数，独立验收覆盖偶页与简单 PAGE 域 | `4896704` | `HeaderTests`，嵌套域文本读回及偶页字段计数 |
| S2-06 前端日志/锁定/XSS | 日志挂载并清理旧内容，事件统一读取锁定状态，交叉验证动态字段和值转义 | `37136b7` | Playwright 桌面/移动端回归，恶意 HTML、日志可见性和锁定检查 |
| S3 双轨实现 | 复杂 OOXML writer 归入共享核心 `structures.py`，CLI 只保留 Handoff 校验和核心调用 | `b31d41d` | 全量回归、DOCX 冒烟测试及架构文档 |
| S3 紧凑语法 | Ruff 统一核心排版、整理导入，并加入独立 CI 检查 | `e2873b4` | Ruff lint/format；16 个核心文件非导入 AST 保持一致 |
| 附加兼容性修复 | 验收内置段落样式时规范化 python-docx 内部/UI 别名，自定义样式继续精确比较 | `1fea324` | `test_builtin_style_alias_round_trips_without_case_folding_custom_names` |

页眉内容与页码当前使用两个独立段落；同一行排版和混合段落的自动改写不属于本轮契约。
未选 XML 子树仍在操作前后重新序列化，以检测 writer 对范围外内容的意外修改。
缓存优化保留逐操作属性范围检查和输出保存后的独立读回验收。

## 最终本地验证

代码验证版本：`e2873b4`；环境为 Windows、Python 3.14。

- `python -m unittest discover -s tests -v`：74 项通过，完整运行约 200.6 秒。
- `python scripts/smoke_test.py`：DOCX 冒烟测试通过，覆盖编号、引用、页码、目录、页眉、模块和三线表等链路。
- `npm run check` 与 `npm test`：JavaScript 检查和 Playwright 桌面/移动端回归通过。
- `python scripts/validate_project.py`：60 个必需文件、1 个 preset、0 个 template workflow 通过。
- Ruff lint、format check 与 `git diff --check` 均通过。

复现命令见 README 的开发验证流程。

GitHub 对 `e2873b4` 的 [CI 运行](https://github.com/xiaoyang732/word-format-master/actions/runs/38393480478)
已完成并通过：Python 3.10、3.11、3.12、3.13、3.14 五个任务，以及 Dashboard UI 和 Ruff 两个任务。

## 性能实测

```powershell
python scripts/benchmark_formatting.py --output .tmp/final-performance.json
python scripts/benchmark_formatting.py --profile --output .tmp/final-performance-profile.json
```

| 样例 | 实测 |
|---|---|
| 合成论文大小 | 96 段 |
| 当前规范展开后的操作数 | 102 |
| 普通执行 `apply_plan` | 6.471 秒 |
| 开启 cProfile 的 `apply_plan` | 7.871 秒，1,999,959 次函数调用 |
| 输出独立复核 | 两次均 `passed` |

原报告基线为约 52.5–54.2 秒、约 3005 万次调用；当前结果显著减少了调用开销。
本轮早前的 profile 实测为 5.296 秒、2,025,823 次调用，显示运行环境和规范调整会影响结果。
上述数字不是同一时刻的受控速度比，不能外推到全部文档。
报告提出的 2–2.5 秒目标尚未达成；当前主要热点仍是原子投影和有效属性读回。

## 页面方向的实际渲染证据

在四节合成文档中分别验证：只设横向、横向加明确 A3 尺寸、横向恢复纵向、未修改的纵向节。
核心验收通过，OOXML 读回尺寸依次为 297×210、420×297、210×297、210×297 mm
（按 0.1 mm 记录，允许 OOXML twip 换算误差）。

本机 `Word.Application` COM 实际对应 WPS Office `12.1.0.25860`。
渲染生成四页，Poppler 对 PDF 的独立尺寸检查依次为：

| 页 | PDF 尺寸（pt） | 方向 |
|---|---|---|
| 1 | 841.9 × 595.3 | A4 横向 |
| 2 | 1190.55 × 841.9 | A3 横向 |
| 3 | 595.3 × 841.9 | A4 纵向 |
| 4 | 595.3 × 841.9 | A4 纵向 |

四页 PNG 均已检查，未发现裁切、重叠、内容丢失或额外空白页。
本地临时证据保存在 `.tmp/geometry-review/`，未提交合成 DOCX、PDF 或页面图片。
Microsoft Word 和 LibreOffice 在本机均不可用，因此报告要求的两种渲染器分别验收尚未完成。
WPS 结果仅证明上述合成样例在该渲染器上的页面几何行为。

## 验收边界

本轮代码缺陷已补充回归并修复；性能目标及 Microsoft Word/LibreOffice 的跨渲染器验收仍保留为后续项目。
所有自动化、性能和渲染结果使用合成样例，尚未形成真实学校模板的完整视觉验收证据。
结构性 `passed` 或 `no_change` 只说明对应的代码检查成立，真实文档交付仍须按选定渲染器检查分页与域刷新结果。

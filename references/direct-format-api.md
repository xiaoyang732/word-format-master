# Direct Format API 2.0

两种入口共享 `scripts/word_format`。AI 负责把人类要求转换为结构化请求；格式引擎不调用大模型、不操作浏览器，也不执行动态生成的 Python 代码。

页眉内容、动态章节标题、奇偶页、分节链接和格式的逐项代码对应见 [页眉契约](header-requirements.md)。支持 `header.*` 独立操作和显式 `section.break.next_page.insert`；新建节后须重新 inspect 并生成页眉计划。

## 命令

```powershell
python scripts/format_cli.py capabilities --output capabilities.json
python scripts/format_cli.py inspect source.docx --output inspect.json
python scripts/format_cli.py plan source.docx result.docx --request request.json --output plan.json
python scripts/format_cli.py apply --plan plan.json --report report.json
python scripts/format_cli.py verify --plan plan.json --report report.json
```

命令失败退出码为 2，并返回 `failed`、`unsupported` 或 `needs_clarification`。已有输出默认拒绝覆盖；只有显式传入 `--overwrite` 才可替换输出。源文件和它的硬链接始终不能作为输出。JSON 路径也不得覆盖源文档、输出文档、请求或计划。

## 请求示例

只把一个段落中的“重要结论”改成 14 磅，不修改其他文字：

```json
{
  "schema_version": "2.0",
  "origin": "direct",
  "operations": [
    {
      "action": "font.size.set",
      "target": {"type": "text_range", "id": "p12", "quote": "重要结论"},
      "params": {"value": 14, "unit": "pt"}
    }
  ]
}
```

段落、章节和整篇的选择示例：

| 范围 | target |
|---|---|
| 一个段落 | `{"type":"paragraph","id":"p12"}` |
| 多个已检查段落 | `{"type":"paragraph","ids":["p12","p13"]}` |
| 某章所有正文 | `{"type":"paragraph","chapter_text":"第三章 实验","role":"body","all":true}` |
| 某章第 2 个正文段落 | `{"type":"paragraph","chapter_id":"p10","role":"body","ordinal":2}` |
| 整篇正文 | `{"type":"paragraph","role":"body","all":true}` |
| 所有顶层段落 | `{"type":"paragraph","all":true}` |
| 表格第 1 行第 2 格文字 | `{"type":"paragraph","table_id":"t0","row":1,"cell":2,"all":true}` |
| 第 1 节页面 | `{"type":"section","number":1}` |
| 指定共享样式 | `{"type":"style","name":"Heading 1"}` |

`ordinal` 是过滤后非空段落的 1 起始序号。`p0` 等 ID 包含空段落和表格段落，必须来自当前 inspect，不应由 AI 猜测。未明确 ID、part、table_id 的普通段落选择只包含顶层段落。字符 start/end 为当前段落可见文字的 0 起始、左闭右开区间；重复 quote 必须给 1 起始 occurrence。存在域、绘图、制表符、换行或修订的复杂文字范围会拒绝。页面位置选择需要渲染锚点，目前不支持。

## 能力与代码对应

`capabilities` 返回每项操作的 action、handler、validator、reader、verifier、参数、允许范围和测试文件。注册表位于 `registry.py`，入口只能执行已登记 action。

| 格式 | 操作 | 唯一写入职责 |
|---|---|---|
| 中文/西文字体 | `font.east_asia.set` / `font.latin.set` | `operations.py` 中各自 handler → `properties.write_font` |
| 字号、粗体、斜体、颜色 | `font.size.set` / `font.bold.set` / `font.italic.set` / `font.color.set` | 独立 handler → 指定 rPr 属性 |
| 对齐、段前、段后 | `paragraph.alignment.set` / `paragraph.spacing_before.set` / `paragraph.spacing_after.set` | 独立 handler → 指定 pPr 属性 |
| 行距 | `paragraph.line_spacing.set` | multiple / exact / at_least，模式与单位一并校验 |
| 首行、左、右、悬挂缩进 | `paragraph.*_indent.set` | chars 或 mm 单一权威值；首行与悬挂互斥 |
| 段落分页控制 | `paragraph.keep_with_next.set` 等 | 每项只写自己的布尔属性 |
| 页面尺寸、边距、页眉页脚距离 | `page.*.set` | 每项只写所选节的对应属性 |
| 表格对齐、标题行重复 | `table.alignment.set` / `table.header_repeat.set` | 所选表格属性 |
| 三线表 | `table.three_line.apply` | 保存原始边框后应用；false 只恢复本工具保存的边框 |
| 样式指派 | `paragraph.style.assign` | 只接受段落样式；可显式创建新样式 |
| 页眉页脚、页码、目录、编号、题注位置、引用 | 登记的 configure / convert / insert 操作 | 明确列出结构影响，独立验收 |

整篇旧规范可用 `{"spec": {...}}` 替代 `operations`；两者不能同时提供。旧规范会展开成同一注册表的操作。网页提交的 Handoff 经 `apply_spec.py` 的身份检查后也走此适配器。

## 保护与验收

计划包含源 SHA-256、已解析对象指纹、输出路径和计划摘要。它用于检测过期和误改，并非签名或身份认证。

普通格式操作在内存中逐项比较整个包，只屏蔽被请求的属性；文字范围允许拆分普通 run，但检查选区外格式、书签、超链接和对象。结构操作单独检查正文文字、对象和无关部件，并只允许声明的转换。保存到临时文件后回读全部操作，再通过结构验收，最后发布输出。失败不会发布半成品。

结构通过不等于视觉通过。Word 字段计算、最终分页、溢出和字形仍需 `render_docx.py` 与逐页检查。缺少渲染或视觉检查时应报告 skipped / pending，不能声明视觉通过。Strict OOXML 写入、宏文档和复杂范围暂不支持；分节仅支持明确的顶层段落后下一页分节。

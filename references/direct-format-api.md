# Direct Format API 2.0

直接请求与格式档案共享 `scripts/word_format`。AI 负责把人类要求转换为结构化请求；格式引擎不调用大模型、不操作浏览器，也不执行动态生成的 Python 代码。

页眉内容、动态章节标题、奇偶页、分节链接和格式的逐项代码对应见 [页眉契约](header-requirements.md)。支持 `header.*` 独立操作和显式 `section.break.next_page.insert`；新建节后须重新 inspect 并生成页眉计划。

## 命令

命令行的标准输出、错误输出和 JSON 报告文件统一使用 UTF-8。通过管道或子进程读取时，应显式用 UTF-8 解码，不依赖 Windows 的本地代码页。

```powershell
python scripts/format_cli.py capabilities --output capabilities.json
python scripts/format_cli.py inspect source.docx --output inspect.json
python scripts/format_cli.py audit source.docx --request request.json --output audit.json
python scripts/format_cli.py plan source.docx result.docx --request request.json --output plan.json
python scripts/format_cli.py apply --plan plan.json --report report.json
python scripts/format_cli.py verify --plan plan.json --report report.json
```

命令失败退出码为 2，并返回 `failed`、`unsupported` 或 `needs_clarification`。已有输出默认拒绝覆盖；只有显式传入 `--overwrite` 才可替换输出。源文件和它的硬链接始终不能作为输出。JSON 路径也不得覆盖源文档、输出文档、请求或计划。

`audit` 接受与 `plan` 相同的 `operations` 或 `spec` 请求，使用同一位置解析器和独立回读器检查文档，不调用写入 handler、不保存 DOCX、不创建输出文档。报告逐项列出对象 ID、文字、期望值、实际有效值和状态，`summary` 汇总通过/失败项。存在偏差时仍可写出完整 JSON 报告，退出码为 2；通过为 0。未选中可执行检查时为 `no_change`。只检查请求覆盖的属性，视觉状态为 `skipped`；这不等于完整论文合规审查。

`verify` 要求报告中的每一项操作/目标与计划完整对应，包含执行时无需改动的项目。缺失、重复、截短或参数变更的检查列表会被拒绝。目录等结构操作可导致输出对象 ID 位移，报告保留源目标和最终输出目标；有效完整报告可以重新排序检查项。独立复检也重新执行结构验收。

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
| 封面大标题 | `{"type":"paragraph","semantic_role":"cover_title","all":true}` |
| 中文摘要标题 | `{"type":"paragraph","semantic_role":"abstract_heading","language":"zh","all":true}` |
| 英文关键词标签 | `{"type":"text_range","semantic_role":"keywords","language":"en","segment":"label","all":true}` |
| 所有顶层段落 | `{"type":"paragraph","all":true}` |
| 表格第 1 行第 2 格文字 | `{"type":"paragraph","table_id":"t0","row":1,"cell":2,"all":true}` |
| 第 1 节页面 | `{"type":"section","number":1}` |
| 指定共享样式 | `{"type":"style","name":"Heading 1"}` |

`ordinal` 是过滤后非空段落的 1 起始序号。`p0` 等 ID 包含空段落和表格段落，必须来自当前 inspect，不应由 AI 猜测。未明确 ID、part、table_id 的普通段落选择只包含顶层段落。字符 start/end 为当前段落可见文字的 0 起始、左闭右开区间；重复 quote 必须给 1 起始 occurrence。存在域、绘图、制表符、换行或修订的复杂文字范围会拒绝。页面位置选择需要渲染锚点，目前不支持。

旧规范可独立控制关键词标签与词条的粗体，并设置目录标题样式：

```json
{
  "spec": {
    "keywords": {"label_bold": true, "value_bold": false},
    "table_of_contents": {
      "enabled": true,
      "title": "目录",
      "title_style": {
        "font_east_asia": "黑体",
        "font_size_pt": 16,
        "bold": true,
        "alignment": "center"
      }
    }
  }
}
```

`label_bold` 仅作用于段首的 `关键词：` / `Key words:` 标签，`value_bold` 作用于标签后的词条。省略两个字段时不改变现有粗体；不要同时使用整段 `keywords.bold` 与局部粗体字段。`table_of_contents.title_style` 的字号、字体和段落属性会在 DOCX 回读时检查。

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

整篇旧规范可用 `{"spec": {...}}` 替代 `operations`；两者不能同时提供。旧规范会展开成同一注册表的操作。`apply_spec.py INPUT.docx OUTPUT.docx --spec spec.json` 也委托同一核心，无需额外会话。要求文档和样例使用 [Format Profile](format-profile.md) 提取证据，再通过 `plan --profile` 编译。

## 保护与验收

计划包含源 SHA-256、已解析对象指纹、输出路径和计划摘要。它用于检测过期和误改，并非签名或身份认证。

普通格式操作在内存中逐项比较整个包，只屏蔽被请求的属性；文字范围允许拆分普通 run，但检查选区外格式、书签、超链接和对象。结构操作单独检查正文文字、对象和无关部件，并只允许声明的转换。保存到临时文件后回读全部操作，再通过结构验收，最后发布输出。失败不会发布半成品。

结构通过不等于视觉通过。Word 字段计算、最终分页、溢出和字形仍需 `render_docx.py` 与逐页检查。缺少渲染或视觉检查时应报告 skipped / pending，不能声明视觉通过。Strict OOXML 写入、宏文档和复杂范围暂不支持；分节仅支持明确的顶层段落后下一页分节。

## 绑定视觉验收

渲染全部页面后，由支持图片输入的当前 AI 逐页检查，并保存 `visual.json`。用户无需导入报告。报告须包含实际模型 ID、输出 DOCX 的 SHA-256、渲染清单文件的 SHA-256、实际渲染方法，以及全部已检查页码：

```json
{
  "status": "passed",
  "model": {"id": "实际检查页面的模型", "supports_image_input": true},
  "output_sha256": "输出DOCX的实际SHA-256",
  "manifest_sha256": "render-manifest.json文件的实际SHA-256",
  "renderer": {"method": "word"},
  "reviewed_pages": [1, 2],
  "findings": []
}
```

示例只代表两页文档。页码必须完整且不重复，不能用小数或布尔值。若发现缺陷，状态为 `failed`，在 findings 中写明页码与问题；修复并重新渲染后生成新的检查报告。

```powershell
python scripts/render_docx.py result.docx --output-dir render --method word
python scripts/format_cli.py verify --plan plan.json --report report.json --visual-report visual.json --render-manifest render/render-manifest.json --output verification.json
```

`word_format/visual.py` 核对清单来源、版本、哈希、渲染器、页数、每张实际图片及完整检查页码；证据过期、缺页或改动时拒绝。渲染器默认遵守计划中的 `verification.render_method`，可用 `--render-method` 明确指定。结果同时报告 `format`、`structure`、`visual`，视觉失败也使总状态失败。

没有传视觉报告时，`visual.status` 为 `skipped`，即使格式与结构通过，也不能描述为视觉通过。模型图片输入能力未确认时同样为 `skipped`，可附原因而不提供渲染清单。程序校验证据完整性，不能代替模型实际查看页面。

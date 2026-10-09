# 页眉要求与代码契约

参考：[《论文的页眉一般写什么？从通用规范到分章设置的全指南》](https://www.global-meetings.com/headerLine/read/1678)，国际会议云，2026-01-19；核对日期：2026-10-09。

这是一份通用排版建议，不能替代学校、期刊规范或用户提供的模板。项目将其落实为可选能力，不强制启用页眉、不自动猜测论文题目或学位信息，也不把字号、位置建议当作全国统一标准。

## 内容规则

| 文档范围 | 文章建议 | 项目表达与执行 |
|---|---|---|
| 摘要、目录、图表清单 | 论文题目或对应部分名称 | 指定节的 `header.content.set`，`mode: text` |
| 正文章节 | 当前章节标题 | `mode: chapter_title`，以实际段落样式生成 `STYLEREF` 字段 |
| 论文题目与章节组合 | “论文题目：章节标题” | `chapter_title` + 明确的 `prefix`；不自动抽取题目 |
| 仅章节编号 | 如“第3章” | 独立节固定文字；不假装从任意手工编号生成动态编号 |
| 参考文献、致谢、附录 | 对应部分名称 | 指定节固定文字，或实际使用同一一级标题样式的动态标题 |
| 学校、机构、学位、姓名、学号 | 仅在规范明确要求时写入 | 明确的固定文字，不猜测这些信息 |
| 双面打印 | 例：奇数页章节名、偶数页论文题目 | 全文奇偶页开关 + `default`/`even` 内容分别设置 |

字体建议为宋体、Times New Roman；正文小四时，页眉可选五号（10.5 pt）或小五（9 pt）。居中或外侧对齐均为可选。外侧对齐用普通/奇数页 `right`、偶数页 `left`。页眉距页面顶端的具体毫米数由规范/模板/用户提供，文章未给出固定值。

## 每项设置的唯一代码与验收

所有下列 action 均在 `scripts/word_format/registry.py` 注册 validator、handler、reader、verifier 和测试路径。`headers.py` 负责节和页型，字体/段落写入复用 `properties.py`，不引入另一套 XML 属性写法。

| 要求/参数 | action | handler（headers.py） | 回读验收 |
|---|---|---|---|
| 固定文字、模块名、题目 | `header.content.set` | `set_content` | literal 与字段列表完全匹配 |
| 当前章标题、标题前缀 | `header.content.set` | `set_content` | `STYLEREF` 指令、样式名、前缀匹配；不验证未计算的显示结果 |
| 链接/取消链接到前一节 | `header.link_previous.set` | `set_link` | headerReference 的实际继承状态 |
| 所选节首页不同 | `header.first_page_different.set` | `set_first_page` | `w:titlePg`；同时影响该节页脚显示 |
| 全文奇偶页不同 | `header.odd_even_different.set` | `set_odd_even` | settings 的 `w:evenAndOddHeaders`；同时影响页脚显示 |
| 中文字体 | `header.font.east_asia.set` | `set_east_asia` | `read_font` 有效中文字体 |
| 西文字体 | `header.font.latin.set` | `set_latin` | `read_font` 有效西文字体 |
| 字号 | `header.font.size.set` | `set_size` | 磅值，5–96 pt，半磅精度 |
| 粗体，`value: boolean` | `header.font.bold.set` | `set_bold` | 每个 run 的有效粗体 |
| 斜体，`value: boolean` | `header.font.italic.set` | `set_italic` | 每个 run 的有效斜体 |
| 颜色，`value: RGB hex / auto` | `header.font.color.set` | `set_color` | 每个 run 的有效颜色 |
| 左/中/右对齐 | `header.alignment.set` | `set_alignment` | 每个页眉段落的有效对齐 |
| 段前，`value: 0–500, unit: pt` | `header.spacing_before.set` | `set_spacing_before` | 指定磅值；段后不变 |
| 段后，`value: 0–500, unit: pt` | `header.spacing_after.set` | `set_spacing_after` | 指定磅值；段前不变 |
| 行距，`kind: multiple / exact / at_least` | `header.line_spacing.set` | `set_line_spacing` | 倍数 0.1–20；固定/最小值 1–500 pt |
| 首行缩进，`value + unit: chars / mm` | `header.first_line_indent.set` | `set_first_indent` | 0–100 chars 或 0–500 mm；与悬挂互斥 |
| 左缩进，`value + unit: chars / mm` | `header.left_indent.set` | `set_left_indent` | 0–100 chars 或 0–500 mm |
| 右缩进，`value + unit: chars / mm` | `header.right_indent.set` | `set_right_indent` | 0–100 chars 或 0–500 mm |
| 悬挂缩进，`value + unit: chars / mm` | `header.hanging_indent.set` | `set_hanging_indent` | 0–100 chars 或 0–500 mm；与首行互斥 |
| 与下段同页，`value: boolean` | `header.keep_with_next.set` | `set_keep_next` | 对应有效布尔属性 |
| 段中不分页，`value: boolean` | `header.keep_together.set` | `set_keep_together` | 对应有效布尔属性 |
| 段前分页，`value: boolean` | `header.page_break_before.set` | `set_page_break` | 对应有效布尔属性 |
| 孤行控制，`value: boolean` | `header.widow_control.set` | `set_widow` | 对应有效布尔属性 |
| 页眉距页面顶端 | `page.header_distance.set` | `operations.set_header_distance` | 所选节 `w:pgMar/@w:header`，mm |
| 下一页分节符 | `section.break.next_page.insert` | `section_breaks.insert` | 指定段落结束节 + 下一节 nextPage；正文与其余包部件不变 |

页型统一用 `variant: default | first | even`，缺省为 `default`。除奇偶页开关外，页眉操作指定 `section` target。`read_header` / `verify_header` 是统一回读接口；正文、所有页脚、未选页型、其他节有效页眉和未请求属性由 `header_projection` 额外保护。后续继承页眉会先复制隔离，防止本次修改扩散。

内容操作会替换所选页型的整个页眉，包括已有文字、域、图片和表格；沿用首段段落属性与首个 run 的字体属性。只改字体/字号应使用对应原子操作，保留正文、页眉内容与对象。新页眉不自动套用任何字体、字号或对齐默认值。

## 直接请求

以下操作把第 2 节的普通/奇数页设为章节标题，偶数页设为论文题目，字号小五，向外对齐：

```json
{
  "schema_version": "2.0",
  "origin": "direct",
  "operations": [
    {"action":"header.odd_even_different.set","target":{"type":"document"},"params":{"value":true}},
    {"action":"header.content.set","target":{"type":"section","number":2},"params":{"mode":"chapter_title","style_name":"Heading 1"}},
    {"action":"header.content.set","target":{"type":"section","number":2},"params":{"mode":"text","text":"用户提供的论文题目","variant":"even"}},
    {"action":"header.font.size.set","target":{"type":"section","number":2},"params":{"value":9,"unit":"pt"}},
    {"action":"header.font.size.set","target":{"type":"section","number":2},"params":{"value":9,"unit":"pt","variant":"even"}},
    {"action":"header.alignment.set","target":{"type":"section","number":2},"params":{"value":"right"}},
    {"action":"header.alignment.set","target":{"type":"section","number":2},"params":{"value":"left","variant":"even"}}
  ]
}
```

`Heading 1` 必须是实际存在且用于正文标题的段落样式。字段采用 Word 的 STYLEREF 查找规则，按当前页面匹配样式；小节标题不应使用同一个样式。模块标题若未采用该样式，应先明确指派样式，或用分节固定文字。Word/LibreOffice 刷新字段和全页检查后才能确认最终显示。

需分节时，先单独向明确段落执行 `section.break.next_page.insert`（`params: {}`），生成中间 DOCX，再 inspect 新文件取得节编号并设置页眉。不得猜测“第三章就是第三节”，不得在旧计划里引用新建节；分节会改变分页，必须明确请求。已有边界重复插入、表格/复杂容器内插入会拒绝。

## 网页与旧规范

网页可选择目标节（空值表示全部）、固定/动态内容、字体字号、距离、首页和偶数页配置。网页仅示意普通页眉；首页、偶数页和分节分页以实际渲染为准。

旧规范新增可选字段：

```json
{
  "headers_footers": {
    "preserve_existing": true,
    "different_odd_even": true,
    "header": {"enabled": true, "mode": "chapter_title", "style_name": "Heading 1", "font_size_pt": 9, "alignment": "right"},
    "even_header": {"enabled": true, "text": "论文题目", "font_size_pt": 9, "alignment": "left"},
    "sections": [
      {"number": 1, "header": {"enabled": true, "text": "摘要"}},
      {"number": 3, "header": {"enabled": true, "text": "参考文献"}}
    ]
  }
}
```

`sections` 中同页型设置优先于全局设置，避免两套规则互相覆盖。`section_number` 可把全局页眉设置限制到一节；全文奇偶页开关仍是全文级。`first_header` 定义每节首页内容；启用首页不同后，设空文字可隐藏首页页眉。

`enabled: false` 的旧全局 `header` 保留“只移除本工具管理文字”的语义；指定节或显式 `header.content.set` 空文字则明确清空该页型。关闭某一页型配置不会自动关闭 Word 的页型开关。所选节模式禁止 `preserve_existing: false` 的全局清理。

模板提取分别记录各节/页型及链接状态，不把不同页眉连接成一个字符串；仅将单段“前缀 + `STYLEREF "样式名"`”提取为动态模式。字段后缀、多段或含对象的组合、复杂字段开关需要保留并报告不支持，不能把后缀移到字段之前。

## 页码与验收边界

文章建议前置部分罗马页码、正文阿拉伯页码、后置部分续编号。该规则是页码要求，与页眉独立；本次页眉操作不修改编号制式、起始页码或页脚内容。当前页码配置仍按已有文档级契约工作，不声称已支持逐节罗马/阿拉伯自动转换。

测试位于 `tests/test_headers.py`：节间隔离、奇偶/首页、字段、局部属性保留、超链接关系、网页同核心展开、显式分节、错误参数和越界写入。结构通过不能替代真实 Word/LibreOffice 全页视觉验收。

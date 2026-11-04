# 段落属性支持矩阵

本表区分原始 XML 留存、有效语义读取和可执行迁移。支持范围以注册操作为准；
不声称覆盖 Word 的所有段落设置或任意 OOXML 扩展。

| 属性 | 原始 XML 留存 | inspect / profile 有效语义 | 可生成规则 | 注册写入 | 保存后独立回读 | 页面验证 |
|---|---|---|---|---|---|---|
| 对齐、段前段后、行距、字符/mm 缩进 | 是 | 是，保留模式和单位 | 是，无法表示或不一致时为 issues | 是 | 是 | 需另行渲染逐页检查 |
| 与下段同页、段中不分页、段前分页、孤行控制 | 是 | 是，含继承及显式 false | 是 | 是 | 是 | 需另行渲染逐页检查 |
| 上下文段距 `contextualSpacing` | 是 | `contextual_spacing: boolean` | 是 | `paragraph.contextual_spacing.set` | 是 | 需另行渲染逐页检查 |
| 文档网格对齐 `snapToGrid` | 是 | `snap_to_grid: boolean` | 是 | `paragraph.snap_to_grid.set` | 是 | 需另行渲染逐页检查 |
| 镜像缩进 `mirrorIndents` | 是 | `mirror_indents: boolean` | 是 | `paragraph.mirror_indents.set` | 是 | 需另行渲染逐页检查 |
| 自定义制表位 `tabs` | 是 | 对齐、整数 twip 位置、引导符；合并继承并应用 clear | 是 | `paragraph.tabs.set` | 是，含引导符及完整位置集合 | 需另行渲染逐页检查 |
| 其他未注册节点、上述四项的扩展属性、无法表示或非法的上述设置 | 快照留存；无关局部修改保留 | 仅保留原始证据或无法表示的值 | 不自动生成，profile issues 要求明确处理 | 无通用任意 XML 写入 | 不宣称可迁移验证 | 未验证 |

三个布尔操作均使用 `{"value": true/false}`，作用于 paragraph 或段落 style。
有效读取会依次考虑直接格式、编号定义、样式链及文档默认值，显式 false 会覆盖继承。
未指定上下文段距、镜像缩进时为 false，未指定文档网格对齐时为 true。

制表位操作示例：

```json
{
  "action": "paragraph.tabs.set",
  "target": {"type": "paragraph", "id": "p1"},
  "params": {
    "value": [
      {"alignment": "right", "position": 1441, "unit": "twips", "leader": "dot"}
    ]
  }
}
```

`value` 是所选对象的完整有效自定义制表位集合，空列表清除已有及继承的自定义制表位；
不会取消 Word 的默认制表间隔。写入时以 `clear` 屏蔽未请求的继承位置。
位置为 -31680 至 31680 的整数 twip，避免经 mm/pt 换算造成舍入损失；不接受重复位置。
`alignment` 支持 `left/center/right/decimal/bar/num/start/end`；
`leader` 支持 `none/dot/hyphen/underscore/heavy/middleDot`。
`unit` 可省略且默认 twips，`leader` 可省略且默认 none。读取结果始终包含两者。
`clear` 是读取及写入继承所用的 OOXML 指令，不是完整集合中可提交的有效制表位。

`analyze_docx.parse_paragraph_properties` 同样读取三个布尔属性和制表位引导符，
其 `tabs` 保留该 XML 层中的 clear 指令；inspect 的 `tabs` 是合并后的有效集合。
分析器的样式有效值与 `inferred_spec` 样式 tokens 也包含这四项，并遵循制表位继承和 clear。
Profile 为正文区域生成对应规则，同时保留原始 pPr；未知属性和非法值仍产生 issues，
不能把已保留 XML 等同于已支持迁移。

旧式 spec 的正文、标题等段落 tokens 也接受 `contextual_spacing/snap_to_grid/mirror_indents/tabs`。
专用 `header.contextual_spacing.set/header.snap_to_grid.set/header.mirror_indents.set/header.tabs.set`
也提供同样的值模型，支持节 target 和 `default/first/even` 页型；后续继承页眉会隔离，
保存后独立回读所选页型并检查其他页眉页脚。页眉 spec 会展开为这些已注册 handler。
自动测试见 `dev/tests/test_paragraph_extended.py`，覆盖语义读取、样式继承、profile 迁移、
原始整数位置、非法请求无输出、错误 writer 拒绝及目录标题的继承制表位替换。

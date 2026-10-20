# Format Profile 1.0

格式档案把来源证据、支持的规则和未决项放在一起，供 AI 审阅并编译成现有格式计划。它不是另一个 DOCX 写入器，也不代表任意 Word 格式都能自动迁移。

## 提取入口

```powershell
python scripts/format_cli.py profile requirements "学校要求.docx" --output profile.json
python scripts/format_cli.py profile template "样例.docx" --output profile.json
```

`requirements` 支持 UTF-8 TXT/MD、DOCX/DOTX 和可提取文字的 PDF。DOCX 保留正文、表格行/单元格、页眉页脚、脚注尾注、批注及可读文本框的文字锚点。PDF 保存页码和页面范围；这不是文字级坐标或 Word 格式推断。扫描页产生 OCR 问题；混合图片的 PDF 页面和含图形/对象的 DOCX 同样会提示视觉复核，避免遗漏截图中的要求。Markdown 标记和说明性文字也可能进入未解析问题，AI 需结合完整原文处理。

`template` 支持 DOCX/DOTX。DOTX 在内存中转换读取，不改原文件。提取正文与表格段落的有效字体/字号/粗斜体/颜色、对齐/段距/行距/缩进/分页控制，按角色、模块、标题级别分组。全部页面节单独记录尺寸、方向、边距、装订线及页眉页脚距离。首页开关按节，奇偶页开关属于全文。表格提取对齐、首行重复及可识别三线表。

## 档案结构

以下为简化示例，`source.sha256` 必须是真实文件摘要：

```json
{
  "schema_version": "1.0",
  "kind": "word-format-profile",
  "source": {"kind": "requirements", "path": "C:/Documents/要求.txt", "sha256": "..."},
  "rules": [
    {
      "id": "r1",
      "target": {"type": "paragraph", "role": "body", "all": true},
      "action": "font.size.set",
      "params": {"value": 12, "unit": "pt"},
      "evidence": [{"anchor": "line:1", "quote": "正文小四", "confidence": 1.0}]
    }
  ],
  "issues": [],
  "coverage": {"represented_rules": 1, "read_blocks": 1, "unresolved": 0},
  "snapshot": {"blocks": [{"anchor": "line:1", "text": "正文小四"}]}
}
```

来源 `kind` 为 `requirements` 或 `template`。规则必须使用 `capabilities` 登记的 action、参数和目标类型。置信度低于 1 的规则需在 `mapping.rules` 中明确目标或注明保留，否则不会自动执行。提取值保留来源锚点，未定义值不会被默认为零或 false。

- `rules`：可表示的执行候选；提取成功不等于已经应用。
- `issues`：同属性冲突、混合有效值、未解析文字、特殊结构或其他不确定项。冲突保留全部 `candidates`，无法表示的观测值放在 `observed`。
- `coverage`：规则数、问题数、读取区块/段落或节数及策略说明。这些计数不是“论文合规百分比”。
- `snapshot`：要求原文；样例则保存段落索引/有效格式/直接 XML、样式、节、包特征、字段和编号等证据。

混合字号或局部强调不能用多数值自动替代。未支持的表格几何、复杂编号、页眉页脚内容、字段、栏及特殊 OOXML 等进入 `issues`。样例中的姓名、题目、正文、引用和页眉专属文字不会复制到目标。

## 问题决策

`resolutions.json` 以问题 ID 为键。决策必须附理由：

```json
{
  "i1": {"decision": "choose", "candidate": 0, "reason": "用户明确选用第一项字号要求"},
  "i2": {
    "decision": "replace",
    "reason": "原文要求一级标题居中，结合上下文明确目标",
    "rules": [
      {
        "action": "paragraph.alignment.set",
        "target": {"type": "paragraph", "role": "heading", "level": 1, "all": true},
        "params": {"value": "center"}
      }
    ]
  },
  "i3": {"decision": "ignore", "reason": "样例未定义该属性，按用户要求保留目标原值"}
}
```

`choose.candidate` 是从 0 开始的候选下标。`replace.rules` 为非空规则数组，编译后 ID 为 `i2-resolved-1` 等。`ignore` 不生成操作，原问题和理由仍在报告中；不能说成已迁移或已验证的要求。示例 ID 必须换成实际档案中的问题 ID。

AI 可依据材料和用户已明确的偏好作决定。需要用户判断的冲突应澄清，不能批量忽略全部问题来强行执行。

## 目标映射

先运行 `inspect` 检查真实目标角色和 ID。`mapping.json` 允许：

```json
{
  "sections": {"1": 1, "2": 3},
  "rules": {
    "r1": {"type": "paragraph", "ids": ["p4", "p5"]},
    "r2": {"preserve": "样例有附录但本目标论文不需要该角色"}
  },
  "preserve": {"p0": "用户要求封面保持目标原样"}
}
```

- `sections`：样例节编号字符串 → 目标节编号；多节样例必须显式覆盖每一节。不会自动创建节。单节样例默认作用于目标第 1 节，需要时也可显式映射。
- `rules`：规则 ID → 新目标选择器，或带理由的 `preserve`。选择器直接指向目标文档，不再套用 `sections` 转换；表格按序号提取，内容/顺序不对应时必须修改映射。
- `preserve`：明确保留的目标段落 ID → 理由。样例模式中，非空正文部件段落必须被规则覆盖或明确保留。这里的覆盖只表示角色/目标已映射，并不证明每个属性完整迁移。

段落角色由共享检测器推断；AI 必须核实标题、正文、摘要、参考文献、题注等实际含义。缺少角色、重复目标、样例不均匀或模块差异不能静默忽略。

## 生成计划与验收

```powershell
python scripts/format_cli.py plan "目标.docx" "输出.docx" --profile profile.json --mapping mapping.json --resolutions resolutions.json --output plan.json
python scripts/format_cli.py apply --plan plan.json --report report.json
python scripts/format_cli.py verify --plan plan.json --report report.json --output verification.json
```

没有问题时可省略决策文件；没有映射时可省略映射文件。`--request-text` 保存人类原始请求；`--request` 接收另一种入口的请求 JSON，不能与 `--profile` 同用。

未决项或目标不能匹配时，`plan` 返回退出码 2 和 `word-format-profile-review`，不生成可执行计划、不写 DOCX。解决后重新运行。成功时生成标准 schema 2.0 计划，操作仍经过 registry 和 `build_plan()`。

计划和执行报告的 `profile_context` 保存来源、档案摘要、提取覆盖计数、逐规则目标/证据、原问题与决策、保留规则和保留段落。格式依据或目标原稿变化后执行拒绝；输出也不能覆盖格式依据。

`verify` 复核每项计划的写入结果和结构。完整符合性还需 AI 对照原文检查覆盖范围，再通过 Word/LibreOffice 渲染逐页检查。未支持项、忽略项及视觉未完成必须列明，不能用结构通过替代所有格式符合。

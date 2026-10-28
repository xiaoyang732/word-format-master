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

样例档案还保存 `snapshot.structure_contract`：按原顺序记录模块、摘要/关键词语言、声明类型、重复次数、章节层级及证据锚点。中文摘要与英文摘要分别记录；诚信承诺、原创性声明、授权书也分别记录。表格或图片位于正文模块之前时保留封面证据；首页题目区记录为 `title_block`，分页封面记录为 `cover`，`separate_page` 保留分页/分节边界，防止擅自把首页题目另做成独立封面。检测结果需要 AI 对照完整样例和页面核实，尤其是自动分页或特殊封面，不能把未识别内容视为不存在。

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

上面的保留示例适用于保留原稿内容的排版任务。按样例生成新论文时，结构必须与样例一致，不能用 `rules.*.preserve`、`mapping.preserve` 或忽略格式问题来豁免缺少/多出的模块。

## 按样例生成论文

结构规则为：**样例有什么就生成什么，不多也不少**。有封面就生成封面；没有英文摘要、英文关键词、诚信承诺书、目录、致谢或附录，就不自动补加。已有模块按样例保留种类、数量、语言和先后顺序，章节层级也保持一致，具体题目、姓名与研究内容替换为新论文的信息。

```powershell
python scripts/format_cli.py check-structure "草稿.docx" --profile profile.json --output structure.json
python scripts/format_cli.py plan "草稿.docx" "输出.docx" --profile profile.json --match-structure --mapping mapping.json --resolutions resolutions.json --output plan.json
```

`check-structure` 只读比较，返回 `expected`、`actual` 和 `differences`；有差异退出码为 2。它重新读取实际样例并检查来源指纹，不能修改快照来掩盖差异。`--match-structure` 用于新论文生成，规划前拒绝结构差异，并在 `profile_context.structure_contract` 和 `verification_spec.document_structure` 中绑定样例结构；写入验收和独立复检会再次核对。该模式遵循样例的实际顺序，不强加通用论文顺序。

工具不会自行编写缺失封面或研究正文，AI 应先按结构清单完成草稿。封面布局、特殊模块和识别歧义还需核对原始内容及逐页渲染，不能仅凭检测结果宣称完全一致。仅做既有论文排版时不默认启用结构重建，保留原稿正文；详见 [论文生成流程](thesis-generation.md)。

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

## 模板格式必须按实际区域取证

`inspect` 为可确认的区域补充 `semantic_role`，不替换原有 `role/module`：
`cover_title`、`cover_subject`、`cover_student`、`cover_college`、`cover_major`、
`cover_advisor`、`cover_student_id`、`cover_date`、`abstract_title`、
`abstract_heading`、`abstract_body`、`keywords`、`reference_heading`。
摘要与关键词使用 `language: zh/en` 区分。无法可靠识别的区域保留原角色，需明确映射。

模板规则按这些区域分组；关键词字体以 `text_range` 的 `segment: label/value`
分别提取和对应，段落属性仍作用于整段。空词条不提供内容字体依据。空摘要占位段的
现有格式保存在快照中并产生待核实项，不能把紧邻标题遗留的格式当作摘要正文要求。

书面要求按显式目标分句，段落之间不沿用上段目标。“加黑”表示粗体，“黑体”表示
字体家族。独立材料（例如优秀论文简介）的条款保留适用范围问题，未决前不并入主论文。
未解析语义仍需逐条处理，不能靠降低问题数量或批量忽略来接受结果。

模板、书面规范与用户要求并存时，先建立“区域—属性—来源锚点—最终采用值”清单，
按用户明确指定 > 模板 > 书面要求处理；模板缺失或仅有空占位时回查规范。
单一 profile 不会自动合并多个来源。合并需记录有依据的 resolution/mapping，
或形成逐项可追溯的注册操作请求，且保留全部来源档案。
`profile_context.coverage.target_properties` 列出每个目标实际检查的属性；
一个段落有任意规则命中不代表全部格式已经覆盖。

段前/段后若使用行单位，以 `{"value": {"value": 0.5, "unit": "lines"}}`
写入对应 spacing 操作；保留 `beforeLines/afterLines`，不按固定字号换算成磅。
点单位仍用数值 `value`。模板中超过可用页宽的首行缩进会产生问题项，
需核对实际页面后决定，不能当作普通规则无条件迁移。

目录写入会为已确认的标题补充实际大纲级别，保留其字体、样式和编号。
完整且独立的旧 TOC 域可替换；跨段域、混有正文或携带分节符的旧目录拒绝自动删除。
验收须有目录范围内的真实标题，并逐项匹配缓存结果，空缓存、占位提示及
“未找到目录项”等错误不能通过。缓存只含标题文字，真实页码仍需 Word/WPS/LibreOffice
更新域并经渲染核对。

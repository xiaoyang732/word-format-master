# 模板驱动排版引擎（1.1）

格式执行代码只解释规则；字体、字号、行距、页边距、编号外观和封面布局放在模板数据中。用户模板不继承内置模板，未提及的属性保留原值。学校之间不通用的参考文献字体、悬挂缩进、章节编号外观也属于模板，不能作为全局常量。

## 入口

从零生成的内容先保存为 DOCX 初稿，再与已有论文走相同的 plan/apply/verify 流程。仅在用户选择内置模板时使用内置 ID：

```powershell
python scripts/format_cli.py template list
python scripts/format_cli.py plan input.docx output.docx --template thesis-standard --output plan.json
python scripts/format_cli.py apply --plan plan.json --report apply.json
python scripts/format_cli.py verify --plan plan.json --report apply.json --output verification.json
python scripts/format_cli.py audit input.docx --template thesis-standard --output audit.json
```

用户带样例、书面规范或两者时，先完整提取证据，再编译独立模板：

```powershell
python scripts/format_cli.py profile template sample.docx --output sample-profile.json
python scripts/format_cli.py profile requirements requirements.docx --document-kind thesis --output requirements-profile.json
python scripts/format_cli.py template compile --profile sample-profile.json --profile requirements-profile.json --id school-thesis --output school-template.json
python scripts/format_cli.py template validate school-template.json
```

编译结果包含 rules、issues、sources、evidence。每项规则保留原文位置、置信度和源文件 SHA-256；完整快照包含直接格式、继承格式、样式 XML、页面、编号和复杂对象信息。同一目标属性冲突时生成 issue，禁止按文件顺序覆盖。解析器不能可靠表达的格式必须留作问题，不能换成默认模板。

## 决策与目标映射

`template compile --resolutions decisions.json` 支持按 issue ID 指定：

- `{"decision":"choose","candidate":0,"reason":"规范明确覆盖样例"}`：选择证据候选。
- `{"decision":"replace","rules":[...],"reason":"从条款补充精确规则"}`：以注册操作表达证据。
- `{"decision":"not_applicable","reason":"此条属于另一个文种"}`：记录不适用依据。
- `{"decision":"ignore","reason":"本轮尚未支持"}`：记录限制，最终只能报告 partial。

解决后的模板通过 `plan --template school-template.json --mapping mapping.json` 绑定到目标。mapping.rules 按规则 ID 指定目标选择器；mapping.sections 使用 `s1:2` 表示第一个证据源的第二节，值为目标节号；单来源可直接使用 `2`。`mapping.preserve` 把目标段落 ID 映射到保留理由。样例模式下有文字的目标段落必须被规则覆盖或明确保留；低置信度、缺失规则目标、未映射的多节布局都会阻止执行。

所有写入仍经过 capability registry。不能用手写 python-docx 作为样例迁移的第二写入引擎。原稿、样例、规范和模板 JSON 均禁止作为输出覆盖。

## 固定结构与可变外观

- 正文章标题绑定 Word Heading N 样式与大纲层级；章节序号使用 `heading.scheme.configure` 的原生 numPr/numbering.xml，不把“第一章”写进文本。
- 摘要、参考文献、致谢等模块标题使用独立标题样式与大纲层级，默认不参与章节序号。目录标题的大纲级别为正文，避免目录收录自身。样例要求模块带编号时必须先补充显式规则或记录限制。
- 参考文献标题与条目是独立语义角色；文献内容与引用事实不在格式任务中重写，GB/T 7714 也不意味着所有学校使用同一种字体。
- 图片、公式、含图文混合对象不作为普通正文套固定行距；需要修改时使用显式目标规则。
- 中文摘要、英文摘要、关键词标签和关键词内容分别选择，避免整段套用同一个字符格式。

## 模板维护与兼容

`templates/builtin/<id>/template.json` 是新入口的内置格式来源。封面使用独立 cover.json，`word_format.cover.render_cover` 只接收布局和真实数据，不填入虚构姓名、题目或日期。封面渲染是新稿构建工具，不能用于重建已有论文的封面。

`assets/presets.json` 和 `templates/compatibility/heading-numbering.json` 保留为旧 spec API 的冻结兼容数据。新功能不再扩展旧枚举，不从兼容文件为用户模板补值。旧示例生成脚本只演示内容构建，最终排版统一调用模板入口。

## 验收状态

verify 分开报告 execution_status 与 compliance。格式回读和对象保护通过，但没有绑定逐页视觉报告时，完整符合状态为 pending；明确跳过要求为 partial；执行或视觉失败为 failed。只有覆盖完整且视觉验收通过才是 passed。渲染成功本身不能代替视觉检查。

模板、来源或规则发生变化必须重新规划。执行时从原始模板重新生成预期操作和覆盖账本，删除操作后重新计算计划哈希也无法绕过完整性检查。

当前提取器会完整记录复杂版式，但自动执行不覆盖任意 OOXML 属性：编号级别的特殊字形/几何、复杂文本框、脚注尾注、图像中的要求等仍会产生显式问题。不得把“已提取”描述为“已完整迁移”。

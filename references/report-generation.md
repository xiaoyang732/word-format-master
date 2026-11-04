# 默认报告生成流程

当用户要求新写一份报告、课程作业报告、课题报告且未指定模板或详细格式要求时，使用项目登记的通用课程报告模板。入口读取 `templates/reports.json` 的 `default`，不依赖以前聊天、`dev/output/` 下的示例、网络临时下载或机器上的绝对路径。历史学校模板及来源原件归档于 `dev/templates/`，不能误选成默认通用版本。此流程不改变已有文档排版、只读检查或学位论文任务的路由。

## 选择与准备

- 沿用用户已说明的主题、用途和素材。主题未知时询问主题，同时可准备模板；字数未指定时根据任务安排合理篇幅，不杜撰硬性字数标准。
- 用户带有其他模板时以其为准；原生 DOCX 可通过 `--sample` 准备。其他 Word 格式先走现有只读转换入口，再使用转换副本。详细书面格式要求按原有 requirements/profile 流程处理，不被本默认值覆盖。
- 学校、姓名、学号、教师和日期等未知信息保留待填。示例数据标明模拟；没有实施的功能写成方案或待实现项。真实测试结果须实际运行，书目和引用依据须核验。内容与篇幅、引用检查还应遵循 [起草与内容核验规范](thesis-generation.md)。

以下命令由 AI 根据安装目录和输出位置使用绝对路径组织。任务目录必须是新目录，已存在时拒绝覆盖。

```powershell
$SKILL_DIR = "D:/project/word-format-master"
$REPORT_TASK = "D:/project/word-format-master/dev/output/新报告任务"
python "$SKILL_DIR/scripts/format_cli.py" report list
python "$SKILL_DIR/scripts/format_cli.py" doctor
python "$SKILL_DIR/scripts/format_cli.py" report prepare "$REPORT_TASK" --topic "用户指定主题"
# 用户提供其他 DOCX 时：report prepare "$REPORT_TASK" --sample "D:/材料/教师模板.docx"
```

产物是 `draft.docx`、`profile.json` 和 `workflow.json`。草稿为模板整包副本，保留封面填写线、表格、节、样式、编号、页眉页脚和动态域；profile 的 source 绑定原模板路径及 SHA-256。登记文件校验模板指纹，模板更新时须核对来源和版式后同步更新指纹。命令本身不编写正文，也不把模板占位稿当作完成的报告。

## 内容填充与格式保护

先读 profile 的 `snapshot.structure_contract`、issues，以及 workflow 中的格式与来源说明，然后在草稿中写入内容。当前通用模板的结构为：封面 → 目录 → 引言 → 任务与要求 → 方案与实施 → 结果与分析 → 总结与体会 → 参考文献 → 附录。保留模板实际的五章及二级标题层级，替换提示和占位书目；不要自动增加摘要、英文摘要或承诺书。若用户明确要求改变结构，先更新独立的任务模板依据，再重新提取 profile，不能把旧模板的严格结构校验直接关闭。

只改对应填写槽与正文，保留包内样式、字体表、节属性和封面布局。封面信息是分栏填写表，**不属于正文数据表**；不要全局套三线表、把所有封面内容居中或用空格重新排齐字段。新增数据表单独设置样式、重复表头和行内不拆分，代码或图片段落也单独处理。

一级标题采用模板的 Word 原生“第一章、第二章……”编号，二级标题采用 `1.1`、`2.1`；参考文献和附录标题不加章号。章节只填写标题正文，不再键入“第一章”。目录保留 TOC 域；图表用 SEQ，页码用 PAGE，文内引用用书签和 REF，参考文献条目用 Word 自动编号。

填充后重新 inspect，按当前对象 ID 编写 `mapping.json` 和有依据的 `resolutions.json`，不复用示例报告的段落 ID，也不一律忽略 profile issues。封面填写表的克隆格式须在 `mapping.rules` 中对相应表格规则登记 `{"preserve":"封面填写表保持原表格及填写线，不套正文数据表规则"}`，避免分析器把填写线误判为三线表。封面、目录、含域、绘图、代码等需要保持的段落用 `mapping.preserve` 明确记录理由；其效果另行回读与视觉核对。缺失属性保留原值，真实冲突作逐项决策，不能伪称已迁移未支持属性。

## 接入统一流水线

```powershell
python "$SKILL_DIR/scripts/run_format_pipeline.py" "$REPORT_TASK/draft.docx" "$REPORT_TASK/报告.docx" --profile "$REPORT_TASK/profile.json" --match-structure --mapping "$REPORT_TASK/mapping.json" --resolutions "$REPORT_TASK/resolutions.json" --report-dir "$REPORT_TASK/验收" --render
```

mapping / resolutions 仅在有对应映射或问题时传入。流水线先 `check-structure`，再执行 `plan --profile --match-structure`、apply、verify、check_numbering，可选渲染。结构不匹配或问题未解决即停止，不回退学位论文格式；原模板和草稿不可被输出覆盖。也可逐条调用同一 CLI，不另建 DOCX 格式写入器。

如后续阶段失败，保留任务目录，修复原因后使用相同命令追加 `--resume`。恢复会检查已完成阶段的文件与证据摘要；改变草稿、格式依据或输出后应另建任务，不能自动覆盖。查看 `summary.json` 获取失败阶段及各项核对状态。

在 Word/WPS 中刷新目录和所有 StoryRanges 的域，并核对实际目录条目、章号、图表引用与正文起始页码。若刷新导致文件变化，须重新基于刷新稿 plan/apply，生成与最终输出哈希一致的证据；不要把刷新前的 apply/verify 当作当前文件验收。

运行 `thesis_quality.py` 核对正文统计与引用映射；只有用户明确给出字数范围时才传字数参数。渲染所有页面后逐页查看封面填写线、段落缩进、表格跨页、代码可读性和分页，创建绑定输出与 render-manifest 哈希的视觉报告，再运行 `format_cli.py verify --visual-report ... --render-manifest ...`。没有逐页检查时视觉为 skipped/pending，渲染成功不等于验收通过。

交付新的 DOCX 和简短的格式/结构/视觉、内容验证说明，列出待填身份信息。不得交付仍含写作提示的模板副本，也不得将模拟测试描述成真实学校系统部署。

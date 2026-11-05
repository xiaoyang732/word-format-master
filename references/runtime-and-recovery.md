# 首次使用与任务恢复

以下命令从安装目录执行，实际任务由 AI 按用户材料填写绝对路径。

## 先检查能力

```powershell
python scripts/format_cli.py doctor
python scripts/format_cli.py doctor --output-dir "D:/任务" --output "D:/任务/环境检查.json"
```

`--output-dir` 应指向已存在目录。自检检查依赖、默认模板及来源/使用说明、目录写入权限、办公软件和 PDF 转图片能力。可用 `--probe-render` 在临时目录创建无个人信息的样例实际转换。默认检测不会启动办公软件或安装依赖。发现程序路径不代表真实转换一定成功。

核心依赖缺失时，自检仍可运行。安装方法为 `python -m pip install -r requirements.txt`。若默认模板或使用说明缺失，重新安装完整发布包，或通过 `report prepare --sample` 指定用户模板；不要删除指纹验证。

## 默认报告起步示例

```powershell
python scripts/format_cli.py report list
python scripts/format_cli.py report prepare "D:/任务/索引实验报告" --topic "数据库索引对查询的影响"
```

准备阶段生成 `draft.docx`、`profile.json` 和 `workflow.json`，此时只是待填草稿。向 AI 提供提纲、实验材料和真实测量结果，要求保留封面、目录、五章正文、文献及附录的结构，替换写作提示，不虚构结果。

AI 填充后按实际对象 ID 组织映射和逐项问题决策，再运行：

```powershell
python scripts/run_format_pipeline.py "D:/任务/索引实验报告/draft.docx" "D:/任务/索引实验报告/报告.docx" --profile "D:/任务/索引实验报告/profile.json" --match-structure --mapping "D:/任务/索引实验报告/mapping.json" --resolutions "D:/任务/索引实验报告/resolutions.json" --report-dir "D:/任务/索引实验报告/验收" --render
```

仅在实际需要时传 mapping/resolutions。其内容需依据当前文档，不可照抄其他任务 ID 或忽略全部问题。身份信息未知时保留待填。首次准备使用新目录，不能覆盖旧任务。

## 中断后继续

失败后保留任务目录，修复导致失败的环境或临时占用问题，再用原命令追加 `--resume`。输出、输入、模板、映射、决策、报告目录和渲染模式须保持一致。已成功的阶段只有在指纹及相关证据一致时才跳过。

- 渲染失败后可复用已通过核对的排版文件。
- 没有 `pipeline-state.json` 的旧任务不能自动恢复。
- 最终文档、原稿、模板或证据已改变时，使用新的输出和任务目录重新处理。
- 已有输出但缺少 apply 报告时，不能推断成功或自动覆盖。
- 活跃任务的锁由操作系统持有，进程退出后释放；不要在另一进程执行时删除锁文件。
- `--render-retries 1` 仅对渲染超时等短暂故障有限重试；`--render-timeout 300` 约束单次渲染子进程。显式 `--render-method word/libreoffice` 不会静默切换后端。

## 查看真实验收状态

`summary.json` 汇总格式、结构、编号、渲染和视觉状态，并标明内容核验未执行。`pipeline-state.json` 保存每阶段状态、尝试次数及产物指纹。失败也会保存汇总。

页面图片已生成仍需逐页查看。完成视觉报告后，可用相同命令追加 `--resume --visual-report "D:/任务/视觉报告.json"` 接入最终 verify。视觉报告须按现有格式绑定最终 DOCX、渲染清单及所有页面；不接受渲染成功代替视觉通过。

字数、引用对应可另用 `thesis_quality.py` 检查；它们不能证明事实、实验结论或文献真实性。Word/WPS 保存刷新域后的文档会改变摘要，须重新为最终文件建立有效验收证据。

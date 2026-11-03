# 开发与历史资料

此目录统一保存项目运行之外的开发资料。运行代码、当前技能说明及运行模板分别保留在根目录的 `scripts/`、`references/` 和 `templates/`。

| 目录 | 内容 | Git 跟踪 |
| --- | --- | --- |
| `docs/` | 开发计划、调查、修复和审查报告 | 是 |
| `reports/` | 需要保留的模板验收摘要 | 是 |
| `templates/` | 未启用的学校模板、来源原件和历史参考 | 是 |
| `tests/` | 自动化回归测试 | 是 |
| `scripts/`、`ci/` | 校验脚本及开发配置 | 按 `.gitignore` 选择跟踪 |
| `output/` | 历次报告、计划 JSON、验收记录、渲染页面等生成产物 | 否，仅本地保留 |
| `artifacts/`、`examples/` | 本地诊断、实验和内容生成示例 | 否，仅本地保留 |

2026-11-03 整理时，原 `_dev/` 改名为 `dev/`，原根目录 `output/` 移至 `dev/output/`，未启用的 `templates/reference/` 学校模板移至 `dev/templates/reference/`。文件保留，未删除报告或来源原件。历史报告、计划和生成脚本中的创建时绝对路径保留为原始记录，不作为当前运行配置；需要重新执行时，应使用当前路径生成新的任务与证据。

当前默认报告模板仍在 `templates/reference/general-course-report/`，由 `templates/reports.json` 登记；学位论文模板及兼容数据也保留在运行模板库。模板调研和历史验收材料不参与默认选择。

在仓库根目录运行：

```powershell
python dev/scripts/validate_project.py
python dev/scripts/smoke_test.py
python -m unittest discover -s dev/tests -v
```

# Word Format Master Performance Implementation Plan

**范围**。按 2026-11-05 优化报告实施第一批改进，重点是已有 DOCX 的格式修改。完整证据、计划绑定、原稿保护、回读和独立复检继续生效。

1. 在 `dev/tests/test_cli_encoding.py` 先测试可选摘要。CLI 使用 `--summary` 配合现有报告路径写完整 JSON，stdout 只给状态、计数和有限异常索引。覆盖 inspect、plan、apply、verify、audit 及缺少报告路径的情况。旧命令和返回码保持兼容。
2. 在 `scripts/word_format/properties.py` 将默认段落样式查找缩到当前样式 XML 的少数候选。保留默认标志语义、重复默认值的最后一项规则和无默认值时的回退。测试样式在同一文档内变更时的读回。
3. 用精简的 `SKILL.md` 区分局部修改、模板排版、只读检查、论文起草和课程报告。命令与多目标示例在 `references/direct-format-api.md`。各步骤优先读取小摘要，按需从证据文件提取细节。
4. 运行完整回归、项目校验、冒烟测试和两种发布包校验。用 96 段样例对比核心耗时和包内容，并检查摘要大小。开发输出放在 `dev/artifacts/`、`dev/cache/` 和 `dev/tmp/`。复核暂存内容后提交。

定向 inspect API、流水线 request 分支与视觉检查等级属于后续批次。

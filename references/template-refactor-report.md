# 模板驱动重构：主仓库接入与验证报告

日期：2026-10-29。代码版本：1.1.0。

## 仓库与基础版本说明

本次正式交付仓库是 `D:/project/word-format-master`，基于其原有 `df15e23` 提交，在 `refactor/template-driven-formatting` 分支上完成重构。

此前修改的是 `C:/Users/yang/Documents/agent local/word-format-master/skills/@user_bb897d7b/docx-format-master` 安装副本，并非直接在 D 盘主仓库上修改。两者属于同一项目，但 D 盘已有更新的格式保留和目录校验修复。因此本次使用三方比较迁移重构增量，没有用整个旧副本替换主仓库。

主仓库原有 Git 历史、LICENSE.txt、远程设置和 `_dev/` 开发布局均沿用。回归测试、校验脚本和测试依赖加入 Git 跟踪；生成文档、诊断材料、缓存及其他开发资源继续忽略。

## 已接入的功能

| 变更 | 行为 |
| --- | --- |
| 独立内置模板 | 三套 `templates/builtin/<id>/template.json`；字体、字号、行距、页边距与编号外观放入数据 |
| 独立封面布局 | `cover.json` 与通用渲染器分离，不填虚构题目、作者、学号或日期 |
| 用户模板编译 | profile → template compile → plan；保存来源指纹、证据、规则、问题及决策，不回退内置格式 |
| 标题与编号 | Word Heading N、大纲与原生多级编号；模块标题与章标题、文献标题与条目分别处理 |
| 内容保护 | 图片、公式与混合对象不批量套正文固定行距；复杂容器保留并进入覆盖账本 |
| 统一执行 | 新稿初稿与现有论文共同使用 `scripts/run_format_pipeline.py --template ...` 或模板 CLI |
| 独立验收 | 执行前重建规则核对计划完整性；执行状态与整体符合状态分开报告 |

流程与参数见 [模板引擎说明](template-engine.md)，设计细节见 [重构方案](template-refactor-plan.md)。

## 主仓库已有修复的保留与组合修正

保留了封面字段和摘要前论文题名识别、中英文摘要与关键词分开匹配、空摘要和异常缩进的证据审查、段落行单位、规范文种范围、真实目录结果与错误目录拒绝、正文首页页码开关，以及内容审校与排版验收的边界。

三方合并后进一步修正：

- 通用语义仅补充缺失字段，不覆盖原有封面和摘要身份；提取章标题时保留层级，避免一级、二级标题合成同一格式组。
- 关键词保留 `label` / `value`，兼容 `content` 别名。
- 目录兼容 `title_style` 与新模板 `title_format`、`field_format`、`break_format`；冲突参数拒绝执行，三类外观均独立回读检查。
- 正文首页页码关闭时隔离继承关系并保留页脚文字和对象，不因自动识别正文而清空无关的前置页脚；验收检查该开关实际生效。
- 另一个文种的规范同时保留原始证据、排除记录和适用性问题，不静默丢弃。
- 测试、开发依赖和兼容封面入口按主仓库目录定位。

## 在 D 盘主仓库的验证结果

| 检查 | 结果 |
| --- | --- |
| 全量 unittest | **162 项通过**，包括原有主仓库回归、16 项模板回归及 5 项组合回归 |
| 项目校验 | 通过，66 个注册操作与内置模板合同有效 |
| DOCX smoke test | 列表、引用、页面、题注、目录、文献、页眉、域、三级标题与三线表通过 |
| Ruff check | 修改涉及的执行核心与入口通过 |
| Ruff format --check | 19 个修改或新增 Python 文件通过 |
| 实际流水线 | 从 D 盘 `scripts/run_format_pipeline.py` 完成 plan / apply / verify |

复现基础检查：

```powershell
python -B -m unittest discover -s _dev/tests -q
python -B _dev/scripts/validate_project.py
python -B _dev/scripts/smoke_test.py
```

开发依赖安装入口为 `_dev/ci/requirements-dev.txt`，Ruff 配置为 `_dev/ci/pyproject.toml`。

## zuoye 案例回归与边界

案例来源为 `C:/Users/yang/Documents/antigravity/zuoye`，输出仅写入主仓库的 `output/template-refactor-regression/`。全部原始材料经前后 SHA-256 核对未变化。

使用真正原稿 `需要修改的论文/JavaEE程序设计课程设计.docx` 测试内置模板入口：执行通过，3 张图片二进制与图片段落行距保留，原有表格保留。整体符合状态为 partial，因为存在保留的复杂容器且尚未进行逐页视觉验收。该文件是默认模板的回归产物，不是学校样例的完整交付稿。

重新从样例和书面要求提取并编译用户模板，得到 538 条合并规则、293 项编译问题；绑定此原稿后有 602 项待处理记录，包括缺失目标、多节关系和格式证据问题。严格模式返回 needs_clarification，没有生成声称“完整匹配样例”的论文。原始 profile、独立模板、问题清单和摘要保存在上述输出目录中。

这些记录应由 AI 先按材料完成能够确定的决策与映射，不表示必须让用户回答数百个问题。复杂文本框、脚注尾注、特殊编号几何和图像中的要求仍有支持边界。提取到了证据，不等于已经能够完整自动迁移任意学校模板；当前也未完成视觉验收。

本次交付包含主仓库中的代码、模板、测试、工作流说明与本报告，Git 提交保留在本地重构分支。

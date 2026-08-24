# 项目审查与上线准备报告

审查日期：2026-08-24
审查范围：接口、程序结构、模板与预设、死代码、安全边界、Windows 运行声明、自动化测试和上线文档。

## 结论

当前代码已完成本轮上线前收口：仓库只保留 1 个已审核的学位论文参数预设，不含废弃模板工作流；格式写入统一经过 Dashboard 确认式 Handoff；模板提取、模块化应用、结构复检和逐页视觉验收已经形成闭环。

项目校验、端到端冒烟测试、Python 编译检查、前端语法检查和差异空白检查均通过。未创建 Git 提交，现有变更保留在工作区供维护者审阅。

## 已修复问题

- 删除网页直接应用接口及其服务端实现，防止绕过用户确认。
- Dashboard 会话不再返回源 DOCX 内容；Handoff 锁定源路径、输出路径和源文件 SHA-256。
- API 增加版本字段和响应头，本地 HTTP 服务增加 Host 白名单。
- 会话增加 TTL；分析任务增加容量、TTL、领取令牌、租约和终态数据释放。
- 新增统一论文模块识别，覆盖封面、摘要、关键词、目录、正文各章、参考文献、致谢和附录。
- 模板提取不再用默认编号、列表缩进或题注位置补齐缺失证据。
- 修复普通导入模板被误标成官方母版模式的问题；提取规范使用参数化写入。
- 渲染清单绑定输出 DOCX、实际渲染器、页数和每页 PNG 哈希，视觉回传必须逐项匹配。
- 清理前端独立应用、Base64 下载、恒假判断、废弃预览和无效样式。
- 删除所有指定废弃出版模板相关内容和字符串。
- CI 固定使用 Windows，运行时不探测其他操作系统路径，也不返回平台字段。
- 依赖固定到明确版本，减少发布环境漂移。

## 当前模块边界

| 模块 | 责任 | 不应承担的责任 |
|---|---|---|
| `analyze_docx.py` | 包结构、有效格式和模板证据提取 | 写入输出文件 |
| `document_structure.py` | 文档模块与段落角色识别 | 字体、页面等格式写入 |
| `apply_spec.py` | 规范校验和确定性 DOCX/OOXML 写入 | HTTP、网页会话、自由文本推断 |
| `dashboard_session.py` | 会话、Handoff、渲染计划和验收回传 | 解析 HTTP 请求 |
| `serve_dashboard.py` | 本地 HTTP 路由、安全门禁和静态资源 | 直接格式写入 |
| `local_ai_analysis.py` | AI 任务领取和结果提交 | 修改 DOCX |
| `render_docx.py` | 文档渲染和不可变清单 | 判断排版是否美观 |
| `verify_output.py` | 结构与视觉报告契约检查 | 生成内容或推断模板规则 |

## 验证结果

```text
Project validation passed: 24 required files, 1 presets, 0 template workflows
DOCX smoke test passed: lists, citations, page size, captions, TOC, references, headers, fields, verification, AI handoff, three-level headings
Python compileall: passed
JavaScript syntax check: passed
Git diff whitespace check: passed
```

冒烟测试覆盖模块顺序、缺失模块验收失败、参数化模板交接、Handoff 防篡改、源数据不泄漏、API 版本、直接应用接口不存在、Host 门禁、任务容量回收、指定渲染器不降级、视觉清单与逐页哈希校验。

## 上线前仍需人工完成

1. 在干净的 Windows 10/11 机器上按 `requirements.txt` 安装并运行 CI 等价检查。
2. 分别用 Microsoft Word 和 LibreOffice 对真实长文档执行全页渲染，人工检查分页、表格、图片、公式、目录和页眉页脚。
3. 对依赖执行漏洞扫描，确认固定版本无已知高危问题。
4. 发布前检查许可证、仓库地址、版本号和安全问题反馈渠道。
5. 后续增强 LibreOffice 下载信任链，优先接入官方独立哈希或签名验证。

## 剩余技术债

- `analyze_docx.py` 与 `apply_spec.py` 仍较大；后续应按样式解析、编号、字段、题注、页眉页脚等领域逐步拆分，但必须先迁移测试再移动代码。
- 当前主要依赖端到端冒烟测试，独立单元测试和恶意 OOXML 测试仍需增加。
- 会话和任务只存内存，不支持服务重启恢复。对本地单任务工具合理，但不适合作为长期多用户服务。
- 浏览器预览不是 Word 分页引擎，不能代替最终本地逐页渲染。

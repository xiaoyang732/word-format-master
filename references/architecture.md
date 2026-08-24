# Architecture and Upgrade Contract

## 目标

本文件定义 Word Format Master 的稳定边界。后续升级可以扩展格式能力，但不得绕过 Dashboard 确认、源文件指纹、确定性写入方法、结构复检或视觉验收契约。

## 处理链

1. `analyze_docx.py` 读取 DOCX/DOTX，输出文档证据和 `inferred_spec`。
2. `serve_dashboard.py` 提供本地接口，`dashboard_session.py` 保存目标文件会话。
3. 用户在 Dashboard 修改规范并确认。
4. `dashboard_session.py` 生成带源路径、输出路径、SHA-256 和应用方法清单的 Handoff。
5. `apply_spec.py` 验证 Handoff 后生成新 DOCX。
6. `analyze_docx.py` 和 `verify_output.py` 重新检查输出。
7. `render_docx.py` 渲染所有页面并生成哈希清单，AI 视觉报告经 `record_visual_verification()` 校验后完成闭环。

## 模块责任

### `document_structure.py`

唯一负责模块名称、模块顺序、标题级别和段落角色的通用识别规则。分析、应用和验收不得各自复制正则表达式。新增模块时必须同步格式契约、模块提取、应用和验收测试。

### `analyze_docx.py`

只产生证据，不写文档。模板提取必须返回实际有效值，并区分不存在与无法确定。不得从预设补齐模板没有定义的值。

### `apply_spec.py`

只执行确定性写入。每个网页可编辑格式路径必须登记在 `SUPPORTED_APPLICATION_PATHS`。新增路径时必须同时增加规范校验、实际写入、输出断言和 Dashboard 能力覆盖检查。

### `serve_dashboard.py`

只负责本地 HTTP、输入限制、任务状态和路由。不得在请求处理器中直接生成可下载 DOCX。生成操作必须通过已确认 Handoff。

### `dashboard_session.py`

拥有会话状态、Handoff 构建、渲染计划和视觉回传校验。新的交接字段应版本化，公开会话响应必须排除源文档字节和内部任务上下文。

### `render_docx.py` 与 `verify_output.py`

渲染器只产生页面证据，验收器只验证证据和规则。两者不得静默替换用户明确指定的渲染器。

## HTTP 接口契约

当前 API 版本为 `1`，通过 `GET /api/capabilities` 的 `api_version` 和所有响应的 `X-Word-Format-API-Version` 公开。

| 方法 | 路径 | 用途 |
|---|---|---|
| GET | `/api/presets` | 读取已审核预设和模板工作流 |
| GET | `/api/capabilities` | 读取可应用格式路径和方法 |
| GET | `/api/session` | 读取当前配置会话，不含源文件字节 |
| GET | `/api/session/status` | 轻量轮询会话与视觉状态 |
| POST | `/api/session/submit` | 提交用户确认的规范并生成 Handoff |
| POST | `/api/session/verification` | 回传与清单绑定的最终视觉报告 |
| POST | `/api/analyze/docx` | 创建模板分析任务 |
| GET | `/api/analysis/tasks/next` | AI 工作者领取任务和令牌 |
| POST | `/api/analysis/tasks/{id}/result` | 使用领取令牌提交分析结果 |
| POST | `/api/validate` | 校验格式规范 |

禁止恢复 `/api/apply`。如需新增会改变输出文件的接口，必须先升级 API 版本、定义确认机制、增加源文件指纹校验并补充拒绝绕过的测试。

## 数据契约版本

- Format Spec：`schema_version: 1.0`。
- AI Handoff：`schema_version: 1.2`。
- Render Manifest：`schema_version: 1.0`。
- Dashboard API：`API_VERSION = 1`。

修改字段语义或删除字段属于不兼容变更，必须提升相应版本。只增加可选字段时仍需保证旧消费者忽略未知字段。

## 升级检查清单

1. 新设置是否有模板证据来源，而不是猜测默认值。
2. 新设置是否有单一、确定性写入方法。
3. `validate_spec()` 是否拒绝错误类型、范围和版本。
4. Dashboard 是否只展示有写入实现的设置。
5. Handoff 是否仍锁定输入、输出和源 SHA-256。
6. 输出是否重新分析并运行结构验收。
7. 视觉报告是否绑定渲染清单和所有页面哈希。
8. 测试是否覆盖成功路径、拒绝路径和旧数据契约。
9. README、SECURITY 和本文件是否同步。
10. Windows CI 与真实 Word 视觉验收是否通过。

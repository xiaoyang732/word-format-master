# Architecture and Upgrade Contract

## 目标

本文件定义 Word Format Master 的稳定边界。网页配置和直接 AI 请求共享同一格式核心。网页模式由人类提交确认；直接模式无需网页。两种模式都遵守源文件指纹、确定性写入和输出复检契约。

## 处理链

1. `analyze_docx.py` 读取 DOCX/DOTX，输出文档证据和 `inferred_spec`。
2. `serve_dashboard.py` 提供本地接口，`dashboard_session.py` 保存目标文件会话。
3. 用户在 Dashboard 修改规范并确认。
4. `dashboard_session.py` 生成带源路径、输出路径、SHA-256 和应用方法清单的 Handoff。
5. `apply_spec.py` 验证 Handoff，`spec.py` 将旧规范转换为操作计划，再由共享核心生成新 DOCX。
6. `analyze_docx.py` 和 `verify_output.py` 重新检查输出。
7. `render_docx.py` 渲染所有页面并生成哈希清单，AI 视觉报告经 `record_visual_verification()` 校验后完成闭环。

直接模式：人类要求 → `format_cli.py inspect` → AI 生成已登记操作 → `plan` → `apply` → `verify`。明确要求直接执行；含糊的位置或冲突参数返回 `needs_clarification`。不得通过 AI 操作网页伪造确认。

### `scripts/word_format`

- `registry.py`：可执行能力表，每项对应参数校验、唯一 handler、独立 reader、verifier 和测试。
- `properties.py`：字体和段落原子 XML 属性的唯一写入所有者。
- `operations.py`：每项操作的命名 handler；复杂结构操作调用已有明确职责的实现。
- `headers.py`：独立节/页型的内容、STYLEREF、链接、格式与读回；`section_breaks.py`：显式下一页分节及保护校验。规则和对应表见 [header-requirements.md](header-requirements.md)。
- `inspect.py` / `selectors.py`：源文件对象索引和严格位置解析。
- `spec.py`：旧 Format Spec 展开器与旧网页字段映射，不能新增第二套格式实现。
- `api.py`：请求、计划指纹、冲突检测、执行和临时文件验收后发布。
- `verification.py`：指定属性的回读、逐操作变更边界和结构操作保护。

只修改一个局部属性时，不改变共享样式或其他属性。共享样式必须显式指定 style target。字符范围只接受普通文字。编号、目录和引用是独立结构操作。详细契约见 [direct-format-api.md](direct-format-api.md)。

## 模块责任

### `document_structure.py`

唯一负责模块名称、模块顺序、标题级别和段落角色的通用识别规则。分析、应用和验收不得各自复制正则表达式。新增模块时必须同步格式契约、模块提取、应用和验收测试。

### `analyze_docx.py`

只产生证据，不写文档。模板提取必须返回实际有效值，并区分不存在与无法确定。不得从预设补齐模板没有定义的值。

### `apply_spec.py`

保留网页 Handoff 身份检查和复杂结构实现，普通属性写入委托共享核心。`SUPPORTED_APPLICATION_PATHS` 由能力适配器生成，新增能力需登记 handler、参数校验、独立 reader/verifier 与回归测试。

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
- Direct Request / Plan / Report：`schema_version: 2.0`。
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

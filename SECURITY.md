# Security Guidelines

## 安全边界

Word Format Master 是 Windows 本地工具，不是公网服务。Dashboard 固定监听 `127.0.0.1`，只接受 `127.0.0.1` 或 `localhost` Host 请求。不要通过端口转发、反向代理或防火墙规则把它暴露到局域网或互联网。

程序处理本地 DOCX/DOTX。除用户在 Dashboard 中主动请求下载 LibreOffice 安装包外，核心分析、写入和验收流程不建立外部连接。

## 已实施措施

### 文档与请求输入

- 仅接受 `.docx` 和 `.dotx`，拒绝旧版 DOC、DOCM 和 DOTM。
- 单个文档上限 30 MiB，请求体上限 42 MiB。
- Base64 使用严格校验；文件名只保留基本名称。
- DOCX ZIP 最多 5000 个部件，解压后总量上限 250 MiB。
- 格式规范要求受支持的 `schema_version` 和明确字段类型。

### 确认式写入

- Dashboard 不提供 `/api/apply`。
- `apply_spec.py` 只接受已确认的 `HANDOFF.json`，并核验源路径、输出路径和源 DOCX SHA-256。
- 直接 AI 入口 `format_cli.py` 不要求网页确认；使用严格操作请求、源指纹和计划摘要。摘要是误改检测，不是数字签名。
- 参数与位置冲突在写入前拒绝。输出先保存临时文件并回读，验收失败不会发布输出。
- 页面不会收到源 DOCX Base64 内容。
- 源文件不会被覆盖，输出路径必须与输入路径不同。
- 网页可编辑设置必须映射到已登记的确定性 DOCX/OOXML 写入方法。
- 格式核心只执行注册操作，不执行 AI 生成的 Python 或浏览器点击。普通格式操作逐项检查无关属性与部件，结构操作单独检查正文和对象保护。

### 会话与分析任务

- Dashboard 会话有效期为 4 小时；过期后清除分析、引用上下文和结果。
- 同时保留的分析任务最多 16 个，任务有效期为 1 小时。
- AI 工作者领取任务时获得一次性 `claim_token`，租约为 10 分钟；超时任务可重新领取。
- 已完成或失败的任务会释放原始文档数据和分析上下文，并在队列需要容量时优先回收。
- API 版本通过 `/api/capabilities` 的 `api_version` 和 `X-Word-Format-API-Version` 响应头公开。

### 本地 HTTP 服务

- 只监听 `127.0.0.1`。
- 校验 Host 请求头，降低本地服务被 DNS rebinding 利用的风险。
- 返回 CSP、`X-Content-Type-Options: nosniff`、`Referrer-Policy: no-referrer` 和 `Cache-Control: no-store`。
- 子进程使用参数数组，不启用 `shell=True`。
- 调试访问日志默认关闭，设置 `DASHBOARD_DEBUG=1` 后才写入标准错误。

### 视觉验收完整性

- `render_docx.py` 为输出 DOCX 和每一页 PNG 记录 SHA-256。
- Dashboard 只接受与当前输出路径、输出哈希、实际渲染器、页数和逐页文件哈希一致的视觉报告。
- 视觉模型未确认支持图片输入时，报告必须为 `skipped`，不能声明通过。
- `pending` 不是通过状态。

### LibreOffice 下载

- 只允许 HTTPS 和 `download.documentfoundation.org`。
- 下载地址必须匹配 Windows x86-64 MSI 规则。
- 校验 Content-Length，下载过程计算 SHA-256，失败时清理临时文件。
- 安装完成后重新探测 `soffice.exe`，运行包保存在 Git 忽略的 `.runtime/` 中。

## 集中配置

限制统一定义在 `scripts/config.py`：

```text
MAX_DOCUMENT_BYTES              30 MiB
MAX_REQUEST_BYTES               42 MiB
MAX_PARTS                       5000
MAX_UNCOMPRESSED                250 MiB
MAX_ANALYSIS_TASKS              16
ANALYSIS_TASK_TTL_SECONDS       3600
ANALYSIS_TASK_LEASE_SECONDS     600
SESSION_TTL_SECONDS             14400
```

## 剩余风险

1. LibreOffice 下载的 SHA-256 是本地计算值，只能证明下载后未再变化；尚未与官方独立哈希或签名进行比对。
2. 本地服务没有用户账户和持久化鉴权，安全前提是只监听回环地址且不被代理到其他网络。
3. 复杂或恶意 OOXML 仍可能触发第三方解析器缺陷。上线使用应保持 Python 依赖、Microsoft Word 和 LibreOffice 在受支持安全版本。
4. Microsoft Word 与 LibreOffice 的渲染差异无法通过结构校验完全消除，最终交付应使用目标环境对应的渲染器完成全页验收。
5. 任务与会话仅存内存，异常退出不会恢复；这避免在磁盘持久化源文档，但不适合长期后台服务。

## 发布检查

```powershell
python scripts/validate_project.py
python scripts/smoke_test.py
python -m compileall -q scripts
node --check assets/ui/app.js
git diff --check
```

发布前还应执行依赖漏洞扫描，并在隔离测试文档上验证 Word 和 LibreOffice 渲染。不要用含敏感内容的真实文档测试未经审核的依赖升级。

## 报告安全问题

请通过私密渠道提供复现步骤、受影响版本和最小样例，不要在公开 Issue 中上传含个人信息或未公开研究内容的文档。

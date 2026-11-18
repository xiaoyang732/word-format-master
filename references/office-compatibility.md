# Microsoft Word 兼容性与验收

Office 验收使用 `--render-method word`。单独渲染时使用 `--method word`，并可传入 `--baseline 原稿.docx` 识别继承的域错误。

- 编号写入按 WordprocessingML Transitional 的容器规则插入。旧编号、书签及关系资源仍受保护。旧稿的编号次序问题可用显式动作 `numbering.order.repair` 修复，目标为 `{"type":"document"}`，参数为 `{}`。仅调整字体时不会自动重排原稿编号。重复 ID 和无效链接须另行处理。
- REF 保留动态域、书签目标、链接和上标。内置 Heading 1–9 的 STYLEREF 使用数字选择器，自定义样式使用真实名称。参考文献的识别不依赖初始是否使用 Bibliography。旧标题编号接口的 `keep` 保留该级现有设置。
- 内置学位论文模板设置 `scope: body`、`start: 1` 和 `show_on_body_first_page: true`。需要时在正文前建立真实分节。目录的 `entry_format` 应用到 Word 的原生 TOC 样式。前置罗马数字和正文阿拉伯数字的完整双序列尚未纳入写入契约。
- 封面按表格标签和填写槽位操作。克隆报告的 `field_matching` 列出已填写及未定位的字段。单独克隆封面时清理无后续正文的尾部分节。已有图形、声明和后续节版式继续保留。上方题注与相邻对象设置同页关系。

渲染、`refresh-fields` 和旧二进制 Word 转换共用 Microsoft Word 自动化层。应用身份包含实际程序、版本、构建号、COM 注册路径和进程 ID。域更新遍历文档故事和支持的文本框，分页采用有限次更新。执行失败时 `refresh-fields` 保留原稿和已有目标文件。

Word 渲染清单 1.1 通过 SHA-256 绑定输出、页面和 `office-evidence.json`。旧清单仍可读取，Office 语义状态单独记录。语义检查读取原生编号、REF、STYLEREF、目录目标逻辑页和正文格式。原稿已有的缺失书签单列为继承问题，出现新增错误或版面检查失败时不得签收。后续保存或改写文档需要重新生成证据。

在 Windows 开发环境快测可运行 `python -B dev/tools/check_office_compatibility.py --workspace <新任务目录>`。该脚本通过官方流水线覆盖 Normal 和 Bibliography 两种文献初态。测试产物和缓存放在 `dev` 或独立工作区。实际构建和语言以每次 Word 清单为准。

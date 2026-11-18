# Microsoft Word 兼容性与验收

Office 验收使用 `--render-method word`。单独渲染时使用 `--method word`，并可传入 `--baseline 原稿.docx` 识别继承的域错误。纯 WPS 环境或双装且 `Word.Application` 被接管的环境可使用独立的 `--method wps`（及 `--method auto` 按 `word -> wps -> libreoffice` 顺序选择），渲染清单诚实记录 `"selected_method": "wps"`，不伪造 Microsoft Word 身份。

- 编号、段落属性、样式、分节、表格属性与单元格边框写入均遵循 WordprocessingML Transitional 的元素序列规则（`numbering`、`abstractNum`、`lvl`、`num`、`numPr`、`style`、`pPr`、`sectPr`、`pBdr`、`tblPr`、`tblBorders`、`tcPr`、`tcBorders`）。旧编号、书签及关系资源仍受保护。旧稿的编号次序问题可用显式动作 `numbering.order.repair` 修复，目标为 `{"type":"document"}`，参数为 `{}`。仅调整字体时不会自动重排原稿编号。重复 ID 和无效链接须另行处理。
- 参考文献与正文自动编号默认写入 `<w:suff w:val="space"/>` 单空格分隔符，且新建编号 `numId` 始终大于 `0`（避免落入 `numId=0` 抑制态）。创建 `Bibliography`、`toc 1`–`toc 9` 等 Word 内置段落样式时使用内置样式标记（不附带 `w:customStyle="1"`），避免中文 Word 将书目重命名为 `书目1`。兼容 Word 保存后不含 `<w:name>` 的匿名段落样式。
- REF 保留动态域、书签目标、链接和上标；段落级字体写入同步覆盖域控制运行（`fldChar`、`instrText`），防止 Word 在 `\* MERGEFORMAT` 更新时退回默认字体导致混合字号（`wdUndefined`）。内置 Heading 1–9 的 STYLEREF 使用数字选择器，自定义样式使用真实名称，并直接以复杂域写入 `\* CHARFORMAT` 及控制运行字体，使落盘 `.docx` 在桌面版 Word 与 WPS 中均无需额外补丁即可保持显式页眉字体；自动化刷新与导出 PDF 时仍兼容存量简单域同步。参考文献的识别不依赖初始是否使用 `Bibliography` 或是否已带 `w:numPr`。旧标题编号接口的 `keep` 保留该级现有设置。
- 内置学位论文模板设置 `scope: body`、`start: 1` 和 `show_on_body_first_page: true`。需要时在正文前建立真实分节。目录的 `entry_format` 应用到 Word 原生 TOC 样式，新建 `toc 1`–`toc 9` 默认初始化首行无缩进及按层阶梯左缩进，并同步 `WFM * Heading` 与未保留标题样式的字符属性及目录项运行字体，防止 Word 更新目录时将标题直接黑体/加粗属性泄漏到目录项。前置罗马数字和正文阿拉伯数字的完整双序列尚未纳入写入契约。
- 封面按表格标签和填写槽位操作，支持“课题题目”及带字号/字体说明的独立标题占位行，克隆时清理悬空批注标记、未配对书签和纯格式说明标注框。克隆报告的 `field_matching` 列出已填写及未定位的字段。单独克隆封面时清理无后续正文的尾部分节。已有图形、声明和后续节版式继续保留。上方题注与相邻对象设置同页关系。

渲染、`refresh-fields` 和旧二进制 Word/WPS 文档转换（`.doc`/`.dot`/`.wps`）支持 Microsoft Word 与独立 WPS Office（`KWPS.Application`）分层自动化。Microsoft Word 模式的应用身份包含实际程序、版本、构建号、COM 注册路径和进程 ID；WPS 模式独立管理新增进程，不关闭用户已打开的窗口。域更新显式遍历正文故事、各节独立页眉/页脚故事以及支持的文本框，分页采用有限次更新。执行失败时 `refresh-fields` 保留原稿和已有目标文件。

Word 渲染清单 1.1 通过 SHA-256 绑定输出、页面和 `office-evidence.json`。旧清单仍可读取，Office 语义状态单独记录。语义检查覆盖原生编号、REF、STYLEREF、目录标题与目录项字体及目标逻辑页，以及正文、章节标题、列表、文献、关键词、摘要和题注的实际字体、字号、对齐、首行/悬挂缩进与行距。原稿已有的缺失书签单列为继承问题，出现新增错误或版面检查失败时不得签收。后续保存或改写文档需要重新生成证据。

在 Windows 开发环境快测可运行 `python -B dev/tools/check_office_compatibility.py --workspace <新任务目录>`。该脚本通过官方流水线覆盖 Normal 和 Bibliography 两种文献初态及图表题注、列表、目录、动态引用。测试产物和缓存放在 `dev` 或独立工作区。实际构建和语言以每次 Word 清单为准。

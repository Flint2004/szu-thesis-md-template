# skills/

给 AI 协作者（Claude Code / Cursor / Aider）或新加入项目的人使用的**可操作指南**。
每份文件都是自含的、聚焦一个常见任务，按需读取即可。

| 文件 | 何时读 |
|------|--------|
| [writing-chapter.md](writing-chapter.md) | 要新建或重写一个章节 md 文件时 |
| [writing-full-thesis.md](writing-full-thesis.md) | 要一次性写完整篇论文（多章并行）时 |
| [citations-and-references.md](citations-and-references.md) | 要加/改参考文献或正文引用时 |
| [figures-and-tables.md](figures-and-tables.md) | 要插入图片、表格、caption 时 |
| [chinese-typography.md](chinese-typography.md) | 要处理引号、标点、空格、数学符号时 |
| [custom-style-blocks.md](custom-style-blocks.md) | 要写中英文题名、目录标题、或其它特殊样式段时 |
| [build-and-verify.md](build-and-verify.md) | 构建失败、docx 效果不对、想加新后处理规则时 |

## 使用方式

AI 协作者被派任务时先**按任务类型**读取对应的 `skills/*.md`，再开始动手。
文件之间相互独立，不读其它 skill 也能正确完成单个任务。

## 通用准则（所有 AI 协作者必读）

1. **先观察，再动手**。写新后处理规则、改样式、解排版问题**之前**，先解包
   相关 docx 看 OOXML 真实结构——不要凭 python-docx 的抽象 API 猜。
   详见 `build-and-verify.md` 的"调试 docx：先解包看 XML"一节。

2. **不臆测字段名**。python-docx 的 `paragraph.style.name` 对 pandoc 自创
   样式（如 `ImageCaption`）会 fallback 成 `Normal`，请用 `_raw_pstyle()`
   直接读 XML 原值；同理 `doc.paragraphs` **不包括表格单元格内段落**——
   这俩坑在本项目历史 commit 里都栽过。

3. **Word bookmark 命名规则**：只允许字母、数字、下划线。横线 `-` 非法，
   会导致跳转失败。本项目所有 bookmark 用 `ref_N` 形式。

4. **Mac Word 兼容性**：内部锚点超链接用元素式 `<w:hyperlink w:anchor>`
   在 Word for Mac 上跳不了，必须转 `HYPERLINK \l "X"` 字段码。`build.py`
   的 `_convert_element_hyperlinks_to_fields` 已处理。

5. **pandoc smart 扩展**会把 ASCII `"..."` 自动转成 Unicode 弯引号，与
   中文引号字符冲突。本项目禁用了（`markdown-smart`），不要随意开回。

6. **占位符替换走 XML 层**，不走 python-docx 段落 API。python-docx 会
   合并/规范化 run，丢失 WPS 生成的多字体多 run 样式。详见
   `build.py` 的 `_fill_placeholders_xml`。

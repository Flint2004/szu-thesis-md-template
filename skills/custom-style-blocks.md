# 自定义样式块（pandoc fenced div）

用于需要**特殊视觉样式**的段落：题名、目录标题等。pandoc 把 `::: {custom-style="X"}...:::` 的内容渲染为段落属性 `pStyle=X`，`scripts/build.py` 的 `_style_custom_blocks` 识别这些样式名后应用对应字号/字体。

## 已定义的样式

### `TOCTitle`

目录大标题。

```markdown
::: {custom-style="TOCTitle"}
目　录
:::
```

渲染：居中 + 三号（16pt）黑体 + 加粗 + 黑色。不进 TOC（脚本把样式降级为 Normal 避免 outline level）。

### `ThesisTitleCN`

中文题名，用于隐藏【摘要】heading 之后。

```markdown
# 【摘要】

::: {custom-style="ThesisTitleCN"}
支持向量机引导的深度特征学习及其泛化性研究
:::

【摘要】正文 …
```

渲染：居中 + 小二号（18pt）华文中宋 + 加粗 + 黑色。

### `ThesisTitleEN`

英文题名，当前模板**不使用**（英文摘要不单独加题名）。保留样式定义供需要的情况。

```markdown
::: {custom-style="ThesisTitleEN"}
Research on …
:::
```

渲染：居中 + 小二号（18pt）Times New Roman + 加粗 + 黑色。

## 如何新增自定义样式

### 步骤 1：在 markdown 中用新样式名

```markdown
::: {custom-style="NoticeBox"}
声明：本文仅供研究使用。
:::
```

### 步骤 2：在 `scripts/build.py` 的 `_style_custom_blocks` 中加分支

```python
elif sn == "NoticeBox":
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in para._element.findall(qn("w:r")):
        _run_set_font(r, eastasia="宋体", ascii_="Times New Roman",
                      size_pt=11, bold=True, color="808080")
    notice_box += 1   # 记得在函数开头声明并在 return 里返回
```

### 步骤 3：重新 build

```bash
uv run python scripts/build.py
```

## 内部机制要点

- pandoc fenced div `::: {custom-style="X"}` 生成 `<w:p><w:pPr><w:pStyle w:val="X"/>...</w:pPr>...</w:p>`
- `X` 不需要在 `reference.docx` 中预定义——python-docx 即使找不到样式名也能读到 pStyle 属性；我们的 `_raw_pstyle` 函数绕过 python-docx 的 Normal fallback 直接读原始 val
- post-process 把 `pStyle=X` 的段落先降级为 Normal，再手动套字体/段落属性；这样样式设置与 reference.docx 解耦

## 为什么不用 markdown heading 实现题名

- Heading 会进 TOC；题名进 TOC 会造成冗余
- Heading 会触发 page break（Heading 1 前自动换页）；题名不应占独立页
- Heading 样式受 reference.docx 默认值约束（如颜色），用 custom-style 更自由

## 已知限制

- pandoc 对**嵌套 fenced div** 支持有限；不要在 `ThesisTitleCN` 里再套子 div
- `custom-style` 段内部的 markdown 格式化（粗体、斜体、math）仍会被解析；如需纯文本，用 raw HTML `<span style="...">...</span>`（但不推荐，可读性差）

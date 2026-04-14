# SZU Undergraduate Thesis Template (Markdown)

深圳大学本科毕业论文 **Markdown 模板** + pandoc + python-docx 后处理流水线。
用 markdown 写正文，`uv run python scripts/build.py` 一键产出符合深大格式要求的 `.docx`。

## 为什么不用 Word 直接写

- markdown 纯文本：可版本控制、可 grep、可 diff、可并行多人协作
- 样式与内容分离：改字号/字体只动 `scripts/build.py`，内容零侵入
- pandoc + python-docx 后处理把"页眉页码/三线表/上标引用/悬挂缩进/分节"等机械格式全自动化，避免手动调 Word 的精神损耗
- 重新构建 30 秒搞定，不会因为反复调格式破坏内容

## 快速开始

```bash
# 前置：安装 uv（https://docs.astral.sh/uv/）
uv sync                                    # 安装依赖（含 pandoc 二进制）
uv run python scripts/make_reference_docx.py   # 生成样式参考文档（仅首次）
uv run python scripts/build.py              # 写完 markdown 后重复这一步
```

产物在 `build/thesis.docx`。打开后按一次 Ctrl+A、F9 可更新目录/页码。

## 目录结构

```
├── thesis/                 # 正文 markdown（按顺序拼接）
│   ├── 00_frontmatter.md   #   目录占位 + 中文摘要 + 关键词
│   ├── 01_intro.md         #   第 1 章
│   ├── 02_related.md       #   第 2 章
│   ├── 03_method.md        #   ...
│   ├── 04_theory.md
│   ├── 05_experiments.md
│   ├── 06_conclusion.md
│   ├── 07_references.md    #   参考文献
│   ├── 08_ack.md           #   致谢
│   ├── 09_abstract_en.md   #   英文摘要 + 关键词
│   └── 10_appendix.md      #   附录（证明等）
├── figures/                # 图片文件
├── build/                  # 构建产物（docx、合并后的 md、reference.docx）
├── scripts/
│   ├── build.py            #   主流水线：合并 md → pandoc → 后处理 docx
│   └── make_reference_docx.py   # 生成 reference.docx 样式源
├── skills/                 # 给 Claude / 协作者的操作指南（见下）
├── pyproject.toml
└── README.md
```

## 写作约定

### 章节标题

```markdown
# N 标题        ← 一级（本章为 "1 绪论"、"2 相关工作" 等）
## N.M 节       ← 二级
### N.M.K 小节  ← 三级
```

不写 "第 N 章"；"N" 后是半角空格，然后中文标题。

参考文献、致谢、附录也是一级标题：

```markdown
# 参考文献
# 致谢
# 附录 A 证明
```

### 中文摘要/英文摘要

两个"隐藏"的一级标题分别驱动 TOC 里的 `【摘要】` / `【Abstract】` 条目，页面上不可见：

```markdown
# 【摘要】          ← 隐藏 H1，占位让 TOC 抓

::: {custom-style="ThesisTitleCN"}
中文题名
:::

【摘要】正文第一段 …

【关键词】关键词 1；关键词 2；关键词 3
```

英文部分不写题名，直接：

```markdown
# 【Abstract】

【Abstract】Body …

【Key words】…
```

### 图片

```markdown
![图 N.M 中文说明，可以很长，pandoc 会渲染在图片下方并自动居中。](../figures/xxx.png)
```

约束：
- caption 必须以 "图 N.M" 开头（中间一个半角空格），否则脚本识别不到
- 图片宽度会被自动限制到 ≤ 9.14 cm（约 60% 页宽），cy 等比
- caption 字号自动压到 9pt 小五、段落居中

### 表格

```markdown
**表 N.M** 表标题写在这一行，加粗「表 N.M」前缀。

| 列 1 | 列 2 | 列 3 |
|:-----|:----:|:---:|
| 数据 | 数据 | 数据 |

表格下一段正文继续写 …
```

约束：
- caption 段紧邻表格上方（pandoc 需要它们之间没有其他段）
- 脚本会识别"段落紧邻其后是 `<w:tbl>`"的段，自动套 booktabs 三线 + 整表居中 + 表头加粗

### 引用与参考文献

正文裸写：

```markdown
本文方法借鉴 [1] 与 [2-4]，与 [6, 9] 有本质差异。
```

脚本自动把 `[N]` / `[N-M]` / `[N, M]` 转成上标超链接，指向参考文献对应条目。

参考文献条目（`thesis/07_references.md`）写法：

```markdown
[1] LECUN Y, BENGIO Y, HINTON G. Deep learning[J]. Nature, 2015, 521(7553): 436-444.

[2] HE K, ZHANG X, REN S, et al. Deep residual learning for image recognition[C]//Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition. 2016: 770-778.
```

脚本自动给每条打上 bookmark（`ref_N`）并按深大规范设置悬挂缩进、0.5 行段前后、9pt 小五字号。GB/T 7714-2015 格式由作者保证。

### 中文引号 —— **必须用中文引号**

写作时请**直接**用中文引号 `"` `"` `'` `'`（U+201C/201D/2018/2019），而不是英文直引号 `"` `'`。

macOS 搜狗/微软拼音默认输入就是中文引号；若拷贝外文源码导致直引号进入文档，使用 `skills/chinese-typography.md` 里给出的 sed 命令或委托 Agent 统一处理。

例外（**保留英文直引号**）：
- 行内/围栏代码块 `` `...` `` / ```` ``` ````
- 数学公式 `$...$`（含导数撇号 $x'$）
- URL / 文件路径
- `::: {custom-style="..."}` 的 pandoc 语法
- 参考文献条目（`07_references.md`）

### 自定义样式块

pandoc fenced div 用于需要特殊视觉样式的段：

```markdown
::: {custom-style="TOCTitle"}
目　录        ← 用一个全角空格分隔；这段自动变成三号黑体加粗居中
:::

::: {custom-style="ThesisTitleCN"}
中文题名      ← 小二号（18pt）华文中宋加粗居中
:::

::: {custom-style="ThesisTitleEN"}
English Title ← 小二号（18pt）Times New Roman 加粗居中
:::
```

### 数学

用 LaTeX 语法，pandoc 自动转 OMML：

```markdown
行内：$\mathbf{w}^\top \mathbf{x} + b$
独占行：
$$L_0(f) \leq \hat{L}_\gamma(f) + \frac{4}{\gamma}\hat{\mathfrak{R}}_S$$
```

脚本会把纯数字 run 强制直立（否则 Word 会按变量斜体化）。

## 构建流水线

`scripts/build.py` 做三步。

**一、合并 markdown**
- 读取 `thesis/` 下按文件名排序的所有 md
- 正文章节：把 `[N]` 引用改写为 pandoc 上标超链接 `^[\[N\]](#ref_N)^`
- 参考文献章节：原样保留（避免把 `[1]` 起首误改）

**二、调用 pandoc 转 docx**
- 用 `reference.docx` 提供样式基础（正文宋体五号、标题黑体等）
- 输入格式：`markdown+tex_math_dollars+raw_tex+fenced_divs`

**三、python-docx 后处理（19 类自动处理）**

| # | 处理 | 说明 |
|---|------|------|
| 1 | 目录字段 | 在「目　录」标题段后注入 `TOC \o "1-3"` 字段 |
| 2 | 自定义样式 div | TOCTitle / ThesisTitleCN / ThesisTitleEN 渲染为对应字号/字体 |
| 3 | 参考文献 bookmark | 为每条 `[N] ...` 段插入 `ref_N` 书签，让上标引用可点击 |
| 4 | Heading 1 样式 | 宋体 + 黑色 + 无缩进（章标题） |
| 5 | Heading 2 样式 | 黑体 + 15pt 小三 + 加粗 + 无缩进 |
| 6 | Heading 3 样式 | 黑体 + 14pt 四号 + 加粗 + 无缩进 |
| 7 | 参考文献 H1 | 特殊：五号（10.5pt）华文楷体加粗顶格 |
| 8 | 致谢 H1 | 特殊：小四号（12pt）黑体加粗居中 |
| 9 | 摘要/关键词 | 华文楷体；【标识符】加粗小四（12pt）；段落顶格（无首行缩进） |
| 10 | 图表 caption 居中 + 9pt 小五 | 按 pStyle `ImageCaption` 或"紧邻 `<w:tbl>`"识别 |
| 11 | 含图片段落居中 | |
| 12 | 表格美化 | 整表居中 + booktabs 三线（上粗 1.5pt、中 0.5pt、下粗 1.5pt）+ 表头加粗 |
| 13 | 元素式超链接 → 字段码 | Mac Word 兼容性，同时上标化引用 |
| 14 | OMML 纯数字直立 | 加 `<m:sty val="p"/>` |
| 15 | 图片宽度统一 | cx 最大 3.6 英寸（约 9.14 cm），cy 等比 |
| 16 | 隐藏 H1 `【摘要】/【Abstract】` | 白字 1pt，保留 outline level 进 TOC |
| 17 | Heading 1 前 page break | 每章另起一页（首个 H1 由 section break 负责） |
| 18 | 分节 + 页码 | 目录 section 用 lowerRoman；正文 section 用 decimal start=1 |
| 19 | 参考文献悬挂缩进 | left=419, hanging=419, 段前后 0.5 行、单倍行距、9pt 小五 |

打印的构建日志形如：

```
[post] toc=True bookmarks=43 centered=25
       h1=11 h2=36 h3=30 refs_h=1 acks_h=1 hidden_h1=2 pbreaks=10 sect=1
       toc_title=1 title_cn=1 title_en=0 abstract=4 refs_entries=43
       captions=25 pics_resized=6 tables=13 hyperlinks=84 math_nums=448
```

## 自定义样式

- 改字号、字体、段前后：编辑 `scripts/build.py` 对应函数（`_restyle_chapter_headings` / `_restyle_abstract_paragraphs` / `_restyle_captions` / ...）
- 改默认正文字体/页边距：编辑 `scripts/make_reference_docx.py`，重新生成 `build/reference.docx`
- 学校模板有细微差异：生成 docx 后用官方 Word 模板套最后一道

## 已知限制

- 封面与诚信声明页需自行在 Word 模板上填写（格式差异过大，本脚本不处理）
- 每章前的 page break 由 Heading 1 触发；若想中途插入额外分页，请在 markdown 用 `::: {custom-style="PageBreak"}` div 或直接插入 Word 分页符
- `reference.docx` 给出的是近似样式；正式投交前建议用学校官方模板二次套版
- pandoc 对 `\newpage` 不生效（已知），依赖本脚本的 Heading 1 page break 机制

## 贡献与定制

欢迎 fork 改成你们学校的模板。主要定制点：

- `scripts/build.py` 顶部常量：bookmark 命名、图片宽度、悬挂缩进参数
- `scripts/make_reference_docx.py`：默认字体/字号/页脚
- `skills/`：给 AI 协作者（Claude Code / Cursor / Aider）的操作指南

## License

MIT

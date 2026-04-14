# 构建与自检

## 标准构建流程

```bash
# 1. 首次：生成 reference.docx（样式源）
uv run python scripts/make_reference_docx.py

# 2. 每次改完 markdown 重跑
uv run python scripts/build.py
```

输出：

```
[merge] -> build/thesis.md          # 合并后的单文件 markdown
[docx]  -> build/thesis.docx         # pandoc 转换
[post]  toc=True bookmarks=43 ...    # python-docx 后处理日志
[done]  -> build/thesis.docx
```

## 日志字段解读

```
[post] toc=True bookmarks=N centered=N
       h1=N h2=N h3=N refs_h=N acks_h=N hidden_h1=N pbreaks=N sect=N
       toc_title=N title_cn=N title_en=N abstract=N refs_entries=N
       captions=N pics_resized=N tables=N hyperlinks=N math_nums=N
```

| 字段 | 期待值 | 含义 |
|------|--------|------|
| `toc=True` | True | TOC 字段注入成功 |
| `bookmarks` | = 参考文献数 | 每条 `[N] ...` 被打 bookmark |
| `centered` | 图+表 caption + 图片段总数 | |
| `h1` | 章数 + 摘要 + Abstract + 参考 + 致谢 + 附录 | 大概 10-12 |
| `h2` | 所有 "N.M 节" | 每章 3-6 个 |
| `h3` | 所有 "N.M.K 小节" | 视需要 |
| `refs_h` | 1 | 参考文献 H1 特殊样式套用 |
| `acks_h` | 1 | 致谢 H1 特殊样式套用 |
| `hidden_h1` | 2 | 【摘要】+【Abstract】两个隐藏 H1 |
| `pbreaks` | h1 - 1 | 章前分页数（首个 H1 由 section break 负责） |
| `sect` | 1 | 分节成功 |
| `title_cn` | 1 | 中文题名块 |
| `title_en` | 0 或 1 | 英文题名（本模板默认 0） |
| `abstract` | 4 | 中摘 + 关键词 + 英摘 + Key words |
| `refs_entries` | = bookmarks | 全部参考文献条目套悬挂缩进 |
| `captions` | 所有图 + 所有表 | |
| `pics_resized` | 超过 9.14 cm 被缩的图数 | |
| `tables` | 表格数 | |
| `hyperlinks` | 正文引用次数 | |
| `math_nums` | OMML 中数字 run 被直立化的数量 | |

若某字段为 0 或异常小，检查对应 markdown 源。

## 打开 docx 后

- Ctrl+A → F9：更新所有字段（目录、页码、交叉引用）
- 检查页码：目录页应是罗马 i, ii；正文从 1 重新开始
- 检查目录：不含"目录"自身；含【摘要】【Abstract】等

## 常见问题

### 构建报错：`pandoc not found`

`pypandoc-binary` 自带 pandoc，应不会出现。若出现：`uv sync` 重装。

### 目录里缺【摘要】或【Abstract】

- 检查 `00_frontmatter.md` 首行是 `# 【摘要】`（不是 `【摘要】` 粗体段）
- 检查 `09_abstract_en.md` 首行是 `# 【Abstract】`

### 引用 `[N]` 点击跳不到参考文献

- 确认 `07_references.md` 里有对应的 `[N] ...` 条目（行首）
- Word for Mac 若仍显示 `h#ref_N`，确认 build.py 的 `_convert_element_hyperlinks_to_fields` 正常运行（日志 `hyperlinks=N > 0`）

### 图片太大或太小

- 太大：已被缩到 9.14 cm；如希望更小，在 markdown 加 `{width=6cm}`
- 太小：原图分辨率不够，脚本不放大

### 表格格式乱

- caption 与表格之间有其它段落？缩到空行相隔
- 表格列对齐符写错？检查 `|:---:|:---|---:|` 语法

### 章节编号错乱

- Heading 1 必须以 "N " 开头（不要 "第 N 章"）
- 检查章节顺序是否与 `CHAPTERS` 列表一致

## 加新后处理规则

所有后处理函数在 `scripts/build.py` 末尾 `post_process_docx` 中调度。加新规则：

1. 写新函数 `_xxx(doc) -> int`
2. 在 `post_process_docx` 里调用并打印日志
3. 重新 build 验证

范式：

```python
def _my_new_rule(doc: Document) -> int:
    fixed = 0
    for para in doc.paragraphs:
        if condition_to_match(para):
            do_something(para)
            fixed += 1
    return fixed
```

## 调试 docx：先解包看 XML

**每次遇到样式/排版异常，第一反应不是改 `build.py`，而是解包 docx 看
真实的 OOXML 结构。** 这是本项目最重要的调试习惯——python-docx 的抽象
有不少陷阱（`style.name` 会 fallback、`doc.paragraphs` 不含表格内段），
只看抽象层会猜错根因。

### 常用命令

```bash
# 解包到临时目录
rm -rf /tmp/extract && mkdir /tmp/extract
unzip -q build/thesis.docx -d /tmp/extract
ls /tmp/extract/word/              # document.xml / footer1.xml / styles.xml / ...

# 不落地直接看
unzip -p build/thesis.docx word/document.xml | less
unzip -p build/thesis.docx word/styles.xml | grep -A3 'w:styleId="Heading1"'

# 查某段文字附近的 pPr/rPr
python3 -c "
import re
x = open('/tmp/extract/word/document.xml').read()
i = x.find('特定中文关键字')
if i > 0:
    start = max(x.rfind('<w:p ', 0, i), x.rfind('<w:p>', 0, i))
    end = x.find('</w:p>', i) + 6
    print(x[start:end])
"
```

### 对比参考稿

若用户给了"改好的 docx"作为目标样式：

```bash
# 解包参考稿
unzip -q /path/to/reference.docx -d /tmp/ref
# 对比同一段落的 XML（比如 H1 第3章）
diff <(python3 extract_para.py /tmp/extract/word/document.xml "第3章") \
     <(python3 extract_para.py /tmp/ref/word/document.xml "第3章")
```

把差异的 attribute / child element 直接抄到 post-process 代码里（黑体/颜色
/ind/sz 等）——这比猜规范更准，也能匹配教师/学院的特殊要求。

### 写新规则前的"先看后写"清单

- [ ] 解包目标 docx，grep 定位问题段的 XML 片段
- [ ] 找该段的 `pStyle val`（区分 pandoc 默认 `Normal` fallback 与真实值）
- [ ] 看它依赖哪个样式（解 `styles.xml` 找同名 style 定义）
- [ ] 看它的 `pPr` / `rPr` 实际写了什么（可能被 reference.docx 继承）
- [ ] 再写代码改它

### 其它快捷调试

- `_raw_pstyle(para)`：python-docx 抽象层读不到真 pStyle 时用这个
- `grep -c 'w:hyperlink' /tmp/extract/word/document.xml`：统计元素式超链接
- `grep -oE 'bookmarkStart[^/]*name="[^"]*"' /tmp/extract/word/document.xml | sort -u`：列出所有书签
- 改完 `build.py` 再跑 build，重复解包 diff 前后差异来验证

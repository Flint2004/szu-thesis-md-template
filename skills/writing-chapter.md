# 写一个章节

## 文件约定

- 所有章节 md 文件放 `thesis/`，命名 `NN_名字.md`（数字前缀决定合并顺序）
- 默认 10 个文件覆盖整篇论文结构：`00_frontmatter` / `01_intro` / … / `10_appendix`
- 如需增删，同步修改 `scripts/build.py` 顶部 `CHAPTERS = [...]` 列表

## 标题层级

```markdown
# N 标题          ← H1，章节号不写"第"字
## N.M 节标题    ← H2
### N.M.K 小节   ← H3（再深一层就不建议再拆 heading）
```

- "N" 后必须一个半角空格
- 参考文献、致谢、附录标题特殊：`# 参考文献` / `# 致谢` / `# 附录 A 证明`（脚本会自动识别并套特殊样式）
- 中英文摘要用隐藏 H1：`# 【摘要】` / `# 【Abstract】`（占位让 TOC 抓，页面上不可见）

## 正文段落

- 写中文段落即可，pandoc 会默认套正文样式（宋体五号等）
- 段落之间**空一行**（markdown 标准）
- **不要**手动加 `\newpage`（在 docx 里 raw_tex 不生效）；换页由 H1 触发
- 段落首行缩进交给 Word 模板，markdown 不管

## 引用外部文献

裸写 `[N]` / `[N-M]` / `[N, M]`，脚本自动转上标超链接：

```markdown
这一做法借鉴了 [3]，并与 [5-7] 的思路有本质差异；综合文献 [8, 12, 15] 可以进一步论证 …
```

## 常见错误

- ❌ `# 第 1 章 绪论` → ✅ `# 1 绪论`
- ❌ `## 1.1. 背景` → ✅ `## 1.1 背景`（不要末尾那个点）
- ❌ 章末加 `\newpage` → ✅ 不加，脚本自动处理
- ❌ 段落里用 `"..."` → ✅ 用 `"..."`（见 chinese-typography.md）
- ❌ 在同一行写两个标题用 `<br>` → ✅ 分两行、空行隔开

## 写完后

不需要手动构建；把草稿交回主流程后由维护者统一 `uv run python scripts/build.py`。
如需自测单章渲染效果，可把单个 md 文件路径喂给 pandoc：

```bash
uv run python -c "import pypandoc; print(pypandoc.convert_file('thesis/03_method.md', 'docx', outputfile='/tmp/ch3.docx'))"
```

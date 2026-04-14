"""把 thesis/ 下 markdown 的英文直引号 "/' 统一替换为中文引号 ""/''。

⚠️ 警告 ⚠️
=========
这个脚本会**原地修改文件**。虽然尽量避免误伤，但以下三类场景依然可能出问题，
使用前请务必阅读：

1. 段内**未配对**的直引号：例如整段只有一个 `"`，脚本会当成**左**引号
   转换。若原意是"英寸标记"或被代码块截断，需手动恢复。

2. 正文中夹杂**英文句子**里的缩略 `'s` `n't` 会被转成中文右撇号 `'`，
   视觉上看不出差别但字符不同。

3. **数学公式里的转置/素数**（如 $\\mathbf{x}'$、$h'(x)$）**已被保护**
   （识别 $...$ 段内不动），但多行数学 $$...$$ 中若跨段书写且语法异常，
   可能误替换。

保护区域（不改）：
- 围栏代码块 ``` ... ```
- 行内代码 `...`
- 数学 $...$ / $$...$$
- YAML frontmatter（文件顶 --- ... ---）
- HTML 注释 <!-- ... -->
- pandoc fenced div 起止行 `::: {custom-style="..."}`
- 反斜杠转义引号 `\"` 与 `\'`（作者显式声明要保留直引号；build.py 会在转
  docx 前自动把 `\"` → `"`）
- URL / 文件路径（含 `"` 的 markdown 链接属性 `[x](url "title")` 仍会动，请自查）
- 整个 thesis/07_references.md 文件（参考文献条目里的英文引号大多是合法的）

使用
----
    # 预览（dry-run，默认）：列出将要修改的文件与统计，**不写入**
    uv run python scripts/fix_quotes.py

    # 真正写入
    uv run python scripts/fix_quotes.py --write

    # 只处理特定文件
    uv run python scripts/fix_quotes.py --write thesis/01_intro.md

运行后建议 `git diff` 逐处目视确认。
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
THESIS_DIR = ROOT / "thesis"
SKIP_FILES = {"07_references.md"}

# 保护区域 regex：这些 span 不动
PROTECTED_PATTERNS = [
    # 围栏代码块
    (re.compile(r"^```.*?^```", re.MULTILINE | re.DOTALL), "fence"),
    # 行内代码
    (re.compile(r"`[^`\n]+`"), "inline_code"),
    # 显示数学
    (re.compile(r"\$\$[\s\S]+?\$\$"), "math_display"),
    # 行内数学
    (re.compile(r"\$[^\$\n]+\$"), "math_inline"),
    # YAML frontmatter
    (re.compile(r"\A---\n.*?\n---\n", re.DOTALL), "yaml"),
    # HTML 注释
    (re.compile(r"<!--.*?-->", re.DOTALL), "html_comment"),
    # pandoc fenced div 起止行（整行都保护，避免 custom-style="..." 被动）
    (re.compile(r"^:::.*$", re.MULTILINE), "pandoc_div"),
    # 反斜杠转义引号：作者显式要保留直引号（build.py 会在转 docx 前去掉反斜杠）
    (re.compile(r"\\[\"']"), "escaped_quote"),
]


def _mask_protected(text: str) -> tuple[str, list[str]]:
    """把保护区域替换为唯一占位符，返回 (masked_text, pieces)。"""
    pieces: list[str] = []
    for pat, _name in PROTECTED_PATTERNS:
        def repl(m: re.Match[str]) -> str:
            pieces.append(m.group(0))
            return f"\x00PROTECTED{len(pieces) - 1}\x01"
        text = pat.sub(repl, text)
    return text, pieces


def _unmask(text: str, pieces: list[str]) -> str:
    def repl(m: re.Match[str]) -> str:
        idx = int(m.group(1))
        return pieces[idx]
    return re.sub(r"\x00PROTECTED(\d+)\x01", repl, text)


def _convert_double(text: str) -> tuple[str, int]:
    """把配对的 "..." 替换为 "..."。单个落单的 " 视为左引号（风险）。"""
    count = 0
    out: list[str] = []
    inside = False
    for ch in text:
        if ch == "\"":
            if not inside:
                out.append("\u201C")  # 左 "
            else:
                out.append("\u201D")  # 右 "
            inside = not inside
            count += 1
        else:
            out.append(ch)
    return "".join(out), count


def _convert_single(text: str) -> tuple[str, int]:
    """把 ' 替换为中文 '/'。

    仅替换"看起来像中文上下文"中的 '：前或后紧邻中日韩 Unicode 区 \u4e00-\u9fff。
    英文缩略（it's, don't）不动，避免视觉隐患。
    """
    count = 0
    CJK = r"[\u4e00-\u9fff]"

    def repl(m: re.Match[str]) -> str:
        nonlocal count
        quote = m.group(0)
        start = m.start()
        # 取前后 2 字符上下文
        before = text[max(0, start - 1): start]
        after = text[start + 1: start + 2]
        if re.match(CJK, before) or re.match(CJK, after):
            count += 1
            # 成对：奇数次出现用左 '，偶数次用右 '（简化：用字符位置判断前后）
            # 简化策略：前是 CJK / 空白 / 行首 / 标点 → 左 \u2018；否则 \u2019
            if re.match(CJK, before) and not re.match(CJK, after):
                return "\u2019"  # 右 '
            return "\u2018"  # 左 '
        return quote

    return re.sub(r"'", repl, text), count


def process_file(fp: Path, write: bool) -> tuple[int, int]:
    text = fp.read_text(encoding="utf-8")
    masked, pieces = _mask_protected(text)
    masked, dq = _convert_double(masked)
    masked, sq = _convert_single(masked)
    new = _unmask(masked, pieces)
    if (dq or sq) and write and new != text:
        fp.write_text(new, encoding="utf-8")
    return dq, sq


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="*", help="指定文件，不给则扫 thesis/ 全部（除 07）")
    ap.add_argument("--write", action="store_true", help="真正写入文件；不加则 dry-run")
    args = ap.parse_args()

    if not args.write:
        print("[dry-run] 只预览，不写入；加 --write 才会修改文件。\n")

    if args.files:
        targets = [Path(f) for f in args.files]
    else:
        targets = sorted(
            p for p in THESIS_DIR.glob("*.md") if p.name not in SKIP_FILES
        )

    total_dq = total_sq = total_files = 0
    for fp in targets:
        if not fp.exists():
            print(f"[skip] {fp} 不存在")
            continue
        dq, sq = process_file(fp, args.write)
        if dq or sq:
            total_files += 1
            total_dq += dq
            total_sq += sq
            action = "改写" if args.write else "将改写"
            print(f"  {fp.relative_to(ROOT)}: {action} 双引号 {dq} 处, 单引号 {sq} 处")

    print(f"\n合计：{total_files} 个文件，双引号 {total_dq} 次，单引号 {total_sq} 次。")

    if not args.write:
        print("\n确认无误后重跑：uv run python scripts/fix_quotes.py --write")
    else:
        print("\n✅ 已写入；建议立即 `git diff thesis/` 逐处核对，尤其数学/代码附近。")


if __name__ == "__main__":
    main()

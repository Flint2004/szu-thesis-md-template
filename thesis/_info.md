---
# 论文封面 + 诚信声明的作者信息。
# 构建时 scripts/build.py 把这里的 YAML 字段注入到
# templates/cover.docx 与 templates/honor.docx 中的 {{key}} 占位符。
#
# 若暂不想附加封面+诚信声明两页，运行：
#     uv run python scripts/build.py --no-cover
title: "冰之妖精对青蛙的可控冻结：一种基于冷气调制的行为学与热力学联合研究"
name: "琪露诺"
major: "冰系魔法学"
college: "雾之湖自然研究院"
student_id: "⑨⑨⑨⑨⑨⑨⑨⑨"
advisor: "帕秋莉·诺蕾姬"
advisor_title: "魔导教授"
date_year: "2099"
date_month: "9"
date_day: "9"
---

此文件仅作**作者信息**之用——不会作为论文章节内容合并。
如需新增字段，同步修改本文件 frontmatter、`templates/cover.docx` 或
`templates/honor.docx` 中的 `{{key}}` 占位符即可。

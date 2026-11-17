---
name: ppt-deck
description: 按共享科研风格或某所高校的既有模板制作可编辑 PPTX，用于组会、学术汇报和研究结果展示；支持从一份 PPTX 提取模板入库、保留封面/目录等页面原型并逐像素复现，也支持按学校、标签、比例和版式检索调用。
metadata:
  version: "0.11.0"
---

# 科研汇报 PPT

执行前读取套件 `shared/styles/current.json`；独立安装读 `assets/styles/current.json`。`--style`、`RESEARCH_STYLE_FILE` 优先于默认文件。标题、正文、说明字号从 `slides` 映射，字体从 `typography` 映射，颜色从 `palette` 映射。

1. 根据用户材料整理大纲，区分已知结果、假设和待补充信息。保留数值、单位和来源，不把示意数据包装成实验结果。
2. 按 [大纲与版式规则](references/ppt-layout-rules.md) 准备 JSON 和本地 PNG/JPEG；图表使用同一风格生成，并检查其 `.style.json` 指纹。已有图像不会被重新着色。
3. 从本技能的 `scripts/build_pptx.py` 生成 PPTX。长 bullet 列表会分页；单条过长或标题超限会报错，先拆分表述而非静默缩小字体。
4. 用可用的 PowerPoint、LibreOffice 或宿主预览渲染检查，修复溢出、遮挡、缺字及低清图。脚本的估算排版和 OOXML 结构检查不能代替视觉检查。无法渲染时明确标为待视觉验收。

```sh
python scripts/build_pptx.py outline.json --output meeting.pptx
python scripts/build_pptx.py outline.json --style custom-style.json --output meeting-custom.pptx
```

命令使用实际绝对路径。依赖为 `python-pptx`、Pillow 和 `jsonschema`；安装包 `assets/requirements.txt` 或源码根目录有完整依赖。交付 PPTX 及 `.style.json`，说明使用的风格和未完成的检查。

## 模板库

用户提到某所学校、课题组或单位的既有 PPT 模板时，先查库再动手，不要凭印象仿造版式。

```sh
python scripts/query_template.py list --school 清华 --tag 答辩 --aspect 16:9
python scripts/query_template.py show tsinghua-thesis-16x9
python scripts/ingest_template.py deck.pptx --school 清华大学 --school-slug tsinghua \
  --id tsinghua-thesis-16x9 --name "清华答辩模板" --tags 学术汇报,答辩 \
  --contributor 姓名 --license "作者已授权本仓库分发"
```

- 入库只保留母版、版式、主题与版式引用的图片；正文和备注被丢弃，库文件可公开分发。`--license` 必须写清授权来源。
- 描述与目录的约定见仓库根 `templates/README.md`；`index.json` 由入库自动重建，`query_template.py index --check` 可发现漂移。

## 页面原型与复现

封面、目录这类页面的设计常常直接画在幻灯片上（校徽图片、装饰线、底纹形状），只靠母版和版式恢复不了。需要逐像素复现时，入库加 `--keep-slides` 把这些页保留为**页面原型**：

```sh
python scripts/ingest_template.py deck.pptx --school 华南理工大学 --school-slug scut \
  --id scut-kaituo-16x9 --keep-slides 1=cover,2=toc --license "作者本人提供"
python scripts/compose_deck.py --template-id scut-kaituo-16x9 --spec compose.json --output output/deck.pptx
```

`compose.json` 的每一页取二者之一：

```json
{"template": "scut-kaituo-16x9", "pages": [
  {"prototype": "cover", "texts": {"2": "新的论文标题", "9": "答 辩 人：张三"}},
  {"layout": 17, "title": "研究内容", "body": ["第一点", "第二点"]}
]}
```

- 原型页保留原幻灯片的全部形状，只改 `texts` 指定的文字槽（键为 shape id 或唯一形状名，id 见 `template.json` 的 `prototypes[].text_slots`）；不写 `texts` 就是原样复现。
- **同一个原型可以复用**：第二次及以后引用会克隆整页（形状、文字、图片关系一并复制），只需改文字。分节目录页就用这个办法——与首页目录同一套版式，只切换当前章节的高亮项，其余格式完全一致；`texts` 的值可写成 `{"text": "...", "color": "#1F8DCF"}` 覆盖颜色。
- 文字槽若是「按文字长度定宽、左对齐」的文本框（老模板常见，靠框刚好包住文字实现视觉居中），换字后会自动改为居中；需要固定方向时把文字写成 `{"text": "...", "align": "left"}`。
- 版式页按模板版式新建。入库时采集的内容页标题样式与**标题形状原始 XML** 会一并复用，因此内容页的标题部分可与原稿逐像素一致。
- 原型页文字槽是固定尺寸，替换文字变长会换行甚至压到相邻元素，交付前必须渲染检查。
- 合成后必须核对：用 `--keep-slides` 保留的原型页应做到渲染逐像素一致，内容页允许版式级差异，并向用户说明差异范围。

## 内容页模块

内容页常用虚线框把版面切成若干模块。先用 `tools/extract_media.py` 把原页图片抽出来，再用 `modules` 描述模块与内容：

```sh
python ../../tools/extract_media.py deck.pptx --slide 3 --output assets/slide03
```

```json
{"layout": 17, "title": "热电转换", "modules": [
  {"rect": [8.2, 26.9, 156.4, 89.8], "padding_mm": 3.0,
   "border": {"color": "#B8B8B7", "width_pt": 2.0, "dash": "round_dot"},
   "items": [
     {"box": [0, 0, 1, 0.10], "text": "小标题", "font_pt": 16, "bold": true, "color": "#1F3864", "align": "center"},
     {"box": [0.02, 0.11, 0.96, 0.74], "image": "assets/slide03/slide03-shape21.png"},
     {"box": [0, 0.87, 1, 0.13], "text": "图注", "font_pt": 10.5, "align": "center"}]}]}
```

- `rect` 用毫米（相对幻灯片左上角）；`items[].box` 是模块内边距以内的相对比例；`image` 路径相对 spec 文件，或由 spec 的 `base_dir` 指定。
- 边框默认是**直角**。原稿常用 `roundRect` 但把圆角值设为 0，渲染就是直角；需要圆角时显式写 `"radius_mm": 3`。
- 合成会强制内容留在边框内侧：越界、文字超出自身盒子、item 相互重叠（除非显式 `"overlay": true`）都会报错，不会静默压到虚线上。
- 版式用 `query_template.py show` 的清单确定索引；同名版式存在于多套母版时必须用索引。批量复现时按页记录真实版式（`tools/analyze_page.py` 会输出每页的 `layout_part` 与标题 XML），不要把探针页的版式和标题位置套到所有页。
- 字体要分别保留 **latin（英文与数字）/ ea（中文）/ cs（复杂文种）** 三套字库；OOXML 里数字跟随 latin，不要用中文字体去渲染数字。项目符号要从段落 `a:pPr` 取 `a:buFont` + `a:buChar`（本库模板的小标题用的是 Wingdings `Ø`，渲染成三角箭头，不是圆点），不能统一写成 `• ` 前缀；`a:buSzPct`、`marL/indent` 一并保留。
- 文本框的 `insets`（内边距）与 `anchor` 要按原稿保留，否则文字会整体偏移几个像素。
- **先判断用途，再决定格式**：每个文本框带 `annotation`（display / heading / subheading / body / body_large / bullet / caption / citation / note / label / page_number），每个 run 还可带细分标注（`emphasis`、`superscript`、`subscript`；引用框拆成 journal / year / volume / pages）。模板入库时会从全篇统计**同一用途的众数格式**，写进 `template.json` 的 `text_styles`；合成时按注释查表统一套用，所以两个引用框不会一个 10pt、一个 12pt，也不会字体各异。
- **同用途同字号、同字库**：字号/字库按注释统一，紧凑文本框不再为"装得下"而各自缩字号；放不下时优先扩框，仍放不下允许溢出，由碰撞避让处理。
- **对齐策略**：小标题、正文、列表一律左对齐；图注/数值标签居中；引用左对齐并锚定在所在文本框左下角（距左边与底边各 2mm，多条引用在同一行等距排开）。
- 碰撞避让：合成前先按原位置检测文字/图片碰撞，先尝试小幅上下移动，必要时缩 1–4pt，最后仍冲突的标记为设计性叠放并记录；尽量不要靠缩字号解决，以免破坏同用途一致性。
- **孤字换行控制**：需要换行时，最后一行不应只剩 1–2 个字；先逐步加宽/平移文本框（最多约 12mm），仍不行才做 ≤1.5pt 的微缩，避免出现"新行只有一个字"。
- **引用带预留**：模块内若有引用，底部保留一条引用带，其他文字不得压入该带；蓝色横幅等细条模块按"不高 16mm"处理，边框内边距收紧且不裁掉条内文字。
- 字号、字库（latin/ea/cs）、加粗、斜体、对齐参与统一；**颜色按「强调色」体系归一**：一般文本 `normal`（黑）、小标题 `weak`（弱强调，主题深蓝）、行内/关键词突出 `strong`（强强调，红）、目录当前项 `accent`（主题强调蓝）。模板入库时从全稿统计出 `text_colors`，合成时按用途套用；源稿里不属于这四类的颜色（例如用于对应图中元素的橙色 `#ED7D31`）保持原样，留给后续的「对应色」机制。
- **对应色（概念，暂未应用）**：解释图片内容时，文字可复用图中该元素的颜色（如图中橙色部分用橙色文字）。`text_colors.corresponding` 已预留字段，当前不做自动匹配。
- 引用/文献用结构化字段，不要手工拼格式：`{"annotation": "citation", "citation": {"journal": "Nat. Mater.", "year": "2021", "volume": "21", "pages": "503"}}` 会按 catalog 渲染成期刊名斜体、年份加粗、卷号斜体、页码常规并居中。引用识别要接受 7–8 位文章编号（如 `2314419`）与 e-locator（如 `e2304567`），否则会漏成 caption 导致格式不一致；引用按单行渲染（不换行），与图下文献条目的原设计一致。
- 短数值/单位标签（如 `9 kPa`、`1.82 mV/K`）归为 `data_label`：加粗 + 强强调色，避免被误判成小标题而变色。
- **上标下标是格式不是字体**：用普通字体 + run 属性 `baseline`（上标 `30000`，下标 `-25000`），不要用 Unicode 上标字符或专用「上标字体」。行内斜体、加粗、颜色用 `paragraphs[].runs` 逐段表达，例如 `{"text": "Fe(CN)"}, {"text": "6", "baseline": -25000}, {"text": "3-", "baseline": 30000}`。

## 双板块与留白率

需要把版面分成左右（或上下）两个板块时，不要手算坐标，用 `panels` 并让渲染器保证留白均衡：

```json
{"layout": 17, "title": "双板块", "panels": {
  "area": [8.0, 26.0, 322.0, 156.0], "direction": "horizontal",
  "ratio": "golden", "gap_mm": 8.0, "padding_mm": 4.0, "whitespace_tolerance": 0.08,
  "panels": [
    {"border": {"color": "#B8B8B7", "width_pt": 2.0, "dash": "round_dot"}, "items": [...]},
    {"border": {"color": "#B8B8B7", "width_pt": 2.0, "dash": "round_dot"}, "items": [...]}]}}
```

- `ratio` 支持 `equal`(1:1)、`golden`(1:1.618)、`sqrt2`、`silver`、`double` 或任意正数；比例指第一个板块与第二个板块的宽（高）之比，不要求 1:1。
- 每个板块的留白率 = 1 − 内容包围盒占比；两者相差超过 `whitespace_tolerance`（默认 0.08）会报错，避免一边塞满一边空荡。不平衡时调整内容或比例，不要靠加空行凑数。
- 双板块同样受边框约束：内容越界、文字溢出、非设计意图的重叠都会报错。
- 模块支持 `"fill"` 底色（可带 `"alpha_pct"` 透明度），用于原稿里的浅色卡片；`fill` 与 `border` 可单独或同时使用。

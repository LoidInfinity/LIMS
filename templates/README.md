# PPT 模板库

按学校分组的 PPT 模板数据库。默认情况下每条记录只保存**母版 + 版式 + 主题 + 版式引用的图片**；正文、备注和只出现在幻灯片上的图片在入库时被丢弃，因此库里的文件可以公开分发，也不会夹带投稿人的课题内容。需要复现封面、目录这类画在幻灯片上的设计时，用 `--keep-slides` 显式保留页面原型（见下文）。

## 结构

```text
templates/
  index.json                  # 目录索引，由 ingest 自动重建，可 diff、可审阅
  <school-slug>/<template-id>/
    template.json             # 学校、标签、比例、主题、版式几何、出处
    template.pptx             # 精简包：0 页幻灯片，母版与版式完整
    preview.png               # 版式缩略图（几何预览，不含原文）
```

- `school-slug`、`template-id` 用小写拉丁字母、数字和连字符：清华 → `tsinghua`，模板 → `tsinghua-thesis-16x9`。中文校名无法自动转拼音，入库时必须给 `--school-slug`。
- `template.json` 必须通过 [`template.schema.json`](../shared/research_style/template.schema.json)。
- `index.json` 与磁盘不一致时，`query_template.py index` 会报错并提示重建。

## 添加一个模板

```powershell
python ppt-deck/scripts/ingest_template.py "答辩模板.pptx" `
  --school 清华大学 --school-slug tsinghua --id tsinghua-thesis-16x9 `
  --name "清华答辩模板" --tags 学术汇报,答辩 `
  --contributor "你的名字" --license "作者已授权本仓库分发"

python ppt-deck/scripts/query_template.py list --school 清华 --tag 答辩
python ppt-deck/scripts/query_template.py show tsinghua-thesis-16x9
```

入库时脚本会：解析母版/版式/主题 → 丢弃幻灯片与备注 → 生成精简包 → 生成几何预览 → 写描述 → 重建索引。同一个 id 已存在时需显式 `--force`。

## 页面原型

封面、目录页的校徽、装饰线、底纹常常直接画在幻灯片上，母版和版式里没有，只靠版式恢复不了。需要逐像素复现时，入库时显式保留这些页：

```powershell
python ppt-deck/scripts/ingest_template.py "开题.pptx" --school 华南理工大学 --school-slug scut `
  --id scut-kaituo-16x9 --keep-slides 1=cover,2=toc --license "作者本人提供"
```

- 保留的页以 `slide-N`/自定义标签记入 `template.json` 的 `prototypes[]`，同时记录每个文字形状的 `shape_id` 与原文，供合成时定位。
- 它们的正文与图片会留在 `template.pptx` 里，**不再满足「无内容」约定**；公开入库前请把示例文字换成占位文字，并在 `--license` 里写清授权。
- 入库还会采集一张内容页的标题样式（字号、加粗、颜色、对齐）存为 `content_style`，并保存标题形状的原始 XML（`title_xml`）；合成内容页时整形状复用，标题部分可与原稿逐像素一致。

## 按模板复现

```powershell
python ppt-deck/scripts/compose_deck.py --template-id scut-kaituo-16x9 --spec compose.json --output output/deck.pptx
```

`compose.json` 每页是「原型页」或「版式页」之一：

```json
{"template": "scut-kaituo-16x9", "pages": [
  {"prototype": "cover", "texts": {"2": "论文标题", "9": "答 辩 人：张三"}},
  {"layout": 17, "title": "研究内容", "body": ["第一点", "第二点"]}
]}
```

- 原型页只改 `texts` 指定的文字槽，不改几何；不写 `texts` 即原样复现（可做到逐像素一致）。每个原型只能用一次。
- 文字槽若是「按文字定宽、左对齐」的旧式文本框，换字后会自动居中，避免视觉上跑偏；需要固定方向时写成 `{"text": "...", "align": "left"}`。
- 版式页用模板版式新建，标题形状取自 `content_style.title_xml`；幻灯片上的自由图形、图片不会自动复现。
- 文字槽尺寸固定，替换后变长会换行或压到相邻元素，交付前必须渲染检查。

## 内容页模块

内容页通常用虚线框分成模块。先把原页图片抽出来，再用 `modules` 描述：

```powershell
python tools/extract_media.py "开题.pptx" --slide 3 --output assets/slide03
```

```json
{"template": "scut-kaituo-16x9", "base_dir": ".", "pages": [
  {"layout": 17, "title": "热电转换", "modules": [
    {"rect": [8.2, 26.9, 156.4, 89.8], "padding_mm": 3.0,
     "border": {"color": "#B8B8B7", "width_pt": 2.0, "dash": "round_dot"},
     "items": [
       {"box": [0, 0, 1, 0.10], "text": "小标题", "font_pt": 16, "bold": true, "color": "#1F3864", "align": "center"},
       {"box": [0.02, 0.11, 0.96, 0.74], "image": "assets/slide03/slide03-shape21.png"},
       {"box": [0, 0.87, 1, 0.13], "text": "图注", "font_pt": 10.5, "align": "center"}]}]}]}
```

- `rect` 为毫米、相对幻灯片左上角；`items[].box` 为模块内边距以内的相对比例。
- 边框默认直角：原稿常是 `roundRect` 且圆角值 0，渲染即直角；需要圆角时显式给 `"radius_mm": 3`。
- 模块支持 `"fill"` 底色（`{"color": "#4472C4", "alpha_pct": 15}` 表达半透明卡片），可与 `border` 单独或同时使用。
- 合成强制内容留在边框内侧：越界、文字超出盒子、item 重叠（除显式 `"overlay": true`）都会报错，不会让文字压到虚线上。
- `image` 路径相对 spec 文件，或用 spec 的 `base_dir` 指定；图片按保持比例缩放并居中。

整页复现可先用分析工具自动起草，再按需微调：

```powershell
python tools/analyze_page.py "开题.pptx" --slide 4 --slide 5 --assets assets/all `
  --spec spec-auto.json --layout 17 --template scut-kaituo-16x9
```

它按 XML 读出每页的虚线框（`prst`/圆角值/线型/线宽）、底色（含透明度）、图片与文本（字号、字体、颜色含 lumMod、项目符号），并逐页记录真实版式与标题 XML；同名版式位于不同母版时索引不同，工具会按页区分。

文本项保留三套字库：`font_latin`（英文与数字）、`font_cjk`（中文）、`font_cs`（复杂文种），主题字体别名（`+mn-lt` 等）会先解析成实际字面名。项目符号取段落的 `a:buFont`/`a:buChar`（本库模板小标题是 Wingdings `Ø`，渲染为三角箭头），连同 `a:buSzPct` 与 `marL/indent` 一起写回，而不是用 `• ` 前缀冒充。

行内格式用 `paragraphs[].runs` 表达（逐段文字/斜体/加粗/颜色/字号），上标下标用 run 的 `baseline` 属性（`30000` / `-25000`）配普通字体，而不是 Unicode 上标字符。文本框 `insets_mm` 与 `anchor` 也按原稿保留。

文字项可带 `role`（heading/subheading/body/caption/citation/note 等）先判定用途再套格式；引用用结构化字段：

```json
{"role": "citation", "citation": {"journal": "Nat. Mater.", "year": "2021", "volume": "21", "pages": "503"}}
```

渲染为期刊名斜体、年份加粗、卷号斜体、页码常规并居中。从源稿复现时在 spec 顶层设 `"role_styles": false`，避免角色默认格式覆盖原稿未显式设置的属性。

## 双板块与留白率

需要把一页分成左右（或上下）两块时，用 `panels` 让版式比例和留白自动受控：

```json
{"layout": 17, "title": "双板块", "panels": {
  "area": [8.0, 26.0, 322.0, 156.0], "direction": "horizontal",
  "ratio": "golden", "gap_mm": 8.0, "padding_mm": 4.0,
  "panels": [{"items": [...]}, {"items": [...]}]}}
```

`ratio` 可取 `equal`、`golden`(1:1.618)、`sqrt2`、`silver`、`double` 或任意正数。每块留白率 = 1 − 内容包围盒占比，两块相差超过 `whitespace_tolerance`（默认 0.08）即报错。

验证渲染差异可用 `tools/` 下两个脚本：`render_pptx.ps1` 用 PowerPoint 导出 PNG，`compare_renders.py` 逐像素比较并输出差异图。

## 规则

1. **只提交你有权分发的模板。** `--license` 要写清授权来源（例如「作者已授权本仓库分发」）；来源不明的模板不要入库。
2. 入库前确认母版里除校徽、页脚等模板元素外没有个人信息或未公开数据。
3. 不要手工编辑 `template.pptx`；需要改动就用 `--force` 重新入库，保证描述与索引一致。
4. 提交前跑 `python -m unittest discover -s tests`，并确认 `query_template.py index --check` 通过。

## 已知边界

- 版式里只记录**显式声明**的占位符几何；从母版继承、但版式未写出的位置目前不会解析（合成时由版式/母版继承补齐）。
- `preview.png` 是几何预览（色块 + 版式名），不是 PowerPoint 渲染图，因此不会泄露原文，也看不出字体细节。
- 一个原型只能用于一页；同一设计要多页复用需要扩展复制逻辑。
- 原型页的文字槽是固定尺寸，不做自动缩放；替换文字过长会换行甚至重叠。
- 内容页模块只复现模块边框与模块内内容；原页里不属于模块的自由图形不会被复制。
- `tools/install.py` 目前只分发四个技能与风格资产，不打包 `templates/`；模板库按仓库路径使用。

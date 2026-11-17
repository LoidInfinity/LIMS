# Research Style Skill Suite

一个以 Agent Skills 标准组织的科研技能套件。对话由宿主提供，技能触发由宿主负责；套件只包含技能说明、可执行脚本和可版本化的风格资产。

当前版本 **0.6.0**。共享配置、提取证据、同步、独立安装、Matplotlib/PPTX/DOCX 生成已实现；模板库支持页面原型（封面/目录逐像素复现）、内容页模块化复现（自动起草几何/底色透明度/字体/项目符号、逐页标题形状复用）与双板块版式（黄金分割等比例 + 留白率校验）；Origin 原生适配器只映射系列样式，尚待有许可证的真实环境验收；详见下方能力范围。

## 目录

```text
style-hub/             # 提取、校验、切换、同步
origin-chart/          # Origin / Matplotlib / plan
ppt-deck/              # 可编辑 PPTX、模板库入库与查询
paper-draft/           # DOCX 草稿与基础引用
templates/             # 按学校分组的 PPT 模板库（精简包、页面原型、索引）
shared/
  research_style/      # 共用 Python 代码和 JSON Schema
  styles/
    current.json       # 当前生效的配置
    presets/           # lab-meeting / thesis / journal
tools/                 # 安装与独立 ZIP 打包
examples/              # 明确标注为合成数据的演示输入
tests/                 # 数据、样式、安装及同步契约测试
```

四个 `SKILL.md` 均使用标准 `name` / `description` 元数据。详细规则在 `references/`，入口脚本在 `scripts/`；没有常驻服务、界面或自建调度器。[Agent Skills 规范](https://agentskills.io/specification)

## 本地运行

Python 3.10+。PowerShell 下，在本仓库目录执行：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe style-hub/scripts/sync_style.py validate

.\.venv\Scripts\python.exe origin-chart/scripts/render_origin.py examples/data.csv --x time --y control treatment --y-error control_sd treatment_sd --backend matplotlib --output output/chart.png --xlabel "Time (h)" --ylabel "Signal (a.u.)"
.\.venv\Scripts\python.exe ppt-deck/scripts/build_pptx.py examples/deck.json --output output/meeting.pptx
.\.venv\Scripts\python.exe paper-draft/scripts/assemble_paper.py examples/paper.json --output output/manuscript.docx
```

macOS/Linux 把 Python 路径替换为 `.venv/bin/python`。输出已有时脚本默认报错，确认替换才加 `--force`。示例 PPT 的图像路径指向 `output/chart.png`，因此先执行绘图。

每份结果附带 `文件名.style.json`，记录后端、配置来源、配置 SHA-256 和限制。相同指纹证明配置一致，不保证不同渲染器的视觉结果完全一致。

## 提取、修改与同步风格

```powershell
.\.venv\Scripts\python.exe style-hub/scripts/extract_style.py history.pptx history.docx figure.svg --output output/my-style.json --id my-lab --name "我的课题组"
.\.venv\Scripts\python.exe style-hub/scripts/sync_style.py diff shared/styles/current.json output/my-style.json
.\.venv\Scripts\python.exe style-hub/scripts/sync_style.py activate output/my-style.json --dry-run
.\.venv\Scripts\python.exe style-hub/scripts/sync_style.py activate output/my-style.json
```

提取器输出候选配置与 `.evidence.json`。不能识别的属性沿用基础配置，报告逐项列出；不会修改历史文件。需要的线型、引用规则等可按 [字段规范](style-hub/references/style-schema.md) 编辑候选 JSON，然后校验。提取本身不激活配置。

切换预设时将 `shared/styles/presets/thesis.json` 等路径传给 `activate`。三个预设是通用示例，不代表具体期刊/学校的官方规则。

读取优先级：`--style FILE` → `RESEARCH_STYLE_FILE` → 源码相邻 `shared/styles/current.json` → 安装包 `assets/styles/current.json`。显式覆盖无效就报错。`activate` 修改持久默认值，不修改环境变量；检查渲染时注意临时覆盖。

## PPT 模板库

`templates/` 按学校分组保存 PPT 模板。每条记录只有母版、版式、主题和版式引用的图片：幻灯片正文、备注和只用于某页的图片在入库时被丢弃，所以库内文件可以公开分发。版式几何写进 `template.json`，构建 PPT 时不必再猜位置。

```powershell
# 入库：一个 PPTX -> templates/<school-slug>/<template-id>/
.\.venv\Scripts\python.exe ppt-deck/scripts/ingest_template.py "答辩模板.pptx" --school 清华大学 --school-slug tsinghua `
  --id tsinghua-thesis-16x9 --name "清华答辩模板" --tags 学术汇报,答辩 --contributor "姓名" --license "作者已授权本仓库分发"

# 查询：按学校/标签/比例/版式类型/关键词筛选
.\.venv\Scripts\python.exe ppt-deck/scripts/query_template.py list --school 清华 --tag 答辩 --aspect 16:9
.\.venv\Scripts\python.exe ppt-deck/scripts/query_template.py show tsinghua-thesis-16x9

# 索引：检查或重建 index.json
.\.venv\Scripts\python.exe ppt-deck/scripts/query_template.py index --check
```

`preview.png` 是版式的几何预览（色块 + 版式名），不是 PowerPoint 渲染图。目录约定、授权要求与已知边界见 [模板库说明](templates/README.md)。

封面、目录的设计若直接画在幻灯片上（校徽图片、装饰线、底纹），母版与版式恢复不了。需要逐像素复现时入库加 `--keep-slides` 保留**页面原型**，再用 `compose_deck.py` 合成：

```powershell
.\.venv\Scripts\python.exe ppt-deck/scripts/ingest_template.py "开题.pptx" --school 华南理工大学 --school-slug scut `
  --id scut-kaituo-16x9 --keep-slides 1=cover,2=toc --license "作者本人提供"
.\.venv\Scripts\python.exe ppt-deck/scripts/compose_deck.py --template-id scut-kaituo-16x9 --spec compose.json --output output/deck.pptx
```

`compose.json` 每页取「原型页」或「版式页」之一：`{"prototype": "cover", "texts": {"2": "标题"}}` 只替换文字槽、几何不动（不写 `texts` 即原样复现），`{"layout": 17, "title": "...", "body": ["..."]}` 用模板版式新建并套用入库采集的内容页标题形状。原型页保留原文与图片，公开入库前需替换为占位文字。

内容页常见虚线模块版式，可先抽出原页图片再用 `modules` 复现：

```powershell
.\.venv\Scripts\python.exe tools/extract_media.py "开题.pptx" --slide 3 --output assets/slide03
# compose.json 中：{"layout": 17, "title": "...", "modules": [{"rect": [x,y,w,h], "padding_mm": 3,
#   "border": {"color": "#B8B8B7", "width_pt": 2, "dash": "round_dot"},
#   "items": [{"box": [0,0,1,0.1], "text": "小标题", ...}, {"box": [...], "image": "assets/slide03/..."}]}]}
```

模块内 `items[].box` 用相对比例定位；合成会拒绝越界、文字超出盒子以及 item 相互重叠（除显式 `"overlay": true`），因此文字不会压到虚线上。模块还支持 `"fill"`（可带 `"alpha_pct"` 透明度）表达原稿的浅色卡片，内容页标题复用入库捕获的标题形状 XML。

整页复现可由 `tools/analyze_page.py` 自动起草：它按 XML 读出每页虚线框的预设几何/圆角值/线型、底色与透明度、图片、文字（字号、**中文/英文/数字三套字库**、颜色含 lumMod、**原始项目符号**——本库模板小标题是 Wingdings `Ø` 三角箭头而非圆点），并逐页记录真实版式与标题 XML（同名版式在不同母版下索引不同，必须逐页区分）。

```powershell
.\.venv\Scripts\python.exe tools/analyze_page.py "开题.pptx" --slide 4 --slide 5 --assets assets/all `
  --spec spec-auto.json --layout 17 --template scut-kaituo-16x9
```

双板块版面用 `panels` 描述，比例与留白自动受控：

```json
{"layout": 17, "title": "双板块", "panels": {
  "area": [8.0, 26.0, 322.0, 156.0], "direction": "horizontal",
  "ratio": "golden", "gap_mm": 8.0, "padding_mm": 4.0,
  "panels": [{"items": [...]}, {"items": [...]}]}}
```

`ratio` 支持 `equal`、`golden`(1:1.618)、`sqrt2`、`silver`、`double` 或任意正数；两块留白率相差超过 `whitespace_tolerance`（默认 0.08）即报错，避免一边塞满一边空荡。

渲染验收用 `tools/render_pptx.ps1`（PowerPoint COM 导出 PNG）与 `tools/compare_renders.py`（逐像素比较，可输出差异图）。本机实测：一份 20 页开题报告入库后合成，封面与目录页渲染结果与原稿 **SHA-256 相同、0 个差异像素**；**全部 20 页的标题区域均为 0 个差异像素**；第 3 页为手工模块复现，第 4–20 页由分析工具自动起草，整页差异 1.3%–18.8%（+第 3 页 22.2%），主要来自行内强调色、项目符号字形与装饰性自由图形。


必须使用安装器或打包 ZIP，使单个技能携带运行代码和风格快照。只复制源码中的一个技能文件夹会丢失共用依赖。Python 库仍需安装，技能 ZIP 不含 Python 解释器。

| 宿主 | 项目目录 | 用户目录 | 依据 |
| --- | --- | --- | --- |
| Codex | `.agents/skills/` | `~/.agents/skills/` | [OpenAI 官方技能文档](https://learn.chatgpt.com/docs/build-skills) |
| Claude Code | `.claude/skills/` | `~/.claude/skills/` | [Claude Code 技能文档](https://code.claude.com/docs/en/skills) |
| dsh | `.dsh/skills/` | `$DSH_HOME/skills` 或 `~/.dsh/skills/` | [dsh-cc 技能发现文档](https://dsh-cc.github.io/guide/skills)；原生技能 provider 支持范围依版本确认 |

```powershell
# 查看目标；随后去掉 --dry-run 执行安装。
.\.venv\Scripts\python.exe tools/install.py --host codex --scope project --project D:/my-research --dry-run
.\.venv\Scripts\python.exe tools/install.py --host claude --scope user --dry-run
.\.venv\Scripts\python.exe tools/install.py --host dsh --scope user --dry-run

# 自定义目录，或仅安装一个可独立运行的技能。
.\.venv\Scripts\python.exe tools/install.py --dest D:/my-research/skills
.\.venv\Scripts\python.exe tools/install.py --dest D:/another-project/skills --skills origin-chart

# 四个独立 ZIP + 一个包含全部技能的 research-style-suite.zip。
.\.venv\Scripts\python.exe tools/package.py --output-dir dist
```

安装器拒绝覆盖同名目录，不会删除已有技能。代码升级可先安装到新目录验证，再自行替换旧版本；风格更新用同步命令即可。完整技能目录应直接放在宿主的技能根目录下，避免额外包一层 `research-style-suite/` 导致只扫描一层的宿主发现不到。

多个宿主的已安装快照需要显式同步：

```powershell
.\.venv\Scripts\python.exe style-hub/scripts/sync_style.py activate output/my-style.json --skills-dir D:/project/.agents/skills --skills-dir D:/project/.claude/skills
.\.venv\Scripts\python.exe style-hub/scripts/sync_style.py check --skills-dir D:/project/.agents/skills --skills-dir D:/project/.claude/skills
```

安装后的 style-hub 不指定 `--skills-dir` 时会更新相邻的已安装套件技能。同步只写匹配套件标识的风格文件；普通写入失败会回滚，跨目录断电/并发事务不作保证。详见 [同步规则](style-hub/references/sync-rules.md)。

宿主可自动选择这四个技能；也可在支持的宿主显式调用，如 Codex 的 `$style-hub`、Claude Code 的 `/style-hub`。保持描述简短，不将固定上下文占比视为跨版本保证。需要只允许手动调用时，Codex 可用 `agents/openai.yaml` 的 `allow_implicit_invocation: false`；本套件默认保留自动触发。[OpenAI 官方说明](https://learn.chatgpt.com/docs/build-skills)

## 能力范围

| 路径 | 已实现 | 当前边界 |
| --- | --- | --- |
| 提取 | PPTX 显式字号/字体、主题色和尺寸；DOCX 命名样式及页面；SVG 显式属性；逐字段证据 | 不完整解析 Office 继承/CSS；不提取 PDF、位图、OPJU；不猜引用体例 |
| Origin | CSV 数据、XY 系列、误差棒、系列色/线型/线宽/标记，保存 OPJU | Windows + Origin 2021+ + originpro；字体、画布、轴、网格、背景等依赖模板；真实集成待验收 |
| Matplotlib | PNG/SVG/PDF，共享 chart/typography/palette | 字体须可用；当前为线性 XY 图，不做统计分析 |
| PPT | 可编辑标题/正文、图片、说明、分页 | 估算文字容量；需 PowerPoint/LibreOffice 视觉验收 |
| 模板库 | PPTX → 母版/版式/主题精简包 + 版式几何 + 几何预览 + 出处；索引漂移检查；按学校/标签/比例/版式/关键词查询；可选保留页面原型 | 只记录版式显式声明的占位符几何，母版继承未解析；预览非渲染图 |
| 模板复现 | 原型页保留原幻灯片全部形状，只改文字槽，可逐像素复现封面/目录；内容页按 XML 自动起草模块规格（框几何/线型/底色透明度/图片，字体分 latin·ea·cs 三套，项目符号取原始 buFont+buChar），标题逐页复用捕获形状；双板块按 equal/golden/sqrt2 等比例切分并校验留白率；渲染与逐像素比对脚本 | 每个原型限用一次；文字槽固定尺寸不做自动缩放；不保留行内强调色（多色 run）；模块外的自由图形（箭头、图示）不会复现 |
| 论文 | DOCX 样式、页边距、行距、页码、基础引用与参考文献 | 非 CSL/Zotero；无图表/公式排版；来源需另行核验 |

Origin 不可用时会明确报错，不自动替换渲染后端。可用 `--backend plan --output output/chart-plan.json` 保存数据与样式计划，或显式指定 `--backend matplotlib`。原生后端要求来自 [Origin 官方 Python 包说明](https://pypi.org/project/originpro/)。

## 验证与维护

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

测试覆盖风格校验与覆盖优先级、非法 CSV、引用顺序及缺失引用、OOXML 格式属性、提取证据、独立安装、快照差异、同步回滚、三个输出的相同风格指纹，以及模板入库的无内容保证、媒体裁剪、目录漂移、查询过滤、页面原型保留、原型文字改写不改几何、原样原型的 XML 往返一致性、内容页标题形状复用、虚线模块布局与越界/重叠拒绝、模块底色透明度、双板块比例与留白率校验、定宽文本框换字居中、分析工具的几何/颜色/素材/字体/项目符号提取。它们不代替 Office/Origin 的视觉与真实宿主自动触发验收。

2026-11-17 本地验证（33 项自动化测试通过）：模板精简包经本机 PowerPoint 打开，0 页、1 套母版、11 个版式，无修复提示；20 页开题报告入库后合成，封面与目录页渲染与原稿 SHA-256 相同（0 差异像素），全部 20 页标题区域 0 差异像素；第 3–20 页均由分析工具起草，整页差异 0.5%–18.8%，无内容与虚线边框相交；小标题的 Wingdings `Ø` 三角箭头与黑体字体已按原始 XML 复现；双板块演示页按 1:1.618 切分（194.1 mm : 119.9 mm），两侧留白率 0.03 / 0.08。2026-10-22 验证：四个技能通过 `skill-creator` 校验；12 项自动化测试通过；示例的两页 PPT 已经本机 PowerPoint 渲染并检查，论文已由 Word 导出 PDF 并检查，修复了母版继承导致的重复项目符号。Origin 原生集成与三种宿主中的实际自动触发尚未验证。Windows Office 可选复验命令如下，输出目录须尚不存在：

```powershell
./tools/verify_office.ps1 -Presentation ./output/meeting.pptx -Document ./output/manuscript.docx -OutputDirectory ./output/review-new
```

修改共用代码只需编辑 `shared/research_style/`，之后重新打包；不要手工维护多个已安装副本的代码。修改 schema 时同步更新三个消费者和测试。建议将 `shared/styles/` 纳入 Git，使风格修改可以 diff、审阅和恢复。

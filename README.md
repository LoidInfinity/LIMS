# Research Style Skill Suite

一个以 Agent Skills 标准组织的科研技能套件。对话由宿主提供，技能触发由宿主负责；套件只包含技能说明、可执行脚本和可版本化的风格资产。

当前版本 **0.1.0**。共享配置、提取证据、同步、独立安装和 Matplotlib/PPTX/DOCX 生成已实现。Origin 原生适配器只映射系列样式，尚待有许可证的真实环境验收；详见下方能力范围。

## 目录

```text
style-hub/             # 提取、校验、切换、同步
origin-chart/          # Origin / Matplotlib / plan
ppt-deck/              # 可编辑 PPTX
paper-draft/           # DOCX 草稿与基础引用
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

## 安装到宿主

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
| 论文 | DOCX 样式、页边距、行距、页码、基础引用与参考文献 | 非 CSL/Zotero；无图表/公式排版；来源需另行核验 |

Origin 不可用时会明确报错，不自动替换渲染后端。可用 `--backend plan --output output/chart-plan.json` 保存数据与样式计划，或显式指定 `--backend matplotlib`。原生后端要求来自 [Origin 官方 Python 包说明](https://pypi.org/project/originpro/)。

## 验证与维护

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

测试覆盖风格校验与覆盖优先级、非法 CSV、引用顺序及缺失引用、OOXML 格式属性、提取证据、独立安装、快照差异、同步回滚和三个输出的相同风格指纹。它们不代替 Office/Origin 的视觉与真实宿主自动触发验收。

2026-10-22 本地验证：四个技能通过 `skill-creator` 校验；12 项自动化测试通过；示例的两页 PPT 已经本机 PowerPoint 渲染并检查，论文已由 Word 导出 PDF 并检查，修复了母版继承导致的重复项目符号。Origin 原生集成与三种宿主中的实际自动触发尚未验证。Windows Office 可选复验命令如下，输出目录须尚不存在：

```powershell
./tools/verify_office.ps1 -Presentation ./output/meeting.pptx -Document ./output/manuscript.docx -OutputDirectory ./output/review-new
```

修改共用代码只需编辑 `shared/research_style/`，之后重新打包；不要手工维护多个已安装副本的代码。修改 schema 时同步更新三个消费者和测试。建议将 `shared/styles/` 纳入 Git，使风格修改可以 diff、审阅和恢复。

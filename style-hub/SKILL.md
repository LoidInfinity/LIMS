---
name: style-hub
description: 从历史 PPTX、DOCX、SVG 提取可追溯的科研风格，管理配色、字体和版式预设，并同步到图表、PPT 与论文技能。适用于提取我的风格、切换风格、检查跨媒介一致性。
metadata:
  version: "0.1.0"
---

# 科研风格中心

把用户风格保存为可校验的 JSON。只处理风格；研究内容、实验结果和参考文献不由本技能生成。

## 工作方式

1. 确定当前技能目录，运行 `scripts/sync_style.py show` 读取实际生效配置。解析顺序是 `--style`、`RESEARCH_STYLE_FILE`、套件根目录 `shared/styles/current.json`、独立安装包 `assets/styles/current.json`。显式路径无效就报错，不悄悄降级。
2. 用户提供历史文件时，用 `scripts/extract_style.py` 生成候选配置和证据报告。读取报告里的 `fields`、`ambiguous`、`retained_from_base`、`warnings`，说明提取到了什么、还有哪些默认值。不要把未观察到的字段描述为用户偏好。
3. 用户要求应用或同步时，校验并激活候选配置；已有授权无需重复确认。只要求提取时保留候选文件，不改变当前风格。
4. 用户要求多个宿主一致时，对明确指定的技能安装目录逐一传入 `--skills-dir`，再执行 `check`。不搜索或改写无关安装。

所有命令相对于本技能目录；宿主应使用实际绝对路径执行，不能假定当前工作目录等于技能目录。

```sh
python scripts/extract_style.py history.pptx history.docx figure.svg --output candidate.json --id my-lab --name "我的课题组"
python scripts/sync_style.py validate --style candidate.json
python scripts/sync_style.py diff base.json candidate.json
python scripts/sync_style.py activate candidate.json --dry-run
python scripts/sync_style.py activate candidate.json --skills-dir /path/to/host/skills
python scripts/sync_style.py check --style candidate.json --skills-dir /path/to/host/skills
```

源码中的预设在 `../shared/styles/presets/`，独立安装后在 `assets/styles/presets/`。切换预设直接将该 JSON 传给 `activate`。

## 约束与按需阅读

- 编辑字段前读 [风格规范](references/style-schema.md)；排查安装、环境变量和同步问题时读 [同步规则](references/sync-rules.md)。
- 提取支持 PPTX、DOCX、SVG 的部分显式属性；不直接解析 PDF、位图或 OPJU。引用体例、图表语义和复杂母版需要用户提供规范或人工检查，不根据外观猜测。
- `activate` 改变持久默认值，忽略临时 `RESEARCH_STYLE_FILE`；临时覆盖仍会影响渲染，应向用户指出冲突。
- 提交/交付风格时给出配置路径、SHA-256 和保留默认值的说明。源文件不修改；生成候选覆盖已有文件需 `--force`。
- Python 3.10+；需要 `jsonschema`。完整依赖见源码根目录或安装包 `assets/requirements.txt`。

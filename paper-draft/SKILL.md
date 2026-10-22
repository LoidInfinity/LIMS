---
name: paper-draft
description: 依据共享科研风格组织和排版论文 DOCX 草稿，将用户提供的章节、摘要及参考文献映射到统一字体、页边距、字号和引用体例。适用于论文草稿组装、学位论文格式统一。
metadata:
  version: "0.1.0"
---

# 论文草稿

执行前读取套件 `shared/styles/current.json`；独立安装读 `assets/styles/current.json`。脚本依次优先使用 `--style` 和 `RESEARCH_STYLE_FILE`。将 `paper`、`typography`、`palette.text/accent` 应用于 DOCX。

1. 使用用户提供的研究材料组织内容。区分已验证结果、作者解释和占位内容；不编造实验、数据、作者或文献。
2. 读取 [内容与引用规则](references/paper-style-rules.md)，把章节和来源写成 JSON。引用以 `[@key]` 表示；仅使用已提供/查证的文献信息。缺失来源先标明并补齐，不绕过引用校验。
3. 运行 `scripts/assemble_paper.py`。支持基本编号引用和作者—年份引用；期刊精确格式需要指定 CSL/官方模板并另行处理，不能把通用格式称为已满足投稿规范。
4. 检查输出中所有引用和参考文献，再用 Word、LibreOffice 或宿主渲染检查换页、标题、字体、页边距。无法渲染时明确标为待视觉验收。

```sh
python scripts/assemble_paper.py manuscript.json --output manuscript.docx
python scripts/assemble_paper.py manuscript.json --style thesis.json --output thesis-draft.docx
```

命令使用实际绝对路径。`--force` 才会覆盖已有输出。交付 DOCX 与 `.style.json`，保留内容 JSON 便于复现；说明文献核验范围与仍待作者确认的内容。依赖为 `python-docx`、`jsonschema`，完整依赖见源码根目录或安装包 `assets/requirements.txt`。

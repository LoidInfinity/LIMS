---
name: ppt-deck
description: 按共享科研风格或某所高校的既有模板制作可编辑 PPTX，用于组会、学术汇报和研究结果展示；支持从一份 PPTX 提取模板入库，并按学校、标签、比例和版式检索调用。
metadata:
  version: "0.2.0"
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
- 模板库目前提供检索与版式几何；按模板版式直接生成幻灯片仍在开发中，生成阶段仍以 `build_pptx.py` 为准，并向用户说明这一限制。

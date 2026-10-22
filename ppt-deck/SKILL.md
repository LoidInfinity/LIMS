---
name: ppt-deck
description: 按共享科研风格制作可编辑 PPTX，用于组会、学术汇报和研究结果展示。将经确认的大纲、文字及本地图表转成统一字体、配色、尺寸和页边距的幻灯片。
metadata:
  version: "0.1.0"
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

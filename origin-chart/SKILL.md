---
name: origin-chart
description: 读取共享科研风格，从 CSV 绘制 XY 曲线、散点连线和误差棒，生成 Origin OPJU 工程或明确选择的 Matplotlib 图像。适用于科研数据绘图与跨媒介配色、线型一致化。
metadata:
  version: "0.1.0"
---

# 科研图表

执行前读取 `shared/styles/current.json`；独立安装时读本技能的 `assets/styles/current.json`。脚本支持 `--style` 和 `RESEARCH_STYLE_FILE` 优先覆盖；不要另外硬编码风格。

1. 明确 CSV 中 X、Y、单位及误差列含义。保留原始行顺序、数值、异常值；未经要求不拟合、平滑、截断坐标或补造统计量。
2. 读取 [数据与后端规则](references/chart-rules.md)。Origin 是默认后端，需要 Windows 和已安装的 Origin。当前原生适配器只自动映射系列颜色、线型、线宽、标记；其余视觉属性由 Origin 模板控制，不能宣称完全同步。
3. 使用本技能目录下 `scripts/render_origin.py` 执行，按用户需要选择可编辑工程或图像。无 Origin 时先说明限制：可生成 `plan` 或经用户选择使用 `matplotlib`，不得把它称为 Origin 输出。
4. 检查图像或 Origin 工程中的坐标、图例、误差棒、文字和线条，读取 `.style.json` 的指纹及警告。只有视觉检查完成后才称为可交付图表。

```sh
python scripts/render_origin.py data.csv --x time --y control treatment --backend origin --output chart.opju
python scripts/render_origin.py data.csv --x time --y control treatment --backend matplotlib --output chart.svg --xlabel "Time (h)" --ylabel "Signal (a.u.)"
python scripts/render_origin.py data.csv --x time --y control --y-error sd --backend plan --output chart-plan.json
```

命令中的输入与输出路径是示例；使用实际绝对路径。`--force` 可覆盖已有输出，不能覆盖输入数据。PNG/SVG/PDF 仅供 Matplotlib 后端；OPJU 仅供 Origin 后端。PNG 可以直接放进 PPT。

依赖与风格必须完整安装。缺少运行代码时用套件 `tools/install.py` 或独立 ZIP 重装，不临时猜测样式。

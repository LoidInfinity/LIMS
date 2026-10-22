# 图表输入与风格映射

输入是带列名的 UTF-8 CSV（可有 BOM）。`--x` 一列，`--y` 一列或多列，`--y-error` 可选且必须与 Y 一一对应。所有选中列必须是有限数值，误差非负；空值、NaN、重复列名或缺列报错。不会丢行、排序或插值。误差的统计定义由用户说明，脚本仅绘制给定大小。

| 后端 | 输出 | 行为 |
| --- | --- | --- |
| origin（默认） | `.opju` | 外部 Python 启动独立 Origin 会话，导入数据并生成工程；需要 Windows、Origin 2021+、originpro。当前仅系列样式自动映射。 |
| matplotlib | `.png/.svg/.pdf` | 全量应用 chart、typography、palette 中的绘图属性。固定画布尺寸，不采用 tight 裁切。字体缺失会报告替换风险。 |
| plan | `.json` | 保存数据、完整风格与坐标信息，不渲染、不启动 Origin；不能称为图表输出。 |

Origin `--template FILE.otp` 可提供自有模板。字体、画布、轴、网格、背景、标题、导出分辨率由模板/Origin 手动设置；`--xlabel/--ylabel/--title` 只在 Matplotlib/plan 中生效。原生适配器尚需在有许可证的 Origin 环境做验收；测试中的 dry plan 不等于原生集成已通过。

Origin 原生接口参考：[官方 external Python 示例](https://docs.originlab.com/externalpython/external-python-code-samples/)、[Plot 属性与 set_cmd](https://docs.originlab.com/originpro/classoriginpro_1_1graph_1_1Plot.html)。运行独立脚本，不从用户当前打开的 Origin 工程执行，也不通过 `op.new()` 清空会话。

交付前观察刻度、单位、图例重叠、中文缺字、误差棒、黑白可辨性；不要用美观理由改变数据。线性 XY 图为当前范围，热图、分面、拟合和统计检验需另外实现。

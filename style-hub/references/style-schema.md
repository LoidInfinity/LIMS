# 风格配置 v1

JSON 是唯一规范存储格式，UTF-8 编码。机器可读 JSON Schema 位于源码 `shared/research_style/style.schema.json`，打包后位于技能 `assets/runtime/research_style/style.schema.json`。不支持 YAML；避免不同解析器的隐式类型差异。

| 字段 | 意义与单位 | 消费者 |
| --- | --- | --- |
| `schema_version` | 固定整数 1；不兼容版本拒绝加载 | 全部 |
| `id`, `name` | 稳定机器标识、显示名称 | 全部 |
| `typography.latin/cjk` | 西文、中文字体名称 | 图表、PPT、论文 |
| `palette.background/text/muted/accent` | `#RRGGBB` 颜色 | 图表/PPT；论文用 text/accent |
| `palette.series` | 1–12 个系列色，按数据列顺序循环 | 图表 |
| `chart.width_mm/height_mm` | 画布毫米 | Matplotlib |
| `chart.font_pt/line_width_pt/marker_size_pt` | 字号、线宽、标记大小，pt | 图表（Origin 字号来自模板） |
| `chart.line_style` | solid/dashed/dotted/dashdot | 图表 |
| `chart.marker` | none/circle/square/triangle | 图表 |
| `chart.grid/dpi` | 网格开关、位图输出分辨率 | Matplotlib |
| `slides.width_mm/height_mm/margin_mm` | 页面尺寸、四边一致留白 | PPT |
| `slides.title_pt/body_pt/caption_pt` | 按用途分级，避免跨媒介复用绝对字号 | PPT |
| `paper.width_mm/height_mm/margins_mm` | 纸张尺寸、上下左右页边距 | DOCX |
| `paper.title_pt/heading_pt/body_pt/line_spacing` | 标题、一级标题、正文、行距倍数 | DOCX |
| `paper.citation_style` | numeric / author-year；仅基础体例 | DOCX |

所有字段必填，未知字段报错。颜色必须六位十六进制；尺寸与字号范围见 schema。校验另检查边距是否留出可用区域、数值是否有限。不要用 null 表示“猜一个”。需要新增字段时同时更新 schema、消费者、参考文档与测试。

三个预设是组会、学位论文和通用投稿示例，不代表任何期刊或学校的强制规范。用户要求官方格式时先查其模板。

## 提取与证据

提取以当前风格为底，修改能观察到的字段。报告保留命中值、观察次数、输入来源和歧义标记；同票按输入顺序取首个值。剩余字段列入 `retained_from_base`，不是从历史文件学得。来源报告独立保存，避免文件路径进入生产风格指纹。

- PPTX：页面尺寸、theme1 的系列色和基础字体、显式运行字体、占位标题/正文的显式字号。不完整解析母版继承。
- DOCX：分节页面尺寸和边距、Normal/Title/Heading1 的显式字体/字号、Normal 相对行距。不推断引用体例或直接格式覆盖。
- SVG：根尺寸、显式文本字体/字号、有色 stroke；不解析 CSS 级联、变换或曲线语义。无单位长度按 CSS 96 DPI 换算。
- 图表线宽/线型等未可靠提取的属性需人工编辑候选 JSON；PDF、位图、OPJU 不在当前提取范围内。

相同输入次序和基础配置得到相同配置与报告，不加入时间戳。样式指纹为排序键后的 JSON SHA-256；每次输出在旁边保存 `.style.json`。不要把指纹一致误认为渲染完全相同：字体、渲染器和模板同样影响视觉结果。

---
name: ppt-deck
description: 按共享科研风格或某所高校的既有模板制作可编辑 PPTX，用于组会、学术汇报和研究结果展示；支持从一份 PPTX 提取模板入库、保留封面/目录等页面原型并逐像素复现，也支持按学校、标签、比例和版式检索调用。
metadata:
  version: "0.3.0"
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

## 页面原型与复现

封面、目录这类页面的设计常常直接画在幻灯片上（校徽图片、装饰线、底纹形状），只靠母版和版式恢复不了。需要逐像素复现时，入库加 `--keep-slides` 把这些页保留为**页面原型**：

```sh
python scripts/ingest_template.py deck.pptx --school 华南理工大学 --school-slug scut \
  --id scut-kaituo-16x9 --keep-slides 1=cover,2=toc --license "作者本人提供"
python scripts/compose_deck.py --template-id scut-kaituo-16x9 --spec compose.json --output output/deck.pptx
```

`compose.json` 的每一页取二者之一：

```json
{"template": "scut-kaituo-16x9", "pages": [
  {"prototype": "cover", "texts": {"2": "新的论文标题", "9": "答 辩 人：张三"}},
  {"layout": 17, "title": "研究内容", "body": ["第一点", "第二点"]}
]}
```

- 原型页保留原幻灯片的全部形状，只改 `texts` 指定的文字槽（键为 shape id 或唯一形状名，id 见 `template.json` 的 `prototypes[].text_slots`）；不写 `texts` 就是原样复现。每个原型只能用一次。
- 版式页按模板版式新建，标题样式取自入库时采集的 `content_style`；幻灯片上的自由图形和图片不会被复现。
- 原型页文字槽是固定尺寸，替换文字变长会换行甚至压到相邻元素，交付前必须渲染检查。
- 合成后必须核对：用 `--keep-slides` 保留的原型页应做到渲染逐像素一致，内容页允许版式级差异，并向用户说明差异范围。

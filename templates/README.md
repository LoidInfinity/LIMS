# PPT 模板库

按学校分组的 PPT 模板数据库。每条记录只保存**母版 + 版式 + 主题 + 版式引用的图片**；正文、备注和只出现在幻灯片上的图片在入库时被丢弃，因此库里的文件可以公开分发，也不会夹带投稿人的课题内容。

## 结构

```text
templates/
  index.json                  # 目录索引，由 ingest 自动重建，可 diff、可审阅
  <school-slug>/<template-id>/
    template.json             # 学校、标签、比例、主题、版式几何、出处
    template.pptx             # 精简包：0 页幻灯片，母版与版式完整
    preview.png               # 版式缩略图（几何预览，不含原文）
```

- `school-slug`、`template-id` 用小写拉丁字母、数字和连字符：清华 → `tsinghua`，模板 → `tsinghua-thesis-16x9`。中文校名无法自动转拼音，入库时必须给 `--school-slug`。
- `template.json` 必须通过 [`template.schema.json`](../shared/research_style/template.schema.json)。
- `index.json` 与磁盘不一致时，`query_template.py index` 会报错并提示重建。

## 添加一个模板

```powershell
python ppt-deck/scripts/ingest_template.py "答辩模板.pptx" `
  --school 清华大学 --school-slug tsinghua --id tsinghua-thesis-16x9 `
  --name "清华答辩模板" --tags 学术汇报,答辩 `
  --contributor "你的名字" --license "作者已授权本仓库分发"

python ppt-deck/scripts/query_template.py list --school 清华 --tag 答辩
python ppt-deck/scripts/query_template.py show tsinghua-thesis-16x9
```

入库时脚本会：解析母版/版式/主题 → 丢弃幻灯片与备注 → 生成精简包 → 生成几何预览 → 写描述 → 重建索引。同一个 id 已存在时需显式 `--force`。

## 规则

1. **只提交你有权分发的模板。** `--license` 要写清授权来源（例如「作者已授权本仓库分发」）；来源不明的模板不要入库。
2. 入库前确认母版里除校徽、页脚等模板元素外没有个人信息或未公开数据。
3. 不要手工编辑 `template.pptx`；需要改动就用 `--force` 重新入库，保证描述与索引一致。
4. 提交前跑 `python -m unittest discover -s tests`，并确认 `query_template.py index --check` 通过。

## 已知边界

- 版式里只记录**显式声明**的占位符几何；从母版继承、但版式未写出的位置目前不会解析。
- `preview.png` 是几何预览（色块 + 版式名），不是 PowerPoint 渲染图，因此不会泄露原文，也看不出字体细节。
- `tools/install.py` 目前只分发四个技能与风格资产，不打包 `templates/`；模板库按仓库路径使用。

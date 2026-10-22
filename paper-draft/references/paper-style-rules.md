# 论文输入与引用

```json
{
  "title": "论文题目",
  "authors": "用户提供的作者",
  "abstract": "摘要，可引用 [@source1]。",
  "sections": [{"heading": "Introduction", "paragraphs": ["基于来源的陈述 [@source1]。"]}],
  "references": [{"id": "source1", "authors": "用户提供的作者", "year": "2026", "title": "用户提供的文献标题", "source": "期刊、卷页或 DOI"}]
}
```

title 和非空 sections 必填；authors/abstract 可选。每个 section 有 heading 和字符串数组 paragraphs。只识别 `[@key]`，其中 key 允许字母、数字、下划线、点、冒号和连字符。引用可用于摘要和正文。暂不支持标题里的引用、脚注、公式、表格、图片和交叉引用；应在下一步编辑或扩展脚本处理。

references 每项必填 id/authors/year/title/source 字符串，ID 唯一。缺失引用报错；没有 references 且没有引用时允许生成草稿。numeric 按首次出现的顺序编号，未引用来源排在其后；author-year 使用用户提供的作者显示字符串。当前为基础体例，非 CSL 引擎，不声称自动满足期刊/学校规范。source 是用户提供的来源文字，不自动联网补文献。

纸张、各边距、三层字号、行距和引用体例读取 paper；中文/西文字体读取 typography，文字/标题色读取 palette。使用 DOCX 命名样式与页码域，便于用户在 Word 中继续编辑。论文纸面背景保持白色。引用文本为静态文本，不是 Zotero/EndNote 引用域。

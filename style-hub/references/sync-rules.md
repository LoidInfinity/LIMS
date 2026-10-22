# 安装与同步规则

源代码目录共享 `shared/styles/current.json` 和 `shared/research_style/`。安装器将每个技能连同运行代码、全部预设、当前配置和套件标识复制成独立包；不依赖软链接或宿主的目录遍历行为。

配置解析顺序：显式 `--style` > 环境变量 `RESEARCH_STYLE_FILE` > 相邻 `shared/styles/current.json` > 技能内 `assets/styles/current.json`。选择后校验失败即报错，不跳到下一个来源。

`activate` 始终更新持久默认位置，忽略临时环境覆盖。源码运行时默认只更新 shared/current；可额外传多个 `--skills-dir` 同步安装包。打包后的 style-hub 默认同时同步相邻的套件技能。只有固定技能名和 `assets/suite.json` 套件标识匹配的目录才接受更新。目标无效时在写入前中止。

`--dry-run` 列出实际目标。每个文件使用同目录临时文件 + replace 原子替换，普通 I/O 失败回滚；跨目录断电一致性和并发写入不受保证，请串行运行。操作后用 `check` 检测所有快照；跨宿主不会后台自动同步。

`activate` 只改风格 JSON，不改脚本、SKILL.md 和用户研究材料。撤销方式是激活之前保存的 JSON/预设或从版本控制恢复。首次提取不会激活候选。安装器拒绝覆盖同名技能；版本升级应先另行打包验证，再由用户选择安装位置。

原始仓库中的单个技能文件夹不是独立分发包，不能直接只复制该文件夹。使用 `tools/install.py --skills NAME` 或 `tools/package.py` 产生的 ZIP；完整仓库按原布局放置则可以直接执行脚本。

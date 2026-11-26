# GitHub 首发检查

- 仓库建议 `brand-market-share-agent`，展示名「品牌市占分析 Agent」。在**解压后的项目根目录**单独初始化仓库，不在原项目或用户目录 `git add .`。
- 已经准备：可执行 CLI、Claude Code 子代理、合成输入、预期报告和 JSON、离线风险联动演示、原版 Excel 上传型看板和两份模拟月表。Agent 本身默认离线；只有原版 HTML 保留原品牌映射、代码、界面文字和第三方资产。不要把该 HTML 描述为完全脱敏。
- 手工复核 `git status --short` 与拟提交内容。不要添加内部文档 ID、凭据、缓存、原始导出或含真实经营数据的输出。README 截图若日后补充，只用合成样例并检查浏览器标签栏等边缘信息。
- **尚未授予项目级开源许可证。** 上传前确认原版 HTML、品牌映射、页面标识和嵌入依赖可对外分发，并保留第三方声明、补齐适用的许可文本或 NOTICE。若添加项目 LICENSE，明确适用范围，不要把原看板及第三方资产一起套用你无权授予的许可。公开可见不等于授权再分发。
- 本地跑 `python -m unittest discover -s tests -v`，试跑 README 命令；推送后确认 GitHub Actions 通过。此包不包含真实连接器，其权限、稳定性和回读策略需在授权环境单独验收。
- 可将 `dashboard/index.html` 放到 Release 或 GitHub Pages。它只展示合成月度数据，不读取原看板 IndexedDB。原版看板的真实底座 JSON、真实数据生成的离线 HTML、原始 ZIP 及历史聚合文件不应加入发布附件。

# 北极星运营

本仓库用于沉淀 WikiFX 自然北极星运营、C 端交易商查询与召回、用户归因字段设计，以及 Telegram India 查询 MVP 的方案、源码和可交付资产。

## 主要内容

- 自然北极星私域运营闭环：Markdown、HTML 与高清流程图。
- 用户归因数据字段图：HTML 与图片版本。
- WikiFX C 端交易商查询与召回：方案文档、流程图和生成脚本。
- Telegram India MVP：Bot 源码、交易商工作簿生成脚本、配置示例和实操手册。
- 数据分析模板与已导出的交付文件。

当前进度、已验证状态和下一步行动见 [PROJECT_STATUS.md](PROJECT_STATUS.md)。

## 安全约定

- 不提交真实 Token、`.env`、私钥或本地虚拟环境。
- 不提交 Telegram 本地数据库和事件明细导出，避免用户行为数据进入 Git 历史。
- `config.example.env` 只保留占位符；真实配置仅在本机或部署平台的 Secret 中维护。


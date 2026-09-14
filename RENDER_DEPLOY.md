# Render 测试部署

本仓库将 Telegram Bot 部署为一个单实例 Render Background Worker。Bot 继续使用 Telegram 长轮询；无需域名或 Webhook。

## 部署前准备

- Render 账号已连接能够访问本私有仓库的 GitHub 账号。
- 公司 Telegram 账号能够在 BotFather 中管理 `@wikifx_broker_check_india_bot`。
- 已取得用于管理员命令的 Telegram 数字 User ID。
- 不要把 Bot Token 写入仓库、Issue、聊天记录或截图。

## 创建 Render Blueprint

1. 登录 Render Dashboard。
2. 选择 **New > Blueprint**。
3. 连接 `hitechhamster/wikifx-telegram-bot`。
4. Branch 选择 `main`。
5. Blueprint 文件保持默认的 `render.yaml`。
6. Render 提示填写 Secret 时：
   - `TELEGRAM_BOT_TOKEN`：在 BotFather 轮换后得到的新 Token。
   - `TELEGRAM_ADMIN_IDS`：一个或多个数字 ID，多个值使用英文逗号分隔。
7. 确认费用后创建服务。Blueprint 会创建一个 Background Worker 和一个 1 GB 持久磁盘。

## 首次启动检查

部署日志应出现：

```text
India MVP bot is running
```

如果出现 Telegram `Conflict` 或类似“another getUpdates request”错误，说明同一 Token 仍被另一个实例使用。停止其他实例，只保留 Render 上的一个 Worker。

如果出现工作簿找不到，检查仓库根目录是否仍包含 `交易商名单.xlsx`，以及 Docker 构建日志是否成功复制仓库内容。

## Telegram 冒烟测试

使用非管理员测试账号依次执行：

```text
/start
/check XM
/check 0001461138
/check unknown-broker-test
```

然后测试：

- 关注交易商
- 取消关注
- 查看我的关注
- 群内 `/check@wikifx_broker_check_india_bot XM`
- Inline 查询 `@wikifx_broker_check_india_bot XM`

使用管理员账号验证隐藏命令：

```text
/stats
/recallpreview
```

确认预览结果后再测试实际推送或召回。

## 持久化和更新

SQLite 数据库保存到 `/data/telegram_mvp.db`，该路径由 Render 持久磁盘保存。代码更新和容器重启不会清除数据库。

每次推送 `main` 分支都会触发 Render 重新部署。部署后检查日志，并至少执行一次 `/check XM`。

## 迁移到公司服务器

公司服务器准备好后，可继续使用仓库中的 `Dockerfile`。将 `/data` 挂载为持久目录，并通过服务器 Secret 或受限环境文件注入同名环境变量。切换时先停止 Render Worker，再启动公司服务器，避免两个实例同时长轮询同一个 Bot Token。

# Telegram印度查询MVP：从空白账号到可查询交易商

## 已锁定的第一阶段

- 先做印度，统一使用英文。
- 巴基斯坦只做约1%的顺带覆盖，不单独建群。
- 印尼等印度MVP通过后复制，再改成Bahasa Indonesia。
- 自建官方频道和官方群，暂不依赖KOL。
- 交易商数据从根目录`交易商名单.xlsx`读取；当前为33,491条全量记录，原5条运营核验数据作为官网、牌照号和别名增强层。
- 保留Telegram User ID、Username和查询历史30天。
- 第一阶段支持用户主动关注、内部新闻推送和关注用户召回；实时新闻/风险变化仍需外部数据源，暂不做自动多语言。

## 资产结构

1. 真人管理员账号：`WikiFX Broker Check | India`
2. 内容频道：`WikiFX Broker Check India`
3. 查询与讨论群：`WikiFX Broker Check India Community`
4. 查询机器人：建议`WikiFXIndiaCheckBot`，具体取决于Username是否可用。

频道负责每天两条内容；群负责交易商查询、讨论和机器人交互。把群设置为频道的Discussion Group，让频道内容自动同步到群。

## 今晚操作顺序

### A. 完善真人账号

进入 `Settings > Edit Profile`：

- Name：`WikiFX Broker Check | India`
- Username：建议`wikifx_check_india`
- Bio：`Official WikiFX broker-risk information account for India. Check regulation, licences and risk signals. Information only.`
- 上传公司批准的方形Logo头像。

进入 `Settings > Privacy and Security`：

- 开启Two-Step Verification并绑定公司恢复邮箱。
- 开启Passcode Lock。
- Phone Number设为Nobody。
- Who can find me by my number设为My Contacts。
- Groups & Channels设为My Contacts。
- 检查Active Sessions并退出陌生设备。

当前是员工私人号码，但你准备后续更换。正式对外发布前，在 `Settings > Edit > Change Number` 更换为公司长期控制的号码；Telegram聊天、群和账号资料会保留，但换号前仍需确认公司内部交接责任。

### B. 创建频道

1. `New Message > New Channel`。
2. Name：`WikiFX Broker Check India`。
3. Description：`Daily broker regulation, licence and risk information for India. Two updates per day. Information only — not investment advice.`
4. 内部测试阶段设为Private。
5. 上传Logo，设置管理员，禁止无关人员发布。

### C. 创建讨论与查询群

1. `New Message > New Group`。
2. Name：`WikiFX Broker Check India Community`。
3. 先添加内部测试人员。
4. 内部测试阶段设为Private。
5. 群简介：`Use /check, then send a broker name or WikiFX ID. Do not share passwords, deposits, account numbers or identity documents.`
6. 频道 `Edit > Discussion > Link Group`，选择该群。

真人账号保持群Owner。机器人暂时只授予以下管理权限：Delete Messages、Ban Users、Pin Messages。不要给Add New Admins权限。

### D. 使用BotFather创建机器人

打开认证的`@BotFather`：

1. `/newbot`
2. Name：`WikiFX Broker Check India`
3. Username：建议`WikiFXIndiaCheckBot`
4. 将Token保存在密码管理工具中，不发群、不截图、不上传代码库。

然后执行 `/mybots`，选择机器人并配置：

- About：`Check broker regulation, licences and risk signals for India. Full evidence on WikiFX.`
- Description：`Send a broker name or WikiFX ID. The bot returns a short WikiFX risk snapshot and links to the full profile and WikiFX App.`
- Userpic：公司批准Logo。
- `/setjoingroups`：Enable。
- `/setprivacy`：Enable。第一阶段用命令和Inline查询，不需要读取群内所有普通消息。
- `/setinline`：Enable；Placeholder填写`Enter broker name or WikiFX ID`。

BotFather默认Commands只保留C端指令：

```text
start - Start and see how to use the bot
check - Check a broker
```

机器人启动后会在所有菜单作用域（包括管理员）固定只显示以上两个C端指令。`/start`返回欢迎介绍卡片，并提供“查找交易商”“我的关注”“打开App”按钮；查询结果卡片提供“关注新闻与风险更新”按钮。`privacy`、`stats`、`export`、`pushnews`、`recall`、`recallpreview`、`pin`、`delete`、`mute`和`unmute`作为隐藏的内部能力继续保留管理员权限校验，不占用C端菜单。召回只面向已主动关注且连续3天未查询的用户，同一用户两次召回至少间隔7天。

### E. 取得你的Telegram数字User ID

`TELEGRAM_ADMIN_IDS`必须填写数字ID，不能填Username。可先运行机器人，在内部测试版临时查看日志，或者使用公司批准的内部方法获取自己的User ID。不要把登录验证码或Bot Token交给第三方机器人。

### F. Windows本地运行

在本文件夹打开PowerShell：

```powershell
python --version
python -m venv .venv
Set-ExecutionPolicy -Scope Process Bypass
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
$env:TELEGRAM_BOT_TOKEN = Read-Host "Paste BotFather token"
$env:TELEGRAM_ADMIN_IDS = Read-Host "Paste your numeric Telegram user ID"
$env:WIKIFX_APP_ONELINK = "https://fxeye.onelink.me/Vm4A/intgbot"
python .\bot.py
```

终端显示`India MVP bot is running`后，不要关闭窗口。

### G. 私聊测试

依次执行：

```text
/start
/privacy
/check XM
/check 0001461138
/check Octa
/check unknown-broker-test
/stats
/export
```

预期结果：

- 全量数据支持名称、公司全称和10位WikiFX ID匹配；原5条增强数据继续支持官网和牌照号匹配。
- 无结果时明确提示未匹配，不猜测交易商。
- 结果用简短卡片显示WikiFX评分、国家/地区、WikiFX ID、监管概要和更新时间。
- 完整内容跳到精确WikiFX交易商页。
- App按钮跳现有WikiFX OneLink。
- 普通用户菜单只显示`/start`和`/check`；管理员可见并可使用运营命令。

### H. 群内测试

1. 把机器人加入测试群。
2. 机器人需要管理功能时，再设为管理员，并只开Delete、Ban、Pin。
3. 测试 `/check@机器人Username XM`。
4. 测试 `@机器人Username XM`。
5. 回复某条测试消息后执行 `/pin`、`/delete`、`/mute 10`和`/unmute`。

## 数据能看到什么

机器人当前可直接记录：

- 哪个Telegram User ID和Username发起查询。
- 查询发生时间、群、市场和来源标记。
- 用户输入的名称、官网或牌照号码。
- 唯一匹配、多结果、无结果。
- 返回了哪个Broker ID。

管理员可以使用 `/stats`查看30天汇总，并用`/export`导出CSV。

当前不能确认：

- 用户是否真正打开WikiFX网页。
- 是否点击App下载按钮。
- 是否安装App。
- 安装后是否进入对应交易商详情页。

原因是普通Telegram URL按钮不会把点击事件回传给机器人。下一步需要：

1. WikiFX服务器提供跟踪跳转链接，先记录`click_id/user_id/broker_id`再跳官网或OneLink。
2. GA4增加`web_deep_view`，携带`broker_id/source_id/page_type`。
3. AppsFlyer OneLink增加`campaign/source/broker_id`参数。
4. App读取深链参数并上报`app_deep_view`。

现有官网已经使用AppsFlyer OneLink做App下载，但是否支持“安装后直接回到指定Broker ID”尚未验证，必须让产品或AppsFlyer管理员确认。

## 一周后真实用户测试前的硬门槛

- 把私人号码换成公司可长期控制的号码，或形成书面交接责任。
- 提供可公开访问的Privacy Policy URL。
- 法务或负责人审核印度风险文案。
- 逐行核验Excel中的交易商数据和更新时间。
- 确认机器人Token只由负责人保管。
- 确认GA4、AppsFlyer和App事件负责人。
- 禁止收益承诺、喊单、入金引导和个性化投资建议。

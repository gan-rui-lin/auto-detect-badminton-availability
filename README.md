# 智慧珞珈场馆监控（API-only）

当前版本仅使用接口数据进行余量分析，不再依赖预约页/详情页的 DOM 结构。

## 功能

- 自动登录获取会话
- 通过接口查询指定球类、指定日期的各场馆各场地小时段余量
- 支持命令行时间范围过滤（例如 18:00-21:00，范围内时段均可）
- 支持命令行场馆映射多选（例如 `--venue 1,3`）
- 支持命令行选择球类（`--sport`，默认羽毛球）
- 支持命令行启用邮箱告警（`--email-alert`，配合 `config.yaml` 的 SMTP 配置）
- 输出精简结果：每个场馆只出现一次
- 支持提示音和 webhook 告警

## 安装

```bash
pip install -r requirements.txt
```

## 配置

主要配置在 config.yaml：

- auth.username / auth.password: 登录账号密码
- booking.sport_key: badminton 或 pingpong
- booking.sport_type_map: 球类到 typeId 的映射
- booking.next_day: 未指定 --date 时是否默认查次日
- monitor.interval_sec: 轮询间隔
- monitor.once_retries: `--once` 模式下额外重试次数（默认 2）
- monitor.once_retry_gap_sec: `--once` 模式重试间隔秒数（默认 1.2）
- monitor.once_refresh_before_retry: `--once` 重试前是否刷新页面
- monitor.dump_last_page_on_exit: 退出前是否保存最后页面调试文件
- monitor.last_page_snapshot_file: 最后页面 HTML 文件路径
- monitor.last_page_screenshot_file: 最后页面截图文件路径
- monitor.last_page_meta_file: 最后页面元数据 JSON 文件路径（含 URL 和时间）
- monitor.appointment_date_offset_days: 未指定 --date 时的日期偏移
- monitor.api_max_courts_per_venue: 每个场馆最多查询场地数
- monitor.consistency_rounds: 每轮检测对同一场馆采样次数（建议 2）
- monitor.consistency_round_gap_sec: 场馆采样间隔秒数
- monitor.venue_code_map: 场馆编号映射（用于 --venue）
- notification.webhook_url: 告警回调地址
- notification.beep: 是否蜂鸣提示
- notification.popup: 是否弹出顶置提醒窗口
- notification.popup_each_hit: 有余量时每轮检测都弹窗（推荐 true）
- notification.popup_title: 弹窗标题
- notification.email.enabled: 是否启用邮箱（可被命令行 `--email-alert` 打开）
- notification.email.smtp_host / smtp_port: SMTP 服务器地址与端口
- notification.email.use_ssl / use_starttls: 邮件加密方式
- notification.email.username / password: SMTP 登录账号密码（部分邮箱需授权码）
- notification.email.from_addr / to_addrs: 发件人与收件人列表
- notification.email.min_interval_sec: 邮件最小发送间隔（代码中强制 > 10 分钟）

### 邮箱配置（以 QQ 邮箱为例）

1. 登录 QQ 邮箱网页端，开启 SMTP 服务。
2. 在 QQ 邮箱中生成 SMTP 授权码。
3. 修改 `config.yaml` 中 `notification.email`：

```yaml
notification:
	email:
		enabled: true
		smtp_host: "smtp.qq.com"
		smtp_port: 465
		use_ssl: true
		use_starttls: false
		username: "你的QQ邮箱@qq.com"
		password: "这里填SMTP授权码，不是QQ登录密码"
		from_addr: "你的QQ邮箱@qq.com"
		to_addrs:
			- "接收提醒的邮箱@qq.com"
		subject_prefix: "[场馆余量提醒]"
		min_interval_sec: 660
```

字段说明：

- `enabled`: 是否启用邮箱通知
- `smtp_host` / `smtp_port`: SMTP 服务器和端口（QQ 常用 `smtp.qq.com:465`）
- `use_ssl` / `use_starttls`: 连接加密方式（465 端口建议 SSL）
- `username`: SMTP 登录账号，通常就是邮箱地址
- `password`: SMTP 授权码（不是邮箱登录密码）
- `from_addr`: 发件人地址，通常与 `username` 相同
- `to_addrs`: 收件人数组，可填多个邮箱
- `min_interval_sec`: 邮件最小发送间隔，代码中会强制大于 600 秒

运行时建议带上 `--email-alert`，即使命令行不带，只要 `notification.email.enabled: true` 也会生效。

## 运行

GUI 启动（Windows 桌面小应用原型）：

```bash
python monitor_gui.py
```

GUI 已支持：

- 参数填写（sport / venue / time-range / date / once / email-alert）
- 场馆下拉框（自动读取 `monitor.venue_code_map`，默认含 1-6 场馆）
- 日期默认今天，并支持日历控件选择（依赖 `tkcalendar`）
- 时间范围用开始/结束两个整数小时下拉（自动拼接为 `HH:00-HH:00`）
- 一键启动/停止监控
- 实时查看日志输出
- 打开最后一次调试快照（HTML/PNG）
- 一键开机自启（写入 Startup 文件夹）

球类参数（默认羽毛球）：

```bash
python monitor_pc.py --once --sport badminton
python monitor_pc.py --once --sport pingpong
```

只检测一次：

```bash
python monitor_pc.py --once
```

指定日期检测（推荐用于验证）：

```bash
python monitor_pc.py --once --date 2026-04-14
```

持续监控：

```bash
python monitor_pc.py
```

按映射编号筛选场馆（例如 1=风雨体育馆）：

```bash
python monitor_pc.py --once --venue 1 --date 2026-04-14
```

按映射编号多选场馆：

```bash
python monitor_pc.py --once --venue 1,3,6 --date 2026-04-14
```

按名称关键字筛选场馆：

```bash
python monitor_pc.py --once --venue 竹园 --date 2026-04-14
```

按名称关键字多选场馆：

```bash
python monitor_pc.py --once --venue 风雨,竹园 --date 2026-04-14
```

按时间范围筛选（只保留范围内小时段）：

```bash
python monitor_pc.py --once --time-range 18:00-21:00 --date 2026-04-14
```

组合使用（场馆 + 时间范围）：

```bash
python monitor_pc.py --once --sport pingpong --venue 1 --time-range 20:00-21:00 --date 2026-04-14
```

组合使用（场馆 + 时间范围 + 邮箱告警）：

```bash
python monitor_pc.py --once --sport pingpong --venue 1 --time-range 20:00-21:00 --date 2026-04-14 --email-alert
```

羽毛球场馆编号映射（默认，可在 `monitor.venue_code_map` 自定义）：

1 = 风雨体育馆

2 = 松园体育馆（工学部体育馆）

3 = 竹园体育馆（信部东区体育馆）

4 = 星湖体育馆（信部西区体育馆）

5 = 卓尔体育馆

6 = 杏林体育馆（医学部体育馆）

切换球类：

1. 命令行使用 `--sport badminton|pingpong`（推荐）
2. 或修改 config.yaml 中 `booking.sport_key`

## 输出说明

日志格式：

场馆名 | 1号场:18:00-19:00、19:00-20:00; 2号场:...

无余量时输出：

当前未发现余量

说明：不传 `--venue` 时默认全选全部场馆。

长时间监控建议：

1. 将 `monitor.interval_sec` 设为你的轮询间隔（例如 300）
2. 将 `notification.popup_each_hit` 设为 `true`
3. 将 `monitor.consistency_rounds` 设为 `2`，减少全查时的瞬时漏报
4. 这样每次检测到有余量都会置顶弹窗，不受冷却时间影响
5. 如需邮件提醒，命令行加 `--email-alert`，并配置 `notification.email`（默认最小间隔 660 秒）

调试建议：

1. `--once` 模式如果担心瞬时漏报，可保留 `monitor.once_retries: 2`（总共检测 3 轮）
2. 程序退出时会保存最后页面到 `page_snapshot.html`、`page_snapshot.png` 和 `page_snapshot.meta.json`
3. 这些文件可直接用于后续 GUI 小应用联调和问题复现

## 后台启动与关闭（Windows）

在 PowerShell 中后台启动（推荐：直接启动 `python`，保存的是 `python.exe` 的真实 PID）：

```powershell
$p = Start-Process python -ArgumentList "monitor_pc.py --sport pingpong --venue 2 --time-range 20:00-21:00 --date 2026-04-14" -WorkingDirectory "." -WindowStyle Hidden -PassThru -RedirectStandardOutput "monitor.out.log" -RedirectStandardError "monitor.err.log"
$p.Id | Set-Content monitor.pid
```

查看是否仍在运行（按 PID）：

```powershell
Get-Process -Id (Get-Content monitor.pid)
```

停止后台进程（按 PID）：

```powershell
Stop-Process -Id (Get-Content monitor.pid) -Force
```

如果 PID 文件丢失，可先查再停：

```powershell
Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match "monitor_pc\.py" } | Select-Object ProcessId,ParentProcessId,Name,CommandLine
Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match "monitor_pc\.py" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
```

注意：不要用 `Start-Process cmd /c ...` 的方式保存 PID。那样保存到 `monitor.pid` 的通常是 `cmd` 的 PID，不一定是实际 `python` 进程的 PID，可能导致“看似已停止但脚本仍在运行”。

## GUI 开发建议

当前 `monitor_gui.py` 是第一版原型，建议按下面顺序继续升级：

1. 将参数保存到独立 profile（多套配置一键切换）
2. 增加“测试登录 / 测试邮件”按钮，减少排障时间
3. 增加历史检测记录面板（按日期/场馆过滤）
4. 将监控核心抽成 service 层，GUI 只负责交互
5. 后续再迁移到 PySide6 以获得更完整的 Windows 桌面体验

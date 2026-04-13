# 智慧珞珈场馆监控（API-only）

当前版本仅使用接口数据进行余量分析，不再依赖预约页/详情页的 DOM 结构。

## 功能

- 自动登录获取会话
- 通过接口查询指定球类、指定日期的各场馆各场地小时段余量
- 支持命令行时间范围过滤（例如 18:00-21:00，范围内时段均可）
- 支持命令行场馆映射多选（例如 `--venue 1,3`）
- 支持命令行选择球类（`--sport`，默认羽毛球）
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

## 运行

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

# 智慧珞珈场馆监控与预约（PC 网页版）

本项目使用 Selenium + OCR 自动登录智慧珞珈网页端，支持：
1. 循环检测是否出现“可预约/余量”文本（告警）
2. 可选自动预约（按配置点击后一天、场地、时间段、提交）
3. 结构化余量分析（按场馆统计上午/下午/晚上可约数量）

## 1. 安装

```bash
pip install -r requirements.txt
```

## 2. 配置

已提供 `config.yaml`（可直接运行）和 `config.example.yaml`（模板）。

关键配置：

- `portal.type_id`: 默认球类 typeId
- `booking.sport_key`: `badminton` 或 `pingpong`
- `booking.sport_type_map`: 球类映射（可扩展，源码也有 `SPORT_TYPE_MAP`）
- `captcha.*`: 验证码识别配置
- `monitor.available_xpath_candidates`: 余量元素 XPath 候选
- `monitor.venue_xpath_template` / `monitor.period_xpath_template`: 场馆与时段结构化分析 XPath
- `monitor.dump_html_on_check`: 导出 `page_snapshot.html` 供 F12 对照调 XPath
- `notification.webhook_url`: 机器人回调地址（可选）

## 3. 运行

只检测一次：

```bash
python monitor_pc.py --once
```

持续监控：

```bash
python monitor_pc.py
```

开启自动预约：将 `config.yaml` 中 `booking.enabled` 改为 `true`。

仅导出页面中含关键字的 XPath 候选（帮助快速校准 XPath）：

```bash
python monitor_pc.py --dump-xpath-only
```

导出的候选文件路径由 `monitor.xpath_dump_file` 控制，默认是 `xpath_candidates.json`。

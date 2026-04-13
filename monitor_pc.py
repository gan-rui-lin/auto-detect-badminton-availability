from __future__ import annotations

import argparse
import ctypes
import json
import logging
import re
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List
from urllib import error as urlerror
from urllib import request as urlrequest

try:
    import yaml
except ImportError as exc:  # pragma: no cover
    raise SystemExit("缺少 PyYAML 依赖。请先执行: python -m pip install -r requirements.txt") from exc

from selenium import webdriver
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    import ddddocr
except ImportError:  # pragma: no cover
    ddddocr = None

try:
    import winsound
except ImportError:  # pragma: no cover
    winsound = None

LOGGER = logging.getLogger("zhlj-monitor")
SPORT_TYPE_MAP = {
    "badminton": 21,
    "pingpong": 22,
}
DEFAULT_VENUE_CODE_MAP = {
    "1": "风雨体育馆",
    "2": "松园体育馆",
    "3": "竹园体育馆",
    "4": "星湖体育馆",
    "5": "卓尔体育馆",
    "6": "杏林体育馆",
}


def setup_logging(level: str) -> None:
    logging.basicConfig(
        level=getattr(logging, level.upper(), logging.INFO),
        format="[%(asctime)s %(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def load_config(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        raise ValueError("配置文件格式错误，根节点必须是对象")
    return data


def normalize_text(text: str) -> str:
    return " ".join(str(text or "").split())


def build_period_label_map(monitor_cfg: Dict[str, Any]) -> Dict[int, str]:
    configured = monitor_cfg.get("period_labels", {})
    if isinstance(configured, dict) and configured:
        parsed: Dict[int, str] = {}
        for key, value in configured.items():
            try:
                idx = int(key)
            except (TypeError, ValueError):
                continue
            label = str(value).strip()
            if idx > 0 and label:
                parsed[idx] = label
        if parsed:
            return parsed

    start_hour = int(monitor_cfg.get("period_start_hour", 8))
    return {i: f"{start_hour + i - 1:02d}:00-{start_hour + i:02d}:00" for i in range(1, 14)}


def parse_periods_arg(periods_arg: str, monitor_cfg: Dict[str, Any]) -> List[int]:
    period_map = build_period_label_map(monitor_cfg)
    label_to_index = {v.replace(" ", ""): k for k, v in period_map.items()}

    raw_tokens = [x.strip() for x in periods_arg.split(",") if x.strip()]
    if not raw_tokens:
        raise ValueError("--periods 不能为空，例如 --periods 10,11 或 --periods 19:00-20:00")

    result: List[int] = []
    for token in raw_tokens:
        if token.isdigit():
            idx = int(token)
        else:
            idx = label_to_index.get(token.replace(" ", ""), 0)
        if idx <= 0:
            raise ValueError(
                f"无法识别时段参数: {token}。请使用时段索引(如 10,11)或时间段(如 19:00-20:00)"
            )
        result.append(idx)

    deduped: List[int] = []
    seen = set()
    for idx in result:
        if idx in seen:
            continue
        seen.add(idx)
        deduped.append(idx)
    return deduped


def parse_time_range(range_text: str) -> tuple[str, str]:
    text = str(range_text or "").strip()
    m = re.match(r"^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$", text)
    if not m:
        raise ValueError("--time-range 格式必须为 HH:MM-HH:MM，例如 18:00-21:00")
    start = m.group(1).zfill(5)
    end = m.group(2).zfill(5)
    if start >= end:
        raise ValueError("--time-range 起始时间必须早于结束时间")
    return start, end


def _resolve_venue_keyword_from_code(cfg: Dict[str, Any], code_text: str) -> str:
    monitor = cfg.setdefault("monitor", {})
    configured = monitor.get("venue_code_map", {})
    code_map: Dict[str, str] = {}
    if isinstance(configured, dict):
        for k, v in configured.items():
            key = str(k).strip()
            val = str(v).strip()
            if key and val:
                code_map[key] = val
    if not code_map:
        code_map = dict(DEFAULT_VENUE_CODE_MAP)

    venue_keyword = code_map.get(code_text, "").strip()
    if not venue_keyword:
        allowed = ", ".join(sorted(code_map.keys()))
        raise ValueError(f"--venue 不在映射中，当前可用编号: {allowed}")
    return venue_keyword


def apply_runtime_overrides(cfg: Dict[str, Any], args: argparse.Namespace) -> None:
    booking = cfg.setdefault("booking", {})
    monitor = cfg.setdefault("monitor", {})

    booking["sport_key"] = str(args.sport or "badminton").strip().lower()

    if args.venue:
        raw_tokens = [x.strip() for x in str(args.venue).split(",") if x.strip()]
        if not raw_tokens:
            raise ValueError("--venue 不能为空，例如 --venue 1,3 或 --venue 风雨,竹园")
        filters: List[str] = []
        for token in raw_tokens:
            if token.isdigit():
                filters.append(_resolve_venue_keyword_from_code(cfg, token))
            else:
                filters.append(token)
        # De-duplicate while preserving order.
        dedup: List[str] = []
        seen = set()
        for f in filters:
            key = normalize_text(f).lower()
            if key in seen:
                continue
            seen.add(key)
            dedup.append(f)
        monitor["venue_name_filters"] = dedup

    if args.time_range:
        start, end = parse_time_range(args.time_range)
        monitor["time_range"] = f"{start}-{end}"

    if args.date:
        date_text = str(args.date).strip()
        try:
            datetime.strptime(date_text, "%Y-%m-%d")
        except ValueError as exc:
            raise ValueError("--date 格式必须为 YYYY-MM-DD，例如 2026-04-14") from exc
        monitor["appointment_date"] = date_text


class ZhihuiLuojiaMonitor:
    def __init__(self, cfg: Dict[str, Any]) -> None:
        self.cfg = cfg
        self.driver: webdriver.Chrome | None = None
        self.ocr = ddddocr.DdddOcr(show_ad=False) if ddddocr is not None else None
        self.last_alert_ts = 0.0
        self.booked_once = False

    @property
    def browser(self) -> Dict[str, Any]:
        return self.cfg.get("browser", {})

    @property
    def portal(self) -> Dict[str, Any]:
        return self.cfg.get("portal", {})

    @property
    def auth(self) -> Dict[str, Any]:
        return self.cfg.get("auth", {})

    @property
    def captcha(self) -> Dict[str, Any]:
        return self.cfg.get("captcha", {})

    @property
    def monitor(self) -> Dict[str, Any]:
        return self.cfg.get("monitor", {})

    @property
    def notification(self) -> Dict[str, Any]:
        return self.cfg.get("notification", {})

    @property
    def booking(self) -> Dict[str, Any]:
        return self.cfg.get("booking", {})

    def start(self) -> None:
        options = webdriver.ChromeOptions()
        if self.browser.get("headless", True):
            options.add_argument("--headless=new")
        options.add_argument("--window-size=1366,2000")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        user_agent = str(self.browser.get("user_agent", "")).strip()
        if user_agent:
            options.add_argument(f"user-agent={user_agent}")

        user_data_dir = str(self.browser.get("user_data_dir", "")).strip()
        if user_data_dir:
            options.add_argument(f"--user-data-dir={user_data_dir}")

        self.driver = webdriver.Chrome(options=options)
        self.driver.set_page_load_timeout(int(self.browser.get("page_load_timeout_sec", 40)))
        self.driver.implicitly_wait(int(self.browser.get("implicit_wait_sec", 2)))

    def close(self) -> None:
        if self.driver is not None:
            self.driver.quit()
            self.driver = None

    def _must_driver(self) -> webdriver.Chrome:
        if self.driver is None:
            raise RuntimeError("浏览器尚未启动")
        return self.driver

    def _find_visible_element(self, xpath: str) -> Any | None:
        driver = self._must_driver()
        elements = driver.find_elements(By.XPATH, xpath)
        for element in elements:
            if element.is_displayed():
                return element
        return None

    def _find_first_visible(self, xpaths: List[str]) -> Any | None:
        for xpath in xpaths:
            if not xpath:
                continue
            element = self._find_visible_element(xpath)
            if element is not None:
                return element
        return None

    def _submit_login(self) -> None:
        driver = self._must_driver()
        submit_xpaths: List[str] = []
        configured = self.auth.get("submit_xpaths", [])
        if isinstance(configured, list):
            submit_xpaths.extend(str(x).strip() for x in configured if str(x).strip())
        submit_xpath = str(self.auth.get("submit_xpath", "")).strip()
        if submit_xpath:
            submit_xpaths.append(submit_xpath)
        submit_xpaths.extend(
            [
                '//*[@id="login_submit"]',
                '//button[@type="submit"]',
                '//input[@type="submit"]',
                '//*[contains(@class, "login-btn")]',
                '//*[contains(@class, "btn-login")]',
            ]
        )

        seen = set()
        for xpath in submit_xpaths:
            if not xpath or xpath in seen:
                continue
            seen.add(xpath)
            try:
                element = self._find_visible_element(xpath)
                if element is None:
                    continue
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
                try:
                    element.click()
                except WebDriverException:
                    driver.execute_script("arguments[0].click();", element)
                LOGGER.info("已点击提交按钮: %s", xpath)
                return
            except WebDriverException:
                continue

        password_candidates = list(self.auth.get("password_xpaths", []))
        password_candidates.append(str(self.auth.get("password_xpath", "")).strip())
        password_candidates.extend(
            [
                '//input[@name="password"]',
                '//input[contains(@placeholder, "密码")]',
                '//input[contains(@type, "password")]',
            ]
        )
        password_element = self._find_first_visible([x for x in password_candidates if x])
        if password_element is not None:
            password_element.send_keys(Keys.ENTER)
            LOGGER.info("提交按钮未命中，已使用回车提交")
            return

        raise TimeoutException("未找到可提交登录的按钮")

    def _is_login_success(self, success_xpath: str, success_url_contains: str) -> bool:
        driver = self._must_driver()
        if success_url_contains and success_url_contains in driver.current_url:
            return True

        if success_xpath:
            try:
                if len(driver.find_elements(By.XPATH, success_xpath)) > 0:
                    return True
            except WebDriverException:
                pass

        extra_url_markers = ["/pc/index", "/mobile/homepage", "/mobile/home"]
        if any(marker in driver.current_url for marker in extra_url_markers):
            return True

        text_markers_cfg = self.auth.get(
            "login_success_text_keywords",
            ["退出登录", "账号管理", "个人数据中心", "应用中心", "办事大厅"],
        )
        text_markers = [str(x).strip() for x in text_markers_cfg if str(x).strip()]
        if text_markers:
            try:
                text = normalize_text(str(driver.execute_script("return document.body ? (document.body.innerText || '') : '';")))
                if any(marker in text for marker in text_markers):
                    return True
            except WebDriverException:
                pass

        return False

    def _fill_captcha_if_needed(self) -> None:
        if not bool(self.captcha.get("enabled", True)):
            return

        input_xpath = str(self.captcha.get("input_xpath", "")).strip()
        image_xpath = str(self.captcha.get("image_xpath", "")).strip()
        refresh_xpath = str(self.captcha.get("refresh_xpath", image_xpath)).strip()
        if not input_xpath or not image_xpath:
            return

        captcha_input = self._find_visible_element(input_xpath)
        captcha_image = self._find_visible_element(image_xpath)
        if captcha_input is None or captcha_image is None:
            return

        if self.ocr is None:
            raise RuntimeError("检测到验证码，但未安装 ddddocr。请先执行: pip install -r requirements.txt")

        max_attempts = int(self.captcha.get("max_attempts_per_login", 3))
        for i in range(1, max_attempts + 1):
            captcha_image = self._find_visible_element(image_xpath)
            if captcha_image is None:
                return

            image_bytes = captcha_image.screenshot_as_png
            guess = self.ocr.classification(image_bytes).strip()
            guess = "".join(ch for ch in guess if ch.isalnum())
            if not guess:
                LOGGER.warning("验证码识别为空，第 %s 次重试", i)
                refresh_element = self._find_visible_element(refresh_xpath)
                if refresh_element is not None:
                    refresh_element.click()
                time.sleep(0.3)
                continue

            captcha_input = self._find_visible_element(input_xpath)
            if captcha_input is None:
                return
            captcha_input.clear()
            captcha_input.send_keys(guess)
            LOGGER.info("验证码识别结果: %s", guess)
            return

        raise RuntimeError("验证码识别失败：达到最大重试次数")

    def _resolve_type_id(self) -> int:
        type_map = dict(SPORT_TYPE_MAP)
        configured = self.booking.get("sport_type_map", {})
        if isinstance(configured, dict):
            for key, value in configured.items():
                type_map[str(key)] = int(value)

        sport_key = str(self.booking.get("sport_key", "")).strip()
        if sport_key:
            if sport_key not in type_map:
                raise ValueError(f"booking.sport_key={sport_key} 未在 sport_type_map 中定义")
            return int(type_map[sport_key])
        return int(self.portal.get("type_id", 21))

    def login(self) -> None:
        driver = self._must_driver()
        login_url = self.portal["login_url"]
        username = str(self.auth.get("username", "")).strip()
        password = str(self.auth.get("password", "")).strip()
        max_login_attempts = int(self.auth.get("max_login_attempts", 6))
        success_xpath = str(self.auth.get("login_success_xpath", "")).strip()
        success_url_contains = str(self.auth.get("login_success_url_contains", "")).strip()
        if not success_url_contains:
            success_url_contains = str(self.portal.get("home_url", "")).strip()
        success_timeout = int(self.auth.get("success_timeout_sec", 20))

        if not username or not password:
            raise ValueError("auth.username 或 auth.password 为空")

        username_xpaths = list(self.auth.get("username_xpaths", []))
        password_xpaths = list(self.auth.get("password_xpaths", []))
        username_xpaths.append(str(self.auth.get("username_xpath", "")).strip())
        password_xpaths.append(str(self.auth.get("password_xpath", "")).strip())
        username_xpaths.extend(
            [
                '//input[@name="username"]',
                '//input[contains(@placeholder, "用户名")]',
                '//input[contains(@id, "user")]',
            ]
        )
        password_xpaths.extend(
            [
                '//input[@name="password"]',
                '//input[contains(@placeholder, "密码")]',
                '//input[contains(@type, "password")]',
            ]
        )

        for attempt in range(1, max_login_attempts + 1):
            LOGGER.info("登录尝试 %s/%s", attempt, max_login_attempts)
            driver.get(login_url)
            try:
                WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                if self._is_login_success(success_xpath, success_url_contains):
                    LOGGER.info("已处于登录态，跳过登录表单")
                    return

                username_el = self._find_first_visible([x for x in username_xpaths if x])
                password_el = self._find_first_visible([x for x in password_xpaths if x])
                if username_el is None or password_el is None:
                    LOGGER.warning("未命中登录输入框，准备重试")
                    continue

                username_el.clear()
                username_el.send_keys(username)
                password_el.clear()
                password_el.send_keys(password)

                self._fill_captcha_if_needed()
                self._submit_login()

                WebDriverWait(driver, success_timeout).until(
                    lambda _d: self._is_login_success(success_xpath, success_url_contains)
                )
                LOGGER.info("登录成功")
                return
            except InvalidSessionIdException:
                LOGGER.warning("浏览器会话中断，正在重建浏览器后重试")
                try:
                    self.close()
                except WebDriverException:
                    pass
                self.start()
            except (TimeoutException, WebDriverException):
                LOGGER.warning("登录未成功，准备重试")

        raise RuntimeError("登录失败：超过最大重试次数")

    def open_reserve_page(self) -> None:
        driver = self._must_driver()
        reserve_url = self.portal["reserve_url_template"].format(type_id=self._resolve_type_id())
        driver.get(reserve_url)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    def _is_reserve_page(self) -> bool:
        return "/pages/index/reserve" in self._must_driver().current_url

    def _ensure_reserve_page(self) -> None:
        if self._is_reserve_page():
            return
        self.open_reserve_page()

    def _appointment_date_str(self) -> str:
        explicit = str(self.monitor.get("appointment_date", "")).strip()
        if explicit:
            return explicit
        offset = int(
            self.monitor.get(
                "appointment_date_offset_days",
                1 if bool(self.booking.get("next_day", False)) else 0,
            )
        )
        return (datetime.now() + timedelta(days=offset)).strftime("%Y-%m-%d")

    def _log_effective_query(self) -> None:
        sport = str(self.booking.get("sport_key", "")).strip() or "badminton"
        date_str = self._appointment_date_str()
        raw_filters = self.monitor.get("venue_name_filters", [])
        venue_filters: List[str] = []
        if isinstance(raw_filters, list):
            venue_filters = [normalize_text(str(x)) for x in raw_filters if normalize_text(str(x))]
        if not venue_filters:
            single = normalize_text(str(self.monitor.get("venue_name_filter", "")))
            if single:
                venue_filters = [single]
        time_range = normalize_text(str(self.monitor.get("time_range", ""))) or "全时段"
        venue_text = ",".join(venue_filters) if venue_filters else "全部场馆"
        LOGGER.info("查询参数: sport=%s date=%s venue=%s time_range=%s", sport, date_str, venue_text, time_range)

    def _api_fetch_json(self, url: str) -> Dict[str, Any]:
        driver = self._must_driver()
        script = r"""
const done = arguments[arguments.length - 1];
const url = arguments[0];

function findJwt() {
  const jwtRe = /^[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+$/;
  const stores = [window.localStorage, window.sessionStorage];
  for (const st of stores) {
    if (!st) continue;
    for (let i = 0; i < st.length; i++) {
      const k = st.key(i);
      const v = st.getItem(k) || '';
      if (jwtRe.test(v)) return v;
      try {
        const obj = JSON.parse(v);
        if (!obj || typeof obj !== 'object') continue;
        for (const key of Object.keys(obj)) {
          const val = String(obj[key] || '');
          if (jwtRe.test(val)) return val;
        }
      } catch (_) {}
    }
  }
  return '';
}

const token = findJwt();
const headers = { 'Accept': '*/*', 'Content-Type': 'application/json' };
if (token) headers['Authorization'] = 'Bearer ' + token;

fetch(url, {
  method: 'GET',
  credentials: 'include',
  headers,
}).then(async (resp) => {
  const text = await resp.text();
  let data = null;
  try { data = JSON.parse(text); } catch (_) {}
  done({ ok: resp.ok, status: resp.status, data, bodyText: text.slice(0, 300) });
}).catch((err) => {
  done({ ok: false, status: 0, error: String(err || '') });
});
"""
        try:
            raw = driver.execute_async_script(script, url)
        except WebDriverException:
            return {}
        if not isinstance(raw, dict):
            return {}
        data = raw.get("data")
        return data if isinstance(data, dict) else {}

    def _period_index_from_label(self, time_label: str) -> int:
        normalized = str(time_label).replace(" ", "")
        period_map = build_period_label_map(self.monitor)
        for index, label in period_map.items():
            if str(label).replace(" ", "") == normalized:
                return int(index)
        return 0

    def _segment_name(self, period_index: int) -> str:
        segments = self.monitor.get(
            "period_segments",
            {"morning": [1, 2, 3, 4, 5], "afternoon": [6, 7, 8, 9], "evening": [10, 11, 12, 13]},
        )
        for name, period_list in segments.items():
            if isinstance(period_list, list) and period_index in [int(x) for x in period_list]:
                return name
        return "other"

    def _segment_name_from_time_label(self, time_label: str) -> str:
        match = re.search(r"(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})", str(time_label))
        if not match:
            return "other"
        start_hour = int(match.group(1))
        if start_hour < 12:
            return "morning"
        if start_hour < 18:
            return "afternoon"
        return "evening"

    def _duration_minutes(self, start_hm: str, end_hm: str) -> int:
        sh, sm = [int(x) for x in start_hm.split(":")]
        eh, em = [int(x) for x in end_hm.split(":")]
        return (eh * 60 + em) - (sh * 60 + sm)

    def _in_time_range(self, start_hm: str, end_hm: str) -> bool:
        raw = str(self.monitor.get("time_range", "")).strip()
        if not raw:
            return True
        try:
            range_start, range_end = parse_time_range(raw)
        except ValueError:
            return True
        return start_hm >= range_start and end_hm <= range_end

    def _venue_name_matched(self, venue_name: str) -> bool:
        raw = self.monitor.get("venue_name_filters", [])
        kws: List[str] = []
        if isinstance(raw, list):
            kws = [normalize_text(str(x)).lower() for x in raw if normalize_text(str(x))]
        if not kws:
            # Backward compatibility for older config keys.
            single = normalize_text(str(self.monitor.get("venue_name_filter", ""))).lower()
            if single:
                kws = [single]

        if not kws:
            return True
        hay = normalize_text(venue_name).lower()
        return any(kw in hay for kw in kws)

    def _structured_availability_via_api(self) -> List[Dict[str, Any]]:
        type_id = self._resolve_type_id()
        date_str = self._appointment_date_str()
        consistency_rounds = max(1, int(self.monitor.get("consistency_rounds", 2)))
        consistency_round_gap_sec = float(self.monitor.get("consistency_round_gap_sec", 0.2))
        list_url = (
            f"https://gym.whu.edu.cn/api/GSStadiums/GetAppointmentList?Version=2"
            f"&SportsTypeId={type_id}&AppointmentDate={date_str}"
        )
        list_payload = self._api_fetch_json(list_url)
        if int(list_payload.get("status", 0)) != 200:
            return []

        resp = list_payload.get("response", {})
        if not isinstance(resp, dict):
            return []
        data = resp.get("data", [])
        if not isinstance(data, list):
            return []

        configured_venue_index = int(self.booking.get("venue_index", 0) or 0)
        max_courts_default = int(self.monitor.get("api_max_courts_per_venue", 24))

        result: List[Dict[str, Any]] = []
        for i, item in enumerate(data, start=1):
            if configured_venue_index > 0 and i != configured_venue_index:
                continue
            if not isinstance(item, dict):
                continue

            venue_name = normalize_text(str(item.get("Title", ""))) or f"venue-{i}"
            if not self._venue_name_matched(venue_name):
                continue
            area_id = int(item.get("StadiumsAreaId", 0) or 0)
            if area_id <= 0:
                continue

            LOGGER.info("API分析场馆[%s]: %s", i, venue_name)
            if consistency_rounds > 1:
                LOGGER.info("场馆[%s]一致性采样轮数: %s", i, consistency_rounds)

            detail_court_slots: Dict[str, set[str]] = {}
            total_capacity = max_courts_default
            # First court probe to discover total capacity.
            first_url = (
                "https://gym.whu.edu.cn/api/GSStadiums/GetAppointmentDetail?Version=3"
                f"&StadiumsAreaId={area_id}&StadiumsAreaNo=1&AppointmentDate={date_str}"
            )
            first_payload = self._api_fetch_json(first_url)
            if int(first_payload.get("status", 0)) == 200:
                first_resp = first_payload.get("response", {})
                if isinstance(first_resp, dict):
                    try:
                        total_capacity = int(first_resp.get("StadiumsArea", {}).get("TotalCapacity", total_capacity) or total_capacity)
                    except (TypeError, ValueError):
                        total_capacity = max_courts_default

            total_capacity = max(1, min(total_capacity, max_courts_default))

            for round_idx in range(1, consistency_rounds + 1):
                for area_no in range(1, total_capacity + 1):
                    detail_url = (
                        "https://gym.whu.edu.cn/api/GSStadiums/GetAppointmentDetail?Version=3"
                        f"&StadiumsAreaId={area_id}&StadiumsAreaNo={area_no}&AppointmentDate={date_str}"
                    )
                    payload = self._api_fetch_json(detail_url)
                    if int(payload.get("status", 0)) != 200:
                        continue

                    p_resp = payload.get("response", {})
                    if not isinstance(p_resp, dict):
                        continue

                    times = p_resp.get("AppointmentTimes", [])
                    if not isinstance(times, list):
                        times = []

                    for t in times:
                        if not isinstance(t, dict):
                            continue
                        can = int(t.get("IsCanAppointment", 0) or 0) == 1
                        remain = int(t.get("RemainingCapacity", 0) or 0)
                        if not can or remain <= 0:
                            continue
                        start = str(t.get("StartTime", "")).strip().zfill(5)
                        end = str(t.get("EndTime", "")).strip().zfill(5)
                        if not start or not end:
                            continue
                        duration = self._duration_minutes(start, end)
                        if duration < 45 or duration > 90:
                            continue
                        if not self._in_time_range(start, end):
                            continue
                        court_key = f"{area_no}号场"
                        detail_court_slots.setdefault(court_key, set()).add(f"{start}-{end}")

                if round_idx < consistency_rounds and consistency_round_gap_sec > 0:
                    time.sleep(consistency_round_gap_sec)

            detail_court_availability: Dict[str, List[str]] = {}
            for court_name, slots_set in detail_court_slots.items():
                slots = sorted(slots_set)
                if slots:
                    detail_court_availability[court_name] = slots

            slot_details: List[Dict[str, Any]] = []
            for court_name, slots in detail_court_availability.items():
                for label in slots:
                    idx = self._period_index_from_label(label)
                    slot_details.append(
                        {
                            "period_index": idx,
                            "time_label": label,
                            "segment": self._segment_name(idx) if idx > 0 else self._segment_name_from_time_label(label),
                            "available": True,
                            "court": court_name,
                            "source": "api_detail",
                        }
                    )

            segment_count: Dict[str, int] = {}
            for slot in slot_details:
                seg = str(slot.get("segment", "other"))
                segment_count[seg] = segment_count.get(seg, 0) + 1

            result.append(
                {
                    "venue_index": i,
                    "venue_name": venue_name,
                    "available_total": len(slot_details),
                    "segment_count": segment_count,
                    "slot_details": slot_details,
                    "court_availability": detail_court_availability,
                }
            )
        return result

    def _structured_availability(self) -> List[Dict[str, Any]]:
        return self._structured_availability_via_api()

    def check_availability(self) -> Dict[str, Any]:
        self._ensure_reserve_page()
        venues = self._structured_availability()

        deduped_by_name: Dict[str, Dict[str, Any]] = {}
        for v in venues:
            key = normalize_text(str(v.get("venue_name", ""))) or str(v.get("venue_index", ""))
            old = deduped_by_name.get(key)
            if old is None or int(v.get("available_total", 0)) > int(old.get("available_total", 0)):
                deduped_by_name[key] = v
        venues = list(deduped_by_name.values())

        positives: List[str] = []
        for venue in venues:
            if int(venue.get("available_total", 0)) <= 0:
                continue
            court_avail = venue.get("court_availability", {})
            court_lines: List[str] = []
            if isinstance(court_avail, dict):
                for court_name in sorted(court_avail.keys(), key=lambda x: int(re.findall(r"\d+", x)[0]) if re.findall(r"\d+", x) else 999):
                    times = [str(x) for x in court_avail.get(court_name, []) if str(x)]
                    if times:
                        court_lines.append(f"{court_name}:{'、'.join(times)}")
            positives.append(f"{venue['venue_name']} | {'; '.join(court_lines) if court_lines else '无'}")

        return {
            "has_available": len(positives) > 0,
            "positives": positives,
            "negatives": [],
            "url": self._must_driver().current_url,
            "venues": venues,
        }

    def log_availability(self, result: Dict[str, Any]) -> None:
        for line in result.get("positives", []):
            LOGGER.info("余量分析: %s", line)
        if not result.get("positives"):
            LOGGER.info("当前未发现余量")

    def _popup_alert(self, message: str) -> None:
        if not bool(self.notification.get("popup", True)):
            return
        title = str(self.notification.get("popup_title", "场馆余量提醒")).strip() or "场馆余量提醒"

        def _show() -> None:
            try:
                # MB_OK(0x0) | MB_ICONWARNING(0x30) | MB_SYSTEMMODAL(0x1000) | MB_TOPMOST(0x40000)
                ctypes.windll.user32.MessageBoxW(0, message, title, 0x0 | 0x30 | 0x1000 | 0x40000)
            except Exception:
                LOGGER.exception("弹窗提醒失败")

        threading.Thread(target=_show, daemon=True).start()

    def alert(self, result: Dict[str, Any]) -> None:
        msg = "检测到场馆有余量: " + " | ".join(result.get("positives", [])[:8])

        # For long-running monitoring, popup should be shown on every detection hit.
        if bool(self.notification.get("popup_each_hit", True)):
            self._popup_alert(msg)

        cooldown = int(self.monitor.get("alert_cooldown_sec", 300))
        now = time.time()
        if now - self.last_alert_ts < cooldown:
            return

        self.last_alert_ts = now
        LOGGER.warning(msg)

        if not bool(self.notification.get("popup_each_hit", True)):
            self._popup_alert(msg)

        if bool(self.notification.get("beep", True)) and winsound is not None:
            winsound.Beep(1300, 600)

        webhook = str(self.notification.get("webhook_url", "")).strip()
        if webhook:
            body = json.dumps(
                {
                    "text": msg,
                    "url": result.get("url", ""),
                    "time": datetime.now().isoformat(timespec="seconds"),
                },
                ensure_ascii=False,
            ).encode("utf-8")
            req = urlrequest.Request(
                webhook,
                data=body,
                headers={"Content-Type": "application/json"},
                method="POST",
            )
            try:
                with urlrequest.urlopen(req, timeout=10) as resp:
                    LOGGER.info("Webhook 已发送，状态码: %s", resp.status)
            except urlerror.URLError as e:
                LOGGER.error("Webhook 发送失败: %s", e)

def run_once(monitor: ZhihuiLuojiaMonitor) -> None:
    monitor.login()
    monitor.open_reserve_page()
    monitor._log_effective_query()
    result = monitor.check_availability()
    monitor.log_availability(result)
    if result.get("has_available", False):
        monitor.alert(result)


def run_loop(monitor: ZhihuiLuojiaMonitor) -> None:
    monitor.login()
    monitor.open_reserve_page()
    monitor._log_effective_query()

    interval = int(monitor.monitor.get("interval_sec", 20))
    refresh_each_cycle = bool(monitor.monitor.get("refresh_each_cycle", True))

    while True:
        result = monitor.check_availability()
        monitor.log_availability(result)
        if result.get("has_available", False):
            monitor.alert(result)
        time.sleep(interval)
        if refresh_each_cycle:
            monitor._must_driver().refresh()


def main() -> None:
    parser = argparse.ArgumentParser(description="智慧珞珈场馆余量监测（API-only）")
    parser.add_argument("--config", default="config.yaml", help="配置文件路径")
    parser.add_argument("--once", action="store_true", help="只检测一次")
    parser.add_argument(
        "--sport",
        choices=["badminton", "pingpong"],
        default="badminton",
        help="球类：badminton(羽毛球) 或 pingpong(乒乓球)，默认 badminton",
    )
    parser.add_argument("--venue", help="场馆筛选，支持多选逗号分隔；可用编号映射或名称关键字，例如 1,3 或 风雨,竹园")
    parser.add_argument("--time-range", help="时段范围过滤，格式 HH:MM-HH:MM，例如 18:00-21:00")
    parser.add_argument("--date", help="查询日期，格式 YYYY-MM-DD，例如 2026-04-14")
    parser.add_argument("--log-level", default="INFO", help="日志级别")
    args = parser.parse_args()

    setup_logging(args.log_level)
    cfg = load_config(Path(args.config).resolve())
    apply_runtime_overrides(cfg, args)

    monitor = ZhihuiLuojiaMonitor(cfg)
    monitor.start()
    try:
        if args.once:
            run_once(monitor)
        else:
            run_loop(monitor)
    finally:
        monitor.close()


if __name__ == "__main__":
    main()

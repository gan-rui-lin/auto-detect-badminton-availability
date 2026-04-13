from __future__ import annotations

import argparse
import json
import logging
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List
from urllib import error as urlerror
from urllib import request as urlrequest

try:
    import yaml
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "缺少 PyYAML 依赖。请先执行: python -m pip install -r requirements.txt"
    ) from exc
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    InvalidSelectorException,
    TimeoutException,
    WebDriverException,
)
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
    return " ".join(text.split())


def contains_any(text: str, keywords: List[str]) -> bool:
    return any(k in text for k in keywords)


def parse_rgb(rgb_text: str) -> tuple[int, int, int]:
    match = re.search(r"rgba?\((\d+),\s*(\d+),\s*(\d+)", rgb_text or "")
    if not match:
        return (0, 0, 0)
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


class ZhihuiLuojiaMonitor:
    def __init__(self, cfg: Dict[str, Any]) -> None:
        self.cfg = cfg
        self.driver: webdriver.Chrome | None = None
        self.ocr = ddddocr.DdddOcr(show_ad=False) if ddddocr is not None else None
        self.last_alert_ts = 0.0
        self.booked_once = False
        self.structured_warned = False

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
    def navigation(self) -> Dict[str, Any]:
        return self.cfg.get("navigation", {})

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

        user_agent = self.browser.get("user_agent", "").strip()
        if user_agent:
            options.add_argument(f"user-agent={user_agent}")

        user_data_dir = self.browser.get("user_data_dir", "").strip()
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

    def _wait_present(self, xpath: str, timeout: int) -> Any:
        driver = self._must_driver()
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )

    def _wait_clickable(self, xpath: str, timeout: int) -> Any:
        driver = self._must_driver()
        return WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )

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
        configured = self.auth.get("submit_xpaths", [])
        submit_xpaths: List[str] = []
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
        deduped = []
        for item in submit_xpaths:
            if item not in seen:
                seen.add(item)
                deduped.append(item)

        for xpath in deduped:
            try:
                element = self._find_visible_element(xpath)
                if element is None:
                    continue
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center',inline:'nearest'});", element
                )
                try:
                    element.click()
                except WebDriverException:
                    driver.execute_script("arguments[0].click();", element)
                LOGGER.info("已点击提交按钮: %s", xpath)
                return
            except (InvalidSelectorException, WebDriverException):
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

        raise TimeoutException("未找到可提交登录的按钮或输入框")

    def _log_login_state(self) -> None:
        driver = self._must_driver()
        LOGGER.warning("当前 URL: %s", driver.current_url)
        LOGGER.warning("当前标题: %s", driver.title)
        script = """
const nodes = [...document.querySelectorAll('button,input[type="submit"],a')].slice(0, 20);
return nodes.map((n) => ({
  tag: n.tagName.toLowerCase(),
  text: (n.innerText || n.value || '').trim().slice(0, 40),
  id: (n.id || '').toString().slice(0, 60),
  cls: (n.className || '').toString().slice(0, 80),
  type: (n.getAttribute('type') || '').toString().slice(0, 20),
}));
"""
        try:
            items = driver.execute_script(script)
            LOGGER.warning("页面可交互候选: %s", json.dumps(items, ensure_ascii=False))
        except WebDriverException:
            LOGGER.warning("无法读取页面候选按钮信息")

    def _is_login_success(self, success_xpath: str, success_url_contains: str) -> bool:
        driver = self._must_driver()
        if success_url_contains and success_url_contains in driver.current_url:
            return True
        if not success_xpath:
            return False
        try:
            return len(driver.find_elements(By.XPATH, success_xpath)) > 0
        except InvalidSelectorException:
            return False

    def _resolve_type_id(self) -> int:
        type_map = dict(SPORT_TYPE_MAP)
        configured = self.booking.get("sport_type_map", {})
        if isinstance(configured, dict):
            for key, value in configured.items():
                type_map[str(key)] = int(value)

        sport_key = str(self.booking.get("sport_key", "")).strip()
        if sport_key:
            if sport_key not in type_map:
                raise ValueError(
                    f"booking.sport_key={sport_key} 未在 sport_type_map 中定义，可修改源码 SPORT_TYPE_MAP 或配置 booking.sport_type_map"
                )
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

        username_xpaths = [x for x in username_xpaths if x]
        password_xpaths = [x for x in password_xpaths if x]

        for attempt in range(1, max_login_attempts + 1):
            LOGGER.info("登录尝试 %s/%s", attempt, max_login_attempts)
            driver.get(login_url)
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body"))
                )
                if self._is_login_success(success_xpath, success_url_contains):
                    LOGGER.info("已处于登录态，跳过登录表单")
                    return

                username_el = self._find_first_visible(username_xpaths)
                password_el = self._find_first_visible(password_xpaths)
                if username_el is None or password_el is None:
                    LOGGER.warning("未命中登录输入框，准备重试")
                    self._log_login_state()
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
            except (TimeoutException, WebDriverException):
                LOGGER.warning("登录未成功，准备重试")
                self._log_login_state()

        raise RuntimeError("登录失败：超过最大重试次数")

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
            raise RuntimeError(
                "检测到验证码，但未安装 ddddocr。请先执行: pip install -r requirements.txt"
            )

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

    def open_reserve_page(self) -> None:
        driver = self._must_driver()
        reserve_url = self.portal["reserve_url_template"].format(
            type_id=self._resolve_type_id()
        )
        driver.get(reserve_url)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        for xpath in self.navigation.get("click_xpaths", []):
            if not xpath:
                continue
            try:
                self._wait_clickable(xpath, 8).click()
                time.sleep(0.4)
            except TimeoutException:
                LOGGER.warning("导航点击失败，跳过 xpath=%s", xpath)

    def _wait_booking_mask_clear(self) -> None:
        mask_xpath = str(
            self.booking.get(
                "mask_xpath",
                '//uni-view[contains(@class,"mask") and contains(@class,"zindex")]',
            )
        ).strip()
        if not mask_xpath:
            return
        driver = self._must_driver()
        try:
            WebDriverWait(driver, int(self.booking.get("mask_timeout_sec", 6))).until(
                EC.invisibility_of_element_located((By.XPATH, mask_xpath))
            )
        except TimeoutException:
            pass

    def _click_booking_xpath(self, xpath: str, label: str, timeout: int = 8) -> None:
        driver = self._must_driver()
        last_error: Exception | None = None
        for _ in range(3):
            self._wait_booking_mask_clear()
            try:
                element = self._wait_clickable(xpath, timeout)
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center',inline:'nearest'});", element
                )
                try:
                    element.click()
                except ElementClickInterceptedException:
                    self._wait_booking_mask_clear()
                    driver.execute_script("arguments[0].click();", element)
                return
            except (TimeoutException, WebDriverException) as exc:
                last_error = exc
                time.sleep(0.35)
        raise RuntimeError(f"{label} 点击失败: {xpath}") from last_error

    def try_book(self) -> bool:
        if not bool(self.booking.get("enabled", False)):
            return False
        if self.booked_once and bool(self.booking.get("book_once", True)):
            LOGGER.info("已完成一次预约，按配置不再重复预约")
            return False

        next_day = bool(self.booking.get("next_day", False))
        next_day_xpath = str(self.booking.get("next_day_xpath", "")).strip()
        venue_xpath = str(self.booking.get("venue_xpath", "")).strip()
        court_number = int(self.booking.get("court_number", 0))
        court_offset = int(self.booking.get("court_row_offset", 1))
        period_indexes = self.booking.get("period_indexes", [])
        period_xpath_template = str(self.booking.get("period_xpath_template", "")).strip()
        court_xpath_template = str(self.booking.get("court_xpath_template", "")).strip()
        submit_order_xpath = str(self.booking.get("submit_order_xpath", "")).strip()
        step_wait = float(self.booking.get("step_wait_sec", 0.4))

        driver = self._must_driver()
        try:
            if next_day and next_day_xpath:
                self._click_booking_xpath(next_day_xpath, "后一天")
                time.sleep(step_wait)

            if venue_xpath:
                self._click_booking_xpath(venue_xpath, "场馆")
                time.sleep(step_wait)

            if court_number > 0 and court_xpath_template:
                court_xpath = court_xpath_template.format(court_row_index=court_number + court_offset)
                self._click_booking_xpath(court_xpath, "场地")
                time.sleep(step_wait)

            if isinstance(period_indexes, list) and period_xpath_template:
                for period in period_indexes:
                    period_xpath = period_xpath_template.format(period_index=int(period))
                    self._click_booking_xpath(period_xpath, f"时段{period}")
                    time.sleep(0.2)

            if submit_order_xpath:
                self._click_booking_xpath(submit_order_xpath, "提交预约")
                LOGGER.warning("已执行预约提交流程，请立即确认页面结果")
            else:
                LOGGER.warning("未配置 submit_order_xpath，已完成预约前步骤但未点击最终提交")

            self.booked_once = True
            return True
        except RuntimeError:
            LOGGER.error("预约失败：关键节点未找到，请检查 booking 下的 XPath 配置")
            LOGGER.error("当前页面 URL: %s", driver.current_url)
            return False

    def _discover_keyword_nodes(self, keywords: List[str]) -> List[Dict[str, str]]:
        driver = self._must_driver()
        script = """
const keywords = arguments[0];
function xpathFor(el) {
  if (el.id) return '//*[@id="' + el.id + '"]';
  const parts = [];
  while (el && el.nodeType === 1) {
    let ix = 1;
    let sib = el.previousElementSibling;
    while (sib) {
      if (sib.tagName === el.tagName) ix += 1;
      sib = sib.previousElementSibling;
    }
    parts.unshift(el.tagName.toLowerCase() + '[' + ix + ']');
    el = el.parentElement;
  }
  return '/' + parts.join('/');
}
const result = [];
const nodes = document.querySelectorAll('*');
for (const el of nodes) {
  const text = (el.innerText || el.textContent || '').trim();
  if (!text || text.length > 40) continue;
  for (const kw of keywords) {
    if (text.includes(kw)) {
      result.push({
        keyword: kw,
        text: text.replace(/\\s+/g, ' '),
        xpath: xpathFor(el),
        className: (el.className || '').toString().slice(0, 120),
      });
      break;
    }
  }
}
return result.slice(0, 600);
"""
        raw = driver.execute_script(script, keywords)
        if isinstance(raw, list):
            return [r for r in raw if isinstance(r, dict)]
        return []

    def _slot_meta(self, element: Any) -> Dict[str, Any]:
        driver = self._must_driver()
        script = """
const e = arguments[0];
const s = getComputedStyle(e);
return {
  text: (e.innerText || e.textContent || '').trim(),
  cls: (e.className || '').toString(),
  pointer: (s.pointerEvents || '').toString(),
  opacity: (s.opacity || '').toString(),
  bg: (s.backgroundColor || '').toString(),
  color: (s.color || '').toString(),
  disabled: !!e.disabled,
};
"""
        result = driver.execute_script(script, element)
        if isinstance(result, dict):
            return result
        return {}

    def _is_slot_available(self, meta: Dict[str, Any]) -> bool:
        text = normalize_text(str(meta.get("text", "")))
        cls = str(meta.get("cls", ""))
        pointer = str(meta.get("pointer", ""))
        opacity = str(meta.get("opacity", "1"))
        disabled = bool(meta.get("disabled", False))
        bg = str(meta.get("bg", ""))
        blob = f"{text} {cls}".lower()

        unavailable_keywords = [str(x).lower() for x in self.monitor.get("unavailable_keywords", [])]
        available_keywords = [str(x).lower() for x in self.monitor.get("available_keywords", [])]
        unavailable_class_keywords = [
            str(x).lower()
            for x in self.monitor.get("unavailable_class_keywords", ["disabled", "gray", "grey", "full", "ban"])
        ]
        available_class_keywords = [
            str(x).lower()
            for x in self.monitor.get("available_class_keywords", ["available", "green", "active"])
        ]

        if any(k in blob for k in unavailable_keywords) or any(k in blob for k in unavailable_class_keywords):
            return False
        if any(k in blob for k in available_keywords) or any(k in blob for k in available_class_keywords):
            return True
        if disabled or pointer == "none":
            return False
        try:
            if float(opacity) < 0.45:
                return False
        except ValueError:
            pass

        r, g, b = parse_rgb(bg)
        if g > r + 18 and g > b + 18:
            return True
        return False

    def _segment_name(self, period_index: int) -> str:
        segments = self.monitor.get(
            "period_segments",
            {"morning": [1, 2, 3, 4, 5], "afternoon": [6, 7, 8, 9], "evening": [10, 11, 12, 13]},
        )
        for name, period_list in segments.items():
            if isinstance(period_list, list) and period_index in [int(x) for x in period_list]:
                return name
        return "other"

    def _segment_count_from_text(self, text: str) -> Dict[str, int]:
        normalized = normalize_text(text)
        labels = {"上午": "morning", "下午": "afternoon", "晚上": "evening"}
        result: Dict[str, int] = {}
        for zh, key in labels.items():
            match = re.search(rf"{zh}\s*[：:]\s*([有无满]|可约|约满|\d+)", normalized)
            if not match:
                continue
            value = match.group(1)
            if value.isdigit():
                result[key] = int(value)
            elif value in ("有", "可约"):
                result[key] = 1
            else:
                result[key] = 0
        return result

    def _structured_availability(self) -> List[Dict[str, Any]]:
        driver = self._must_driver()
        venue_template = str(self.monitor.get("venue_xpath_template", "")).strip()
        period_template = str(
            self.monitor.get("period_xpath_template", self.booking.get("period_xpath_template", ""))
        ).strip()
        period_indexes = self.monitor.get("period_indexes", list(range(1, 14)))
        max_venues = int(self.monitor.get("max_venues", 6))
        analysis_wait_sec = float(self.monitor.get("analysis_wait_sec", 0.25))
        court_number = int(self.booking.get("court_number", 0))
        court_offset = int(self.booking.get("court_row_offset", 1))
        court_template = str(self.booking.get("court_xpath_template", "")).strip()

        if not venue_template or not period_template:
            return []

        result: List[Dict[str, Any]] = []
        for venue_index in range(1, max_venues + 1):
            venue_xpath = venue_template.format(venue_index=venue_index)
            venue_element = self._find_visible_element(venue_xpath)
            if venue_element is None:
                continue

            venue_name = normalize_text(venue_element.text) or f"venue-{venue_index}"
            text_segment_count = self._segment_count_from_text(venue_name)
            try:
                driver.execute_script(
                    "arguments[0].scrollIntoView({block:'center',inline:'nearest'});", venue_element
                )
                venue_element.click()
            except WebDriverException:
                driver.execute_script("arguments[0].click();", venue_element)
            time.sleep(analysis_wait_sec)

            if text_segment_count:
                result.append(
                    {
                        "venue_index": venue_index,
                        "venue_name": venue_name,
                        "available_total": sum(text_segment_count.values()),
                        "segment_count": text_segment_count,
                        "slot_details": [],
                    }
                )
                continue

            if court_number > 0 and court_template:
                try:
                    court_xpath = court_template.format(court_row_index=court_number + court_offset)
                    self._wait_clickable(court_xpath, 5).click()
                    time.sleep(analysis_wait_sec)
                except TimeoutException:
                    pass

            slot_details: List[Dict[str, Any]] = []
            for period in period_indexes:
                period_idx = int(period)
                slot_xpath = period_template.format(period_index=period_idx)
                slot_element = self._find_visible_element(slot_xpath)
                if slot_element is None:
                    continue
                meta = self._slot_meta(slot_element)
                available = self._is_slot_available(meta)
                slot_details.append(
                    {
                        "period_index": period_idx,
                        "segment": self._segment_name(period_idx),
                        "available": available,
                        "text": normalize_text(str(meta.get("text", ""))),
                    }
                )

            if not slot_details:
                continue

            segment_count: Dict[str, int] = {}
            for item in slot_details:
                if item["available"]:
                    segment_count[item["segment"]] = segment_count.get(item["segment"], 0) + 1

            result.append(
                {
                    "venue_index": venue_index,
                    "venue_name": venue_name,
                    "available_total": sum(segment_count.values()),
                    "segment_count": segment_count,
                    "slot_details": slot_details,
                }
            )
        return result

    def dump_xpath_candidates(self) -> None:
        keywords = list(self.monitor.get("available_keywords", [])) + list(
            self.monitor.get("unavailable_keywords", [])
        )
        candidates = self._discover_keyword_nodes(keywords)
        payload = {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "url": self._must_driver().current_url,
            "count": len(candidates),
            "candidates": candidates,
        }
        out_path = Path(str(self.monitor.get("xpath_dump_file", "xpath_candidates.json"))).resolve()
        out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        LOGGER.info("已导出 XPath 候选到 %s", out_path)

    def dump_html_snapshot(self) -> None:
        if not bool(self.monitor.get("dump_html_on_check", False)):
            return
        out_path = Path(str(self.monitor.get("html_snapshot_file", "page_snapshot.html"))).resolve()
        html = self._must_driver().page_source
        out_path.write_text(html, encoding="utf-8")
        LOGGER.info("已导出 HTML 快照到 %s", out_path)

    def check_availability(self) -> Dict[str, Any]:
        driver = self._must_driver()
        self.dump_html_snapshot()
        venues = self._structured_availability()
        if venues:
            self.structured_warned = False
            positives = []
            for venue in venues:
                morning = venue["segment_count"].get("morning", 0)
                afternoon = venue["segment_count"].get("afternoon", 0)
                evening = venue["segment_count"].get("evening", 0)
                positives.append(
                    f"{venue['venue_name']} | 上午:{'有' if morning > 0 else '无'}({morning}) 下午:{'有' if afternoon > 0 else '无'}({afternoon}) 晚上:{'有' if evening > 0 else '无'}({evening}) 总计:{venue['available_total']}"
                )
            has_available = any(v["available_total"] > 0 for v in venues)
            return {
                "has_available": has_available,
                "positives": positives,
                "negatives": [],
                "url": driver.current_url,
                "venues": venues,
            }
        if not self.structured_warned and str(self.monitor.get("venue_xpath_template", "")).strip():
            LOGGER.warning("结构化分析未命中场馆，请检查 monitor.venue_xpath_template 与 period_xpath_template")
            self.structured_warned = True

        available_keywords = list(self.monitor.get("available_keywords", []))
        unavailable_keywords = list(self.monitor.get("unavailable_keywords", []))

        snippets: List[str] = []
        for xpath in self.monitor.get("available_xpath_candidates", []):
            try:
                elements = driver.find_elements(By.XPATH, xpath)
                snippets.extend(normalize_text(el.text) for el in elements if el.text.strip())
            except InvalidSelectorException:
                LOGGER.warning("无效 XPath: %s", xpath)

        discovered = self._discover_keyword_nodes(available_keywords + unavailable_keywords)
        snippets.extend(normalize_text(item["text"]) for item in discovered if item.get("text"))

        positives: List[str] = []
        negatives: List[str] = []
        for text in snippets:
            has_pos = contains_any(text, available_keywords)
            has_neg = contains_any(text, unavailable_keywords)
            if has_pos and not has_neg:
                positives.append(text)
            elif has_neg:
                negatives.append(text)

        unique_pos = sorted(set(positives))
        unique_neg = sorted(set(negatives))
        has_available = len(unique_pos) > 0
        return {
            "has_available": has_available,
            "positives": unique_pos,
            "negatives": unique_neg,
            "url": driver.current_url,
            "venues": [],
        }

    def log_availability(self, result: Dict[str, Any]) -> None:
        venues = result.get("venues", [])
        if isinstance(venues, list) and venues:
            for line in result.get("positives", []):
                LOGGER.info("余量分析: %s", line)

    def alert(self, result: Dict[str, Any]) -> None:
        cooldown = int(self.monitor.get("alert_cooldown_sec", 300))
        now = time.time()
        if now - self.last_alert_ts < cooldown:
            return

        self.last_alert_ts = now
        msg = "检测到场馆有余量: " + " | ".join(result["positives"][:8])
        LOGGER.warning(msg)

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


def run_once(monitor: ZhihuiLuojiaMonitor, dump_xpath_only: bool) -> None:
    monitor.login()
    monitor.open_reserve_page()
    if bool(monitor.monitor.get("dump_xpath_on_start", True)) or dump_xpath_only:
        monitor.dump_xpath_candidates()
    if dump_xpath_only:
        return
    result = monitor.check_availability()
    monitor.log_availability(result)
    if result["has_available"]:
        monitor.alert(result)
        monitor.try_book()
    else:
        LOGGER.info("当前未发现余量")


def run_loop(monitor: ZhihuiLuojiaMonitor) -> None:
    monitor.login()
    monitor.open_reserve_page()
    if bool(monitor.monitor.get("dump_xpath_on_start", True)):
        monitor.dump_xpath_candidates()

    interval = int(monitor.monitor.get("interval_sec", 20))
    refresh_each_cycle = bool(monitor.monitor.get("refresh_each_cycle", True))
    navigate_each_cycle = bool(monitor.monitor.get("navigate_each_cycle", False))

    while True:
        result = monitor.check_availability()
        monitor.log_availability(result)
        if result["has_available"]:
            monitor.alert(result)
            monitor.try_book()
        else:
            LOGGER.info("当前未发现余量")
        time.sleep(interval)
        if navigate_each_cycle:
            monitor.open_reserve_page()
        elif refresh_each_cycle:
            monitor._must_driver().refresh()


def main() -> None:
    parser = argparse.ArgumentParser(description="智慧珞珈场馆余量监测与自动预约（PC 网页版）")
    parser.add_argument("--config", default="config.yaml", help="配置文件路径")
    parser.add_argument("--once", action="store_true", help="只检测一次")
    parser.add_argument("--dump-xpath-only", action="store_true", help="只导出 XPath 候选后退出")
    parser.add_argument("--log-level", default="INFO", help="日志级别")
    args = parser.parse_args()

    setup_logging(args.log_level)
    cfg = load_config(Path(args.config).resolve())
    monitor = ZhihuiLuojiaMonitor(cfg)
    monitor.start()
    try:
        if args.once or args.dump_xpath_only:
            run_once(monitor, dump_xpath_only=args.dump_xpath_only)
        else:
            run_loop(monitor)
    finally:
        monitor.close()


if __name__ == "__main__":
    main()

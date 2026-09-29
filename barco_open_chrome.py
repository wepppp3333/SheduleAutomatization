from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, time as datetime_time
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook
import math
import re
import json
import time
import sys
import traceback
import atexit
import os
from difflib import SequenceMatcher


BASE_DIR = Path(__file__).resolve().parent
ARTIFACTS_DIR = BASE_DIR / "automation_artifacts"
SCREENSHOTS_DIR = ARTIFACTS_DIR / "screenshots"
LOG_PATH = ARTIFACTS_DIR / "barco_automation.log"
SCHEDULE_JSON_PATH = ARTIFACTS_DIR / "schedule.json"

ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)


def find_excel_file():
    preferred_patterns = [
        "Рассписание*.xlsx",
        "Рассписание*.xlsm",
        "Рассписание*.xls",
        "Расписание*.xlsx",
        "Расписание*.xlsm",
        "Расписание*.xls",
    ]
    for pattern in preferred_patterns:
        matches = sorted(BASE_DIR.glob(pattern))
        if matches:
            return matches[0]
    raise FileNotFoundError(
        f"Excel файл с именем 'Рассписание' не найден в папке проекта: {BASE_DIR}"
    )


def normalize_excel_time(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None

    if isinstance(value, str):
        match = re.fullmatch(r"\s*(\d{1,2}):(\d{2})(?::\d{2})?\s*", value)
        if not match:
            return None
        hour, minute = map(int, match.groups())
    elif isinstance(value, datetime):
        hour, minute = value.hour, value.minute
    elif isinstance(value, datetime_time):
        hour, minute = value.hour, value.minute
    elif isinstance(value, (int, float)):
        total_minutes = round((float(value) % 1) * 24 * 60) % (24 * 60)
        hour, minute = divmod(total_minutes, 60)
    else:
        return None

    if not (0 <= hour <= 23 and 0 <= minute <= 59):
        return None
    return f"{hour:02d}:{minute:02d}"


def cell_has_red_font(cell):
    color = cell.font.color
    if color is None:
        return False
    if color.type == "rgb" and color.rgb:
        return color.rgb.upper()[-6:] == "FF0000"
    return color.type == "indexed" and color.indexed == 10


def _css_px_to_float(value):
    try:
        return float(str(value).replace("px", "").strip())
    except Exception:
        return 0.0


def click_top_slot(driver, day_view):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", day_view)
    driver.execute_script(
        """
const day = arguments[0];
const rect = day.getBoundingClientRect();
const x = rect.width * 0.6;
const y = 6;
const clientX = rect.left + x;
const clientY = rect.top + y;
const target = document.elementFromPoint(clientX, clientY) || day;
target.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, clientX, clientY}));
""",
        day_view,
    )


def click_time_slot(driver, day_view, time_str):
    hour, minute = [int(x) for x in time_str.split(":")]

    for _ in range(3):
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", day_view)
            result = driver.execute_script(
                """
const day = arguments[0];
const hour = arguments[1];
const minute = arguments[2];
const lines = day.querySelectorAll('.hourLine');
if (lines.length < 2) return {ok:false, reason:'hourLine<2'};
const top0 = parseFloat(getComputedStyle(lines[0]).top);
const top1 = parseFloat(getComputedStyle(lines[1]).top);
const step = (top1 > top0) ? (top1 - top0) : 80;
const y = top0 + (hour * step) + (minute / 60) * step + 2;
const rect = day.getBoundingClientRect();
const x = Math.min(rect.width - 2, Math.max(2, rect.width * 0.6));
const clampedY = Math.min(rect.height - 2, Math.max(2, y));
const clientX = rect.left + x;
const clientY = rect.top + clampedY;
const target = document.elementFromPoint(clientX, clientY) || day;
target.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, clientX, clientY}));
return {ok:true, clientX, clientY, x, y: clampedY};
""",
                day_view,
                hour,
                minute,
            )
            if not result or not result.get("ok"):
                reason = result.get("reason") if isinstance(result, dict) else "unknown"
                raise RuntimeError(f"JS click failed: {reason}")
            return result.get("x"), result.get("y")
        except StaleElementReferenceException:
            time.sleep(0.3)
            continue
        except Exception:
            time.sleep(0.3)
            continue

    # Last resort: ActionChains if JS failed
    try:
        return ActionChains(driver).move_to_element(day_view).click().perform()
    except Exception as e:
        raise RuntimeError(f"Click time slot failed after retries: {e}")


def _wait_popover(driver, timeout_sec=2):
    return WebDriverWait(driver, timeout_sec).until(
        EC.visibility_of_element_located((By.ID, "showPlaceHolderPopover"))
    )


def open_show_popover(driver, day_view, time_str):
    # Try multiple click strategies to open the popover
    for _ in range(3):
        try:
            _wait_popover(driver, timeout_sec=1.5)
            return True
        except Exception:
            pass

        try:
            click_top_slot(driver, day_view)
        except Exception:
            pass

        try:
            _wait_popover(driver, timeout_sec=1.5)
            return True
        except Exception:
            pass

        try:
            click_time_slot(driver, day_view, time_str)
        except Exception:
            pass

        try:
            _wait_popover(driver, timeout_sec=1.5)
            return True
        except Exception:
            pass

        try:
            placeholder = day_view.find_element(By.CLASS_NAME, "showPlaceHolder")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", placeholder)
            try:
                placeholder.click()
            except Exception:
                driver.execute_script("arguments[0].click();", placeholder)
        except Exception:
            pass

        time.sleep(0.3)

    return False


def hover_element(driver, element):
    try:
        ActionChains(driver).move_to_element(element).perform()
    except Exception:
        try:
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
        except Exception:
            pass


def scroll_timeline_to_top(driver):
    try:
        driver.execute_script(
            """
const area = document.querySelector('.timLineViewArea');
if (area) { area.scrollTop = 0; }
window.scrollTo(0, 0);
"""
        )
    except Exception:
        pass


def normalize_title(text):
    if text is None:
        return ""
    t = text.lower()
    t = re.sub(r"[^a-zа-я0-9\s]+", " ", t, flags=re.IGNORECASE)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def titles_match(expected, actual):
    e = normalize_title(expected)
    a = normalize_title(actual)
    if not e or not a:
        return False
    if e in a or a in e:
        return True
    # Require shared words when exact contains check fails.
    e_words = [w for w in e.split() if len(w) > 2]
    a_words = [w for w in a.split() if len(w) > 2]
    common = set(e_words) & set(a_words)
    if e_words and len(common) >= max(1, len(e_words) - 1):
        return True
    # Try dropping last letter in last word (Ушаков/Ушакова)
    e_parts = e.split()
    if e_parts:
        e_last = e_parts[-1]
        if len(e_last) > 3:
            e_parts[-1] = e_last[:-1]
            e2 = " ".join(e_parts)
            if e2 in a:
                return True
    return False


def title_similarity(expected, actual):
    e = normalize_title(expected)
    a = normalize_title(actual)
    if not e or not a:
        return 0.0
    if e in a or a in e:
        return 1.0

    e_words = [w for w in e.split() if len(w) > 2]
    a_words = [w for w in a.split() if len(w) > 2]
    overlap = 0.0
    if e_words:
        overlap = len(set(e_words) & set(a_words)) / len(set(e_words))

    seq_ratio = SequenceMatcher(None, e, a).ratio()
    return max(overlap, seq_ratio)


def wait_for_show_block(driver, index, title, timeout_sec=8):
    end_at = time.time() + timeout_sec
    while time.time() < end_at:
        try:
            day_views = driver.find_elements(By.CLASS_NAME, "dayView")
            if index >= len(day_views):
                time.sleep(0.3)
                continue
            day_view = day_views[index]
            show_blocks = day_view.find_elements(By.CLASS_NAME, "rowItem")
            for block in show_blocks:
                try:
                    title_div = block.find_element(By.CLASS_NAME, "title")
                    if titles_match(title, title_div.text):
                        return block
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(0.3)
    return None


def open_menu_show(driver, wait, target_block):
    for _ in range(3):
        try:
            hover_element(driver, target_block)
            try:
                target_block.click()
            except Exception:
                pass
            menu_show = wait.until(EC.presence_of_element_located((By.ID, "menuShow")))
            if not menu_show.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_show)
            menu_show = wait.until(EC.element_to_be_clickable((By.ID, "menuShow")))
            try:
                menu_show.click()
            except Exception:
                driver.execute_script("arguments[0].click();", menu_show)
            return
        except Exception:
            time.sleep(0.5)
            continue
    # JS fallback if still not found
    clicked = driver.execute_script(
        """
const el = document.getElementById('menuShow');
if (!el) return false;
el.click();
return true;
"""
    )
    if not clicked:
        raise RuntimeError("menuShow not found after retries")


def click_move_to(driver, wait, target_block):
    for _ in range(5):
        try:
            open_menu_show(driver, wait, target_block)
            time.sleep(0.3)
            move_candidates = driver.find_elements(By.ID, "moveTo")
            for candidate in move_candidates:
                try:
                    if not candidate.is_displayed():
                        continue
                    cls = (candidate.get_attribute("class") or "").lower()
                    if "disabled" in cls:
                        continue
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", candidate)
                    try:
                        candidate.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", candidate)
                    return True
                except Exception:
                    continue

            clicked = driver.execute_script(
                """
const nodes = Array.from(document.querySelectorAll('#moveTo'));
for (const n of nodes) {
  const style = window.getComputedStyle(n);
  if (style.display === 'none' || style.visibility === 'hidden') continue;
  if (n.classList.contains('disabled')) continue;
  n.click();
  return true;
}
return false;
"""
            )
            if clicked:
                return True
        except Exception:
            pass
        time.sleep(0.4)
    return False


def click_visible_id(driver, element_id, retries=4):
    for _ in range(retries):
        try:
            candidates = driver.find_elements(By.ID, element_id)
            for candidate in candidates:
                try:
                    if not candidate.is_displayed():
                        continue
                    cls = (candidate.get_attribute("class") or "").lower()
                    if "disabled" in cls:
                        continue
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", candidate)
                    try:
                        candidate.click()
                    except Exception:
                        driver.execute_script("arguments[0].click();", candidate)
                    return True
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(0.3)
    return False


def clear_blocking_modal_backdrop(driver):
    try:
        driver.execute_script(
            """
document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
"""
        )
    except Exception:
        pass


def close_datetime_modal(driver):
    try:
        # Prefer explicit close controls if available.
        close_btns = driver.find_elements(By.CSS_SELECTOR, "#dateTimeModal .close, #dateTimeModal [data-dismiss='modal']")
        for btn in close_btns:
            try:
                if btn.is_displayed():
                    btn.click()
                    return
            except Exception:
                continue
        # Fallback: ESC key and forced hide.
        driver.find_element(By.TAG_NAME, "body").send_keys("\uE00C")
        driver.execute_script(
            """
const modal = document.getElementById('dateTimeModal');
if (modal) {
  modal.classList.remove('in');
  modal.style.display = 'none';
}
"""
        )
    except Exception:
        pass
    clear_blocking_modal_backdrop(driver)


def round_minute_to_grid(minute):
    rounded = int(round(int(minute) / 3) * 3)
    return min(57, max(0, rounded))


def format_time_12h(hour, minute):
    suffix = "AM" if hour < 12 else "PM"
    display_hour = hour % 12 or 12
    return f"{display_hour:02d}:{minute:02d} {suffix}"


def get_visible_date_headers(driver):
    headers = driver.find_elements(By.CLASS_NAME, "dayHeader")
    result = []
    for index, header in enumerate(headers):
        try:
            text = header.find_element(By.CLASS_NAME, "date").text.strip()
            result.append((index, datetime.strptime(text, "%d/%m/%Y"), text))
        except (ValueError, StaleElementReferenceException):
            continue
    return result


def find_date_column(driver, wait, date_str, max_week_changes=16):
    target_date = datetime.strptime(date_str, "%d.%m.%Y")

    for _ in range(max_week_changes + 1):
        headers = get_visible_date_headers(driver)
        for index, header_date, _ in headers:
            if header_date.date() == target_date.date():
                return index

        if not headers:
            raise RuntimeError("На странице не найдены dayHeader")

        first_before = headers[0][2]
        navigation_class = "nextHeader" if target_date > headers[-1][1] else "prevHeader"
        navigation = wait.until(
            EC.element_to_be_clickable((By.CLASS_NAME, navigation_class))
        )
        navigation.click()
        wait.until(
            lambda d: get_visible_date_headers(d)
            and get_visible_date_headers(d)[0][2] != first_before
        )

    raise RuntimeError(f"Дата {date_str} не появилась после перелистывания недель")


def get_day_view(driver, index):
    day_views = driver.find_elements(By.CLASS_NAME, "dayView")
    if index >= len(day_views):
        raise RuntimeError(
            f"dayView с индексом {index} не найден; всего колонок: {len(day_views)}"
        )
    return day_views[index]


def row_title(row):
    try:
        return row.find_element(By.CLASS_NAME, "title").text.strip()
    except Exception:
        return ""


def row_start_time(row):
    try:
        return row.find_element(By.CLASS_NAME, "startTime").text.strip()
    except Exception:
        return ""


def normalize_barco_time(value):
    cleaned = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip().upper()
    try:
        return datetime.strptime(cleaned, "%I:%M %p").strftime("%I:%M %p")
    except ValueError:
        return cleaned


def find_show_row(driver, day_index, title, start_time=None):
    day_view = get_day_view(driver, day_index)
    for row in day_view.find_elements(By.CLASS_NAME, "rowItem"):
        if not titles_match(title, row_title(row)):
            continue
        if start_time is not None and normalize_barco_time(
            row_start_time(row)
        ) != normalize_barco_time(start_time):
            continue
        return row
    return None


def get_matching_show_rows(driver, day_index, title):
    day_view = get_day_view(driver, day_index)
    return [
        row
        for row in day_view.find_elements(By.CLASS_NAME, "rowItem")
        if titles_match(title, row_title(row))
    ]


def wait_for_new_show_row(driver, day_index, title, existing_row_ids, timeout=12):
    def find_new_row(current_driver):
        for row in get_matching_show_rows(current_driver, day_index, title):
            if row.id not in existing_row_ids:
                return row
        return False

    return WebDriverWait(driver, timeout).until(find_new_row)


def show_exists(driver, day_index, title, hour, minute):
    return find_show_row(
        driver,
        day_index,
        title,
        format_time_12h(hour, minute),
    ) is not None


def click_free_hour_line(driver, day_index):
    day_view = get_day_view(driver, day_index)
    hour_lines = day_view.find_elements(By.CLASS_NAME, "hourLine")
    if len(hour_lines) < 24:
        raise RuntimeError(f"Ожидалось 25 hourLine, найдено {len(hour_lines)}")

    free_index = driver.execute_script(
        """
const day = arguments[0];
const preferred = [5, 4, 6, 3, 7, 2, 1, 0];
const rows = Array.from(day.querySelectorAll('.rowItem')).map(row => ({
  top: parseFloat(getComputedStyle(row).top) || 0,
  bottom: (parseFloat(getComputedStyle(row).top) || 0) + row.getBoundingClientRect().height
}));
const past = day.querySelector('.pastTime');
const pastHeight = past && getComputedStyle(past).display !== 'none'
  ? past.getBoundingClientRect().height : 0;
const lines = day.querySelectorAll('.hourLine');
for (const index of preferred) {
  const top = parseFloat(getComputedStyle(lines[index]).top) || 0;
  const occupied = rows.some(row => top >= row.top - 2 && top <= row.bottom + 2);
  if (top > pastHeight + 2 && !occupied) return index;
}
return -1;
""",
        day_view,
    )
    if free_index < 0:
        raise RuntimeError("Не найден свободный временный час в начале дня")

    target_line = hour_lines[free_index]
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});", target_line
    )
    try:
        target_line.click()
    except Exception:
        driver.execute_script("arguments[0].click();", target_line)

    return free_index, format_time_12h(free_index, 0)


def choose_show_in_popover(driver, wait, title):
    popover = wait.until(
        EC.visibility_of_element_located((By.ID, "showPlaceHolderPopover"))
    )
    caret = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "#showPlaceHolderPopover .caretBtn"))
    )
    caret.click()

    show_list = wait.until(
        EC.visibility_of_element_located((By.ID, "listOfShows"))
    )
    links = show_list.find_elements(By.TAG_NAME, "a")
    candidates = [(title_similarity(title, link.text), link) for link in links]
    if not candidates:
        raise RuntimeError("Список фильмов пуст")

    score, target = max(candidates, key=lambda item: item[0])
    if score < 0.55:
        available = [link.text.strip() for link in links if link.text.strip()]
        raise RuntimeError(
            f"Фильм '{title}' не найден. Доступные фильмы: {available}"
        )

    selected_title = target.text.strip()
    target.click()
    wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#showPlaceHolderPopover .ok")
        )
    ).click()
    print(f"Выбран фильм '{selected_title}' (совпадение {score:.2f})")


def find_visible_element(driver, by, value):
    for element in driver.find_elements(by, value):
        try:
            if element.is_displayed():
                return element
        except StaleElementReferenceException:
            continue
    return False


def open_move_dialog(driver, wait, day_index, row):
    title = row_title(row)
    start_time = row_start_time(row)
    last_error = None

    for attempt in range(1, 4):
        try:
            current_row = find_show_row(driver, day_index, title, start_time)
            if current_row is None:
                raise RuntimeError(
                    f"Блок '{title}' в {start_time} исчез после перерисовки"
                )

            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center'});", current_row
            )
            ActionChains(driver).move_to_element(current_row).perform()
            move_button = current_row.find_element(By.CLASS_NAME, "moveRowBtn")
            try:
                move_button.click()
            except Exception:
                driver.execute_script("arguments[0].click();", move_button)

            menu_show = WebDriverWait(driver, 4).until(
                lambda d: find_visible_element(d, By.ID, "menuShow")
            )
            try:
                menu_show.click()
            except Exception:
                driver.execute_script("arguments[0].click();", menu_show)

            move_to = WebDriverWait(driver, 4).until(
                lambda d: find_visible_element(d, By.ID, "moveTo")
            )
            try:
                move_to.click()
            except Exception:
                driver.execute_script("arguments[0].click();", move_to)

            return wait.until(
                EC.visibility_of_element_located((By.ID, "dateTimeModal"))
            )
        except Exception as error:
            last_error = error
            print(
                f"Повтор открытия меню переноса {attempt}/3 "
                f"для '{title}' в {start_time}"
            )
            time.sleep(0.7)

    raise RuntimeError(
        f"Не удалось открыть меню переноса для '{title}' в {start_time}"
    ) from last_error


def click_exact_text(elements, expected, description):
    for element in elements:
        if element.text.strip() == expected:
            element.click()
            return
    raise RuntimeError(f"{description} '{expected}' не найден")


def set_modal_datetime(driver, wait, modal, date_str, hour, minute):
    target_day = str(int(date_str.split(".")[0]))
    day_cells = modal.find_elements(
        By.CSS_SELECTOR,
        ".datepicker-days td.day:not(.old):not(.new):not(.notSelectable)",
    )
    click_exact_text(day_cells, target_day, "День календаря")

    modal.find_element(By.CLASS_NAME, "timepicker-hour").click()
    hours = wait.until(
        EC.visibility_of_all_elements_located(
            (By.CSS_SELECTOR, "#dateTimeModal .timepicker-hours .hour")
        )
    )
    click_exact_text(hours, f"{hour:02d}", "Час")

    modal.find_element(By.CLASS_NAME, "timepicker-minute").click()
    minutes = wait.until(
        EC.visibility_of_all_elements_located(
            (By.CSS_SELECTOR, "#dateTimeModal .timepicker-minutes .minute")
        )
    )
    minute_grid_value = (minute // 3) * 3
    click_exact_text(minutes, f"{minute_grid_value:02d}", "Минута")
    for _ in range(minute - minute_grid_value):
        increment = wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "#dateTimeModal [data-action='incrementMinutes']")
            )
        )
        increment.click()

    modal.find_element(By.CLASS_NAME, "timepicker-second").click()
    seconds = wait.until(
        EC.visibility_of_all_elements_located(
            (By.CSS_SELECTOR, "#dateTimeModal .timepicker-seconds .second")
        )
    )
    click_exact_text(seconds, "00", "Секунда")

    confirm = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#dateTimeModal #confirmDateTimeBtn")
        )
    )
    confirm.click()
    wait.until(EC.invisibility_of_element_located((By.ID, "dateTimeModal")))


def schedule_show(driver, wait, day_index, show):
    title = show["title"]
    hour, minute = [int(part) for part in show["time"].split(":")]
    expected_time = format_time_12h(hour, minute)

    if show_exists(driver, day_index, title, hour, minute):
        print(f"Сеанс уже существует: '{title}' в {expected_time}. Пропускаем.")
        return

    existing_row_ids = {
        row.id for row in get_matching_show_rows(driver, day_index, title)
    }

    click_free_hour_line(driver, day_index)
    choose_show_in_popover(driver, wait, title)

    row = wait_for_new_show_row(
        driver,
        day_index,
        title,
        existing_row_ids,
    )
    print(
        f"Создан временный блок '{row_title(row)}' в {row_start_time(row) or 'неизвестное время'}"
    )
    modal = open_move_dialog(driver, wait, day_index, row)
    set_modal_datetime(driver, wait, modal, show["date"], hour, minute)

    wait.until(lambda d: show_exists(d, day_index, title, hour, minute))
    print(f"Фильм '{title}' установлен на {show['date']} {expected_time}")


class Tee:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for stream in self.streams:
            stream.write(data)
            stream.flush()

    def flush(self):
        for stream in self.streams:
            stream.flush()


log_file = LOG_PATH.open("a", encoding="utf-8")
sys.stdout = Tee(sys.__stdout__, log_file)
sys.stderr = Tee(sys.__stderr__, log_file)
print(f"\n===== Start run: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====")


def _close_log_file():
    if not log_file.closed:
        log_file.close()


atexit.register(_close_log_file)


def log_exception(context):
    print(f"❗ {context}:")
    print(traceback.format_exc())


def _global_excepthook(exc_type, exc_value, exc_tb):
    print("❗ Необработанная ошибка:")
    print("".join(traceback.format_exception(exc_type, exc_value, exc_tb)))


def search_elements(elementsClass):
    elements = driver.find_element(elementsClass)
    return elements

sys.excepthook = _global_excepthook


# Загрузка exel
# Удаление старого schedule.json если он существует
if SCHEDULE_JSON_PATH.exists():
    SCHEDULE_JSON_PATH.unlink()
    print("🗑️ Старый файл schedule.json удалён")
else:
    print("Старый json не нашли")

excel_path = find_excel_file()
print(f"Excel для загрузки: {excel_path}")

workbook = load_workbook(excel_path, data_only=True)
worksheet = workbook.active

schedule = []
current_date = None

for row in worksheet.iter_rows():
   first_cell = row[0]
   second_cell = row[1]
   first_col = first_cell.value
   second_col = second_cell.value

   if isinstance(first_col,str):
      try:
         parsed_date = datetime.strptime(first_col.strip(), "%d.%m.%Y")
         current_date = parsed_date.strftime("%d.%m.%Y")
      except ValueError:
         pass

   elif isinstance(first_col,datetime):
      current_date = first_col.strftime("%d.%m.%Y")

   show_time = normalize_excel_time(first_col)
   if show_time and second_col is not None and current_date:
        if cell_has_red_font(first_cell) or cell_has_red_font(second_cell):
            print(
                f"Пропущена красная строка: {show_time} — {str(second_col).strip()}"
            )
            continue
        raw_title = str(second_col).strip()
        schedule.append({
            "date": current_date,
            "time": show_time,
            "title": re.split(r"\s+\d+D|,\s*\d+\+?", raw_title)[0]
        })      
   

json_path = SCHEDULE_JSON_PATH
with open(json_path, "w", encoding="utf-8") as f:
   json.dump(schedule, f, ensure_ascii=False, indent=2)

print(f"✅ Готово! Сохранено {len(schedule)} фильмов в файл {json_path}")

options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")

driver = None
env_driver_path = os.getenv("CHROMEDRIVER_PATH")
fallback_driver_paths = [
    Path(r"C:\Users\Ust-Kinel\Desktop\autometization\chromedriver-win64\chromedriver.exe"),
    Path("/opt/homebrew/bin/chromedriver"),
]

if env_driver_path:
    fallback_driver_paths.insert(0, Path(env_driver_path))

try:
    # Selenium Manager подбирает совместимый драйвер под текущий Chrome.
    print("Пробуем запуск Chrome через Selenium Manager (автоподбор драйвера)...")
    driver = webdriver.Chrome(options=options)
    print("✅ Chrome запущен через Selenium Manager.")
except Exception as e:
    print(f"⚠️ Selenium Manager не сработал: {e}")
    for candidate in fallback_driver_paths:
        if not candidate.exists():
            continue
        try:
            print(f"Пробуем локальный ChromeDriver: {candidate}")
            driver = webdriver.Chrome(service=Service(str(candidate)), options=options)
            print(f"✅ Chrome запущен с локальным ChromeDriver: {candidate}")
            break
        except Exception as fallback_error:
            print(f"⚠️ Не удалось запустить через {candidate}: {fallback_error}")

if driver is None:
    raise RuntimeError(
        "Не удалось запустить Chrome. Обновите ChromeDriver до версии вашего Chrome "
        "или задайте корректный путь в переменной CHROMEDRIVER_PATH."
    )

driver.get("https://192.168.100.2:43744")

wait = WebDriverWait(driver, 10)

try:
    # Ждем и нажимаем кнопку "Подробно" (details-button)
    details_button = wait.until(EC.element_to_be_clickable((By.ID, "details-button")))
    details_button.click()

    # Ждем и нажимаем ссылку "Продолжить" (proceed-link)
    proceed_link = wait.until(EC.element_to_be_clickable((By.ID, "proceed-link")))
    proceed_link.click()
except Exception as e:
    # Показать ошибку в alert в браузере
    error_message = str(e).replace('"', '\\"')
    driver.execute_script(f'alert("Ошибка: {error_message}");')
    time.sleep(10)  # чтобы успеть увидеть alert


username_input = wait.until(EC.presence_of_element_located((By.ID, "loginUsername")))
username_input.send_keys("admin")
password_input = wait.until(EC.presence_of_element_located((By.ID, "loginPass")))
password_input.send_keys("Admin1234")

login_button = wait.until(EC.element_to_be_clickable((By.ID, "loginSubmit")))
login_button.click()

time.sleep(10)
driver.get("https://192.168.100.2:43744/#sms/scheduler")

date_time = "На 10 секунд"
print("Встал на ожидание", date_time)
time.sleep(10)
try: 
  lock_app = wait.until(EC.presence_of_element_located((By.ID, "lockApp")))
  if "lockAppRed" in lock_app.get_attribute("class"):
     lock_app.click()
     print("Кнопка с lockAppRed найдена и нажата.")
  else: 
     print("Кнопка есть но класс lockAppRed отсутсвует - не нажимаем")
except Exception as e: 
   print(f"Ошибка при проверке lockApp: {e}")



# Новый код с циклом
# Загружаем расписание
with open(SCHEDULE_JSON_PATH, "r", encoding="utf-8") as f:
    schedule_data = json.load(f)

# Группируем по датам
grouped_schedule = defaultdict(list)
for item in schedule_data:
    grouped_schedule[item["date"]].append(item)

automation_failed = False
try:
    for date, shows in grouped_schedule.items():
        print(f"\nОбрабатываем дату: {date}")
        day_index = find_date_column(driver, wait, date)
        print(f"Дата {date} найдена, индекс колонки: {day_index}")

        for show in shows:
            print(f"Добавляем фильм: {show['title']} в {show['time']}")
            try:
                schedule_show(driver, wait, day_index, show)
            except Exception:
                automation_failed = True
                log_exception(
                    f"Ошибка добавления '{show['title']}' "
                    f"на {show['date']} {show['time']}"
                )
                screenshot_name = re.sub(
                    r'[\\/:*?"<>|]+',
                    "_",
                    f"{show['date']}_{show['time']}_{show['title']}",
                )
                try:
                    driver.save_screenshot(
                        str(SCREENSHOTS_DIR / f"error_{screenshot_name}.png")
                    )
                except Exception:
                    pass
                close_datetime_modal(driver)
                break

        if automation_failed:
            break
finally:
    if automation_failed:
        print("Автоматизация остановлена после первой ошибки, чтобы не создавать неверные сеансы.")
    else:
        print("Расписание обработано без ошибок.")
    time.sleep(3)
    driver.quit()

sys.exit(1 if automation_failed else 0)


# Legacy flow retained temporarily for reference; it is unreachable after sys.exit above.
day_headers = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayHeader")))

for date, shows in grouped_schedule.items():
    print(f"\n📅 Обрабатываем дату: {date}")

    day = str(int(date.split(".")[0]))
    # time_hour = 
    # Ищем нужный dayHeader по дате
    found_index = None
    
    for i in range(len(day_headers)):
        try:
            header = day_headers[i]
            header_date_text = header.find_element(By.CLASS_NAME, "date").text.strip()
            if header_date_text.replace("/", ".") == date:
                found_index = i
                header.click()
                print(f"✅ Найдена дата {date} в расписании, индекс: {i}")
                break
        except StaleElementReferenceException:
            day_headers = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayHeader")))
            continue

    if found_index is None:
        print(f"⚠️ Дата {date} не найдена на странице. Пропускаем.")
        # driver.find_element(By.CLASS_NAME,"nextHeader").click()
        # day_headers = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayHeader")))
        continue
    
    # //*[@id="schedulerTimeViewInner"]/div[2]/div[4]/div[7]/div[3]
   #  day_view = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))[found_index]

    # scroll_timeline_to_top(driver)

    for show in shows: 
        movie_name = show["title"].strip().lower()
        hour_time = show["time"].split(":")[0]
        minuts_time = show["time"].split(":")[1]
        # day = str(int(date.split(".")[0]))
        print(f"🎬 Добавляем фильм: {show['title']} в {show['time']}")
        print(f"found_index{found_index}")
        time.sleep(2)
        day_view = driver.find_elements(By.CLASS_NAME,"dayView")[found_index]
        hour_lines = day_view.find_elements(By.CLASS_NAME,"hourLine")
        hour_lines[5].click()

        time.sleep(2)
        caret_btn = driver.find_element(By.CLASS_NAME,"caretBtn").click()
        print(f"Клик по кнопке произошел")

        time.sleep(2)
        list_Of_Shows = driver.find_element(By.ID,"listOfShows")
        links = list_Of_Shows.find_elements(By.TAG_NAME, "a")
        target = None
        for a in links:
            text_value = a.text.strip().lower()
            if movie_name in text_value:
                target = a
                
                print(f"🎬 Найден фильм в списке {text_value} наименование в exel {movie_name}")
                break
            print(f"🎬 Наименования в списке выбора фильмов {text_value}")
        
        # Нашли фильм в списке выбрали его 
        time.sleep(2)
        target.click()
        popover_title = driver.find_element(By.ID,"showPlaceHolderPopover")
        ok_btn = popover_title.find_element(By.CLASS_NAME,"ok").click()

        # Ищем фильм для перемещения
        time.sleep(2)
        row_items = day_view.find_elements(By.CLASS_NAME,"rowItem")

        row_items_target = None

        for el in row_items:
            
            title_items = el.find_element(By.CLASS_NAME,"title")
            value = title_items.text.strip().lower()

            if movie_name in value:
                row_items_target = el
                break
        
        row_items_target.click()

        time.sleep(5)
        menu_Show = driver.find_element(By.ID,"menuShow").click()
        time.sleep(7)
        move_to = driver.find_element(By.ID,"moveTo").click()


        # Работа с перемещением с календарем
        print(f"Нужный день {day}")
        time.sleep(2)
        table_condensed = driver.find_element(By.CLASS_NAME,"datepicker-days")
        time.sleep(7)
        day_shedule = table_condensed.find_elements(By.CLASS_NAME,"day")

        for dayShedule in day_shedule:
            print(f"Зашел в выбор дня в рассписании")   
            print(f"cell:", dayShedule.text.strip(), dayShedule.get_attribute("class")) 
            cls = dayShedule.get_attribute("class")
            txt = dayShedule.text.strip()

            if txt != day:
                continue
            if "notSelectable" in cls:
                continue

            print(f"Найденный день в календаре {el}")   
            dayShedule.click()
            break

        time.sleep(2)
        showHours = driver.find_element(By.CLASS_NAME,"timepicker-hour").click()
        timepicker = driver.find_element(By.CLASS_NAME,"timepicker")
        hour_arr = timepicker.find_elements(By.CLASS_NAME,"hour")

        for hour in hour_arr:
            value_hour = hour.text.strip()

            if value_hour != hour_time:
                continue

            hour.click()
            break

        # Минуты в этом пикере идут с шагом 3, поэтому округляем к ближайшему значению.
        target_minute_int = int(minuts_time)
        rounded_minute = int(round(target_minute_int / 3) * 3)
        rounded_minute = min(57, max(0, rounded_minute))
        rounded_minute_str = f"{rounded_minute:02d}"
        print(f"Минуты из Excel: {minuts_time}, ставим: {rounded_minute_str}")

        time.sleep(2)
        driver.find_element(By.CLASS_NAME, "timepicker-minute").click()
        minute_cells = driver.find_elements(By.CLASS_NAME, "minute")

        minute_selected = False
        for minute_cell in minute_cells:
            if minute_cell.text.strip() == rounded_minute_str:
                minute_cell.click()
                minute_selected = True
                break

        if not minute_selected:
            print(f"Не нашли минуту {rounded_minute_str} в списке, пробуем через increment/decrement")
            for _ in range(25):
                current_min = driver.find_element(By.CLASS_NAME, "timepicker-minute").text.strip()
                if current_min == rounded_minute_str:
                    minute_selected = True
                    break
                if int(current_min) < rounded_minute:
                    driver.find_element(By.CSS_SELECTOR, "[data-action='incrementMinutes']").click()
                else:
                    driver.find_element(By.CSS_SELECTOR, "[data-action='decrementMinutes']").click()
                time.sleep(0.1)


        # Сохраняем рассписание
        # dateTimeModal = driver.find_element(By.ID,"dateTimeModal")
        time.sleep(2)
        driver.find_element(By.ID,"confirmDateTimeBtn").click()
        print(f" Фильм добавлен {movie_name} время {hour_time} минуты {minuts_time}")
        print(f" Ушел на паузу 20 секунд")
        time.sleep(5)
        # print(f"Длинна",len(driver.find_elements(By.CLASS_NAME,"dayView")))


    # for show in shows:
    #     print(f"🎬 Добавляем фильм: {show['title']} в {show['time']}")

    #     try:
    #         clear_blocking_modal_backdrop(driver)
    #         # Обновляем day_view и кликаем по таймлайну в нужное время
    #         day_views = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))
    #         day_view = day_views[found_index]
    #         scroll_timeline_to_top(driver)
    #         ok = open_show_popover(driver, day_view, show["time"])
    #         if not ok:
    #             raise RuntimeError("Поповер не открылся после клика по таймлайну")
    #     except Exception as e:
    #         print(f"❗ Ошибка при клике на таймлайн: {e}")
    #         try:
    #             screenshot_name = re.sub(r'[\\/:*?"<>|]+', "_", f"{show['date']}_{show['time']}_{show['title']}")
    #             driver.save_screenshot(str(SCREENSHOTS_DIR / f"error_timeline_{screenshot_name}.png"))
    #         except Exception:
    #             pass
    #         continue

    #     # Выбор фильма из выпадающего списка
    #     try:
    #         print(f"❗ Выбираем фильм из выпадающего списка")
    #         caret_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "caretBtn")))
    #         try:
    #             caret_btn.click()
    #         except Exception:
    #             driver.execute_script("arguments[0].click();", caret_btn)
    #         show_list = wait.until(EC.presence_of_element_located((By.ID, "listOfShows")))
    #         show_items = show_list.find_elements(By.TAG_NAME, "li")

    #         best_item = None
    #         best_score = 0.0
    #         for item in show_items:
    #             score = title_similarity(show["title"], item.text)
    #             if score > best_score:
    #                 best_score = score
    #                 best_item = item

    #         if best_item is None or best_score < 0.55:
    #             print(f"❗ Фильм '{show['title']}' не найден в списке")
    #             available_titles = [normalize_title(i.text) for i in show_items if i.text.strip()]
    #             print(f"Доступные в списке (normalized): {available_titles}")
    #             continue
    #         best_link = None
    #         try:
    #             best_link = best_item.find_element(By.TAG_NAME, "a")
    #         except Exception:
    #             best_link = best_item
    #         try:
    #             best_link.click()
    #         except Exception:
    #             driver.execute_script("arguments[0].click();", best_link)
    #         matched_label = (best_link.text or best_item.text or "").strip()
    #         print(f"Совпадение фильма: '{show['title']}' -> '{matched_label}' (score={best_score:.2f})")

    #         ok_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".popover-inner .ok.btn")))
    #         try:
    #             ok_button.click()
    #         except Exception:
    #             driver.execute_script("arguments[0].click();", ok_button)
    #     except Exception as e:
    #         print(f"❗ Ошибка при выборе фильма: {e}")
    #         continue


    #     try:
    #            target_block = wait_for_show_block(driver, found_index, show["title"], timeout_sec=8)
    #            if not target_block:
    #               print(f"❗ Блок с фильмом '{show['title']}' не найден.")
    #               continue
    #            print(f"❗ Блок с фильмом '{show['title']}' найден.")
    #     except Exception as e:
    #            print(f"❗ Ошибка при поиске блока с фильмом: {e}")
    #            continue

    #     try:
    #            time.sleep(10)
    #            hover_element(driver, target_block)
    #            move_btn = target_block.find_element(By.CLASS_NAME, "moveRowBtn")
    #            driver.execute_script("arguments[0].scrollIntoView(true);", move_btn)
    #            try:
    #               wait.until(EC.element_to_be_clickable(move_btn)).click()
    #            except Exception:
    #               driver.execute_script("arguments[0].click();", move_btn)
    #            print("✅ Клик по moveRowBtn прошёл")
    #     except Exception as e:
    #            print(f"❗ Ошибка при клике по moveRowBtn: {e}")
    #            continue

    #     time.sleep(10)

    #     try:
    #            open_menu_show(driver, wait, target_block)
    #            print("✅ Клик по menuShow прошёл")
    #     except Exception as e:
    #            print(f"❗ Ошибка при работе с menuShow: {e}")
    #            screenshot_name = re.sub(r'[\\/:*?"<>|]+', "_", show["title"])
    #            driver.save_screenshot(str(SCREENSHOTS_DIR / f"error_menuShow_{screenshot_name}.png"))
    #            print("Встал на ожидание на 10 секунд для проверки")
    #            time.sleep(10)
    #            continue

    #     time.sleep(1)

    #     try:
    #            clicked = click_move_to(driver, wait, target_block)
    #            if not clicked:
    #                raise RuntimeError("moveTo not clickable after retries")
    #            print("✅ Клик по moveTo прошёл")
    #     except Exception as e:
    #            print(f"❗ Ошибка при клике по moveTo: {e}")
    #            continue


    #     # Календарь
    #     try: 
    #         time.sleep(5)
    #         wait.until(EC.presence_of_element_located((By.ID, "dateTimeModal")))
    #         day_cells = driver.find_elements(By.CLASS_NAME, "day")
    #         target_day = date.split(".")[0]
    #         if target_day.startswith("0"):
    #             target_day = target_day[1:]
    #         print(target_day + " ДЕНЬ")
    #         print("ДЕНЬ")

    #         for cell in day_cells:
    #             if cell.text.strip() == target_day and "notSelectable" not in cell.get_attribute("class"):
    #                 cell.click()
    #                 break
    #     except Exception as e:
    #         print(f"❗ Ошибка при выборе дня в календаре: {e}")
    #         continue

    #     # Время
    #     try:
    #         time.sleep(5)
    #         hour_str, minute_str = show["time"].split(":")
    #         # Час
    #         wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "timepicker-hour"))).click()
    #         hour_set = False
    #         for cell in driver.find_elements(By.CLASS_NAME, "hour"):
    #             if cell.text.strip() == hour_str:
    #                 cell.click()
    #                 hour_set = True
    #                 break
    #         if not hour_set:
    #             raise RuntimeError(f"Час {hour_str} не найден в timepicker")

    #         # Минуты
    #         wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "timepicker-minute"))).click()
    #         minute_set = False
    #         for cell in driver.find_elements(By.CLASS_NAME, "minute"):
    #             if cell.text.strip() == minute_str:
    #                 cell.click()
    #                 minute_set = True
    #                 break
    #         if not minute_set:
    #             raise RuntimeError(f"Минута {minute_str} не найдена в timepicker")
    #     except Exception as e:
    #         print(f"❗ Ошибка при установке времени: {e}")
    #         close_datetime_modal(driver)
    #         continue
    #         # Код ИИ
    #     # Подтверждение
    #     try:
    #         clicked = click_visible_id(driver, "confirmDateTimeBtn", retries=5)
    #         if not clicked:
    #             raise RuntimeError("confirmDateTimeBtn not clickable")
    #         try:
    #             WebDriverWait(driver, 5).until(
    #                 EC.invisibility_of_element_located((By.ID, "dateTimeModal"))
    #             )
    #         except Exception:
    #             # If modal still visible, try one more click.
    #             if not click_visible_id(driver, "confirmDateTimeBtn", retries=2):
    #                 raise RuntimeError("confirmDateTimeBtn clicked but modal did not close")
    #         print(f"✅ Фильм '{show['title']}' добавлен в расписание.")
    #     except Exception as e:
    #         print(f"❗ Ошибка при подтверждении времени: {e}")
    #         close_datetime_modal(driver)
    #         continue
    #     time.sleep(10)
    #     print(f"✅ Встал на паузу на 10 секунд")
    #     scroll_timeline_to_top(driver)
    #     time.sleep(10)

   


time.sleep(3)
driver.quit()

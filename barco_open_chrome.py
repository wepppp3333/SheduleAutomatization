from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from pathlib import Path
from collections import defaultdict
import pandas as pd
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

df = pd.read_excel(excel_path,header=None)

schedule = []
current_date = None

for i in range(len(df)):
   first_col = df.iloc[i,0]
   second_col = df.iloc[i,1]

   if isinstance(first_col,str):
      try:
         parsed_date = datetime.strptime(first_col.strip(), "%d.%m.%Y")
         current_date = parsed_date.strftime("%d.%m.%Y")
      except ValueError:
         pass

   elif isinstance(first_col,datetime):
      current_date = first_col.strftime("%d.%m.%Y")


   if isinstance(first_col,str) and ":" in first_col and pd.notna(second_col) and current_date:
        schedule.append({
            "date": current_date,
            "time": first_col.strip(),
            "title": re.split(r"\s+\d+D|,\s*\d+\+?", second_col.strip())[0]
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



time.sleep(15)
# Новый код с циклом
# Загружаем расписание
with open(SCHEDULE_JSON_PATH, "r", encoding="utf-8") as f:
    schedule_data = json.load(f)

# Группируем по датам
grouped_schedule = defaultdict(list)
for item in schedule_data:
    grouped_schedule[item["date"]].append(item)

day_headers = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayHeader")))

for date, shows in grouped_schedule.items():
    print(f"\n📅 Обрабатываем дату: {date}")

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
    
    time.sleep(10)
   #  day_view = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))[found_index]

    # scroll_timeline_to_top(driver)

    for show in shows: 
        print(f"🎬 Добавляем фильм: {show['title']} в {show['time']}")
        print(f"found_index{found_index}")
        day_view = driver.find_elements(By.CLASS_NAME,"dayView")
        print(f"day_view{day_view}")

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

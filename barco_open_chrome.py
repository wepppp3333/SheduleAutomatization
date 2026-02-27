from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
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


def click_time_slot(driver, day_view, time_str):
    hour_lines = day_view.find_elements(By.CLASS_NAME, "hourLine")
    if len(hour_lines) < 2:
        raise RuntimeError("hourLine недостаточно для расчета позиции клика")

    top0 = _css_px_to_float(hour_lines[0].value_of_css_property("top"))
    top1 = _css_px_to_float(hour_lines[1].value_of_css_property("top"))
    step = top1 - top0 if top1 > top0 else 80.0

    hour, minute = [int(x) for x in time_str.split(":")]
    y = top0 + (hour * step) + (minute / 60.0) * step + 2

    height = day_view.size.get("height", 0)
    width = day_view.size.get("width", 0)
    if height:
        y = max(2, min(y, height - 2))
    x = 60
    if width:
        x = max(2, min(width * 0.6, width - 2))

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", day_view)

    try:
        ActionChains(driver).move_to_element_with_offset(day_view, x, y).click().perform()
    except Exception:
        driver.execute_script(
            """
const el = arguments[0];
const x = arguments[1];
const y = arguments[2];
const rect = el.getBoundingClientRect();
const clientX = rect.left + x;
const clientY = rect.top + y;
const target = document.elementFromPoint(clientX, clientY);
if (target) {
  target.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, clientX, clientY}));
}
""",
            day_view,
            x,
            y,
        )

    return x, y


def open_show_popover(driver, wait, day_view):
    try:
        return wait.until(EC.visibility_of_element_located((By.ID, "showPlaceHolderPopover")))
    except Exception:
        try:
            placeholder = day_view.find_element(By.CLASS_NAME, "showPlaceHolder")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", placeholder)
            try:
                placeholder.click()
            except Exception:
                driver.execute_script("arguments[0].click();", placeholder)
        except Exception:
            pass
    return wait.until(EC.visibility_of_element_located((By.ID, "showPlaceHolderPopover")))


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

# Находим все dayHeader
day_headers = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayHeader")))
day_views = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))

for date, shows in grouped_schedule.items():
    print(f"\n📅 Обрабатываем дату: {date}")

    # Ищем нужный dayHeader по дате
    found_index = None
    for i, header in enumerate(day_headers):
        header_date_text = header.find_element(By.CLASS_NAME, "date").text.strip()
        if header_date_text.replace("/", ".") == date:
            found_index = i
            header.click()
            print(f"✅ Найдена дата {date} в расписании, индекс: {i}")
            break

    if found_index is None:
        print(f"⚠️ Дата {date} не найдена на странице. Пропускаем.")
        continue
    
    time.sleep(10)
   #  day_view = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))[found_index]


    for show in shows:
        print(f"🎬 Добавляем фильм: {show['title']} в {show['time']}")

        try:
            # Обновляем day_view и кликаем по таймлайну в нужное время
            day_views = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))
            day_view = day_views[found_index]
            click_time_slot(driver, day_view, show["time"])
            open_show_popover(driver, wait, day_view)
        except Exception as e:
            print(f"❗ Ошибка при клике на таймлайн: {e}")
            continue

        # Выбор фильма из выпадающего списка
        try:
            print(f"❗ Выбираем фильм из выпадающего списка")
            caret_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "caretBtn")))
            try:
                caret_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", caret_btn)
            show_list = wait.until(EC.presence_of_element_located((By.ID, "listOfShows")))
            show_items = show_list.find_elements(By.TAG_NAME, "li")

            found = False
            for item in show_items:
                if show["title"] in item.text:
                    item.click()
                    found = True
                    break
            if not found:
                print(f"❗ Фильм '{show['title']}' не найден в списке")
                continue

            ok_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".popover-inner .ok.btn")))
            try:
                ok_button.click()
            except Exception:
                driver.execute_script("arguments[0].click();", ok_button)
        except Exception as e:
            print(f"❗ Ошибка при выборе фильма: {e}")
            continue

        # Ищем добавленный блок
      #   try:
      #       show_blocks = day_view.find_elements(By.CLASS_NAME, "rowItem")
      #       target_block = None
      #       for block in show_blocks:
      #           try:
      #               title_div = block.find_element(By.CLASS_NAME, "title")
      #               if show["title"] in title_div.text:
      #                   target_block = block
      #                   break
      #           except:
      #               continue

      #       if not target_block:
      #           print(f"❗ Блок с фильмом '{show['title']}' не найден.")
      #           continue
      #       print(f"❗ Блок с фильмом '{show['title']}' найден.")
      #       time.sleep(10)   
      #       move_btn = target_block.find_element(By.CLASS_NAME, "moveRowBtn")
      #       driver.execute_script("arguments[0].scrollIntoView(true);", move_btn)

      #       wait.until(EC.element_to_be_clickable(move_btn)).click()
      #       print("✅ Клик по moveRowBtn прошёл")
      #       time.sleep(10)
      #       # ⏱ Ждём появления меню
      #       menu_show = wait.until(EC.element_to_be_clickable((By.ID, "menuShow")))
      #       print("✅ menuShow найден")
      #       try:
      #          menu_show.click()
      #       except:
      #          driver.execute_script("arguments[0].click();", menu_show)
      #       print("✅ Клик по menuShow прошёл")
      #       print("✅ menuShow найден")
      #       time.sleep(5)
      #       move_to = wait.until(EC.element_to_be_clickable((By.ID, "moveTo")))
      #       move_to.click()
      #       print("✅ Клик по moveTo прошёл")             
      #       # move_btn = target_block.find_element(By.CLASS_NAME, "moveRowBtn")
      #       # driver.execute_script("arguments[0].scrollIntoView(true);", move_btn)
      #       # move_btn.click()

      #       # menu_show = wait.until(EC.element_to_be_clickable((By.ID, "menuShow")))
      #       # menu_show.click()
      #       # move_to = wait.until(EC.element_to_be_clickable((By.ID, "moveTo")))
      #       # move_to.click()
      #   except Exception as e:
      #       print(f"❗ Ошибка при поиске или нажатии moveRoxwBtn/menuShow/moveTo: {e}")
      #       continue

        try:
               show_blocks = day_view.find_elements(By.CLASS_NAME, "rowItem")
               target_block = None
               for block in show_blocks:
                  try:
                        title_div = block.find_element(By.CLASS_NAME, "title")
                        if show["title"] in title_div.text:
                           target_block = block
                           break
                  except:
                        continue

               if not target_block:
                  print(f"❗ Блок с фильмом '{show['title']}' не найден.")
                  continue
               print(f"❗ Блок с фильмом '{show['title']}' найден.")
        except Exception as e:
               print(f"❗ Ошибка при поиске блока с фильмом: {e}")
               continue

        try:
               move_btn = target_block.find_element(By.CLASS_NAME, "moveRowBtn")
               driver.execute_script("arguments[0].scrollIntoView(true);", move_btn)
               wait.until(EC.element_to_be_clickable(move_btn)).click()
               print("✅ Клик по moveRowBtn прошёл")
        except Exception as e:
               print(f"❗ Ошибка при клике по moveRowBtn: {e}")
               continue

        time.sleep(10)

        try:
               menu_show = wait.until(EC.element_to_be_clickable((By.ID, "menuShow")))
               print("✅ menuShow найден")
               try:
                  menu_show.click()
               except Exception as e_click:
                  print(f"❗ Простой клик по menuShow не удался, пробуем через JS: {e_click}")
                  driver.execute_script("arguments[0].click();", menu_show)
               print("✅ Клик по menuShow прошёл")
        except Exception as e:
               print(f"❗ Ошибка при работе с menuShow: {e}")
               screenshot_name = re.sub(r'[\\/:*?"<>|]+', "_", show["title"])
               driver.save_screenshot(str(SCREENSHOTS_DIR / f"error_menuShow_{screenshot_name}.png"))
               print("Встал на ожидание на 100 секунд для проверки")
               time.sleep(100)
               continue

        time.sleep(5)

        try:
               move_to = wait.until(EC.element_to_be_clickable((By.ID, "moveTo")))
               move_to.click()
               print("✅ Клик по moveTo прошёл")
        except Exception as e:
               print(f"❗ Ошибка при клике по moveTo: {e}")
               continue


        # Календарь
        try: 
            time.sleep(5)
            wait.until(EC.presence_of_element_located((By.ID, "dateTimeModal")))
            day_cells = driver.find_elements(By.CLASS_NAME, "day")
            target_day = date.split(".")[0]
            if target_day.startswith("0"):
                target_day = target_day[1:]
            print(target_day + " ДЕНЬ")
            print("ДЕНЬ")

            for cell in day_cells:
                if cell.text.strip() == target_day and "notSelectable" not in cell.get_attribute("class"):
                    cell.click()
                    break
        except Exception as e:
            print(f"❗ Ошибка при выборе дня в календаре: {e}")
            continue

        # Время
        try:
            time.sleep(5)
            hour_str, minute_str = show["time"].split(":")
            # Час
            wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "timepicker-hour"))).click()
            for cell in driver.find_elements(By.CLASS_NAME, "hour"):
                if cell.text.strip() == hour_str:
                    cell.click()
                    break

            # Минуты
            wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "timepicker-minute"))).click()
            for cell in driver.find_elements(By.CLASS_NAME, "minute"):
                if cell.text.strip() == minute_str:
                    cell.click()
                    break
        except Exception as e:
            print(f"❗ Ошибка при установке времени: {e}")
            continue

        # Подтверждение
        try:
            confirm_btn = wait.until(EC.element_to_be_clickable((By.ID, "confirmDateTimeBtn")))
            confirm_btn.click()
            print(f"✅ Фильм '{show['title']}' добавлен в расписание.")
        except Exception as e:
            print(f"❗ Ошибка при подтверждении времени: {e}")
            continue
        time.sleep(10)
        print(f"✅ Встал на паузу на 10 секунд")



# Старый код
# # Поиск сегодняшней даты 
# today = datetime.now().strftime("%d/%m/%Y")
# print ("сегодняшняя дата ", today)

# day_headers_shelder = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayHeader")))


# today_index = None
# for i,day_header in enumerate(day_headers_shelder):
#    date_element = day_header.find_element(By.CLASS_NAME, "date")
#    date_text = date_element.text.strip()
#    if date_text == today:
#       today_index = i
#       print(f"найдена сегодняшняя дата: {date_text}, кликаем.... Индекс: {today_index}")
#       day_header.click()
#       break
#    else:
#       print("\033[91mСегодняшняя дата не найдена на странице.\033[0m")

# # Поиск нужного столбца
# if today_index is None:
#    print("Сегодняшняя дата не найдена!")
# else: 
#     day_views = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, "dayView")))

#     today_day_view = day_views[today_index]

#     hour_lines = today_day_view.find_elements(By.CLASS_NAME, "hourLine")
#     if hour_lines:
#        last_hour_line = hour_lines[-2]
#        print(f"Нажимаем на предпоследний hourLine с индексом {len(hour_lines)-2}")
#        driver.execute_script("arguments[0].scrollIntoView(true);", last_hour_line)
#        last_hour_line.click()
#     else:
#        print("В dayView нет элементов hourLine")

# time.sleep(15)
# # Добавляем фильм в рассписание
# time.sleep(15)

# caret_btn = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "caretBtn")))
# caret_btn.click()
# print("Клик по списку фильмов")

# show_list = wait.until(EC.presence_of_element_located((By.ID, "listOfShows")))
# show_items = show_list.find_elements(By.TAG_NAME, "li")

# found = False

# for item in show_items:
#    text = item.text.strip()
#    if "Три богатыря" in text:
#       print(f"Найден пункт: {text}, кликаем")
#       item.click()
#       found = True
#       break
#    else:
#       print("\033[91mФильм 'Три богатыря' не найден в списке!\033[0m")

# try: 
#    ok_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".popover-inner .ok.btn")))
#    ok_button.click()
#    print("Кнопка OK нажата")    
# except Exception as e: 
#    print(f"\033[91mОшибка при нажатии на OK: {e}\033[0m")



# # Далее находим наш сформированный фильм
# try:
#     print("Ищем блок с фильмом 'Три богатыря' в dayView...")
#     show_blocks = today_day_view.find_elements(By.CLASS_NAME, "rowItem")
#     found_block = None

#     for block in show_blocks:
#         try:
#             title_div = block.find_element(By.CLASS_NAME, "title")
#             if "Три богатыря" in title_div.text:
#                 found_block = block
#                 break
#         except:
#             continue

#     if found_block:
#         print("✅ Блок с фильмом найден!")

#         move_btn = found_block.find_element(By.CLASS_NAME, "moveRowBtn")
#         driver.execute_script("arguments[0].scrollIntoView(true);", move_btn)
#         move_btn.click()
#         print("✅ Нажата кнопка moveRowBtn")

#         menu_show = wait.until(EC.element_to_be_clickable((By.ID, "menuShow")))
#         menu_show.click()
#         print("✅ Нажата кнопка menuShow")

#     else:
#         print("\033[91m❗ Блок с фильмом 'Три богатыря' не найден в dayView!\033[0m")

# except Exception as e:
#     print(f"\033[91m❗ Ошибка при попытке нажать moveRowBtn или menuShow: {e}\033[0m")

# try:
#     move_to = wait.until(EC.element_to_be_clickable((By.ID, "moveTo")))
#     move_to.click()
#     print("Клик по Move To выполнен")
# except Exception as e:
#    print(f"Ошибка при клике по Move To: {e}")
#    driver.quit()
#    exit()

# # Ждем модального окна на с календарем
# try:
#    wait.until(EC.presence_of_element_located((By.ID, "dateTimeModal")))
#    print("Окно выбора даты появилось")
# except Exception as e:
#    print(f"Модальное окно не появлось: {e}")
#    driver.quit()
#    exit()

# today = str(datetime.today().day)


# time.sleep(15)
# # Ищем даты подходящие 
# try:
#     all_days = driver.find_elements(By.CLASS_NAME, "day")
#     clicked = False

#     for day in all_days:
#         class_attr = day.get_attribute("class")
#         day_text = day.text.strip()
#         print(f"День который нашел: {day_text}")
#         if day_text == today and "notSelectable" not in class_attr and "new" not in class_attr:
#             day.click()
#             print(f"✅ Клик по дню {today} выполнен")
#             clicked = True
#             break

#     if not clicked:
#         print(f"❗ Не удалось найти подходящий день {today} для клика")
#         driver.quit()
#         exit()
# except Exception as e:
#    print(f"Ошибка при выборе даты: {e}")
#    driver.quit()
#    exit()
   
# # Установка времени и минут 
# try:
#     show_hours = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "timepicker-hour")))
#     show_hours.click()
#     print("🔽 Раскрыли выбор часов")

#     hour_table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "timepicker-hours")))
#     hour_cells = hour_table.find_elements(By.CLASS_NAME, "hour")

#     for cell in hour_cells:
#        if cell.text.strip() == "22":
#           cell.click()
#           print("Установлен час: 22")
#           break
#     else:
#        print("Час 22 не найден")
# except Exception as e:
#    print(f"Ошибка при установке часа: {e}")
#    driver.quit()
#    exit()

# try:
#     show_minutes = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "timepicker-minute")))
#     show_minutes.click()
#     print("🔽 Раскрыли выбор минут")

#     minute_table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "timepicker-minutes")))
#     minute_cells = minute_table.find_elements(By.CLASS_NAME, "minute")

#     for cell in minute_cells:
#        if cell.text.strip() == "15":
#           cell.click()
#           print("Установлены минуты: 15")
#           break
#     else:
#        print("минуты 15 не найдены")
# except Exception as e:
#    print(f"Ошибка при установке минут: {e}")
#    driver.quit()
#    exit()

# # Установка времени
# try:
#     confirm_btn = wait.until(EC.element_to_be_clickable((By.ID, "confirmDateTimeBtn")))
#     confirm_btn.click()
#     print("✅ Время подтверждено, нажата кнопка confirmDateTimeBtn")
# except Exception as e:
#     print(f"❗ Ошибка при нажатии confirmDateTimeBtn: {e}")
#     driver.quit()
#     exit()
   

time.sleep(200)
driver.quit()

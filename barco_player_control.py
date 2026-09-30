import argparse
import atexit
import json
import os
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait


BASE_DIR = Path(__file__).resolve().parent
ARTIFACTS_DIR = BASE_DIR / "automation_artifacts"
LOG_PATH = ARTIFACTS_DIR / "barco_player_control.log"
DEFAULT_BARCO_URL = "https://192.168.100.2:43744"

PLAYER_MODE_NORMAL = 0
PLAYER_MODE_SCHEDULER = 1
PLAYER_STATE_CLEARED = 0
PLAYER_STATE_STOPPED = 4


class Tee:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for stream in self.streams:
            if not stream.closed:
                stream.write(data)
                stream.flush()

    def flush(self):
        for stream in self.streams:
            if not stream.closed:
                stream.flush()


def configure_logging():
    ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    log_file = LOG_PATH.open("a", encoding="utf-8")
    sys.stdout = Tee(original_stdout, log_file)
    sys.stderr = Tee(original_stderr, log_file)
    print(f"\n===== Player control run: {datetime.now():%Y-%m-%d %H:%M:%S} =====")

    def close_log():
        sys.stdout = original_stdout
        sys.stderr = original_stderr
        if not log_file.closed:
            log_file.close()

    atexit.register(close_log)
    return close_log


def build_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--ignore-certificate-errors")

    candidates = []
    if os.getenv("CHROMEDRIVER_PATH"):
        candidates.append(Path(os.environ["CHROMEDRIVER_PATH"]))
    candidates.append(BASE_DIR / "drivers" / "chromedriver-win64" / "chromedriver.exe")

    for candidate in candidates:
        if not candidate.exists():
            continue
        try:
            print(f"Запуск ChromeDriver: {candidate}")
            return webdriver.Chrome(service=Service(str(candidate)), options=options)
        except Exception as error:
            print(f"ChromeDriver {candidate} не запустился: {error}")

    print("Запуск Chrome через Selenium Manager")
    return webdriver.Chrome(options=options)


def wait_for_model(driver, timeout=20):
    WebDriverWait(driver, timeout).until(
        lambda current: current.execute_script(
            "return typeof window.g_MainStatusModel !== 'undefined' && "
            "window.g_MainStatusModel !== null;"
        )
    )


def main_status(driver):
    wait_for_model(driver)
    return driver.execute_script(
        """
const model = window.g_MainStatusModel;
return {
    playerMode: Number(model.get('playerMode')),
    playerState: Number(model.get('playerState')),
    lampOn: Boolean(model.get('isProjectorLampOn')),
    dowserClosed: Boolean(model.get('isProjectorDowserClosed'))
};
"""
    )


def wait_for_status(driver, predicate, description, timeout=30):
    def condition(current):
        status = main_status(current)
        return status if predicate(status) else False

    status = WebDriverWait(driver, timeout, poll_frequency=0.5).until(condition)
    print(f"Подтверждено: {description}. Состояние: {status}")
    return status


def send_sms_command(driver, method, *parameters, timeout=20):
    driver.set_script_timeout(timeout)
    raw_result = driver.execute_async_script(
        """
const method = arguments[0];
const parameters = arguments[1];
const done = arguments[arguments.length - 1];

if (!window.SmsComm || typeof window.SmsComm[method] !== 'function') {
    done(JSON.stringify({success: false, clientError: 'SmsComm method unavailable'}));
    return;
}

const callback = {
    callback: function(raw) {
        done(typeof raw === 'string' ? raw : JSON.stringify(raw));
    }
};

try {
    window.SmsComm[method].apply(window.SmsComm, parameters.concat([callback]));
} catch (error) {
    done(JSON.stringify({success: false, clientError: String(error)}));
}
""",
        method,
        list(parameters),
    )

    try:
        result = json.loads(raw_result) if isinstance(raw_result, str) else raw_result
    except json.JSONDecodeError as error:
        raise RuntimeError(f"Некорректный ответ команды {method}: {raw_result}") from error

    if not isinstance(result, dict) or not result.get("success"):
        raise RuntimeError(f"Команда {method} завершилась ошибкой: {result}")
    print(f"Команда {method} выполнена: {result}")
    return result


def login(driver, base_url):
    username = os.getenv("BARCO_USERNAME", "admin")
    password = os.getenv("BARCO_PASSWORD", "Admin1234")

    driver.get(base_url)
    wait = WebDriverWait(driver, 20)
    username_input = wait.until(lambda current: current.find_element(By.ID, "loginUsername"))
    password_input = wait.until(lambda current: current.find_element(By.ID, "loginPass"))
    username_input.send_keys(username)
    password_input.send_keys(password)
    driver.find_element(By.ID, "loginSubmit").click()
    wait_for_model(driver)
    print("Авторизация Barco выполнена")


def unlock_controls_if_needed(driver):
    try:
        lock = driver.find_element(By.ID, "lockApp")
        if "lockAppRed" in lock.get_attribute("class").split():
            lock.click()
            WebDriverWait(driver, 10).until(
                lambda current: "lockAppRed"
                not in current.find_element(By.ID, "lockApp").get_attribute("class").split()
            )
            print("Панель управления разблокирована")
    except Exception as error:
        print(f"Проверка блокировки пропущена: {error}")


def shutdown_and_schedule(driver, base_url):
    unlock_controls_if_needed(driver)
    status = main_status(driver)
    print(f"Начальное состояние: {status}")

    if status["playerMode"] == PLAYER_MODE_SCHEDULER:
        send_sms_command(driver, "changeMode", PLAYER_MODE_NORMAL)
        status = wait_for_status(
            driver,
            lambda value: value["playerMode"] == PLAYER_MODE_NORMAL,
            "Scheduler выключен",
        )
    else:
        print("Scheduler уже выключен")

    if status["playerState"] not in (PLAYER_STATE_CLEARED, PLAYER_STATE_STOPPED):
        send_sms_command(driver, "stop")
        wait_for_status(
            driver,
            lambda value: value["playerState"] in (
                PLAYER_STATE_CLEARED,
                PLAYER_STATE_STOPPED,
            ),
            "воспроизведение остановлено",
        )
    else:
        print("Player уже остановлен или очищен")

    status = main_status(driver)

    if not status["dowserClosed"]:
        send_sms_command(driver, "setDowser", True)
        wait_for_status(
            driver,
            lambda value: value["dowserClosed"],
            "заслонка закрыта",
        )
    else:
        print("Заслонка уже закрыта")

    status = main_status(driver)
    if status["lampOn"]:
        send_sms_command(driver, "setLamp", False)
        wait_for_status(
            driver,
            lambda value: not value["lampOn"],
            "лампа выключена",
            timeout=60,
        )
    else:
        print("Лампа уже выключена")

    status = main_status(driver)
    if status["playerMode"] != PLAYER_MODE_SCHEDULER:
        send_sms_command(driver, "changeMode", PLAYER_MODE_SCHEDULER)
        status = wait_for_status(
            driver,
            lambda value: value["playerMode"] == PLAYER_MODE_SCHEDULER,
            "Scheduler включен",
        )

    if status["lampOn"] or not status["dowserClosed"]:
        raise RuntimeError(f"Небезопасное итоговое состояние: {status}")
    print(f"Операция завершена успешно. Итоговое состояние: {status}")


def parse_args():
    parser = argparse.ArgumentParser(description="Barco player and projector control")
    parser.add_argument(
        "action",
        choices=["shutdown-and-schedule"],
        help="Safe projector control action",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    close_log = configure_logging()
    driver = None
    try:
        base_url = os.getenv("BARCO_URL", DEFAULT_BARCO_URL).rstrip("/")
        driver = build_driver()
        login(driver, base_url)
        if args.action == "shutdown-and-schedule":
            shutdown_and_schedule(driver, base_url)
    except Exception:
        print("Необработанная ошибка управления Barco:")
        traceback.print_exc()
        if driver is not None:
            screenshot_path = ARTIFACTS_DIR / "barco_player_control_error.png"
            try:
                driver.save_screenshot(str(screenshot_path))
                print(f"Скриншот ошибки: {screenshot_path}")
            except Exception as screenshot_error:
                print(f"Не удалось сохранить скриншот: {screenshot_error}")
        raise
    finally:
        if driver is not None:
            driver.quit()
        close_log()


if __name__ == "__main__":
    main()

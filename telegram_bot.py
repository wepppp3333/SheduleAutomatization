import json
import logging
import os
from pathlib import Path

import httpx
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
)


BASE_DIR = Path(__file__).resolve().parent
BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
ALLOWED_USER_IDS = {
    int(value.strip())
    for value in os.getenv("TELEGRAM_ALLOWED_USER_IDS", "").split(",")
    if value.strip().isdigit()
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logging.getLogger("httpx").setLevel(logging.WARNING)
logger = logging.getLogger("barco-telegram-bot")


def load_cinemas():
    with (BASE_DIR / "cinemas.json").open(encoding="utf-8") as file:
        return json.load(file)


CINEMAS = load_cinemas()


def is_allowed(update: Update):
    user = update.effective_user
    return bool(user and user.id in ALLOWED_USER_IDS)


async def deny(update: Update):
    user_id = update.effective_user.id if update.effective_user else "unknown"
    logger.warning("Denied Telegram user: %s", user_id)
    if update.callback_query:
        await update.callback_query.answer("Нет доступа", show_alert=True)
    elif update.effective_message:
        await update.effective_message.reply_text("Нет доступа.")


def main_keyboard():
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("Сформировать расписание", callback_data="run_menu")],
            [
                InlineKeyboardButton(
                    "Выключить фильм и включить очередь",
                    callback_data="shutdown_schedule_menu",
                )
            ],
            [
                InlineKeyboardButton(
                    "Отключить очередь и проектор",
                    callback_data="disable_projector_menu",
                )
            ],
            [InlineKeyboardButton("Статус", callback_data="status_menu")],
        ]
    )


def cinema_keyboard(action):
    rows = [
        [InlineKeyboardButton(config["label"], callback_data=f"{action}:{key}")]
        for key, config in CINEMAS.items()
    ]
    rows.append([InlineKeyboardButton("Назад", callback_data="main")])
    return InlineKeyboardMarkup(rows)


def get_cinema_credentials(cinema_key):
    cinema = CINEMAS.get(cinema_key)
    if not cinema:
        raise ValueError("Кинотеатр не найден")

    api_token = os.getenv(cinema["api_token_env"], "")
    if not api_token:
        raise RuntimeError(
            f"Не задана переменная {cinema['api_token_env']} для {cinema['label']}"
        )
    return cinema, api_token


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return
    await update.effective_message.reply_text(
        "Управление расписанием Barco", reply_markup=main_keyboard()
    )


async def whoami(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if user:
        await update.effective_message.reply_text(f"Ваш Telegram user ID: {user.id}")


async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await deny(update)
        return

    query = update.callback_query
    await query.answer()
    action = query.data

    if action == "main":
        await query.edit_message_text(
            "Управление расписанием Barco", reply_markup=main_keyboard()
        )
        return

    if action == "run_menu":
        await query.edit_message_text(
            "Выберите кинотеатр:", reply_markup=cinema_keyboard("confirm")
        )
        return

    if action == "status_menu":
        await query.edit_message_text(
            "Статус какого кинотеатра проверить?",
            reply_markup=cinema_keyboard("status"),
        )
        return

    if action == "shutdown_schedule_menu":
        await query.edit_message_text(
            "В каком кинотеатре остановить фильм, выключить проектор "
            "и снова включить очередь?",
            reply_markup=cinema_keyboard("shutdown_schedule_confirm"),
        )
        return

    if action == "disable_projector_menu":
        await query.edit_message_text(
            "В каком кинотеатре отключить очередь, закрыть заслонку "
            "и выключить лампу?",
            reply_markup=cinema_keyboard("disable_projector_confirm"),
        )
        return

    command, cinema_key = action.split(":", 1)
    cinema = CINEMAS.get(cinema_key)
    if not cinema:
        await query.edit_message_text("Кинотеатр не найден.")
        return

    if command == "confirm":
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "Подтвердить запуск", callback_data=f"run:{cinema_key}"
                    )
                ],
                [InlineKeyboardButton("Отмена", callback_data="main")],
            ]
        )
        await query.edit_message_text(
            f"Запустить формирование расписания: {cinema['label']}?",
            reply_markup=keyboard,
        )
        return


    if command == "shutdown_schedule_confirm":
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "Подтвердить выключение",
                        callback_data=f"shutdown_schedule:{cinema_key}",
                    )
                ],
                [InlineKeyboardButton("Отмена", callback_data="main")],
            ]
        )
        await query.edit_message_text(
            f"Выполнить в кинотеатре {cinema['label']}?\n\n"
            "Будет выключена очередь, остановлен фильм, закрыта заслонка, "
            "выключена лампа и снова включена очередь.",
            reply_markup=keyboard,
        )
        return

    if command == "disable_projector_confirm":
        keyboard = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "Подтвердить отключение",
                        callback_data=f"disable_projector:{cinema_key}",
                    )
                ],
                [InlineKeyboardButton("Отмена", callback_data="main")],
            ]
        )
        await query.edit_message_text(
            f"Выполнить в кинотеатре {cinema['label']}?\n\n"
            "Очередь останется выключенной, заслонка будет закрыта, "
            "лампа выключена. Команда Stop отправлена не будет.",
            reply_markup=keyboard,
        )
        return

    try:
        cinema, api_token = get_cinema_credentials(cinema_key)
        headers = {"X-API-Key": api_token}
        async with httpx.AsyncClient(timeout=15) as client:
            if command == "run":
                response = await client.post(
                    f"{cinema['api_url'].rstrip('/')}/run-schedule", headers=headers
                )
            elif command == "shutdown_schedule":
                response = await client.post(
                    f"{cinema['api_url'].rstrip('/')}/player/shutdown-and-schedule",
                    headers=headers,
                )
            elif command == "disable_projector":
                response = await client.post(
                    f"{cinema['api_url'].rstrip('/')}/player/disable-schedule-and-projector",
                    headers=headers,
                )
            elif command == "status":
                response = await client.get(
                    f"{cinema['api_url'].rstrip('/')}/status", headers=headers
                )
            else:
                raise ValueError("Неизвестная команда")

        payload = response.json()
        if response.status_code >= 400:
            await query.edit_message_text(
                f"Ошибка {cinema['label']}: HTTP {response.status_code}\n{payload}"
            )
        elif command == "run":
            await query.edit_message_text(
                f"Расписание запущено: {cinema['label']}\n"
                f"Job ID: {payload.get('job_id')}"
            )
        elif command == "shutdown_schedule":
            await query.edit_message_text(
                f"Выключение запущено: {cinema['label']}\n"
                f"Job ID: {payload.get('job_id')}\n"
                "Результат можно проверить кнопкой «Статус»."
            )
        elif command == "disable_projector":
            await query.edit_message_text(
                f"Отключение очереди и проектора запущено: {cinema['label']}\n"
                f"Job ID: {payload.get('job_id')}\n"
                "Результат можно проверить кнопкой «Статус»."
            )
        else:
            await query.edit_message_text(
                f"Статус {cinema['label']}: {payload.get('status')}\n"
                f"Операция: {payload.get('action') or 'не указана'}\n"
                f"Job ID: {payload.get('job_id') or 'нет'}\n"
                f"Код завершения: {payload.get('exit_code')}"
            )
    except (httpx.HTTPError, RuntimeError, ValueError) as error:
        logger.exception("Telegram command failed")
        await query.edit_message_text(f"Не удалось выполнить команду: {error}")


def main():
    if not BOT_TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN is not configured")
    if not ALLOWED_USER_IDS:
        logger.warning(
            "TELEGRAM_ALLOWED_USER_IDS is empty; only /whoami can be used"
        )

    application = Application.builder().token(BOT_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("whoami", whoami))
    application.add_handler(CallbackQueryHandler(button))
    application.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()

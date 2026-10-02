#!/bin/bash

script_path="$0"
if [[ -L "$script_path" ]]; then
    script_path="$(readlink "$script_path")"
fi
project_dir="$(cd "$(dirname "$script_path")" && pwd -P)"
cd "$project_dir" || exit 1

if pgrep -f '[p]ython.*telegram_bot.py' >/dev/null; then
    echo "Barco Telegram bot is already running."
    read -r -p "Press Enter to close this window..."
    exit 0
fi

if [[ ! -f .env || ! -x .venv/bin/python ]]; then
    echo "Missing .env or .venv/bin/python in $project_dir"
    read -r -p "Press Enter to close this window..."
    exit 1
fi

set -a
source .env
set +a

if [[ -z "$TELEGRAM_BOT_TOKEN" || -z "$TELEGRAM_ALLOWED_USER_IDS" ]]; then
    echo "Missing TELEGRAM_BOT_TOKEN or TELEGRAM_ALLOWED_USER_IDS in .env"
    read -r -p "Press Enter to close this window..."
    exit 1
fi

echo "Starting Barco Telegram bot. Leave this window open; press Ctrl+C to stop."
exec .venv/bin/python telegram_bot.py

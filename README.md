# SheduleAutomatization

Автоматизация загрузки расписания из Excel в Barco Scheduler.

## Установка

```powershell
py -m pip install -r requirements.txt
```

## Локальный запуск

```powershell
py barco_open_chrome.py
```

## Удалённый запуск через Tailscale

Задайте отдельный длинный токен на компьютере кинотеатра:

```powershell
[Environment]::SetEnvironmentVariable(
    "BARCO_API_TOKEN",
    "ЗАМЕНИТЕ_НА_ДЛИННЫЙ_СЛУЧАЙНЫЙ_ТОКЕН",
    "User"
)
```

Откройте новое окно PowerShell и запустите handler:

```powershell
.\start_automation_server.ps1
```

Проверка с другого компьютера в той же сети Tailscale:

```powershell
curl http://TAILSCALE_IP:8080/health
```

Запуск расписания:

```powershell
Invoke-RestMethod `
  -Method Post `
  -Uri "http://TAILSCALE_IP:8080/run-schedule" `
  -Headers @{"X-API-Key"="ВАШ_ТОКЕН"}
```

Статус последнего задания:

```powershell
Invoke-RestMethod `
  -Uri "http://TAILSCALE_IP:8080/status" `
  -Headers @{"X-API-Key"="ВАШ_ТОКЕН"}
```

Безопасно остановить фильм, закрыть заслонку, выключить лампу и снова включить
очередь:

```powershell
Invoke-RestMethod `
  -Method Post `
  -Uri "http://TAILSCALE_IP:8080/player/shutdown-and-schedule" `
  -Headers @{"X-API-Key"="ВАШ_ТОКЕН"}
```

Операция выполняется асинхронно. Проверяйте результат через `/status` и журнал
`automation_artifacts/barco_player_control.log`.

## Telegram-бот

Создайте бота через `@BotFather`, затем задайте токен и API-токен кинотеатра:

```powershell
[Environment]::SetEnvironmentVariable("TELEGRAM_BOT_TOKEN", "ТОКЕН_ОТ_BOTFATHER", "User")
[Environment]::SetEnvironmentVariable("LUKOYANOV_API_TOKEN", "API_ТОКЕН_ЛУКОЯНОВА", "User")
```

В новом PowerShell запустите:

```powershell
.\start_telegram_bot.ps1
```

Отправьте боту `/whoami`, сохраните полученный числовой ID и остановите бота.
Разрешите этому пользователю управление:

```powershell
[Environment]::SetEnvironmentVariable("TELEGRAM_ALLOWED_USER_IDS", "ВАШ_TELEGRAM_ID", "User")
```

После перезапуска команда `/start` покажет кнопки формирования расписания,
безопасного выключения фильма и проверки статуса. Опасные операции требуют
отдельного подтверждения.

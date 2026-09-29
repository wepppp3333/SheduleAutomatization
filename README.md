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

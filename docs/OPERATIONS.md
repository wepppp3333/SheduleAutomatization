# Памятка по запуску Barco и Telegram-бота

Актуально на 30.09.2026. Команды ниже выполняются вручную; автоматический запуск после перезагрузки не настроен этой памяткой.

## 1. Что где работает

Телефон с Telegram → бот на Mac → Tailscale → API на Windows в кинотеатре → Chrome/Selenium → проектор Barco.

На Mac должен работать один экземпляр `telegram_bot.py`. На каждом компьютере кинотеатра должен работать свой `start_automation_server.ps1`. Компьютеры должны быть включены и не находиться в режиме сна. Для браузерной автоматизации оставляйте пользовательский сеанс Windows открытым.

| Кинотеатр | Компьютер | Tailscale IP | Папка проекта на Windows |
|---|---|---|---|
| Лукоянов | DESKTOP-LURMO7V | 100.118.9.80 | `C:\Users\Lukoyanov\Desktop\test\SheduleAutomatization` |
| Кинель | DESKTOP-KILFC2J | 100.78.107.67 | `C:\Users\Ust-Kinel\Desktop\barco_automatization\SheduleAutomatization` |
| Новоспасское | DESKTOP-KCJPV2S | 100.107.6.2 | `C:\Users\october\Desktop\test\SheduleAutomatization` |

Проект на Mac: `/Users/fadin/Desktop/SheduleAutomatization`.

## 2. Обычный запуск после включения компьютеров

1. Включите Tailscale на Mac и компьютерах кинотеатров.
2. На каждом нужном Windows-компьютере запустите API по разделу 3.
3. На Mac включите VPSUS для связи с Telegram и проверьте маршруты по разделу 5.
4. На Mac запустите бота по разделу 4.
5. В Telegram откройте `@barco_close_bot`, отправьте `/start`, выберите «Статус» и нужный кинотеатр.

`idle` означает, что API готов и после его запуска ещё не было задания. Это не проверка соединения с самим проектором.

## 3. Запуск API на Windows в кинотеатре

Откройте PowerShell и перейдите в папку соответствующего кинотеатра из таблицы. Например, Новоспасское:

```powershell
cd C:\Users\october\Desktop\test\SheduleAutomatization
```

Далее одинаково для всех трёх:

```powershell
$env:BARCO_API_TOKEN = [Environment]::GetEnvironmentVariable("BARCO_API_TOKEN", "User")
if (-not $env:BARCO_API_TOKEN) { throw "API token is missing" }
powershell -ExecutionPolicy Bypass -File .\start_automation_server.ps1
```

Нормальный результат: `Application startup complete` и `Uvicorn running on http://0.0.0.0:8080`.

Оставьте окно открытым. `0.0.0.0` означает прослушивание всех локальных интерфейсов; для обращения используйте Tailscale IP из таблицы. Доступность извне зависит также от брандмауэра и правил Tailscale.

Для штатной остановки API нажмите Ctrl+C в его окне, предварительно убедившись, что задание не выполняется. Остановка API не является командой остановки фильма.

## 4. Запуск Telegram-бота на Mac

Откройте Terminal. Сначала проверьте, нет ли уже работающего бота:

```bash
pgrep -fl '[t]elegram_bot.py'
```

Если процесс уже есть, второй экземпляр не запускайте. Если вывода нет:

```bash
cd /Users/fadin/Desktop/SheduleAutomatization
set -a
source .env
set +a
.venv/bin/python telegram_bot.py
```

Ожидаемый результат: `Application started`. Оставьте Terminal открытым. Чтобы остановить бота, нажмите Ctrl+C в этом окне. Для перезапуска снова выполните блок выше. После изменения `.env` или `cinemas.json` бот требует перезапуска.

Бот работает на Mac, а не на телефоне. Если Mac выключен, спит или процесс завершился, ответы прекращаются. Запуск через временную сессию помощника не заменяет постоянную службу.

При отсутствии зависимостей в локальном окружении:

```bash
cd /Users/fadin/Desktop/SheduleAutomatization
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

## 5. VPSUS и Tailscale на Mac

В текущей конфигурации VPSUS нужен для Telegram, но маршруты к кинотеатрам иногда уходят через другой интерфейс. Проверка:

```bash
route -n get 100.118.9.80
route -n get 100.78.107.67
route -n get 100.107.6.2
```

Найдите интерфейс Tailscale по локальному адресу Mac `100.120.123.116`:

```bash
ifconfig | awk '/^[a-zA-Z0-9]+:/{iface=$1} /inet 100\.120\.123\.116 /{print iface, $2}'
```

На момент настройки это `utun6`, но номер может измениться. Если вывод показывает другой интерфейс, замените `utun6` ниже на него. Если адрес не найден, сначала проверьте подключение Tailscale.

Добавление отсутствующих маршрутов, выполняется на Mac:

```bash
sudo route -n add -host 100.118.9.80 -interface utun6
sudo route -n add -host 100.78.107.67 -interface utun6
sudo route -n add -host 100.107.6.2 -interface utun6
```

При `File exists` проверьте текущий маршрут. Если он неправильный, используйте `change`, например:

```bash
sudo route -n change -host 100.107.6.2 -interface utun6
```

Введите пароль Mac в системном запросе. После перезагрузки или переподключения VPN маршруты нужно проверять повторно. `utun` сам по себе не доказывает, что это Tailscale: сверяйте его с адресом Mac выше.

## 6. Кнопки Telegram

Отправьте `/start`, выберите действие, кинотеатр и подтвердите операцию, если бот запросит подтверждение.

| Кнопка | Что делает |
|---|---|
| Сформировать расписание | Читает Excel на компьютере выбранного кинотеатра и добавляет сеансы в Barco |
| Выключить фильм и включить очередь | Отключает Scheduler, останавливает активный фильм, закрывает заслонку, выключает лампу, снова включает Scheduler |
| Отключить очередь и проектор | Отключает Scheduler, закрывает заслонку и выключает лампу; команду Stop не отправляет |
| Остановить фильм и выключить очередь | Отключает Scheduler, останавливает активный фильм, закрывает заслонку и выключает лампу; Scheduler остаётся выключенным |
| Статус | Показывает состояние последнего задания API |

Отдельной кнопки «только включить очередь» сейчас нет. Включение Scheduler разрешает автоматическую работу по расписанию.

После сообщения о запуске проверяйте «Статус»: принятие задания ещё не означает успешное выполнение.

| Статус | Значение |
|---|---|
| idle | С момента запуска API заданий не было |
| running | Задание выполняется |
| success | Скрипт завершился с кодом 0 |
| failed | Скрипт завершился с ошибкой; нужен журнал |

Статус хранится в памяти и сбрасывается при перезапуске API. Он не показывает текущее физическое состояние проектора. Проверка доступности API Новоспасского подтвердила подключение, но сама по себе не является тестом управления его проектором.

## 7. Подготовка Excel

Поместите актуальное `Рассписание.xlsx` в корень проекта именно на компьютере нужного кинотеатра. Бот не передаёт Excel с Mac и не загружает его из Telegram.

Оставляйте один актуальный входной файл с префиксом `Рассписание` или `Расписание`, чтобы избежать выбора старой копии. Даты идут отдельными строками-маркерами; строки сеансов содержат время `HH:MM` в первом столбце и название фильма во втором.

Парсер пропускает строки с распознанным красным цветом шрифта в ячейке времени или названия. Красная заливка и условное форматирование не равнозначны этому правилу. Проверяйте сообщения «Пропущена красная строка» и полученный `automation_artifacts/schedule.json`.

Ручной запуск расписания на Windows из папки проекта:

```powershell
py barco_open_chrome.py
```

Это реальное изменение расписания. Не запускайте его одновременно с заданием API: защита API не охватывает отдельные ручные процессы.

## 8. Обновление проекта

На нужном Windows-компьютере дождитесь завершения задания, остановите API через Ctrl+C, перейдите в папку проекта и выполните:

```powershell
git pull
```

Продолжайте только если обновление прошло без ошибок:

```powershell
git log -1 --oneline
py -m pip install -r requirements.txt
```

Затем снова запустите API по разделу 3. Не ориентируйтесь на один навсегда фиксированный хеш коммита: версия меняется по мере доработок.

Если `git pull` сообщает о локальных изменениях или конфликтах, сохраните вывод для разбора. Не используйте `git reset --hard` для обычного обновления.

На Mac: остановите бота в его Terminal, выполните `git pull` из папки проекта, при необходимости обновите зависимости через `.venv/bin/python -m pip install -r requirements.txt` и запустите бота по разделу 4.

### Лукоянов: автоматическое восстановление Tailscale-канала

Если Mac перестаёт получать статус Лукоянова, а исходящий `tailscale ping` с Лукоянова до Mac восстанавливает связь, обновите проект на Лукоянове и один раз сохраните адрес Mac:

```powershell
cd C:\Users\Lukoyanov\Desktop\test\SheduleAutomatization
git pull
[Environment]::SetEnvironmentVariable("BARCO_TAILSCALE_KEEPALIVE_IP", "100.120.123.116", "User")
```

Перезапустите API: остановите прежнее окно через Ctrl+C, затем в новом PowerShell выполните:

```powershell
cd C:\Users\Lukoyanov\Desktop\test\SheduleAutomatization
$env:BARCO_API_TOKEN = [Environment]::GetEnvironmentVariable("BARCO_API_TOKEN", "User")
$env:BARCO_TAILSCALE_KEEPALIVE_IP = [Environment]::GetEnvironmentVariable("BARCO_TAILSCALE_KEEPALIVE_IP", "User")
powershell -ExecutionPolicy Bypass -File .\start_automation_server.ps1
```

Строка `Tailscale keepalive enabled` подтверждает запуск фонового ping раз в минуту. Он работает только пока запущен API и останавливается вместе с ним. Переменная опциональная: на других компьютерах её не задавайте без аналогичной проблемы. Если Tailscale IP Mac изменится, обновите значение.

Это обход нестабильного канала, а не доказательство причины. Для диагностики сети Лукоянова выполните `tailscale netcheck` на Windows и сравните результаты `tailscale ping 100.120.123.116` и доступа Mac к `/health` до/после ping.

## 9. Безопасная проверка связи

На компьютере кинотеатра, в отдельном PowerShell:

```powershell
Invoke-RestMethod http://127.0.0.1:8080/health
```

Ожидается `status: ok` и имя компьютера. Проверка статуса с токеном:

```powershell
$env:BARCO_API_TOKEN = [Environment]::GetEnvironmentVariable("BARCO_API_TOKEN", "User")
Invoke-RestMethod -Uri http://127.0.0.1:8080/status -Headers @{"X-API-Key"=$env:BARCO_API_TOKEN}
```

На Mac проверка всех трёх компьютеров:

```bash
curl --connect-timeout 5 --max-time 10 http://100.118.9.80:8080/health
curl --connect-timeout 5 --max-time 10 http://100.78.107.67:8080/health
curl --connect-timeout 5 --max-time 10 http://100.107.6.2:8080/health
```

GET `/health` и `/status` не запускают управление. POST `/run-schedule` и POST `/player/...` выполняют реальные действия, поэтому для проверки связи они не нужны.

## 10. Где искать ошибки

Все пути ниже относительно папки проекта на компьютере кинотеатра:

| Файл | Содержание |
|---|---|
| `automation_artifacts/barco_automation.log` | Журнал формирования расписания |
| `automation_artifacts/barco_player_control.log` | Журнал управления Player/Control |
| `automation_artifacts/schedule.json` | Последняя выборка из Excel |
| `automation_artifacts/barco_player_control_error.png` | Скриншот ошибки управления, если удалось сохранить |
| `automation_artifacts/screenshots/` | Скриншоты обработанных ошибок расписания |

Последние записи в PowerShell:

```powershell
Get-Content .\automation_artifacts\barco_automation.log -Tail 100
Get-Content .\automation_artifacts\barco_player_control.log -Tail 100
```

API пишет свой вывод в окно PowerShell, бот — в Terminal Mac. Отдельный постоянный файловый журнал этих двух процессов текущими командами не создаётся.

Для разбора ошибки сохраните название кинотеатра, действие, время запуска, Job ID, статус и конец соответствующего журнала.

## 11. Частые проблемы

| Симптом | Что проверить |
|---|---|
| Бот не отвечает | Процесс на Mac, отсутствие сна, доступ к Telegram через VPSUS |
| Один кинотеатр не отвечает | Его Windows/API, Tailscale, маршрут с Mac, брандмауэр |
| `py` не найден | Установку Python и новый PowerShell после установки; `python --version` |
| `No module named ...` | `py -m pip install -r requirements.txt` в папке проекта; на Mac использовать `.venv/bin/python` |
| `requirements.txt` не найден | Текущую папку; не запускать установку из `C:\Windows\System32` |
| Скрипты PowerShell запрещены | Запуск `powershell -ExecutionPolicy Bypass -File .\start_automation_server.ps1` |
| `Invalid API token` / 401 | Совпадение токена Windows с токеном этого кинотеатра в `.env` Mac; перезапустить процессы после изменения |
| Порт 8080 занят / 10048 | Уже запущенный API; второй экземпляр не нужен |
| Timeout | Маршрут, Tailscale и брандмауэр; сначала локальный `/health`, затем проверка с Mac |
| 404 на новом маршруте | Версию кода и перезапуск API; один `git pull` работающий процесс не обновляет |
| 405 на GET управляющего маршрута | Ожидаемо: маршрут принимает POST; действие не выполнялось |
| 409 | Уже выполняется задание; дождаться его результата |
| Ошибка TLS в `git pull` | Обновление не произошло; проверить сеть/VPN. Запуск API после ошибки использует старый код |
| ChromeDriver не соответствует Chrome | Версию Chrome; комплектный драйвер рассчитан на Chrome 153 |

Кто слушает порт 8080 на Windows:

```powershell
Get-NetTCPConnection -LocalPort 8080 -State Listen | Select-Object LocalAddress, LocalPort, OwningProcess
```

Чтобы остановить старый сервер, предпочтительно найдите его окно и нажмите Ctrl+C. Не завершайте все процессы Python подряд.

Переопределение ChromeDriver перед запуском API или скрипта:

```powershell
$env:CHROMEDRIVER_PATH="C:\path\to\chromedriver.exe"
```

## 12. Где хранятся настройки

`cinemas.json` хранит названия и адреса. На Mac локальный `.env` хранит `TELEGRAM_BOT_TOKEN`, `TELEGRAM_ALLOWED_USER_IDS`, `LUKOYANOV_API_TOKEN`, `KINEL_API_TOKEN`, `NOVOSPASSKOYE_API_TOKEN`.

На каждом Windows-компьютере свой токен сохранён в пользовательской переменной `BARCO_API_TOKEN`. Она должна соответствовать переменной этого кинотеатра на Mac. Не создавайте новый токен при каждом запуске. `.env` исключён из Git, поэтому клонирование проекта не переносит секреты.

Для управления Player/Control доступны переменные `BARCO_URL`, `BARCO_USERNAME`, `BARCO_PASSWORD`. Скрипт расписания пока имеет отдельную настройку подключения в коде; нельзя считать, что эти переменные перенастраивают оба скрипта. Обычный адрес проектора в локальной сети кинотеатра: `https://192.168.100.2:43744`.

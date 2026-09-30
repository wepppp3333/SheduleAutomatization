# What this is

Automation script that reads a cinema schedule from Excel and enters shows into a Barco scheduler through Selenium and Chrome. The current runtime target is a Windows workstation that can reach the private Barco URL.

# Directory Map

- `barco_open_chrome.py` -> main automation entry point; Excel parsing, logging, Chrome startup, login, and scheduler automation.
- `barco_player_control.py` -> state-aware Player/Control automation for safe shutdown with either Scheduler restored or left disabled.
- `automation_server.py` -> authenticated FastAPI handler for remote runs over Tailscale; prevents concurrent automation jobs.
- `start_automation_server.ps1` -> discovers the computer's Tailscale IPv4 address and starts the handler on all local interfaces, port 8080.
- `telegram_bot.py` -> allowlisted Telegram control panel for starting and checking cinema automation jobs.
- `cinemas.json` -> non-secret cinema labels, Tailscale API URLs, and names of token environment variables.
- `start_telegram_bot.ps1` -> validates required environment variables and starts the Telegram bot.
- `requirements.txt` -> Python runtime dependencies for automation and the API handler.
- `README.md` -> minimal repository title; not yet a setup guide.
- `test.html` -> unrelated HTML scratch file; not used by the Python automation.
- `automation_artifacts/` -> runtime output created automatically; contains log, generated JSON, and screenshots.
- `automation_artifacts/barco_automation.log` -> append-only console and exception log.
- `automation_artifacts/barco_player_control.log` -> append-only log for Player/Control operations and confirmed hardware states.
- `automation_artifacts/schedule.json` -> normalized schedule generated from the Excel input on every run.
- `automation_artifacts/screenshots/` -> screenshots captured by selected failure handlers.
- `drivers/chromedriver-win64/chromedriver.exe` -> bundled Windows ChromeDriver 153 used before Selenium Manager.
- `Рассписание*.xlsx` or `Расписание*.xlsx` -> expected Excel input in the repository root; not committed by default.
- `__pycache__/` -> generated Python bytecode; not application source.
- `docs/ARCHITECTURE.md` -> detailed data flow, Selenium sequence, selectors, and known risks.

# Commands

Install runtime dependencies:

```powershell
py -m pip install -r requirements.txt
```

Run on the Windows workstation with access to Barco:

```powershell
py barco_open_chrome.py
```

Override the bundled ChromeDriver when testing another Chrome version:

```powershell
$env:CHROMEDRIVER_PATH="C:\path\to\chromedriver.exe"
py barco_open_chrome.py
```

Check syntax without running browser automation:

```powershell
py -m py_compile barco_open_chrome.py barco_player_control.py automation_server.py telegram_bot.py
```

There is currently no automated test, build, migration, or deployment command.

Start the Tailscale-only API handler after setting `BARCO_API_TOKEN`:

```powershell
.\start_automation_server.ps1
```

Start the Telegram bot after loading its local environment variables:

```bash
set -a; source .env; set +a
.venv/bin/python telegram_bot.py
```

# Rules & Gotchas

- Run the real automation only on a machine that can reach `https://192.168.100.2:43744`.
- The script has import-time side effects: it creates artifact directories, rewrites `schedule.json`, starts Chrome, and operates the scheduler.
- Keep the Excel input in the project root and name it with the `Рассписание` or `Расписание` prefix.
- Excel rows are parsed as date markers followed by rows whose first column contains `HH:MM`; titles come from the second column.
- The active scheduler uses helper functions and a fail-fast grouped loop; the previous procedural implementation remains unreachable below `sys.exit` until the new flow is verified.
- Selenium elements become stale after Barco redraws the schedule. Re-query `dayHeader`, `dayView`, `rowItem`, and modal controls after state-changing clicks.
- Prefer `WebDriverWait` for visible/clickable state. Existing `time.sleep` calls are temporary stabilization and should not be multiplied without evidence.
- Minute choices use a three-minute grid. The active code rounds source minutes to the nearest available value.
- Do not assume `.click()` returns an element; Selenium `.click()` returns `None`.
- `div` and `span` values are read through `.text` or `get_attribute(...)`, not `.value`.
- Logs must remain enabled; failures are diagnosed from `automation_artifacts/barco_automation.log` on the remote workstation.
- Player/Control buttons are toggles. Never click lamp, dowser, or Scheduler blindly; read `g_MainStatusModel`, issue only the necessary transition, and wait for the confirmed target state.
- The safe shutdown flow must end with `playerMode == 1`, `isProjectorLampOn == false`, and `isProjectorDowserClosed == true`.
- The disable-projector flow does not send Stop and must end with Scheduler disabled, lamp off, and dowser closed.
- The stop-and-disable flow sends Stop once when needed and must end with playback stopped, Scheduler disabled, lamp off, and dowser closed.
- Credentials and the private URL are currently hard-coded. Do not publish real replacements or add new secrets to Git.
- On the current Mac host, `VPSUS` is needed for Telegram but conflicts with Tailscale's `100.64.0.0/10` route. Both cinema IPs need explicit host routes through the active Tailscale `utun` interface.
- `selenium` and `openpyxl` are required, but no dependency lock file exists yet.
- The bundled driver supports Chrome 153. If Chrome updates to another major version, replace the binary or set `CHROMEDRIVER_PATH`; Selenium Manager is the fallback.
- The tracked `__pycache__/barco_open_chrome.cpython-314.pyc` is generated output and should not be treated as source.

# Docs Links

- Architecture and runtime flow: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

Update this map whenever project structure or commands change.

## 2026-09-30 — Multi-cinema remote automation

### Done
- Added authenticated FastAPI handler (`automation_server.py`) with run/status endpoints and single-job protection.
- Added Windows launcher; Uvicorn now listens on `0.0.0.0:8080` and advertises the machine's Tailscale IP.
- Added allowlisted Telegram bot with cinema selection, confirmation, run, and status actions.
- Configured `cinemas.json` for Lukoyanov (`100.118.9.80`) and Kinel (`100.78.107.67`).
- Added environment-based API tokens; `.env` remains local and ignored by Git.
- Added Kinel to the bot and verified both cinema agents can operate after their API servers are started.
- Added Windows firewall rules and disabled Tailscale shields-up on Kinel during diagnosis.
- Configured Tailscale Serve on Kinel, but direct port `8080` remains the bot's configured endpoint.
- Fixed Mac routing so both cinema IPs use the Tailscale interface while `VPSUS` stays connected for Telegram access.
- Telegram bot was running on the Mac as `telegram_bot.py` at handoff time.

### Next Steps
- Automate Mac host-route restoration after reboot or VPN reconnect; current `/32` routes are temporary.
- Configure both Windows API launchers as startup services/tasks so Uvicorn survives logout and reboot.
- Move the Telegram bot to an always-on VPS joined to the same tailnet, removing dependency on the Mac and `VPSUS`.
- Rotate the Telegram bot token because it was exposed in chat, then update the local `.env`.
- Retry `git pull` on Kinel after its intermittent Windows TLS/schannel errors are resolved.
- Remove the tracked generated `__pycache__/barco_open_chrome.cpython-314.pyc` in a dedicated cleanup change.

### Gotchas
- `VPSUS` is required for this Mac to reach Telegram, but it installs a conflicting route for `100.64.0.0/10`.
- Keep `VPSUS` and Tailscale enabled together, with explicit host routes for both cinema IPs through Tailscale.
- Tailscale ping can succeed even when normal TCP follows the wrong macOS route; verify with `route -n get <IP>`.
- A cinema showing `connection refused` usually means its Uvicorn window/server is no longer running.
- Never trigger `/run-schedule` merely to test connectivity; use `/health` and authenticated `/status`.
## 2026-10-01 — Three-cinema operations and handoff

### Done
- Added Novospasskoye (`100.107.6.2`, `DESKTOP-KCJPV2S`) to `cinemas.json`; pushed as `c28b8be` and configured its separate API token locally.
- Verified Novospasskoye `/health` and authenticated `/status` returned HTTP 200; status was `idle`.
- Added `docs/OPERATIONS.md` with startup, Excel, VPN, API, Telegram, log, and troubleshooting instructions; linked it from `README.md` and `AGENTS.md`.
- Recovered Lukoyanov (`100.118.9.80`) after its API stopped; verified Tailscale ping and authenticated `/status` returned HTTP 200 with `idle`.
- Confirmed local Telegram bot process `telegram_bot.py` was running at last check; no projector actions were sent during connection checks.

### Next Steps
- Commit and push the new operations guide, map update, and these notes; keep the unrelated tracked `__pycache__/barco_open_chrome.cpython-314.pyc` change untouched.
- Configure Windows API and Mac Telegram bot as persistent services if uptime without open terminals is needed.
- Automate Mac Tailscale host-route restoration after VPN reconnect; verify current `utun` number first.
- Rotate the exposed Telegram token and update only local secrets when practical.

### Gotchas
- `/health` and authorized `/status` are read-only checks. `idle` means no job since API restart, not verified projector state.
- An outgoing `tailscale ping` from Windows to Mac restored connectivity in Novospasskoye and Lukoyanov when Mac-to-Windows ping had timed out.
- The API server exits when its Windows PowerShell window receives Ctrl+C; a local `127.0.0.1:8080/health` check isolates this from Tailscale issues.
- Mac bot startup requires `.venv/bin/python`, because system `python3` lacks `httpx`.
- LoreBase vault exists at `/Users/fadin/vault` but is not a Git repository; sync notes there without a vault commit.

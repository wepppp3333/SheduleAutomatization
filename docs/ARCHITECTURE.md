# Architecture

## Purpose

The project automates creation and rescheduling of cinema shows in a Barco web scheduler. It converts a human-maintained Excel schedule into normalized JSON, then drives the Barco UI with Selenium.

## Runtime Environment

- Python 3 on Windows for production execution.
- Google Chrome controlled through Selenium.
- Selenium Manager is attempted first; `CHROMEDRIVER_PATH` and known local paths are fallbacks.
- The Barco scheduler is available only on the private network at `https://192.168.100.2:43744`.
- The local development machine may validate syntax and edit code but cannot verify the private UI end to end.

## Data Flow

1. Resolve the repository root from `Path(__file__).resolve().parent`.
2. Create `automation_artifacts/` and `automation_artifacts/screenshots/`.
3. Redirect stdout and stderr through `Tee` into `automation_artifacts/barco_automation.log`.
4. Find the first Excel file matching `Рассписание*` or `Расписание*` with an Excel extension.
5. Read the workbook with `pandas.read_excel(..., header=None)`.
6. Treat recognized `DD.MM.YYYY` values in column 1 as the current date.
7. Treat rows with a colon in column 1 as show rows and read the title from column 2.
8. Strip format/rating suffixes from titles and write normalized records to `automation_artifacts/schedule.json`.
9. Group JSON records by date.
10. Open Chrome, pass the certificate warning, log in, and navigate to `#sms/scheduler`.
11. Locate the scheduler column for each date and process each show in that column.
12. Add the title in an initial free timeline slot, find the created `rowItem`, open its move menu, and set the desired date/time.
13. Confirm the date/time modal and continue to the next show.

Normalized record shape:

```json
{
  "date": "09.03.2026",
  "time": "18:25",
  "title": "Красавица"
}
```

## Main Code Areas

### Paths and input discovery

Top-level constants define artifact paths. `find_excel_file()` searches only the repository root and raises `FileNotFoundError` when no matching workbook exists.

### Timeline helpers

- `click_top_slot()` clicks near the top of a `dayView` through JavaScript.
- `click_time_slot()` calculates a vertical position from `hourLine` spacing and dispatches a JavaScript click.
- `open_show_popover()` retries top-slot, time-slot, and placeholder clicks.
- `scroll_timeline_to_top()` resets the scheduler and window scroll positions.

### Title matching

- `normalize_title()` lowercases text and removes punctuation/duplicate whitespace.
- `titles_match()` performs containment and word-overlap checks.
- `title_similarity()` combines word overlap with `SequenceMatcher` similarity.
- `wait_for_show_block()` repeatedly re-queries `dayView` and `rowItem` to survive DOM redraws.

The active flow uses similarity scoring for dropdown selection and normalized title matching when locating scheduled rows.

### Move menu and modal helpers

- `open_menu_show()` retries hover/click behavior for `menuShow`.
- `click_move_to()` selects a visible, non-disabled `moveTo` element and has a JavaScript fallback.
- `click_visible_id()` clicks the visible element when duplicate IDs exist in hidden UI fragments.
- `clear_blocking_modal_backdrop()` and `close_datetime_modal()` recover from partially closed Bootstrap modals.

### Logging

`Tee` mirrors console output to the log file. `sys.excepthook` records unhandled tracebacks. Runtime logs are appended across runs and each run starts with a timestamp marker.

## Active Selenium Sequence

The active code currently follows this sequence for every show:

1. Find the target date column, paging backward or forward by week when needed.
2. Check whether the same title and rounded time already exist; skip duplicates.
3. Choose a free early `hourLine` and create a temporary show placeholder.
4. Open `.caretBtn`, score links under `#listOfShows`, and select the best title match.
5. Confirm the show in `#showPlaceHolderPopover`.
6. Find the temporary `.rowItem` by title and temporary start time.
7. Click `.moveRowBtn`, then visible `#menuShow`, then visible `#moveTo`.
8. Find the target day in `.datepicker-days`, rejecting `old`, `new`, and `notSelectable` cells.
9. Select the target hour, rounded three-minute value, and seconds `00`.
10. Confirm the timepicker, then click `#confirmDateTimeBtn`.
11. Verify that the title appears in the target column at the expected time.
12. Stop on the first error and save a traceback and screenshot to avoid cascading bad entries.

## Player Control Sequence

`barco_player_control.py shutdown-and-schedule` performs an idempotent shutdown
sequence through Barco's authenticated browser session. It calls the same
`SmsComm` commands used by the UI, avoiding dependency on version-specific DOM:

1. Open Player and read `g_MainStatusModel`.
2. If Scheduler mode is active, send `changeMode(0)` and confirm `playerMode == 0`.
3. If playback is active, send `stop` once and wait for Cleared or Stopped state.
4. Send `setDowser(true)` only when the dowser is currently open.
5. Turn the lamp off only when it is currently on.
6. Send `changeMode(1)` and confirm `playerMode == 1`.
7. Fail unless the final state has Scheduler enabled, lamp off, and dowser closed.

The FastAPI endpoint is `POST /player/shutdown-and-schedule`. It shares the same
single-job lock and `/status` response as schedule generation. Runtime details are
written to `automation_artifacts/barco_player_control.log`.

## Important Selectors

| Purpose | Selector |
| --- | --- |
| Date headers | `.dayHeader` and child `.date` |
| Date columns | `.dayView` |
| Timeline lines | `.hourLine` |
| Add-show dropdown | `.caretBtn`, `#listOfShows` |
| Add-show confirmation | `#showPlaceHolderPopover .ok` |
| Scheduled rows | `.rowItem` and child `.title` |
| Row move menu | `#menuShow`, `#moveTo` |
| Calendar | `.datepicker-days`, `.day` |
| Time picker | `.timepicker`, `.timepicker-hour`, `.timepicker-minute` |
| Final confirmation | `#confirmDateTimeBtn` |
| Player stop | `#btnStop` |
| Scheduler mode | `#btnScheduler` |
| Projector lamp | `#btnLamp` |
| Projector dowser | `#btnDowser` |

## Known Risks

- The file mixes reusable helpers, active procedural code, and a large commented legacy implementation. It is easy to edit the wrong flow.
- Many state transitions rely on fixed sleeps; rendering time varies across machines and runs.
- `dayHeader` and `dayView` ordering is assumed to match.
- Minute rounding changes requested times: for example, `18:25` becomes `18:24` with nearest-step rounding.
- The URL and credentials are embedded in source code.
- Browser cleanup is not protected by a top-level `try/finally`, so crashes may leave Chrome running.
- There are no automated tests or fixture pages for selector behavior.
- Runtime artifacts and bytecode are not consistently ignored by Git.
- The previous procedural implementation remains below an explicit `sys.exit` until the new flow passes a real projector test; it should then be deleted.

## Safe Change Strategy

1. Make one UI behavior change at a time.
2. Run `py -m py_compile barco_open_chrome.py` locally.
3. Pull the change on the Windows workstation.
4. Test with a small Excel file containing one or two shows.
5. Inspect `automation_artifacts/barco_automation.log` and screenshots before changing selectors again.
6. Re-query Selenium elements after any click that redraws the scheduler or opens/closes a modal.

## Recommended Next Refactor

- Move configuration and credentials to environment variables.
- Split Excel parsing, browser setup, and show scheduling into functions.
- Delete the commented legacy flow after the active implementation is stable.
- Replace fixed sleeps with explicit waits around the exact modal/dropdown states.
- Add `requirements.txt`, `.gitignore`, and focused tests for Excel parsing/title matching/minute rounding.

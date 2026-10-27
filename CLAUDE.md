# EHS Room Displays — context for Claude sessions

Classroom wall displays for Edwardsville High School (EHS), Edwardsville IL. Owner: Adam Koester,
CTE Department Chair (GitHub `arkoester`). Live site: https://arkoester.github.io/EHSMonitorControl/
Every screen in the building opens `room.html`, picks its room once, and then runs itself.

## Standing rules (from the owner — do not relitigate)

1. **Everything must scale to 100+ screens and ~130 teachers.** No per-screen or per-teacher setup
   beyond "pick your room once". Demo looks are fine; nothing else that won't scale.
2. **No filename conventions for teachers.** Scope & sequence files are matched by the course name
   *inside* the file (row 1), never by filename.
3. **Nothing on a PC, nobody logged in, no human in the daily loop.** The compile runs in GitHub
   Actions. A design that needs a computer left on, or Adam doing a daily step, is rejected.
4. **Announcements are published verbatim** from the school's page. Never summarize or reword them.
5. **Keep O365 connections minimal.** The Microsoft 365 connector needs district admin consent
   (confirmed unavailable 2026-10-27). SharePoint is only a drop folder for teachers.
6. **The public `data.json` never carries** section numbers, enrollment, teacher numbers, staff
   logins (userids), or file paths. Teacher names appear as "A. Koester".
7. **Publish Now means now**: a manual run must publish immediately (Actions > Run workflow).
8. **Teachers' lives easy** beats elegant engineering. Office/chair edits should work from a phone.
9. **Never go blank**: a failed course keeps yesterday's rows; a failed scrape keeps yesterday's
   announcements; a fatal registry error publishes nothing.
10. When asked for a bug/enhancement pass, keep running passes until a pass returns nothing.

## What runs where

| Piece | Where | Notes |
|---|---|---|
| `room.html` | every screen | engine + UI; per-screen memory in localStorage (`ehs_room`, `ehs_scale`); manual mode (⋮ menu) and "Bell today" are per-screen, ephemeral |
| `index.html` | hub | room list/search from `data.json` |
| `config.js` | site | only hand-edited site file |
| `compile.py` | GitHub Actions | reads `src/`, scrapes announcements, writes `data.json` + `version.txt` + `compile-report.txt` |
| `.github/workflows/publish.yml` | GitHub Actions | 10:15 UTC nightly; weekdays every 30 min 12:00–15:30 UTC and hourly 16–21 UTC; `workflow_dispatch` = Publish Now |
| `src/registry/EHS_room_display_data.xlsx` | repo | PUBLIC copy of the registry: Settings, Weeks, Bell, DayTypes, Rooms (no teacher numbers/notes/Teachers tab). Master copy lives in Adam's Drive: Claude Sandbox › items for future › EHS Displays |
| `src/sns/*.xlsx` | repo | teachers' scope & sequence workbooks (district template). Course name in row 1 |
| `src/day-types.txt` | repo | `DATE  DAY TYPE  note` lines; phone-editable half days / closures |
| `compile-report.txt` | repo | last run's issues; also `src/sns/_READ ME - problems found last night.txt` when sns files have problems |

Screens poll `version.txt` (tiny) and only fetch `data.json` when it changes. Polling schedule and
bandwidth math are in DISTRICT-PIPELINE.md ("Scale notes"). GitHub Pages caches ~10 minutes.

## Data model that room.html expects (`data.json`)

`generated` (ISO local time), `settings` {key: value}, `bell` [{day_type, period, label, start, end}],
`daytypes` [{date|weekday-name, day_type, note}], `rooms` [{room, period, course, course_title, teacher,
day_type}], `courses` {course: [{unit, starts, learning_target_s, success_criteria, current, notes}]},
`announcements` [{date, title, text}]. Exact-date daytypes override weekday rows; a period with no
course row → announcements. Day types without bell rows (e.g. `No School`) → announcements all day.

## Bell schedules (real EHS, 2026-27)

Regular (M/T/Th/F), Wednesday (early release, wired as a weekday row), Half Day (the ½-day
pep-assembly bell: 15-min periods, Pep Rally 9:21–10:15; used for homecoming/holiday assemblies and
quarter-end half days). First day 2026-08-19; quarter anchors in Settings; 39-week `Weeks` tab.

## Scope & sequence rules

Units = every row whose Chapter/Unit (B) cell has text (12-row template blocks or 2-row compact).
Time Frame (A) → start date: `Weeks N` via Weeks tab, with "month wins" when the teacher also wrote
a month; `M/D/YYYY`; `Q1–4`; `Semester`; month names (first/mid/late Monday). Units are kept
monotonic. C-cell lines starting "I can" → success criteria (promoted to targets if that's all there
is). Notes (H) starting `now`/`x`/`current` = pacing override. Matching: normalized exact → prefix →
word-subset (unique only); `Content source` column forces a path.

## How to work in this repo

- Test a compiler change with `pip install openpyxl && python3 compile.py --site . --no-fetch`
  (the sandbox can't reach the school site; Actions can). Check `compile-report.txt`.
- For `room.html`, load it against the real `data.json` in a headless browser and watch the console;
  `window.SignageEngine.decide(date, data, room)` is unit-testable.
- Don't hand-commit `data.json`/`version.txt`/`compile-report.txt`; the workflow regenerates them.
  Commit inputs (`src/`), code, docs.
- Registry edits: edit the master in Drive, then save the stripped copy into `src/registry/`
  (drop Teacher number + Notes columns and the Teachers tab).
- Debugging order when a wall is wrong: `compile-report.txt` → registry Rooms tab → the S&S file →
  the screen's ⋮ status line (data age, last check, overrides).

## Open items (as of 2026-10-27)

- **Teacher intake:** teachers drop S&S files in the department channel folder `Displays/sns`
  (SharePoint). Nothing automated can read it yet. Planned: Power Automate standard flow emails
  new/changed files to a dedicated Gmail drop box; a workflow step pulls attachments over IMAP into
  `src/sns/` (app password in repo secret `MAIL_PASSWORD`). Until then Adam uploads files by hand.
- Half-day bell for quarter-end days with no rally: add rows under a new day-type name when times
  are known; it appears in the screen's "Bell today" buttons automatically.
- Staff click-to-open icon; screen purchases (~50", under $500 each) once the pipeline is proven.
- Retire leftovers: the old PC tasks were deleted; `C:\ExamView\ehs-display` should be gone too.

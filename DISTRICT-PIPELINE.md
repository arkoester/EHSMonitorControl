# EHS Room Displays - District Pipeline (all-Microsoft)

The production path. Teachers edit locked `userid.xlsx` files in a shared
SharePoint library; a scheduled script compiles everything each night into one
`data.json` beside the website; every screen reads that file automatically.
No Google anywhere, no site edits ever, and nothing a teacher can break.

## The shape of it

```
SharePoint library "CTE Displays"            One always-on PC (4:15 AM)          The site
  EHS_room_display_data.xlsx  (office)  -->  Compile-Displays.ps1  -->  data.json --> every screen
  teachers\akoester.xlsx      (teacher)          |                        (GitHub Pages now,
  teachers\sbhooshan1.xlsx    (teacher)          +-> compile-report.txt    district server later)
  teachers\sbhooshan2.xlsx    (teacher)              (back into the library)
```

- The **registry** (`EHS_room_display_data.xlsx`) is office-owned: bell, day
  types, and the Rooms tab, which maps room + period -> course + **Teacher
  userid**. The userid is the join key AND the filename, so same-surname
  teachers can never be confused: `sbhooshan1.xlsx` and `sbhooshan2.xlsx`
  are different people by definition.
- Each **teacher file** has one tab per course. Only the yellow cells are
  editable; tabs cannot be renamed, added, or deleted (structure lock). It
  works in desktop Excel and Excel in the browser, autosaves, co-authors.
- The **compiler** validates before it publishes: missing files, renamed-or-
  missing tabs, unreadable dates, live units with blank targets, and courses
  with nothing active today all land in `compile-report.txt` in the library.
  A course that fails keeps **yesterday's data on the wall** - screens never
  go blank because of a bad edit.

## One-time setup

1. **Library.** Create the SharePoint document library, add the registry and
   a `teachers` folder. Everyone gets access to the library; each teacher
   only ever opens their own file (protection does the rest).
2. **Provision teacher files.** Fill `roster.csv` (courses copy-pasted from
   the registry's Course column - that exact match is the system), then run
   `Provision-Teachers.ps1` on any PC with Excel. It stamps one locked
   `userid.xlsx` per teacher; re-run any time for new hires (existing files
   are never touched).
3. **The compiler PC.** Any always-on district PC, logged into the account
   that will run the task:
   - Sync the library in OneDrive; right-click the folder -> **Always keep on
     this device** (files-on-demand placeholders break scripts).
   - `Install-Module ImportExcel -Scope CurrentUser` (reads xlsx without
     Excel - safe for headless scheduled runs; Excel itself is only needed
     for the one-time provisioning above).
   - `git clone` the site repo to `C:\ehs-display`, push once by hand so the
     credential manager stores the GitHub token.
   - Edit the CONFIG block at the top of `Compile-Displays.ps1`.
   - Schedule it **as that user, "run whether user is logged on or not"**:
     `schtasks /Create /TN "EHS Displays Nightly" /SC DAILY /ST 04:15 /TR "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ehs-display\Compile-Displays.ps1"`
     (Do not run it as SYSTEM - SYSTEM has no OneDrive sync.)
4. **First compile.** Double-click `Publish Now.bat`. Screens flip from
   sample data to live data within a minute of the push landing.

## Daily life

- **Teachers:** open your file from the library, type in yellow cells, close.
  It is on your wall the next school day. That is the whole job. Running
  ahead or behind? Put an `x` in Current (teacher file) or `now` in the
  unit's Notes cell (scope & sequence file) - the calendar yields to you
  until you clear it. Need the wall to show something else *right now*
  (assembly ran long, the bell is off)? Tap the corner &#8942; on the
  screen: Auto / Targets / Announcements / Video. Manual mode shows an
  orange MANUAL badge, wears off at the next bell or after an hour, and
  never survives into the next day.
- **Office:** bell changes, day types, room moves, and the broadcast banner
  live in the registry. They also ride the nightly compile - for a
  same-morning change (surprise late start at 6:45), edit the registry and
  double-click **Publish Now**; screens update within about a minute.
- **Daily announcements, hands-free:** the compiler scrapes the school's
  public Daily Announcements page (ehs.ecusd7.org/link-2) every night, and
  every idle screen rotates through the items - the page's own date line,
  each bold lead as the title, the rest as the body. Nobody maintains a
  slide deck. If the page can't be read or parses to zero items, screens
  keep yesterday's announcements and the report says so; a same-morning
  repost on the website rides Publish Now like everything else. The
  deck_url Slides embed remains as a fallback when the feed is empty.
- **Friday video announcements, hands-free too:** `video_url` points at
  EHS Broadcasting's uploads playlist, so the Friday video window
  automatically plays the club's newest morning show - no weekly link
  chasing, nothing to update. When the club starts broadcasting live,
  flip one pair of settings: clear `video_url`, put the channel ID
  (UC9JqMASFrPtxEgmjsFainRg) in `video_channel_id`, and the live stream
  takes over the window. Kiosk sound note: launch the browser with
  `--autoplay-policy=no-user-gesture-required` so the video plays with
  audio and no click.
- **You:** skim `compile-report.txt` in the library with your coffee. It
  names the file, the tab, and the unit for every issue.

## Scope & sequence as a source (no second document)

A course can skip the teacher file entirely and compile straight from a
district scope & sequence workbook: put the file's library path in the Rooms
tab's **Content source** column (e.g. `sns\AP_Macro_Scope_Sequence.xlsx`).
Rules, all validated nightly:

- One course per S&S file; the compiler reads the first sheet, so the tab
  name doesn't matter here.
- The first line of each Chapter/Unit cell becomes the wall title.
- In the Learning Target(s) cell, lines starting "I can" (or "PF: I can")
  become success criteria and the rest become targets; a cell that is ALL
  I-cans promotes them to the targets - that IS the unit's wall content.
- "Weeks N" time frames schedule against the registry's **Weeks** tab
  (school week -> Monday date, filled once a year by the office). This
  matters: plain calendar arithmetic put spring units 4-7 weeks early on
  the pilot file because school weeks pause for breaks. Real date ranges
  in Time Frame also work; anything unreadable is named in the report.

- **Pacing override:** running ahead or behind the plan? Type `now` (or
  `x`) at the start of that unit's Notes cell (column H) and the wall
  jumps to it - next school day via the nightly compile, or about a
  minute via Publish Now. Clear it to hand control back to the calendar.
  The report announces an active override ("showing Unit 2 while the
  calendar says Unit 1") and flags it if two units are marked at once.

Teachers who keep a good S&S never touch a second document; the wall is a
byproduct of the plan. The locked userid.xlsx remains for everyone else.

## Semester refresh: the Schedule Matrix is the registry's source

The SIS "Schedule Matrix Report" PDF is parsed into clean rows by
`parse_schedule_matrix.py` (kept in this bundle). Twice a year:

1. Export the new Schedule Matrix Report to PDF.
2. `pdftotext Schedule_Matrix_Report.pdf matrix.txt`
3. `python3 parse_schedule_matrix.py matrix.txt schedule_matrix_full.csv`
   (or hand the PDF to Claude, which runs the same script)
4. Regenerate the registry's Rooms tab from the semester's rows.

Course titles are canonicalized against the 2026-27 EHS Course Handbook by
course code (`handbook_names.csv`, joined on the section number's leading
digits), so SIS shorthand like "Consumer Education EB" or "Hon Business
Management" lands on the wall as the handbook's "Early Bird Consumer
Education" and "Honors Business Management" - and, since the course name is
the join key to teacher files and S&S docs, everything matches by the same
canonical name. `course_name_corrections.csv` logs every change each run.

The parser understands the report's quirks: wrapped course names split
across lines, each semester cell printing once per quarter band
(deduplicated), day-label debris inside the grid, PE's letter-coded
sections, club/caseload duty cells, LCCC cohort sections, and the CAVC/CEO
off-campus list blocks (skipped with a warning - they aren't wall rooms).
Every assignment carries teacher, teacher number, period, term, course,
section, room, and enrollment, so the same CSV also feeds
`roster_skeleton.csv` - add userids and it becomes the provisioning
roster. Co-taught periods (two courses, one room, same hour) are flagged
"shares period" in the Rooms tab; the screen shows the first row.

## Migrating to district hosting later

Designed as a two-line change on the compiler PC, invisible to teachers:
set `$PublishMode = 'copy'` and `$WebRootPath` to the district web path IT
gives you, and copy the site files (`index.html`, `room.html`, `config.js`)
there once. Screens get re-pointed to the district URL one time (or IT
redirects the old one). Nothing else in the pipeline changes.

## Troubleshooting

- **Compile ran but data is stale on a screen:** screens re-check data.json
  every ~60 s and hard-reload nightly; check the screen's clock and network.
- **`Import-Excel` not found:** the module install ran under a different
  user. Install under the task's account.
- **Git push fails at 4 AM:** the stored token expired; push once by hand.
- **A teacher "lost" a tab:** they can't - structure is locked. If a tab is
  missing, the file predates that course; re-run provisioning after adding
  the course to their roster line (or unprotect with the office password and
  copy a tab by hand).
- **Two teachers share a course name:** fine - courses are read per
  (userid, course), and each room row points at its own teacher's file.

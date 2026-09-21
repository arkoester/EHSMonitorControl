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

- The **registry** (`EHS_room_display_data.xlsx`) is office-owned. Its
  **Rooms** tab is *generated* from the Schedule Matrix (room + period ->
  course + teacher + teacher number) and can be regenerated every semester
  with nothing lost, because nothing hand-typed lives there. Its
  **Teachers** tab is the one hand-maintained sheet: teacher number, name,
  and **userid** - typed once, kept forever. The userid is the teacher file's
  name, so same-surname teachers can never be confused: `sbhooshan1.xlsx`
  and `sbhooshan2.xlsx` are different people by definition.
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
2. **Provision teacher files.** Fill the userid column of the registry's
   Teachers tab, then run `Provision-Teachers.ps1` on any PC with Excel.
   It reads the Teachers tab directly (no roster file) and stamps one locked
   `userid.xlsx` per teacher with a tab for every course from both
   semesters. Re-run any time: new teachers get a file, existing files get
   any missing course tab added, nothing else is touched.
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
   Then a second, morning schedule - because the office posts today's
   announcements *after* the 4:15 run, and "hands-free" only holds if the
   compiler looks again once they have. Every 30 minutes, 7:00-9:30:
     `schtasks /Create /TN "EHS Displays Morning" /SC DAILY /ST 07:00 /RI 30 /DU 02:30 /TR "powershell -NoProfile -ExecutionPolicy Bypass -File C:\ehs-display\Compile-Displays.ps1"`
   Runs take about a minute and only push when something changed, so the
   extra runs cost nothing. If the office posts at a fixed time, one run
   ten minutes after it is enough.
4. **First compile.** Double-click `Publish Now.bat`. Screens flip from
   sample data to live data within about ten minutes of the push landing
   (GitHub's CDN caches the file for up to ten minutes).

## Daily life

- **Teachers:** open your file from the library, type in yellow cells, close.
  It is on your wall the next school day. That is the whole job. Running
  ahead or behind? Put an `x` in Current (teacher file) or `now` in the
  unit's Notes cell (scope & sequence file) - the calendar yields to you
  until you clear it. Need the wall to show something else *right now*
  (assembly ran long, the bell is off)? Tap the corner &#8942; on the
  screen: Auto / Targets / Announcements / Video. Choosing Targets also
  offers **Hour** and **Unit** pickers, so you can put any period's
  learning targets - and any unit, in or out of calendar order - on the
  wall at any time of day, including after school with no class running.
  **Text size** (A-/A+) is there too and is saved on that screen for good.
  Manual mode shows an orange MANUAL badge, wears off at the next bell or
  after an hour, and never survives into the next day.
- **Office:** bell changes, day types, room moves, and the broadcast banner
  live in the registry. They also ride the nightly compile - for a
  same-morning change (surprise late start at 6:45), edit the registry and
  double-click **Publish Now**; screens update within about ten minutes
  (the CDN cache window - see Scale notes).
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
district scope & sequence workbook. **Filenames do not matter.** Every
`.xlsx` in the library's `sns\` folder is opened and asked which course it
is: the compiler reads the course name from row 1 of the first sheet (the
cell beside the template's "Course Name:" label) and matches it to the
Course column in the Rooms tab. The match is forgiving: parentheticals and
dash suffixes are ignored, so "AP Macroeconomics (LCCC ECON 151) - Full
Year" and "AP Macro" both resolve to AP Macroeconomics. Teachers drop their
file in the folder however they named it; nothing else is required of them.

If the compiler finds a problem with an sns file, it also writes
`sns\_READ ME - problems found last night.txt` right there in the folder,
naming the file and the fix, so teachers can self-serve without waiting
for the chair to relay the report.

Resolution order per course, every night:
1. an explicit path in the Rooms tab's **Content source** column (override),
2. an `sns\` file whose internal course name matches,
3. `teachers\<userid>.xlsx` with a tab named for the course.

The report names any `sns\` file whose course name matches nothing in the
registry (a one-cell fix), any file it cannot open, and any duplicate where
two files claim the same course.

Rules, all validated nightly:

- One course per S&S file; the compiler reads the first sheet, so the tab
  name doesn't matter here.
- The first line of each Chapter/Unit cell becomes the wall title.
- In the Learning Target(s) cell, lines starting "I can" (or "PF: I can")
  become success criteria and the rest become targets; a cell that is ALL
  I-cans promotes them to the targets - that IS the unit's wall content.
- Time Frame accepts, in order of preference: **Weeks N** (against the
  registry's Weeks tab, which is what keeps spring units from drifting),
  a real **date** or date range, **Q1-Q4** / Quarter N, **Semester 1/2**,
  or a **month name** ("late Oct", "Nov-Dec" -> the first Monday of that
  month). Anything else is named in the report and in the sns folder note.

- **Pacing override:** running ahead or behind the plan? Type `now` (or
  `x`) at the start of that unit's Notes cell (column H) and the wall
  jumps to it - next school day via the nightly compile, or about a
  minute via Publish Now. Clear it to hand control back to the calendar.
  The report announces an active override ("showing Unit 2 while the
  calendar says Unit 1") and flags it if two units are marked at once.

Teachers who keep a good S&S never touch a second document; the wall is a
byproduct of the plan. The locked userid.xlsx remains for everyone else.

## The district calendar is already loaded

From ecusd7.org/district/calendar (2026-27): first student day **Aug 19,
2026** (half day), last **May 25, 2027**, early release every Wednesday.
The registry ships with all 29 no-school dates as DayTypes rows (institutes,
Labor Day, conferences, Thanksgiving, holiday break, MLK, Presidents Day,
spring break, Memorial Day) and a **Weeks** tab of all 39 school weeks -
week number to Monday date, skipping the two full break weeks. That tab is
what makes a teacher's "Weeks 20" land in January instead of drifting.
Not yet loaded: the three half days (Aug 19, Oct 23, May 25) run the normal
bell until someone publishes half-day bell times; add them as a Half Day
day type plus its Bell rows.

## Semester refresh: the Schedule Matrix is the registry's source

The SIS "Schedule Matrix Report" PDF is parsed into clean rows by
`parse_schedule_matrix.py` (kept in this bundle). Twice a year:

1. Export the new Schedule Matrix Report to PDF.
2. `pdftotext Schedule_Matrix_Report.pdf matrix.txt`
3. `python3 parse_schedule_matrix.py matrix.txt schedule_matrix_full.csv`
   (or hand the PDF to Claude, which runs the same script)
4. Regenerate the registry's Rooms tab from the semester's rows. This is
   safe: userids live in the Teachers tab and are untouched.

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
section, room, and enrollment, so the same CSV also rebuilds the
registry's Teachers tab (both semesters' courses per teacher). Co-taught periods (two courses, one room, same hour) are flagged
"shares period" in the Rooms tab; the screen shows the first row.

## What the public site does and does not publish

`data.json` is served from a public web page, so the compiler publishes only
what a screen renders: room, period, course, course title, teacher (as
"A. Koester"), day type. Everything else stays in the registry:

| Dropped from data.json | Why | Where it still lives |
|---|---|---|
| `notes` (section no., enrolled/cap, "shares period") | enrollment counts on a public URL | Rooms tab |
| `teacher_userid` | a staff login name | Rooms tab |
| `source_sheet_id`, `content_source` | internal file wiring | Rooms tab |

**DEBUGGING NOTE - read this before chasing a wrong course on a wall.**
Section numbers and enrollment are intentionally NOT in the published file.
When a screen shows something unexpected, look in this order:
1. `compile-report.txt` in the library - names the file, tab, and unit for
   every problem from last night's run.
2. The registry's **Rooms** tab - full detail: `sect 11103.1 - 32/34`,
   `shares period`, the teacher userid, and any Content source path.
3. `schedule_matrix_full.csv` - the raw SIS export the Rooms tab was built
   from, including both semesters and the CAVC/LCCC rows.
Do not "fix" this by republishing the dropped fields; the registry is the
debugging surface by design.

Other exposure notes: the repo is public, so room/period/course and teacher
names are readable by anyone with the URL - catalog-level information, but
worth knowing. The registry's `video_url` and `deck_url` are single cells
that control every screen in the building: keep the registry root
office-only and share `sns\` for editing, not the library root. The
compiler PC holds a GitHub token that can rewrite the site - lock the
screen, and scope the token to this one repo.

## Scale notes: what 100+ screens actually do

Worked out before the fleet exists, because these are the failures that
only show up at scale.

**Bandwidth, and why Publish Now means now.** Screens do a two-step check.
Every poll fetches `version.txt` - a few bytes, cache-busted so GitHub's CDN
can never hide a new push. Only when that value changes does a screen
download `data.json` (keyed by version, so it is fresh too). Result: an
unchanged day costs the whole building a few megabytes, and a **Publish Now
reaches every awake screen within one poll interval plus GitHub's ~1-minute
build** - about 90 seconds on a weekday morning. Verified: a change landed
on a test wall 6 seconds after it was published, with exactly one download.

**The polling schedule** (settings `refresh_seconds` and
`refresh_seconds_day`, defaults 60 and 120):

| When | Screens check for new data |
|---|---|
| Weekdays 6:00 AM - end of 2nd hour (9:30) | every `refresh_seconds` (60 s) - nearly all updates land here |
| Weekdays 9:30 AM - 5:00 PM | every `refresh_seconds_day` (120 s) |
| Weekday evenings and overnight | asleep until 6:00 AM |
| Weekends 7:00 AM - 5:00 PM | every two hours |
| Weekends otherwise | asleep until 7:00 AM (6:00 if Monday is next) |

The nightly 24-hour page reload still happens on top of this and picks up
site changes. A screen that is asleep is still displaying; it is only not
asking for updates.

**A network blip must never blank a wall.** Fixed a real bug: a single
failed refresh used to swap the screen to the built-in sample data (green
dot and all). Now a screen that has ever loaded real data keeps it through
any outage - the corner dot goes amber after three missed refreshes and
the content stays put. Sample data only ever appears on a screen that has
never seen a `data.json`.

**Clocks.** The page trusts the device clock, and across 100 devices some
will be wrong. Each refresh now reads the server's time from the response
and corrects the page's clock automatically; the health note names skews
over five minutes so someone fixes the device. A device set to the wrong
**time zone** gets a visible note too - periods would be an hour off and
nothing else would look wrong.

**The compiler dies silently.** If the nightly run stops, screens would
show last week's targets forever with a green dot. Now any screen whose
data is more than 36 hours old shows "Wall data is N days old - the nightly
compile has not run." That plus `compile-report.txt`'s timestamp is your
proof-of-life.

**The health note** is the small amber pill bottom-left. Priority order:
wrong time zone > clock skew > stale data > sample data. It is meant to be
read from across the room by whoever walks in.

**Friday at 7:15.** Every room in the video window starts a YouTube stream
at the same moment. The building already streams the Friday show this way
today, so this is a known, accepted load - noted here only so nobody
rediscovers it as a surprise.

**Rolling out site changes.** Screens load `room.html` once and reload
nightly (staggered up to 30 minutes), so a new version of the page reaches
the building over one night, not in a thundering herd. `data.json` changes
propagate on the refresh interval; site changes propagate overnight.

**Device requirements, stated once:** Central time zone, an accurate
clock (both now self-reported on screen), a browser that keeps
localStorage between boots, and no sleep timer on the display.

## If a file gets overwritten or deleted

SharePoint has this covered; nothing extra was built. Overwritten file:
open it in SharePoint, **... > Version history**, restore. Deleted file:
the library's **Recycle bin**, then the site collection recycle bin (about
93 days in total). The wall keeps showing the last good compile meanwhile.

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

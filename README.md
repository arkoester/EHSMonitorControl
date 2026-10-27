# EHS Room Displays

Classroom wall displays for Edwardsville High School. One web page serves
every room: the Friday morning show in its window, each period's learning
targets and success criteria at the bell, and the school's daily
announcements the rest of the time. Screens remember their own room, follow
the real bell (including early-release Wednesdays and the district
calendar), and correct their own clocks.

**Hub:** `index.html` - every room, searchable, grouped by wing.
**Screen:** `room.html` - the page every kiosk opens; first open asks which room.

## What lives here

The public site (`index.html`, `room.html`, `config.js`), the compiled
`data.json` + `version.txt` the screens read, and the compiler itself:
`compile.py`, its inputs under `src/` (the office registry and the
teachers' scope & sequence workbooks), and the GitHub Actions workflow
that runs it on a schedule. The published data is deliberately slim -
room, period, course, teacher - with section numbers, enrollment, teacher
numbers and staff logins kept in the department's full registry, not here.

## Where the content comes from

Teachers' scope & sequence workbooks sit in `src/sns/` (any filename -
the course name inside the file is what counts). GitHub Actions runs
`compile.py` every morning and through the school day: it validates,
compiles `data.json`, and commits it - no PC, no login, nothing to keep
running. **Publish Now** = Actions tab > Publish displays > Run workflow.
The office maintains bell times, day types, and the room map in the
registry workbook under `src/registry/`. Daily announcements are read
from the school website; the Friday show plays straight from the EHS
Broadcasting channel.

Setup, daily operation, the polling schedule, debugging order, and the
security notes are all in **DISTRICT-PIPELINE.md**.

## Screen controls

Tap the faint corner button on any screen: Auto / Targets / Announcements /
Video, an hour and unit picker, and text size. Manual mode shows an orange
badge and wears off at the next bell.

## Demo

Open the hub and press **Demo tour** to walk a full school day on sample
data with the clock-jump buttons.

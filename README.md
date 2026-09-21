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

Only the public site: `index.html`, `room.html`, `config.js`, and the
compiled `data.json` + `version.txt` the nightly compiler pushes. Nothing
here is edited by hand after setup. The published data is deliberately
slim - room, period, course, teacher - with section numbers, enrollment,
and staff logins kept in the district's SharePoint registry, not on the
public internet.

## Where the content comes from

Teachers keep their scope & sequence workbooks (or a locked per-teacher
file) in a SharePoint library. A scheduled script on a district PC reads
them every night, validates, compiles `data.json`, and pushes it here.
The office maintains bell times, day types, and the room map in one
registry workbook in the same library. Daily announcements are read from
the school website; the Friday show plays straight from the EHS
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

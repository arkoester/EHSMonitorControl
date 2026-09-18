# EHS Room Display — setup guide (CTE pilot, 15 rooms)

Two files do the work:

- `room.html` — the page every screen shows. One copy serves every room; the room is chosen by URL: `room.html?room=A124`.
- `EHS_room_display_data.xlsx` — the data workbook. Import it into Google Sheets once; after that, everything is edited in the Sheet.

The page decides what to show every second: the live announcements stream during the first-hour window, the period's learning targets and success criteria at the bell, and the announcements deck the rest of the time. Teacher edits in the Sheet reach the wall within about a minute.

## 1. Try it right now (no setup)

Open `room.html?room=A124&demo=1` in any browser. The demo strip at the top jumps the clock to each state. Built-in sample data is used until a sheet ID is added.

Open `index.html` for the hub: one tile per room plus the demo tour. Any `?sheet=` you put on the hub URL is passed along to every room link.

## 2. Data

**District mode (production, all-Microsoft):** teachers edit locked `userid.xlsx` files in a SharePoint library; a nightly script compiles everything to a `data.json` beside this site, and every page finds it automatically - nothing to configure anywhere. Full setup and daily life: **DISTRICT-PIPELINE.md**. Until the first compile runs, screens quietly show the built-in sample data.

### Google Sheet mode (demo / dev, 10 minutes)

1. In Google Drive: New > File upload > `EHS_room_display_data.xlsx`, then open it and choose File > Save as Google Sheets. Keep the tab names.
2. Share > General access > Anyone with the link > Viewer. The screens are not signed in, so they need this. If district policy blocks link sharing, use the Apps Script option in section 6 instead.
3. Copy the sheet ID from the URL: the long string between `/d/` and `/edit`.
4. Open `room.html` in a text editor, find `CONFIG` near the bottom, and paste the ID into `sheetId`.

The real EHS bell is loaded for both day patterns - Regular (Mon, Tue, Thu, Fri) and Wednesday, hours 0-6 - with lunch waves treated as one long 4th hour (5th on Wednesdays). Before the first live test, fix the remaining placeholders: the Rooms tab (A124 carries placeholder hour assignments; A116, A122, A123, A125, and A210 are scaffolded blank - fill in each schedule) and each course tab's unit start dates and wording.

**Go live without touching GitHub:** append `&sheet=<yourSheetID>` to any room or hub URL - the page reads the override at load, so the kiosk bookmark carries the data source. (Baking the ID into CONFIG inside room.html also works; the URL wins if both are set.) If the sheet is ever unreachable, screens fall back to the built-in sample data and show a red status dot instead of going blank.

## 3. Hosting the page (IT)

Upload `index.html`, `room.html`, and `config.js` together - the hub is the folder's front page (e.g. `https://<user>.github.io/ehs-display/`), every tile opens `room.html` beside it, and `config.js` is the only file you ever edit.

Going live is one edit: open `config.js` and paste the Google Sheet ID between the quotes (share the Sheet as Anyone with the link - Viewer first). No URL parameters, no other file changes. `?sheet=<ID>` on any URL still works as a temporary override for testing before you commit the ID.

Rolling out to many screens: give every kiosk, screensaver, and browser homepage the same URL - `room.html`, no parameters. On first open the screen asks "Which room is this screen?"; one tap and that device remembers its room from then on. `?room=` in a URL previews another room without changing a screen's memory, and the small corner &#8942; button re-opens the picker if a screen ever moves rooms. The hub's room list and teacher names come straight from the Rooms tab, so adding a room to the Sheet adds it everywhere.

Put `room.html` anywhere the screens can reach over HTTPS: GitHub Pages, the district web server, or an internal web share. The content filter must allow that host, `docs.google.com` (sheet reads and the Slides embed), and `youtube.com` for the players' device group.

Suggested URL shape: `https://<host>/display/room.html?room=A124`

## 4. Panel side (ActivPanel 7)

The panel is the secondary surface, so the temp build can start with no app at all: bookmark `room.html` (no parameters) in the AP7's built-in browser and open it whenever the panel is on its Android side - the first open asks which room the screen is, then it remembers. The page still switches between video, targets, and deck by itself; an app only automates when it appears.

If and when the APK route opens up, in order of least hassle:

1. Ask IT whether the panels are enrolled in Promethean Panel Management or any MDM. If so, the Fully Kiosk APK can be pushed from the console to all six panels at once, no USB.
2. A one-time USB sideload needs "install unknown apps" enabled in the panel's Android settings - usually an IT-permission question rather than a technical one.
3. Skip the panel app for the pilot entirely and let the workstation carry everything.

### Fully Kiosk settings (whichever install route lands)

1. Install Fully Kiosk Browser from Google Play, or sideload the APK if the panel has no Google services. Buy the PLUS license per panel.
2. Settings > Web Content > Start URL: the room's page URL.
3. Settings > Device Management > Launch on boot: on. Settings > Screensaver (PLUS): Screensaver URL = the same page URL; Screensaver timer = 120 seconds.
4. Settings > Web Content > Autoplay videos: on (needed for stream audio).
5. Do not turn on "single app mode" or "return to app after idle": that would pull the panel away from a teacher's HDMI source. Signage only needs to own the Android home screen.

## 5. Workstation side (Windows screensaver)

The workstation is the live HDMI source all day, so the page runs there as a screensaver: idle shows the page, the teacher moves the mouse and their desktop is back.

Zero-install interim: make `room.html` (no parameters) the workstation browser's homepage and press F11 for full screen before the bell or whenever the room is idle. The first open asks which room the screen is, then it remembers. Nothing to deploy; works today.

For hands-off automation, two routes, IT's choice:

- WebView2 screensaver (recommended): a small `.scr` that opens the room URL full screen and exits on input. Deploy through Group Policy like any screensaver, set the timeout to match the passing period (5 minutes), and keep "On resume, display logon screen" as policy requires. This build is the next deliverable.
- Xibo for Windows: free, installs `Xibo.scr`, and can show a single web page layout pointed at the room URL. Needs a self-hosted Xibo CMS on a district VM.

PowerPoint in slideshow mode suppresses the screensaver, so slides left on screen are not covered.

## 6. Office: deck, stream, broadcast

- Deck: build the announcements in Google Slides. File > Share > Publish to web > Embed, copy the `src` URL, paste it into Settings > `deck_url`. The page sets auto-advance and loop itself.
- Stream: paste the school YouTube channel ID into Settings > `video_channel_id`. The page embeds whatever that channel is streaming live during the window in `video_start`–`video_end`. If nothing is live, the embed shows YouTube's waiting screen.
- Broadcast: type a message into Settings > `broadcast`; every idle screen shows it within a minute. Clear the cell to remove it. `broadcast_style` = `urgent` makes it full screen. This is for operations, not life safety.

### Apps Script alternative (if link sharing is blocked) and heartbeat

Extensions > Apps Script in the Sheet, paste this, then Deploy > New deployment > Web app, Execute as Me, Access: Anyone. Put the web app URL into `CONFIG.jsonUrl` (data) and into Settings > `heartbeat_url` (status).

```javascript
function doGet(e) {
  var ss = SpreadsheetApp.getActive();
  var p = e.parameter || {};
  if (p.mode) {                      // heartbeat from a screen
    var lock = LockService.getScriptLock();          // fleet-safe: serialize Status writes
    try { lock.waitLock(5000); } catch (err) { return ContentService.createTextOutput('busy'); }
    try {
      var st = ss.getSheetByName('Status') || ss.insertSheet('Status');
      if (st.getLastRow() === 0) st.appendRow(['Room', 'Last seen', 'Mode', 'Period', 'Course', 'Reason']);
      var rows = st.getDataRange().getValues(), r = -1;
      for (var i = 1; i < rows.length; i++) if (rows[i][0] === p.room) r = i + 1;
      var row = [p.room, new Date(), p.mode, p.period || '', p.course || '', p.reason || ''];
      if (r > 0) st.getRange(r, 1, 1, row.length).setValues([row]); else st.appendRow(row);
      return ContentService.createTextOutput('ok');
    } finally { lock.releaseLock(); }
  }
  function tab(name) {               // sheet -> array of objects
    var sh = ss.getSheetByName(name); if (!sh) return [];
    var v = sh.getDataRange().getDisplayValues(), start = 0;
    for (var i = 0; i < v.length; i++) if (v[i].filter(String).length > 2) { start = i; break; }
    var h = v[start].map(function (x) { return String(x).trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, ''); });
    return v.slice(start + 1).filter(function (r) { return r.join('').trim(); }).map(function (r) { var o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o; });
  }
  var settings = {}; tab('Settings').forEach(function (r) { settings[r.key] = r.value; });
  var rooms = tab('Rooms'), courses = {};
  rooms.forEach(function (r) { if (r.course && String(r.room).toLowerCase() === String(p.room || '').toLowerCase()) courses[r.course] = tab(r.course); });
  var out = { settings: settings, bell: tab('Bell'), daytypes: tab('DayTypes'), rooms: rooms, courses: courses };
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
```

The `Status` tab then shows each room's last check-in, mode, and any reason a screen fell back to the deck (for example "no current unit for Networking").

## 7. Teacher workflow

- Your course tab is the only thing you touch. One row per unit; one line per bullet.
- Wording changes appear on the wall within a minute.
- To move to the next unit, either set its Starts date or put an `x` in Current. Clear the `x` when you want dates to take over again.
- Nothing shows during instruction. The card appears at the bell for `targets_minutes` (default 8), then the deck returns; a teacher who wants it up longer can open the same URL on the desktop.

## 8. Day types: Wednesdays, late starts, no school

Two bell patterns ship loaded: Regular (Mon, Tue, Thu, Fri) and Wednesday. The Wednesday pattern repeats by itself because DayTypes carries a row with `Wednesday` in the Date column - any weekday name there applies every week. An exact date beats a weekday row, so a one-off schedule (late start, finals, assembly) is one dated DayTypes row plus its Bell rows, added any time before that morning. A dated day type with no Bell rows (like the Labor Day example) means announcements only, all day. The video window is fixed by clock time; if the stream moves on a special day, change `video_start` and `video_end` that morning or clear them.

## Built to scale

The pilot and the whole building run the same files; scale is a data question, not a rebuild.

- **One registry.** The Rooms tab is the only list of rooms: the hub, every screen's picker, and every schedule derive from it. Adding a room to the Sheet adds it everywhere.
- **Teacher-owned content (federation).** A course's LT/SC tab can live in the teacher's own spreadsheet: they share it Anyone with the link - Viewer, name the tab exactly the course name, and the office puts that spreadsheet's ID in the Rooms tab's `Source sheet ID` column for that course's rows. The central sheet stays office-owned (bell, day types, room map); content stays teacher-owned. Blank source = the central sheet, so the pilot needs nothing extra.
- **Fleet data path.** Direct Sheet reads are fine for a wing. For 100+ screens, switch `config.js` to the Apps Script `jsonUrl`: one request per screen per refresh instead of six, and it works even where district policy blocks link-sharing. (To federate through Apps Script later, swap `SpreadsheetApp.getActive()` for `openById` per course - one line.)
- **No thundering herd.** Screens jitter each refresh by up to 15 seconds and their nightly reload by up to 30 minutes, so a building never hits the Sheet or the network in lockstep. For a full fleet, set `refresh_seconds` to 120-300.
- **Per-screen identity without device management.** Screens remember their room locally; nothing central tracks devices. In Fully Kiosk, leave "clear cache/storage on start" OFF so the memory survives reboots - or skip memory entirely by giving each device a start URL with `?room=` from your MDM.
- **Heartbeat at scale.** The Apps Script Status write is lock-protected, so a hundred screens checking in at once won't collide.

## Known limits of the pilot build

- Each screen reads five small tabs about once a minute (settings, bell, day types, rooms, and its courses). Fifteen rooms is trivial; at building scale, raise `refresh_seconds` to 120.
- The page needs a browser with `fetch` and CSS grid (Chrome 57 or newer). An AP7 whose WebView cannot update may need the Xibo or Fully route tested first.
- No CAP emergency alerting. The broadcast cell is an operations tool.

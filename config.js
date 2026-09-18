/* ============================================================
   EHS Room Displays - one-time site configuration.
   This is the only file you ever edit to go live.

   1. Import EHS_room_display_data.xlsx into Google Sheets.
   2. Share the Sheet: Anyone with the link - Viewer.
   3. Paste the Sheet ID below (the long string in its URL
      between /d/ and /edit) and re-upload this file.

   Leave sheetId blank to run on the built-in sample data.
   ============================================================ */
window.EHS_CONFIG = {
  /* District (all-Microsoft) mode needs NOTHING set here: the nightly
     compiler writes data.json next to the site and every page finds it
     automatically. The settings below are for the Google demo/dev mode. */
  sheetId: "",

  /* Optional: Apps Script web app URL (see README section 6).
     Use this instead of sheetId if district policy blocks
     link-sharing. Overrides sheetId when set. */
  jsonUrl: "",

  /* Optional: force every screen to one room. Kiosk fleets
     leave this blank - each screen picks and remembers its
     own room on first open. */
  defaultRoom: ""
};

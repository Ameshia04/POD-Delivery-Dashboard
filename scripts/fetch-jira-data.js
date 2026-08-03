/**
 * team-pulse-apps-script.gs
 *
 * Reference script for the Delivery Dashboard's Team Pulse card. This is NOT
 * run by Node or the GitHub Action -- it runs inside Google Apps Script,
 * bound to a Google Sheet, and is what the dashboard's Team Pulse card
 * (index.html) talks to over plain HTTP.
 *
 * WHY BOTH READ AND SAVE GO THROUGH doGet: Google Apps Script Web Apps
 * respond to requests at script.google.com with a redirect to
 * script.googleusercontent.com. When a browser's fetch() follows a redirect
 * after a POST, it drops the POST body -- so doPost never actually receives
 * the data it's expecting, and fails. Routing "save" through a GET request
 * with the data as query parameters (?action=save&rating=...&notes=...)
 * sidesteps this entirely, since a GET has no body to lose. There's no
 * doPost function at all here anymore.
 *
 * ONE-TIME SETUP:
 *   1. Create a new Google Sheet (e.g. "Team Pulse Log"). On its first tab,
 *      add this exact header row:
 *        Date | Time | Project | Rating | Notes | Submitted By
 *   2. In that Sheet: Extensions > Apps Script. Delete the placeholder
 *      "myFunction" code and paste everything below it in this file.
 *   3. If your tab isn't named "Sheet1", change SHEET_NAME below to match.
 *   4. Deploy > New deployment > gear icon > select type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Click Deploy, authorize the script when Google prompts you, then
 *      copy the "Web app URL" it gives you (it ends in /exec).
 *   5. Paste that URL into index.html as TEAM_PULSE_APPS_SCRIPT_URL, near
 *      the top of the <script> block (search the file for that name). Do
 *      this in BOTH index.html copies you maintain (the top-level one and
 *      whichever mirror folder you're syncing right now).
 *   6. Test it: open the dashboard, pick a project (or leave "All
 *      Projects"), fill in the Team Pulse card, and click Save. A new row
 *      should appear in the Sheet within a couple seconds.
 *
 * IF YOU EDIT THIS SCRIPT LATER (including this fix, if you're pasting this
 * over an older version): editing the code alone does not update the live
 * URL's behavior. Go to Deploy > Manage deployments > edit (pencil icon) >
 * Version: "New version" > Deploy. The /exec URL itself does not change, so
 * you do not need to update TEAM_PULSE_APPS_SCRIPT_URL in index.html again.
 *
 * SECURITY NOTE: "Who has access: Anyone" means anyone who has this exact
 * URL can call it -- there's no login check on top of it. That's an
 * acceptable tradeoff for a low-stakes internal signal like this one, but
 * don't reuse this pattern for anything sensitive. The URL itself is long
 * and effectively unguessable, so this is "security through obscurity," not
 * real authentication.
 */

const SHEET_NAME = 'Sheet1'; // change to match your tab's actual name

/**
 * GET /exec?project=INV  ->  the most recent Team Pulse entry for INV.
 * GET /exec?project=All  ->  the most recent org-wide entry (or omit the
 * query param entirely for the same result). Falls back sensibly: most
 * recent match for the requested project, else most recent "All" entry,
 * else the most recent entry of any kind -- so the dashboard always shows
 * *something* if the exact project hasn't been logged yet.
 *
 * GET /exec?action=save&project=INV&rating=green&notes=...&submittedBy=...
 * -> appends a new row instead (see saveEntry() below and the header
 * comment for why this is a GET, not a POST).
 */
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (e.parameter.action === 'save') {
    return saveEntry(e, sheet);
  }

  const rows = sheet.getDataRange().getValues();
  const requestedProject = (e.parameter.project || 'All');

  let match = null, allMatch = null, anyMatch = null;
  for (let i = rows.length - 1; i >= 1; i--) { // skip header row (index 0)
    const row = rows[i];
    if (!anyMatch) anyMatch = row;
    if (!allMatch && row[2] === 'All') allMatch = row;
    if (!match && row[2] === requestedProject) match = row;
  }
  const chosen = match || allMatch || anyMatch;
  if (!chosen) return respond({ rating: null });

  return respond({
    date: chosen[0],
    time: chosen[1],
    project: chosen[2],
    rating: chosen[3],
    notes: chosen[4],
    submittedBy: chosen[5],
    savedAt: `${chosen[0]} ${chosen[1]}`,
  });
}

/** Appends one new row with a fresh server-side timestamp and echoes back
 * what was saved (so the dashboard can render it immediately). Called from
 * doGet when ?action=save is present -- see the header comment for why
 * saving is a GET with query params instead of a POST with a body. */
function saveEntry(e, sheet) {
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const date = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const time = Utilities.formatDate(now, tz, 'HH:mm:ss');
  const project = e.parameter.project || 'All';
  const rating = e.parameter.rating || '';
  const notes = e.parameter.notes || '';
  const submittedBy = e.parameter.submittedBy || '';

  sheet.appendRow([date, time, project, rating, notes, submittedBy]);

  return respond({ date, time, project, rating, notes, submittedBy, savedAt: `${date} ${time}` });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

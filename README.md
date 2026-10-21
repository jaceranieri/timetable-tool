# timetable-tool
A timetable helper wizard tool that assists Primary schools creating their timetable

## Layout

| Folder | Apps Script project | What it does |
|---|---|---|
| `wizard/` | The standalone web app project | `Index.html` is the setup wizard; `Code.gs` copies the master template and builds the class and specialist tabs |
| `template/` | The project bound to the master template spreadsheet | `TemplateSync.gs` adds the **Timetable Admin** menu and does the sync, lock and clash formatting; `TimetableTools.gs` adds the checks, copy helpers and Google Docs. Both are copied into every generated timetable along with the spreadsheet. |

`PLAN.md` lists the planned fixes and features in priority order, and
`template/SHEET_CHANGES.md` the edits made in the template spreadsheet itself.

`wizard/LoadTest.gs` is a development-only load test; don't include it in the
live wizard project.

## Running the tests

The `tests/` folder runs both scripts against an in-memory stand-in for
Google Sheets, so changes can be checked before they're pasted into Apps
Script. It needs Node.js 18 or later and no other packages:

```
npm test
```

`tests/fixtures/Code.before-phase4.gs` is the previous generator, kept so the
tests can check the new one produces the same timetable with far fewer
spreadsheet calls.

## Deploying the wizard

Deploy → Manage deployments:

- **Execute as:** User accessing the web app
- **Who has access:** Anyone within Sydney Catholic Schools

In the editor, under **Services**, add the **Google Sheets API** (identifier
`Sheets`). The wizard uses it to create all the copied tabs in a few
requests, which is much faster for large schools; without it, tabs are
copied one at a time.

Each user needs view access to the master template, so share it with the
domain as **Viewer**.

### What users see the first time

Because the wizard runs as each user, Google asks them to allow it to use
their Google Drive and Sheets the first time they open the wizard's link
(and again if the wizard ever needs a new permission). This happens before
the wizard loads, so mention it in the message that shares the link, for
example: *"The first time you open it, Google will ask you to allow the
timetable tool to access your Drive and Sheets. This is needed to create
your timetable in your own Drive."*

Generated timetables are saved in a **Timetables** folder in the user's My
Drive (created the first time), named
`<school name> - Timetable - <date>`.

Unfinished setups are saved for each user (Apps Script User Properties) and
offered back the next time they open the wizard. A setup can also be saved
to a file and loaded later, for example to start next year's timetable; the
file holds the same name/value list the timetable keeps in its DATA tab
(`RawData_JSON`).

The `appsscript.json` manifests in this repo reflect those settings. Check
them against the ones in the Apps Script editor (Project Settings → "Show
appsscript.json manifest file") before pushing them.

## The template's script project

In the template's script editor (Extensions → Apps Script), under
**Services**, add the **Drive API** (identifier `Drive`, version v3). The
Timetable Admin menu uses it to make Google Docs of a timetable or of the
compliance summary. The setting is copied into every generated timetable.
The first time someone uses the menu in a timetable, Google asks them to
allow it (including Drive access, to save the Docs next to the timetable).

## Testing against a copy of the template

The wizard copies the live master template unless a `TEMPLATE_ID` script
property is set. To test against a copy:

1. Make a copy of the master template (this also copies its bound script).
2. In the wizard project, open Project Settings → Script Properties and add
   `TEMPLATE_ID` with the copy's file ID (the long ID in its URL).
3. Delete the property to go back to the live template.

Script Properties apply to every deployment of the project, so remove the
property before releasing to schools.

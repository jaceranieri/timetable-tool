# timetable-tool
A timetable helper wizard tool that assists Primary schools creating their timetable

## Layout

| Folder | Apps Script project | What it does |
|---|---|---|
| `wizard/` | The standalone web app project | `Index.html` is the setup wizard; `Code.gs` copies the master template and builds the class and specialist tabs |
| `template/` | The project bound to the master template spreadsheet | `TemplateSync.gs` adds the **Timetable Admin** menu (sync, lock, clash formatting). It is copied into every generated timetable along with the spreadsheet. |

`PLAN.md` lists the planned fixes and features in priority order.

## Deploying the wizard

Deploy → Manage deployments:

- **Execute as:** User accessing the web app
- **Who has access:** Anyone within Sydney Catholic Schools

Each user needs view access to the master template, so share it with the
domain as **Viewer**.

The `appsscript.json` manifests in this repo reflect those settings. Check
them against the ones in the Apps Script editor (Project Settings → "Show
appsscript.json manifest file") before pushing them.

## Testing against a copy of the template

The wizard copies the live master template unless a `TEMPLATE_ID` script
property is set. To test against a copy:

1. Make a copy of the master template (this also copies its bound script).
2. In the wizard project, open Project Settings → Script Properties and add
   `TEMPLATE_ID` with the copy's file ID (the long ID in its URL).
3. Delete the property to go back to the live template.

Script Properties apply to every deployment of the project, so remove the
property before releasing to schools.

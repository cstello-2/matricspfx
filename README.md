# Metrics web part (SPFx)

A SharePoint Framework web part that builds dashboards from a **declarative JSON config**: KPI tiles, charts, and relational tables over **SharePoint lists**. Inline data and REST endpoints also work as sources.

The web part is a small app with a header and three tabs: **List**, **Metrics** and **Data**.

- **List**: a Jira-style view of one SharePoint list. You choose which columns to show and in what order. It has search, quick filters, sortable columns, status lozenges, priority icons, tag pills, people avatars, relative dates and keyboard navigation. Clicking an item, or **New item**, opens a **full-screen form**:
  - every field is a labeled input (text, dropdowns, date pickers, a people picker), saved with a **Save** bar;
  - the status is a lozenge dropdown;
  - the item's SharePoint **comments** are at the bottom.
- **Metrics**: the dashboard described below.
- **Data**: the rows the metrics are built from, one sub-tab per data source (or derived dataset). It has search, sorting, paging and **Export CSV**, so people can check where a number comes from. Rows from SharePoint sources open the same form.
- **Auto-tagging**: site owners, list owners and anyone you allow can fill a tags column with AI. It uses any OpenAI-compatible embeddings API (plus an optional chat model) to match each item or document, including the text of `.docx`, Markdown and text files, against the available tags.

The dashboard features:
- **Relational**: you can join lists on lookup or ID fields (left, inner, or one-to-many), stack datasets on top of each other, and union them.
- **Metrics**: computed fields, `groupBy` with date buckets, aggregates (`sum`, `avg`, `countIf`, `median`, …), and `having`, `sort`, and `limit`.
- **Charts** use Apache ECharts: bar, column, line, area, pie, donut, scatter, gauge, funnel, heatmap, sankey, treemap, and radar. They use a colorblind-safe palette and follow the site's light or dark theme. Every chart has a table view, and an `echarts` escape hatch lets you set any raw ECharts option.
- **Tables** support sorting, search, paging, data bars, status icons, links, and **master/detail** drill-down.
- **Safe expressions**: expressions are evaluated by a small parser, not `eval`, so the web part needs no custom-script permission and is CSP-friendly.

You don't need a tenant to develop it. The **local workbench** runs the same components with live reload against a mock SharePoint REST API.

```
src/
  core/           pure TypeScript engine (no SPFx imports): types, expressions, datasets, data sources, validation
  components/     React UI shared by the web part and the workbench (dashboard, widgets, JSON editor)
  webparts/metrics/  thin SPFx shell: property pane, JSON editor panel, SPHttpClient adapter, theme
  samples/        example configs (also shown in the workbench)
workbench/        Vite dev harness + SharePoint REST mock + fixture lists
  tests/          Vitest unit tests (engine, expressions, validation, theme mapping, OData mock, chart options)
```

## Getting started

Requires Node 22 (or 24).

```bash
npm install
```

This also installs the workbench's own dev tooling (Vite and Vitest live in `workbench/`, kept apart from the SPFx toolchain).

### Run locally with live reload (no tenant needed)

```bash
npm run workbench
```

Then open http://localhost:4400.

- On the left is a JSON editor that validates as you type: JSON syntax, unknown datasets or sources, and broken expressions are underlined. On the right is the rendered web part. While the JSON is invalid, the last valid config keeps rendering.
- In the toolbar you can pick a **sample**, a **site theme**, a **section background** (none, neutral, soft, strong), a **width** (full, one column, two-thirds, half, one-third, phone), whether you're **signed in as** a site owner or a plain member (to test who sees auto-tagging), and read or edit mode.
- Site themes are generated the way the SharePoint theme designer generates them, with Fluent UI's `ThemeGenerator`. Section backgrounds are derived with `getNeutralVariant`, `getSoftVariant` and `getStrongVariant` from `@fluentui/scheme-utilities`, the same functions SharePoint uses for section backgrounds.
- The workbench renders the same `MetricsWebPartView` component as the web part, so what you see is the web part's own UI.
- The **Requests** panel shows every mocked SharePoint REST call, with timing and item counts.
- **Seed data** switches the preview to a browser for the mock's data. It covers every fixture list and the users, with an items table (search, lookups and people resolved to names, comment counts) and a schema table.
  - Items you changed in the app (edits, new comments) are highlighted, and "Only edited items" filters to them.
  - **Discard edits** reloads the fixture files and drops all in-memory changes. It's served from `/__sp-mock/seed` and `/__sp-mock/reset`.
- Source changes in `src/` hot-reload. Edits to fixtures in `workbench/fixtures/` reload the data automatically.
- Your edits are stored per sample in the browser. **Reset** restores the sample.
- If a sample file changes on disk while you have edits, a banner offers to load the updated sample or keep your edits. Samples you haven't edited update by themselves.

#### The SharePoint mock

`workbench/sp-mock/plugin.ts` emulates the list REST endpoints for any site URL, addressing lists by `getbytitle('X')` or `lists(guid'…')`. It reads them from `workbench/fixtures/lists/*.json`:

```json
{
  "title": "Tasks",
  "id": "8a4f…",
  "fields": {
    "Title": { "type": "Text", "required": true },
    "Status": { "type": "Choice", "choices": ["Open", "Done"] },
    "Project": { "type": "Lookup", "list": "Projects", "field": "Title" },
    "AssignedTo": { "type": "User", "title": "Assigned to" },
    "Due": { "type": "DateTime", "dateOnly": true }
  },
  "items": [{ "Id": 1, "Title": "Build API", "Status": "Open", "ProjectId": 3, "AssignedToId": 12 }],
  "comments": { "1": [{ "id": 1, "text": "On it", "createdDate": "2026-11-01T09:00:00Z", "authorId": 12 }] }
}
```

The endpoints it serves:

| Endpoint | Supports |
|---|---|
| list (`…/lists/getbytitle('X')`) | GET list info |
| `/fields` | GET the field schema |
| `/items` | GET with `$select` (including `Lookup/Field`), `$expand` for lookup and person columns (people come from `fixtures/users.json`), `$filter`, `$orderby`, `$top`, and paging through `odata.nextLink` |
| `/items(id)` | GET one item; POST with `X-HTTP-Method: MERGE` to update it |
| `/items(id)/Comments()` | GET and POST comments |

The mock also answers the site-level calls auto-tagging needs:
- `/_api/web/EffectiveBasePermissions` and the list's `/EffectiveBasePermissions`: full control as owner, contribute as member;
- `/_api/web/currentuser/groups`;
- `GetFileByServerRelativeUrl('…')/$value`, which serves files from `workbench/fixtures/files/<Library>/`.

A **mock AI endpoint** at `/mock-ai/v1` implements `/embeddings` and `/chat/completions`. Its embeddings are deterministic hashed word features, which is enough to try auto-tagging offline. To use a real model, point `autotag.api.baseUrl` at your service. The `documents` sample is a "Policies" document library with real `.docx`, Markdown, text and PDF files, ready to auto-tag.

Notes on the mock:
- **Filters**: `$filter` supports `eq ne gt ge lt le and or not`, `substringof`, `startswith`, `endswith` and `datetime'…'`.
- **Behavior**: it adds 120–450 ms of simulated latency. Its errors look like SharePoint's (unknown list, invalid filter, select without expand, unknown or read-only fields, invalid choice values).
- **Edits**: edits and new comments are kept in memory until the dev server restarts or a fixture file changes. Created and modified-by values use "Dev User".

To add a list, drop a JSON file in `workbench/fixtures/lists/`. To regenerate the demo data (Projects, Tasks, Tickets, users), run:

```bash
npm run fixtures
```

### Tests

```bash
npm run test:unit
```

### Build and deploy

```bash
npm run build
```

This runs the SPFx production build and writes `sharepoint/solution/metricspfx.sppkg`. Upload it to your tenant (or site collection) app catalog, then add the **Metrics** web part to a page; it's under *Advanced*. If you do get a tenant, `npm start` runs the standard SPFx hosted-workbench flow.

## Using the web part

In edit mode, the web part offers **Open JSON editor** (a large panel with the same validating editor) or **Use starter example**. In the property pane you choose where the config lives:

- **In this web part**: the JSON is stored in the web part's properties.
- **JSON file**: a server-relative or absolute URL to a `.json` file, for example in Site Assets. Several pages can then share one definition, and you can version it in a document library.

Comments and trailing commas are allowed in the JSON. For autocomplete in VS Code, point `$schema` at [`src/core/schema/metrics.schema.json`](src/core/schema/metrics.schema.json).

### Theme

The web part declares `supportsThemeVariants` and reads its theme from SharePoint's `ThemeProvider` service. It re-renders when `themeChangedEvent` fires, which happens when the site theme or the section background changes.

- **Dashboard colors**: text, subtext, dividers, card surface, accent, links, hover color and font all come from the section theme's `semanticColors`, `palette` and `fonts`.
- **Dark sections**: inverted (dark or strong) sections switch to the dark steps of the chart palette.
- **Fluent UI controls**: the placeholder buttons are wrapped in Fluent's `ThemeProvider` with the section theme. The editor panel uses the site theme, like SharePoint's own panels.
- **Chart colors**: categorical series use a validated colorblind-safe palette by default. Set `"theme": { "palette": "theme" }` to lead with the site color and use the site's theme shades for heatmaps and funnels.

## Config reference

```jsonc
{
  "title": "Project portfolio",
  "refreshSeconds": 300,                // optional auto-refresh
  "dataSources": { … },                 // raw data
  "datasets":   { … },                  // derived data (joins, metrics)
  "list":       { … },                  // the List tab (optional)
  "dataView":   { … } | false,          // the Data tab (shown by default)
  "widgets":    [ … ],                  // the Metrics tab, in a 12-column grid
  "layout": { "columns": 12, "gap": 16 },
  "theme": { "palette": "default" | "theme" | ["#hex", …], "cardStyle": "outlined" | "elevated" | "flat" },
                                        // colors otherwise come from the SharePoint section theme
  "locale": "en-GB"
}
```

### Data sources

| type | keys |
|---|---|
| `spList` | `list` (title) or `listId`, `webUrl` (another site), `select` (use `Lookup/Field` for lookup and person columns; `$expand` is derived), `filter` (OData, runs on the server), `orderBy`, `top` (default 5000, pages automatically), `pageSize`, `rename` |
| `inline` | `rows: [ {…}, … ]` |
| `rest` | `url`, `path` (dot path to the array), `headers`, `sharepoint` (use SharePoint auth; on by default for `/_api/` URLs) |

Lookup and person columns arrive as objects: `"select": ["AssignedTo/Title"]` becomes `AssignedTo.Title` in expressions. Multi-value fields arrive as arrays.

### Datasets

```json
"projectStats": {
  "from": "projects",                                  // a source or another dataset; ["a", "b"] = union
  "join": [{ "source": "taskTotals", "on": ["Id", "ProjectId"], "as": "t", "type": "left" }],
  "filter": "Status != 'Cancelled'",
  "compute": { "progress": "safeDiv(t.done, t.tasks, 0)", "daysLeft": "daysBetween($today, DueDate)" },
  "groupBy": ["Department", { "field": "DueDate", "bucket": "month", "as": "month" }],
  "aggregate": { "budget": "sum(Budget)", "n": "count()", "late": "countIf(daysLeft < 0)" },
  "having": "n > 1",
  "sort": ["-budget", "month"],
  "limit": 10
}
```

Joins:
- A plain join multiplies rows, like SQL.
- `"many": true` instead nests all matches as an array. You can then aggregate it in `compute`, for example `sum(tasks.Hours)` or `count(tasks)`.
- A multi-value lookup on the left side matches any of its IDs.

Widgets can reference a dataset by name or define one inline (`"dataset": { "from": …, … }`).

### Expressions

The language looks like JavaScript, with SQL-friendly extras:

- **Literals and fields**: `Hours * 1.2`, `Status == 'Open' && !Done`, `Priority in ['High', 'Urgent']`, `a ?? 0`, `x > 5 ? 'big' : 'small'`, `project.Title`, `` `Field With Spaces` ``.
- **Variables**: `$today`, `$now`, `$user.email`, `$user.name`, `$webUrl`.
- **Scalar functions**: `if coalesce isEmpty round floor ceil abs min max clamp safeDiv pct number lower upper trim len contains startsWith endsWith replace left right split join concat date year month day weekday quarter addDays addMonths daysBetween startOf bucket formatDate`.
- **Aggregates**: `sum avg min max count countIf distinct median percentile stdev first last`.
  - In `aggregate`, a KPI `value`, or a text `{{ }}` placeholder, they run over the rows of the group: `sum(Hours) / count()`.
  - On an array (from a `many` join) they aggregate that array: `sum(tasks.Hours)`.
- **Dates**: subtracting two dates gives days; adding a number to a date adds days.

### List tab

```json
"list": {
  "source": "tasks",
  "title": "Tasks",
  "keyPrefix": "TASK",
  "columns": ["Title", "Status", "Priority", "AssignedTo", "Project", { "field": "EstimateHours", "title": "Estimate" }, "Modified"],
  "sort": "-Modified",
  "quickFilters": ["Status", "Priority", "AssignedTo"],
  "statusColors": { "Not started": "neutral", "In progress": "info", "Blocked": "critical", "Done": "good" },
  "form": {
    "status": "Status",
    "main": [
      "Description",
      { "title": "Effort", "columns": 2, "fields": ["EstimateHours", "ActualHours"] }
    ],
    "side": [
      { "title": "Details", "fields": ["Status", "AssignedTo", "Priority", "Project"] },
      { "title": "Dates", "fields": ["CompletedDate"] },
      { "title": "Labels", "fields": ["Tags"] }
    ],
    "sideWidth": 380,
    "readOnly": ["Project"],
    "comments": true
  }
}
```

**The list itself**
- **Source**: `source` names an `spList` data source. The tab reads the list's field schema, so column headers, editors and choice options come from SharePoint. You don't need to repeat them in the config.
- **Columns**: `columns` sets which columns appear and their order, using the same column objects as table widgets. Lookup and person fields can be listed by name (`"Project"`, `"AssignedTo"`) or by sub-field (`"Project.Code"`). The `select`/`$expand` query is built for you.
- **Keys**: every item gets a Jira-style key (`keyPrefix` + ID).
- **Lozenges**: choice values listed in `statusColors`, and the `form.status` field, show as lozenges. The levels are `good`, `info`, `warning`, `serious`, `critical` and `neutral`.
- **Options**: `quickFilters` adds filter dropdowns above the list. `filter` is an expression that limits which items show. `search` (default on) and `pageSize` (default 25) control search and paging.

**The item form**
- **Opening**: clicking a row (or pressing Enter or Space on it) opens the full-screen form, and **New item** opens it empty. Rows on the Data tab that come from an `spList` source open the same form.
- **Keyboard**: arrow keys move between rows.
- **Deep links**: the URL gets `?item-<id>=<itemId>` (or `<source>:<itemId>` for Data tab sources, `new` for a new item), so items can be linked directly. The browser's back button closes the form.
- **Layout**: the form has a main column and a right column, both set in config.
  - `form.main` lists what goes under the title, in order. `form.side` lists the right column.
  - Each entry is a field name or a section: `{ "title": "Effort", "columns": 2, "fields": [...] }`. A section gets a heading, and `columns: 2` puts its fields side by side (one column on narrow screens).
  - In the right column each section is a card. Loose field names are grouped into a **Details** card.
  - Multi-line fields in the main column get a large editor. The `form.status` field shows as a lozenge picker wherever it's placed, and goes first in the right column if neither list mentions it.
  - `sideWidth` (260 to 640, default 380) sets the right column's width in pixels. Below 900px wide the columns stack.
  - Defaults: `form.description` is shorthand for `"main": ["<field>"]`, and `form.fields` is the older name for `side`. Without `side`, the right column shows the list's columns, or every field for sources without a configured form. Comments stay at the bottom of the main column, and who created and last updated the item at the bottom of the right one.
- **Editing**: every editable field is a labeled input.
  - Text, multi-line text, number, currency, yes/no, choice, multi-choice, date, lookup, multi-lookup, and person or group fields (a people picker over the site's users) can all be edited.
  - Changes collect in a draft. A **Save** bar shows how many are unsaved, and **Discard** reverts them. **Ctrl+S** saves.
  - Required fields are checked before saving.
  - Closing or moving to another item with unsaved changes asks for confirmation.
  - Saves use the list item REST API with the current user's permissions.
  - Rich-text and read-only fields are shown, not edited. `form.readOnly` disables editing entirely (`true`) or for the listed fields.
- **Comments**: these are the list item's own SharePoint comments (the same ones shown in the SharePoint list UI), read and posted through `/items(id)/Comments()`. Set `form.comments` to `false` to hide them.

Edits on the List tab refresh the Metrics tab's data.

### Data tab

The Data tab shows every data source by default. Configure it to choose and order what viewers see, including calculated datasets, with friendlier titles and columns:

```json
"dataView": {
  "title": "Data",
  "sources": [
    { "name": "tasks", "title": "Tasks", "columns": ["Id", "Title", "Status", { "field": "AssignedTo.Title", "title": "Assigned to" }, { "field": "Created", "format": "date" }] },
    "projects",
    { "name": "projectStats", "title": "Project summary (calculated)" }
  ],
  "pageSize": 50,
  "export": true
}
```

- **Columns**: use the same column objects as table widgets (`title`, `format`, `width`, `align`, `bar`, `link`, `status`). Without `columns`, every field is shown except SharePoint system fields and internal join objects.
- **Loading**: rows are loaded with the viewer's permissions, through the same queries the metrics use.
- **Export**: **Export CSV** downloads what the source contains, with formatted values and a byte-order mark so Excel opens UTF-8 correctly. Set `"export": false` to hide the button.
- **Hiding the tab**: `"dataView": false`.

### Auto-tagging

```json
"list": {
  "source": "policies",
  "autotag": {
    "field": "Topics",
    "api": { "baseUrl": "https://ai.example.com/v1", "embeddingModel": "text-embedding-3-small", "chatModel": "gpt-4o-mini" },
    "tagDescriptions": { "Security": "passwords, MFA, phishing, VPN", "Travel": "flights, hotels, expenses for trips" },
    "sourceFields": ["Title", "Summary"],
    "includeFileText": true,
    "maxTags": 3,
    "minScore": 0.25,
    "mode": "review",
    "allow": ["owners", "listOwners", "Policy Editors", "ann@contoso.com"]
  }
}
```

**How suggestions are made**
1. **Tag vocabulary.** It comes from the tags column: a choice column's choices, a lookup column's list items, or `tags` for plain text columns. Each tag is embedded once per session, together with its description from `tagDescriptions` or `tags[].description`. Good descriptions are the biggest quality lever.
2. **Item text.** Each item's text is built from `sourceFields` and capped at `maxChars`. In document libraries it also includes the file's text when `includeFileText` is on: `.txt`, `.md`, `.csv`, `.json`, `.html`/`.aspx` and `.docx` (read in the browser, no extra libraries). Other files, such as PDFs, fall back to their metadata.
3. **Ranking.** Tags are ranked by cosine similarity. Tags below `minScore` and tags the item already has are dropped. If `api.chatModel` is set, the chat model picks at most `maxTags` from the closest candidates; otherwise the top `maxTags` are used.

**Where it shows up**
- **List toolbar: Auto-tag.** Runs on the items without tags, or on everything currently shown (filters apply), up to `maxItems` per run.
  - In `review` mode you see every suggestion with its score, can untick any of them, then apply.
  - In `auto` mode the suggestions are applied straight away and you see the result per item.
- **Item form: Suggest tags.** Shows chips; clicking one adds that tag.
- **Saving.** By default new tags are merged with existing ones; set `merge: false` to replace them. Saves use the normal item update API with the user's own permissions.

**Who can use it.** `allow` controls this. The default `["owners", "listOwners"]` means people with Manage Web on the site or Manage Lists on the list. You can add login names, emails and SharePoint group names. Everyone else doesn't see the feature.

This check only decides who sees the button; it is not a security boundary. Anyone who can edit items can change tags anyway, and your public AI endpoint is reachable by anyone.

**The API** is called straight from the browser: `POST {baseUrl}/embeddings` and, optionally, `POST {baseUrl}/chat/completions`. No key is sent. The endpoint must allow CORS requests from your SharePoint origin.

### Widgets

All widgets accept `title`, `subtitle`, `span` (1–12; defaults are kpi 3, chart 6, table 12), `height`, and `dataset`. Below 560 px of width, widgets stack.

**kpi**

```json
{ "type": "kpi", "dataset": "lastMonth", "label": "Revenue", "value": "sum(Revenue)", "format": "currency",
  "compare": { "dataset": "prevMonth", "value": "sum(Revenue)", "label": "vs last month", "goodWhen": "up" },
  "target": 50000,
  "thresholds": [{ "when": "value >= target", "level": "good" }, { "when": "true", "level": "warning", "label": "Below target" }],
  "sparkline": { "dataset": "byMonth", "x": "month", "y": "revenue" } }
```

The levels are `good`, `warning`, `serious`, `critical`, and `neutral`. Each one always has its own icon, so status never relies on color alone.

**chart** (`chart`: `bar` horizontal · `column` · `line` · `area` · `pie` · `donut` · `scatter` · `gauge` · `funnel` · `heatmap` · `sankey` · `treemap` · `radar`)

| chart | keys |
|---|---|
| bar / column / line / area / radar | `x`, `y` (a field, a list, or `{field, name, type, color, stack}` for combo charts), `series` (pivot one field into series), `stacked`, `smooth` |
| pie / donut / funnel / treemap | `x` (label), `y` (value). Pies fold everything past 8 slices into "Other". |
| scatter | `x`, `y` (numeric), `series` |
| gauge | `value` (an aggregate expression), `min`, `max` |
| heatmap | `x`, `series` (rows), `value` |
| sankey | `from`, `to`, `value`. Union several grouped datasets to get multi-stage flows. |

Common options are `format`, `labels` (`auto`, `true`, or `false`), `showLegend`, `colors`, and `echarts` (raw options, deep-merged). There is deliberately no dual y-axis: plot two measures on different scales as two charts.

**table**

```json
{ "type": "table", "dataset": "projectStats", "search": true, "pageSize": 10,
  "columns": ["Title", { "field": "progress", "format": "percent", "bar": true },
              { "field": "Status", "status": [{ "when": "value == 'At risk'", "level": "critical" }] },
              { "field": "Title", "link": "concat($webUrl, '/Lists/Projects/DispForm.aspx?ID=', Id)" }],
  "detail": { "dataset": "tasks", "on": ["Id", "ProjectId"], "title": "Tasks", "columns": ["Title", "Status"] } }
```

**text**

```json
{ "type": "text", "dataset": "open", "text": "There are **{{ count() }}** open tickets worth {{ sum(Value) | currency }}." }
```

Formats: `number`, `integer`, `decimal`, `percent`, `currency`, `compact`, `date`, `datetime`, `text`, or `{ "type": "currency", "currency": "EUR", "decimals": 0, "prefix": "", "suffix": " h" }`.

## Notes and limits

- Data is read with the current user's permissions, through `SPHttpClient`.
- `filter` on an `spList` runs in SharePoint. On lists over 5000 items, filter on indexed columns, or SharePoint rejects the query. Dataset `filter`, `groupBy` and the other transforms run in the browser after loading, up to `top` items.
- The List tab loads up to `top` items of its source (default 5000) and filters, sorts and pages them in the browser.
- Item edits use `IF-MATCH: *`, so the last save wins; there is no conflict detection between two people editing the same item.
- **Write format.** Saves and creates use SharePoint's verbose format:
  - the body carries `__metadata.type` set to the list's `ListItemEntityTypeFullName`;
  - multi-choice, multi-lookup and multi-person values are sent as typed `Collection(Edm.String/Int32)` `results`.
  - The workbench mock enforces the same rules, rejecting bare arrays and wrong entity types, so format mistakes show up locally.
- **Schema and files.**
  - The list schema is read without `$select`, so every field type returns its own properties.
  - Document text is read with `GetFileByServerRelativePath(decodedurl=…)`, which handles `%` and `#` in file names.
- **New items in libraries.** **New item** is hidden for document libraries, because SharePoint only creates library items by uploading a file.
- The item form doesn't edit rich-text notes, attachments or managed metadata yet. The people picker searches the site's users (`/_api/web/siteusers`), so people who have never visited the site won't appear until they're added to it.
- Auto-tagging reads `.docx`, text and HTML files, but not PDFs or other Office formats. Item and document text is sent to the configured AI endpoint, so point it at a service you're allowed to send that content to.
- The workbench mock implements the common subset of the list items REST API. It doesn't cover CAML, `RenderListDataAsStream`, or the 5000-item threshold.
- The production web part bundle is about 1.6 MB (minified), mostly ECharts, CodeMirror and Fluent UI. Lazy-loading the edit-mode JSON editor would be the first optimization.

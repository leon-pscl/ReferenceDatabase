# Reference Database

A Google Sheets + Apps Script tool that syncs your reference library into a spreadsheet, automatically creates Google Docs for each reference, and lets you edit notes that sync back into the sheet.

**Template:** https://docs.google.com/spreadsheets/d/1Vt1OyLrniMwJVqdyX9dizRBKMi0jiOr9CvDs5ytEt6M/edit?usp=sharing

## Features

- **Zotero sync** -- pulls metadata and formatted citations from your Zotero group library
- **One-click Doc creation** -- generates a Google Doc per reference with title, authors, abstract, and editable sections
- **Doc-to-sheet sync** -- edits you make in a reference's Doc (Problem, Solution, Comments, etc.) are pulled back into the sheet
- **Column visibility manager** -- hide/show columns or apply presets ("Essentials", "Citation info", "Everything")
- **Manual import** -- browse your Zotero library from a sidebar, or paste a Zotero item link to add a single reference
- **Native Google Sheets Table** -- wraps the data in a filterable, sortable Table (requires Google Sheets API advanced service)

---

## Initial Setup

### 1. Duplicate the template

Open the template link above, then go to **File > Make a copy** to create your own editable version.

### 2. Open the Apps Script editor

In your copy, go to **Extensions > Apps Script**. This opens the script editor where the code lives.

### 3. Add Script Properties

Go to **Project Settings > Script Properties** and click **Add script property** for each of these:

| Property | Required | Description |
|---|---|---|
| `DOCS_FOLDER_ID` | **Yes** | The Google Drive folder ID where Docs will be created. You can find this in the folder's URL: `drive.google.com/drive/folders/THIS_PART_HERE` |
| `ZOTERO_API_KEY` | Yes (for Zotero) | Your personal Zotero API key. Generate one at https://www.zotero.org/settings/keys |
| `ZOTERO_GROUP_ID` | Yes (for Zotero) | The numeric ID of your Zotero group library. Found in the group URL: `zotero.org/groups/YOUR_GROUP_ID` |

> The `DOCS_FOLDER_ID` value is the one censored in the screenshot below.

![DOCS_FOLDER_ID location](Screenshot%202026-08-17%20123644.png)

### 4. Set up the header row

Go back to the spreadsheet, **refresh the page**, and you should see a new menu: **Ref DB Sync**.

1. Click **Ref DB Sync > Set up Header Row** -- this writes all column headers and hides the internal "Zotero Key" column.
2. Click **Ref DB Sync > Run full sync now** -- this pulls all items from your Zotero library, creates rows, and generates Docs.

The first run will ask you to authorize the script. Grant the permissions when prompted.

### 5. (Optional) Enable Google Sheets Table

For built-in per-column filter/sort dropdowns:

1. Open the Apps Script editor.
2. Go to **Services (+)** > **Google Sheets API** > **Add**.
3. Then click **Ref DB Sync > Convert References range to filterable Table**.

---

## Using the Script

### Daily workflow

| Step | Menu item | What it does |
|---|---|---|
| 1 | **Pull from Zotero** | Fetches latest metadata from Zotero, updates existing rows, appends new ones |
| 2 | **Create Docs for new rows** | Generates a Google Doc for any row that doesn't have one yet |
| 3 | **Pull notes from Docs** | Reads edited sections in each Doc and syncs values back to the sheet |

Or just click **Run full sync now** to do all three steps at once.

### Importing references

- **Browse & import references** -- opens a sidebar where you can search your Zotero library (personal or group) and click "Import" on any item.
- **Import a reference (paste Zotero link)** -- paste a Zotero item link (e.g., `https://www.zotero.org/users/12345/items/ABCD1234`) to add a single reference.

### Managing columns

Click **Manage visible columns** to show/hide Zotero-sourced columns. You can also use the one-click presets:
- **Essentials** -- Title, Authors, Year, URL, DOI
- **Citation info** -- Title, Authors, Year, Citation Key, Full Citation
- **Everything** -- all Zotero-sourced columns

### Creating a new sheet for another reference set

1. Duplicate the current sheet tab (right-click the tab > Duplicate).
2. In the duplicate, select all rows with content, right-click > **Delete Rows**.
3. The "Ref DB Sync" menu and script will work on the new sheet automatically.

---

## How It Works

The script is a single Google Apps Script project (`Code.gs`) with two HTML sidebar files. Here's what happens under the hood:

### Sync pipeline

```
Zotero API  ──>  Google Sheet  ──>  Google Docs  ──>  Google Sheet
  (pull)          (upsert rows)    (create + edit)    (sync notes back)
```

1. **Zotero -> Sheet** (`syncZoteroToSheet`): The script fetches all items from your Zotero group library via the Zotero Web API (`/groups/{id}/items/top`), paginating 100 items at a time. For each item, it extracts metadata (title, authors, year, DOI, abstract, citation, etc.) and either updates an existing row (matched by the `Zotero Key` column) or appends a new row. The `MANUAL_COLUMNS` list protects user-edited fields (Status, Item Type, Comments, etc.) from being overwritten.

2. **Sheet -> Docs** (`createDocsForNewRows`): For every row that has a title but no Doc Link yet, the script calls `DocumentApp.create()` to make a new Google Doc. The Doc is moved into a subfolder (named after the sheet tab) inside your `DOCS_FOLDER_ID` folder. The Doc is populated with a title, metadata block, abstract, and empty editable sections (Problem, Current Solution, Lacking in Current Solution, Comments). The Doc's URL is written back into the "Doc Link" column as a smart chip (if the Sheets API advanced service is enabled) or a plain URL.

3. **Docs -> Sheet** (`syncDocsToSheet`): The script opens each Doc, walks its headings, and reads the text under each `HEADING2` that matches a synced field name. If the text in the Doc differs from what's in the sheet, it updates the sheet cell. This means you can edit Problem, Solution, etc. directly in the Doc and they'll sync back on the next run.

### Column system

The script maintains a canonical column order defined in the `HEADERS` array (line 33 of `Code.gs`). Columns are looked up **by name**, so you can reorder them in the sheet without breaking anything. The `COLUMN_PRESETS` object (line 123) defines which columns each preset shows.

### Sidebar (Zotero browser)

`Sidebar.html` is a client-side UI that calls server-side functions (`getZoteroLibraryOptions`, `getZoteroCollections`, `searchZoteroItems`, `importZoteroItem`) via `google.script.run`. It supports browsing both your personal Zotero library and any group library your API key has access to. Collections are rendered as a nested tree, and items are displayed as cards with an "Import" button.

### Column Manager

`ColumnManager.html` displays checkboxes for each Zotero-sourced column. It reads current visibility state from the sheet, lets you toggle columns, and applies changes via `applyColumnVisibility`. Preset buttons call `applyColumnPreset` to show/hide columns in bulk.

---

## File Structure

```
CPE029_RefDatabase/
  README.md
  AppsScript/
    Code.gs           # All script logic (sync, import, column management, Doc creation)
    Sidebar.html       # Zotero browse/import sidebar UI
    ColumnManager.html # Column visibility checkbox dialog
```

---

## Notes

- The `Zotero Key` column is hidden and used internally to match sheet rows to Zotero items. Don't rename it.
- `MANUAL_COLUMNS` are never overwritten by the Zotero sync. You're free to fill them in.
- If the Google Sheets API advanced service is not enabled, the script falls back to plain URLs for Doc links and skips Table creation.
- The script stores only one piece of cached data (the Zotero user ID, cached for 6 hours). Everything else lives in the sheet or your Drive.

---

## Developer Notes: Customizing the Script

If you're modifying this script for your own use, here are the key things to know.

### Config constants to edit

All configuration lives at the top of `Code.gs` (lines 28-130). These are the values you'll most likely want to change:

| Constant | What it controls | Where |
|---|---|---|
| `HEADERS` | Canonical column order. Script looks up columns **by name**, so order in the sheet doesn't matter — but the names must match exactly. | Line 33 |
| `MANUAL_COLUMNS` | Columns the Zotero sync will never overwrite. Add any column you want to fill in manually. | Line 70 |
| `DOC_SYNCED_FIELDS` | Subset of manual columns that get editable `HEADING2` sections inside each Google Doc. These are the fields that sync back from Docs to Sheet. | Line 84 |
| `CITATION_STYLE` | Zotero citation style for formatted bibliographies. Any style ID Zotero supports (e.g. `'apa'`, `'mla'`, `'chicago-note-bibliography'`). | Line 92 |
| `ZOTERO_SOURCED_COLUMNS` | Columns pulled from Zotero. Used by the column manager presets. | Line 97 |
| `COLUMN_PRESETS` | One-click presets in the column visibility dialog. Each key is a preset name, value is an array of column names to show. | Line 123 |
| `TABLE_NAME` | Name given to the native Google Sheets Table. | Line 130 |

### Adding a new column

1. Add the column name to `HEADERS` in the desired position.
2. If it's a **Zotero-sourced** column, also add it to `ZOTERO_SOURCED_COLUMNS`.
3. If it's a **manual** column (user-edited, never overwritten by sync), add it to `MANUAL_COLUMNS`.
4. If it should have an editable section in the Google Doc, add it to `DOC_SYNCED_FIELDS`.
5. If you want a data dropdown, add a validation in `applyValidations()` (line 172).
6. Run **Ref DB Sync > Reorder columns to match canonical order** to reposition existing columns in the sheet.

### Changing the dropdown options

Edit the `requireValueInList(...)` arrays inside `applyValidations()` (line 172). The current values are:

- **Status**: `'Dropped'`, `'Cited in Paper'`, `'N/A'`, `'Used as dataset'`
- **Item Type**: `'Academic Study'`, `'Article/Report'`, `'Dataset'`, `'Standard'`
- **Added By**: `'Baldovino'`, `'Pascual'`, `'Potestades'`, `'Verdejo'`

### Updating the Doc template

The Doc structure is built by `populateDocTemplate()` (line 895). It writes:

1. Title (as `TITLE` heading)
2. Metadata block (Authors, Year, URL, DOI)
3. Abstract (as `HEADING2` + paragraph)
4. One `HEADING2` section per entry in `DOC_SYNCED_FIELDS`

To add a new editable section to Docs, add the column name to `DOC_SYNCED_FIELDS` — the template and sync logic will pick it up automatically.

### After making changes

1. **Refresh the spreadsheet** so the menu rebuilds.
2. Run **Set up header row** again if you changed `HEADERS` (this overwrites the header row).
3. Run **Reorder columns to match canonical order** to move columns to match your new `HEADERS` array.
4. Run **Run full sync now** to verify nothing breaks.

### Things to be careful about

- **Never rename `Zotero Key`** — it's the unique identifier that matches rows to Zotero items. Renaming it will break the sync.
- **`MANUAL_COLUMNS` vs `DOC_SYNCED_FIELDS`** — a column can be in `MANUAL_COLUMNS` without being in `DOC_SYNCED_FIELDS`. The former protects the column from overwrite during Zotero sync; the latter controls whether it gets a section in the generated Doc.
- **Column names are case-sensitive** and must match exactly across `HEADERS`, `MANUAL_COLUMNS`, `DOC_SYNCED_FIELDS`, and `ZOTERO_SOURCED_COLUMNS`.
- The script fetches items in batches of 100. If your Zotero library has thousands of items, expect the first sync to take a while (Zotero API rate limits apply).
- If you add a column to `HEADERS` but forget to add it to `ZOTERO_SOURCED_COLUMNS`, it won't appear in the column manager presets — but it will still exist in the sheet.

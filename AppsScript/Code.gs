/**
 * ============================================================================
 * ZOTERO <-> GOOGLE SHEETS <-> GOOGLE DOCS REFERENCE DATABASE SYNC
 * ============================================================================
 * Recreates a Notion-style "reference database" where:
 *   - Each row = one Zotero item, metadata auto-synced from a Zotero library
 *   - Each row gets its own Google Doc (created automatically) for notes
 *   - Manual note fields written in the Doc get synced back into the Sheet
 *
 * SETUP (see SETUP.md for full walkthrough):
 *   1. Extensions > Apps Script > paste this file in as Code.gs, and add
 *      Sidebar.html as a second file (HTML type)
 *   2. Project Settings > Script Properties, add:
 *        ZOTERO_API_KEY   = your personal Zotero API key
 *        ZOTERO_GROUP_ID  = (optional) group library ID for full sync; omit to import manually via sidebar
 *        DOCS_FOLDER_ID   = (optional) Drive folder ID where new Docs are created
 *   3. Run `setupSheet` once to write the header row
 *   4. Run `runFullSync` once manually to test, grant permissions when prompted
 *   5. Use the "Ref DB Sync" custom menu, or install a time-driven trigger on
 *      runFullSync (Triggers > Add Trigger > time-driven > every hour)
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------------

// Canonical header row, in order. Script looks columns up BY NAME, so you can
// reorder columns in the sheet later without breaking anything -- just don't
// rename them (or update HEADERS to match if you do). Run
// "Reorder columns to match canonical order" from the menu after editing this
// to physically move columns in an existing sheet to match.
const HEADERS = [
  'Doc Link',
  'Title',
  'Year',
  'URL',
  'Problem',
  'Current Solution',
  'Lacking in Current Solution',
  'Comments',
  'Added By',
  'Status',
  'Item Type',
  // -- everything else, in original relative order --
  'Zotero Key',        // unique ID used to match Zotero items to rows, keep hidden
  'In-Text Citation',
  'Authors',
  'Editors',
  'Date',
  'Date Added',
  'Date Modified',
  'Publication',
  'Proceedings Title',
  'Series Title',
  'Short Title',
  'Place',
  'DOI',
  'Abstract',
  'Citation Key',
  'Extra',
  'Full Citation',
  'Tags',
  'Collections',
  'Zotero URI',
];

// Columns the script will NEVER overwrite once a row exists (user-owned).
const MANUAL_COLUMNS = [
  'Status',
  'Item Type',
  'Added By',
  'Problem',
  'Current Solution',
  'Lacking in Current Solution',
  'Comments',
  'In-Text Citation',
];

// Manual columns that get created as editable sections inside each new Doc,
// and read back into the Sheet by syncDocsToSheet(). Order = order in the Doc.
const DOC_SYNCED_FIELDS = [
  'Problem',
  'Current Solution',
  'Lacking in Current Solution',
  'Comments',
];

const CITATION_STYLE = 'apa'; // any style id Zotero supports, e.g. 'apa', 'mla', 'chicago-note-bibliography'

// The Zotero-sourced columns, in display order -- these are the ones the
// column visibility manager and presets below operate on. MANUAL_COLUMNS,
// 'Zotero Key', and 'Doc Link' are never touched by presets.
const ZOTERO_SOURCED_COLUMNS = [
  'Title',
  'URL',
  'Authors',
  'Editors',
  'Year',
  'Date',
  'Date Added',
  'Date Modified',
  'Publication',
  'Proceedings Title',
  'Series Title',
  'Short Title',
  'Place',
  'DOI',
  'Abstract',
  'Citation Key',
  'Extra',
  'Full Citation',
  'Tags',
  'Collections',
  'Zotero URI',
];

// One-click presets for the column manager dialog. Edit freely -- any field
// not listed for a preset gets hidden when that preset is applied.
const COLUMN_PRESETS = {
  'Essentials': ['Title', 'Authors', 'Year', 'URL', 'DOI'],
  'Citation info': ['Title', 'Authors', 'Year', 'Citation Key', 'Full Citation'],
  'Everything': ZOTERO_SOURCED_COLUMNS.slice(),
};

// Name given to the native Google Sheets Table wrapping the References range.
const TABLE_NAME = 'ReferencesTable';

// ---------------------------------------------------------------------------
// MENU / ENTRY POINTS
// ---------------------------------------------------------------------------

function onInstall() {
  onOpen();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ref DB Sync')
    .addItem('Run full sync now', 'runFullSync')
    .addSeparator()
    .addItem('1. Pull from Zotero', 'syncZoteroToSheet')
    .addItem('2. Create Docs for new rows', 'createDocsForNewRows')
    .addItem('3. Pull notes from Docs', 'syncDocsToSheet')
    .addSeparator()
    .addItem('Browse & import references...', 'showZoteroSidebar')
    .addItem('Import a reference (paste Zotero link)...', 'importReferenceFromZoteroLink')
    .addSeparator()
    .addItem('Manage visible columns...', 'showColumnManager')
    .addItem('Convert References range to filterable Table', 'convertReferencesRangeToTable')
    .addSeparator()
    .addItem('List groups my API key can see', 'listMyZoteroGroups')
    .addItem('Set up header row', 'setupSheet')
    .addItem('Reorder columns to match canonical order', 'reorderColumnsToCanonicalOrder')
    .addToUi();
}

/** Run this once to write the header row into a fresh sheet. */
function setupSheet() {
  const sheet = getSheet();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  const zoteroKeyCol = HEADERS.indexOf('Zotero Key') + 1;
  sheet.hideColumns(zoteroKeyCol);
  SpreadsheetApp.getUi().alert('Header row created. Now set your Script Properties and run "Run full sync now".');
}

function applyValidations(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const statusCol = HEADERS.indexOf('Status') + 1;
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Dropped', 'Cited in Paper', 'N/A', 'Used as dataset'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, lastRow - 1, 1).setDataValidation(statusValidation);

    const itemTypeCol = HEADERS.indexOf('Item Type') + 1;
    const itemTypeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Academic Study', 'Article/Report', 'Dataset', 'Standard'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, itemTypeCol, lastRow - 1, 1).setDataValidation(itemTypeValidation);

    const addedByCol = HEADERS.indexOf('Added By') + 1;
    const addedByValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Baldovino', 'Pascual', 'Potestades', 'Verdejo'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, addedByCol, lastRow - 1, 1).setDataValidation(addedByValidation);
  } catch (e) {
    Logger.log('Skipped validations (table may have typed columns): ' + e);
  }
}

/** Retrofits any existing Doc Link cells that are still plain URL text into smart chips. */


/*
 * HEADERS array above -- lets you edit HEADERS and apply it to a sheet that
 * already has live data, instead of only affecting brand-new sheets.
 * Any column present in the sheet but not listed in HEADERS is left in place.
 */
function reorderColumnsToCanonicalOrder() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    ui.alert('No header row found -- run "Set up header row" first.');
    return;
  }

  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const unknown = [];
  const moves = headerRow.reduce(function (arr, h, i) {
    const to = HEADERS.indexOf(h);
    if (to === -1) {
      if (h) unknown.push(h);
      return arr; // leave columns not in HEADERS untouched
    }
    arr.push({ from: i + 1, to: to + 1 });
    return arr;
  }, []);
  moves.sort(function (a, b) { return a.to - b.to; });

  moves.forEach(function (move) {
    if (move.from !== move.to) {
      sheet.moveColumns(sheet.getRange(1, move.from, sheet.getMaxRows()), move.to);
      // Column indices shift after every move -- adjust remaining "from" positions.
      moves.forEach(function (other) {
        if (other.from < move.from) other.from += 1;
      });
    }
  });

  let msg = 'Columns reordered to match the canonical order.';
  if (unknown.length) {
    msg += '\n\nLeft these columns where they were (not found in HEADERS): ' + unknown.join(', ');
  }
  ui.alert(msg);
}


function runFullSync() {
  const apiKey = getProp('ZOTERO_API_KEY');
  const groupId = getProp('ZOTERO_GROUP_ID', true);
  if (!groupId) {
    SpreadsheetApp.getUi().alert('No ZOTERO_GROUP_ID set — skipping sync. Use "Browse & import references..." to add items manually.');
    return;
  }
  const items = fetchAllZoteroItems(apiKey, 'groups', groupId);
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  sheets.forEach(function (sheet) {
    if (sheet.getLastColumn() === 0) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasKey = false;
    headers.forEach(function (h) { if (h === 'Zotero Key') hasKey = true; });
    if (!hasKey) return;

    syncItemsToSheet(sheet, items, apiKey, libraryType, libraryId);
    createDocsForNewRows(sheet);
    syncDocsToSheet(sheet);
    try { ensureReferencesTable(sheet); } catch (e) { Logger.log('Skipped Table refresh: ' + e); }
  });

  SpreadsheetApp.getUi().alert('Sync complete -- all sheets updated.');
}

function syncItemsToSheet(sheet, items, apiKey, libraryType, libraryId) {
  const colIndex = getColumnIndexMap(sheet);
  const existing = getExistingRowsByZoteroKey(sheet, colIndex);
  const rowsToAppend = [];

  items.forEach(function (item) {
    if (!item.data || item.data.itemType === 'attachment' || item.data.itemType === 'note') return;
    const rowValues = buildRowFromZoteroItem(item, apiKey, libraryType, libraryId);
    const key = item.key;
    if (existing[key]) {
      writeZoteroColumnsToRow(sheet, existing[key].rowNum, colIndex, rowValues);
    } else {
      rowsToAppend.push(buildFullRowArray(colIndex, rowValues));
    }
  });

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, HEADERS.length).setValues(rowsToAppend);
  }
  applyValidations(sheet);
  try { ensureReferencesTable(sheet); } catch (e) { Logger.log('Skipped table expand: ' + e); }
  Logger.log('Synced ' + sheet.getName() + ': ' + items.length + ' items, ' + rowsToAppend.length + ' new.');
}

// ---------------------------------------------------------------------------
// STEP 0: DISCOVER YOUR ZOTERO GROUP ID (helper, run manually once)
// ---------------------------------------------------------------------------

function listMyZoteroGroups() {
  const apiKey = getProp('ZOTERO_API_KEY');
  const info = getZoteroKeyInfo(apiKey);
  const groups = (info.access && info.access.groups) || {};
  const lines = ['Your userID: ' + info.userID, '', 'Groups this key can access:'];
  Object.keys(groups).forEach(function (gid) {
    if (gid === 'all') return;
    lines.push('  Group ID ' + gid + ' -> permissions: ' + JSON.stringify(groups[gid]));
  });
  if (lines.length === 3) {
    lines.push('  (none found -- edit your key at zotero.org/settings/keys and grant group access)');
  }
  SpreadsheetApp.getUi().alert(lines.join('\n'));
  Logger.log(lines.join('\n'));
}

/** Fetches /keys/{apiKey}, used to resolve the personal userID and sanity-check the key. */
function getZoteroKeyInfo(apiKey) {
  const res = UrlFetchApp.fetch('https://api.zotero.org/keys/' + apiKey, {
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Zotero API error ' + res.getResponseCode() + ' while resolving your API key.');
  }
  return JSON.parse(res.getContentText());
}

/** Resolves and caches the personal library userID tied to ZOTERO_API_KEY. */
function getZoteroUserId(apiKey) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('zotero_user_id');
  if (cached) return cached;
  const info = getZoteroKeyInfo(apiKey);
  if (!info.userID) {
    throw new Error('Could not resolve a Zotero userID from your API key.');
  }
  cache.put('zotero_user_id', String(info.userID), 21600); // 6 hours
  return String(info.userID);
}

// ---------------------------------------------------------------------------
// SIDEBAR: BROWSE + SEARCH ZOTERO, CLICK TO IMPORT
// (Replaces having to paste a Zotero item link by hand.)
// ---------------------------------------------------------------------------

function showZoteroSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(900)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import from Zotero');
}

/**
 * Called from Sidebar.html. Returns which libraries are available to browse,
 * so the client can populate the "search in" dropdown without hardcoding IDs.
 */
function getZoteroLibraryOptions() {
  const apiKey = getProp('ZOTERO_API_KEY');
  const options = [];

  try {
    const userId = getZoteroUserId(apiKey);
    options.push({ label: 'My Library', libraryType: 'users', libraryId: userId });
  } catch (e) {
    Logger.log('Could not resolve personal library: ' + e);
  }

  const groupId = getProp('ZOTERO_GROUP_ID', true);
  if (groupId) {
    options.push({ label: 'Group Library (' + groupId + ')', libraryType: 'groups', libraryId: groupId });
  }

  return options;
}

function getZoteroCollections(libraryType, libraryId) {
  const apiKey = getProp('ZOTERO_API_KEY');
  const all = [];
  let start = 0;
  while (true) {
    const url =
      'https://api.zotero.org/' + libraryType + '/' + libraryId + '/collections' +
      '?format=json&limit=100&start=' + start;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Zotero-API-Key': apiKey },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      throw new Error('Zotero API error ' + res.getResponseCode() + ': ' + res.getContentText());
    }
    const batch = JSON.parse(res.getContentText());
    all.push.apply(all, batch);
    if (batch.length < 100) break;
    start += 100;
  }
  return all.map(function (c) {
    return { key: c.data.key, name: c.data.name, parentCollection: c.data.parentCollection || null };
  });
}

function getZoteroCollectionItems(libraryType, libraryId, collectionKey) {
  const apiKey = getProp('ZOTERO_API_KEY');
  const all = [];
  let start = 0;
  while (true) {
    const basePath = collectionKey === '__all__'
      ? '/' + libraryType + '/' + libraryId + '/items/top'
      : '/' + libraryType + '/' + libraryId + '/collections/' + collectionKey + '/items/top';
    const url =
      'https://api.zotero.org' + basePath +
      '?format=json&include=data&limit=100&start=' + start;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Zotero-API-Key': apiKey },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      throw new Error('Zotero API error ' + res.getResponseCode() + ': ' + res.getContentText());
    }
    const batch = JSON.parse(res.getContentText());
    all.push.apply(all, batch);
    if (batch.length < 100) break;
    start += 100;
  }
  return all
    .filter(function (item) {
      return item.data && item.data.itemType !== 'attachment' && item.data.itemType !== 'note';
    })
    .map(function (item) {
      var d = item.data;
      var authors = (d.creators || [])
        .filter(function (c) { return c.creatorType === 'author'; })
        .map(function (c) { return c.lastName || c.name || ''; })
        .join(', ');
      var year = '';
      if (d.date) {
        var m = d.date.match(/(\d{4})/);
        if (m) year = m[1];
      }
      return { key: item.key, title: d.title || '(untitled)', authors: authors, year: year, itemType: d.itemType };
    });
}

/**
 * Called from Sidebar.html as the user types. Returns a lightweight list of
 * matching items for display -- NOT the full Zotero record (kept small so the
 * sidebar stays fast). importZoteroItem() re-fetches the full item on import.
 */
function searchZoteroItems(query, libraryType, libraryId) {
  if (!query || query.trim().length < 2) return [];
  const apiKey = getProp('ZOTERO_API_KEY');

  const url =
    'https://api.zotero.org/' + libraryType + '/' + libraryId + '/items' +
    '?q=' + encodeURIComponent(query.trim()) +
    '&qmode=titleCreatorYear&format=json&include=data' +
    '&limit=25';

  const res = UrlFetchApp.fetch(url, {
    headers: { 'Zotero-API-Key': apiKey },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Zotero API error ' + res.getResponseCode() + ': ' + res.getContentText());
  }

  const items = JSON.parse(res.getContentText());
  return items
    .filter(function (item) {
      return item.data && item.data.itemType !== 'attachment' && item.data.itemType !== 'note';
    })
    .map(function (item) {
      const authors = (item.data.creators || [])
        .filter(function (c) { return c.creatorType === 'author'; })
        .map(formatCreatorName)
        .join('; ');
      const year = (item.data.date || '').match(/\d{4}/);
      return {
        key: item.key,
        title: item.data.title || '(untitled)',
        authors: authors,
        year: year ? year[0] : '',
        itemType: item.data.itemType,
      };
    });
}

/**
 * Called from Sidebar.html when the user clicks "Import" on a result row.
 * Re-fetches the full item (with formatted bibliography) and adds it as a
 * new Sheet row, same as the paste-a-link flow.
 */
function importZoteroItem(libraryType, libraryId, itemKey) {
  const apiKey = getProp('ZOTERO_API_KEY');
  const url =
    'https://api.zotero.org/' + libraryType + '/' + libraryId + '/items/' + itemKey +
    '?format=json&include=data,bib&style=' + CITATION_STYLE;
  const res = UrlFetchApp.fetch(url, {
    headers: { 'Zotero-API-Key': apiKey },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Zotero API error ' + res.getResponseCode() + ' -- could not fetch that item.');
  }
  const item = JSON.parse(res.getContentText());

  // Personal-library items get a "personal:" prefix so the group sync never matches
  // or overwrites them. Group-library items keep their raw key.
  const keyPrefix = libraryType === 'users' ? 'personal:' : '';
  const uri = 'https://www.zotero.org/' + libraryType + '/' + libraryId + '/items/' + itemKey;

  const result = addZoteroItemToSheet(item, keyPrefix, uri);
  return {
    status: result.status, // 'added' or 'duplicate'
    title: item.data.title || '(untitled)',
  };
}

// ---------------------------------------------------------------------------
// COLUMN MANAGER: HIDE/SHOW ZOTERO-SOURCED COLUMNS FROM A CHECKBOX DIALOG
// ---------------------------------------------------------------------------

function showColumnManager() {
  const html = HtmlService.createHtmlOutputFromFile('ColumnManager')
    .setWidth(340)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage visible columns');
}

/** Called from ColumnManager.html to build the checkbox list + current state. */
function getColumnVisibilityState() {
  const sheet = getSheet();
  const colIndex = getColumnIndexMap(sheet);
  return {
    fields: ZOTERO_SOURCED_COLUMNS.map(function (field) {
      let hidden = false;
      try {
        hidden = sheet.isColumnHiddenByUser(colIndex[field] + 1);
      } catch (e) {
        hidden = false;
      }
      return { field: field, hidden: hidden };
    }),
    presets: Object.keys(COLUMN_PRESETS),
  };
}

/** Called from ColumnManager.html "Apply" button. hiddenFields = array of field names to hide. */
function applyColumnVisibility(hiddenFields) {
  const sheet = getSheet();
  const colIndex = getColumnIndexMap(sheet);
  const hiddenSet = {};
  (hiddenFields || []).forEach(function (f) { hiddenSet[f] = true; });

  ZOTERO_SOURCED_COLUMNS.forEach(function (field) {
    const col = colIndex[field] + 1; // 1-based
    if (hiddenSet[field]) {
      sheet.hideColumns(col);
    } else {
      sheet.showColumns(col);
    }
  });
  return getColumnVisibilityState();
}

/** Called from ColumnManager.html when a preset button is clicked. */
function applyColumnPreset(presetName) {
  const preset = COLUMN_PRESETS[presetName];
  if (!preset) throw new Error('Unknown preset: ' + presetName);
  const toHide = ZOTERO_SOURCED_COLUMNS.filter(function (field) {
    return preset.indexOf(field) === -1;
  });
  return applyColumnVisibility(toHide);
}

// ---------------------------------------------------------------------------
// NATIVE TABLE: WRAP THE REFERENCES RANGE IN A GOOGLE SHEETS TABLE
// Gives built-in per-column sort/filter dropdowns in the sheet UI, on top of
// (not instead of) the column manager above.
//
// Requires the "Google Sheets API" advanced service to be enabled:
// Apps Script editor > Services (+) > Google Sheets API > Add.
// ---------------------------------------------------------------------------

function convertReferencesRangeToTable() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = ensureReferencesTable();
    if (result && result.action === 'skipped') {
      ui.alert('Nothing to do -- there are no data rows yet (' + result.reason + ').');
      return;
    }
    ui.alert(
      (result.action === 'created' ? 'Created' : 'Updated') + ' the "' + TABLE_NAME + '" Table' +
        (result.tableId ? ' (id ' + result.tableId + ')' : '') + '.\n\n' +
      'IMPORTANT: this was written via the Sheets REST API, which does NOT push live into ' +
      'a tab you already had open. If you don\'t see it, reload this browser tab now -- ' +
      'the Table will be there after refresh.'
    );
  } catch (e) {
    ui.alert(
      'Could not create the Table: ' + e.message + '\n\n' +
      'Make sure the "Google Sheets API" advanced service is enabled first: ' +
      'Apps Script editor > Services (+ icon) > Google Sheets API > Add.'
    );
  }
}

/**
 * Creates the References Table if it doesn't exist yet, or resizes it to cover
 * all current rows/columns. Returns details about what happened so callers can
 * show a real confirmation instead of a blind "Done".
 */
function ensureReferencesTable(optSheet) {
  const sheet = optSheet || getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { action: 'skipped', reason: 'no data rows yet' };

  const sheetId = sheet.getSheetId();
  const range = {
    sheetId: sheetId,
    startRowIndex: 0,
    endRowIndex: lastRow,
    startColumnIndex: 0,
    endColumnIndex: lastCol,
  };

  const meta = Sheets.Spreadsheets.get(ss.getId());
  const targetSheetMeta = (meta.sheets || []).filter(function (s) {
    return s.properties.sheetId === sheetId;
  })[0];
  const existingTable = targetSheetMeta && (targetSheetMeta.tables || []).filter(function (t) {
    return t.name === TABLE_NAME;
  })[0];

  const requests = existingTable
    ? [{ updateTable: { table: { tableId: existingTable.tableId, range: range }, fields: 'range' } }]
    : [{ addTable: { table: { name: TABLE_NAME, range: range } } }];

  const response = Sheets.Spreadsheets.batchUpdate({ requests: requests }, ss.getId());
  Logger.log('ensureReferencesTable response: ' + JSON.stringify(response));

  const reply = response && response.replies && response.replies[0];
  const returnedTableId =
    (reply && reply.addTable && reply.addTable.table && reply.addTable.table.tableId) ||
    (existingTable && existingTable.tableId) ||
    null;

  return {
    action: existingTable ? 'updated' : 'created',
    tableId: returnedTableId,
    range: range,
  };
}

// ---------------------------------------------------------------------------
// MANUAL IMPORT (FALLBACK): PASTE A ZOTERO WEB LINK TO IMPORT ONE ITEM
// (Sheet only -- does NOT copy the item into the shared Zotero group.)
// ---------------------------------------------------------------------------

function importReferenceFromZoteroLink() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    'Import from Zotero',
    'Tip: the "Browse & import references..." menu item lets you search instead of pasting a link.\n\n' +
      'In a browser, open zotero.org/mylibrary, find the reference, open its "..." menu ' +
      'and choose "Copy Item Link" (or "Copy Persistent Link"). Paste that link below:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const pasted = resp.getResponseText().trim();
  if (!pasted) return;

  const match = pasted.match(/zotero\.org\/(users|groups)\/(\d+)\/items\/([A-Za-z0-9]+)/i);
  if (!match) {
    ui.alert(
      'Could not find a valid Zotero item link in that text. It should look like:\n' +
      'https://www.zotero.org/users/12345/items/ABCD1234'
    );
    return;
  }
  const libraryType = match[1].toLowerCase(); // 'users' or 'groups'
  const libraryId = match[2];
  const itemKey = match[3];

  let result;
  try {
    result = importZoteroItem(libraryType, libraryId, itemKey);
  } catch (e) {
    ui.alert(e.message);
    return;
  }
  ui.alert(
    result.status === 'duplicate'
      ? 'That reference is already in the Sheet.'
      : 'Imported "' + result.title + '" into the Sheet.'
  );
}

/** Adds one Zotero item (from either library type) as a new Sheet row only. */
function addZoteroItemToSheet(item, keyPrefix, uriOverride) {
  const sheet = getSheet();
  const colIndex = getColumnIndexMap(sheet);
  const zoteroKey = keyPrefix + item.key;

  const existing = getExistingRowsByZoteroKey(sheet, colIndex);
  if (existing[zoteroKey]) {
    return { status: 'duplicate' };
  }

  const values = buildRowFromZoteroItem(item, null, null, null);
  values['Zotero Key'] = zoteroKey;
  values['Zotero URI'] = uriOverride;

  const row = buildFullRowArray(colIndex, values);
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, HEADERS.length).setValues([row]);
  try { ensureReferencesTable(sheet); } catch (e) { Logger.log('Skipped table expand: ' + e); }
  return { status: 'added', rowNum: startRow };
}

// ---------------------------------------------------------------------------
// STEP 1: PULL FROM ZOTERO -> UPSERT INTO SHEET
// ---------------------------------------------------------------------------

function syncZoteroToSheet() {
  const apiKey = getProp('ZOTERO_API_KEY');
  const groupId = getProp('ZOTERO_GROUP_ID', true);
  if (!groupId) {
    SpreadsheetApp.getUi().alert('No ZOTERO_GROUP_ID set — skipping sync. Use "Browse & import references..." to add items manually.');
    return;
  }
  const items = fetchAllZoteroItems(apiKey, 'groups', groupId);
  syncItemsToSheet(getSheet(), items, apiKey, 'groups', groupId);
}

function fetchAllZoteroItems(apiKey, libraryType, libraryId) {
  const items = [];
  const pageSize = 100;
  let start = 0;
  while (true) {
    const url =
      'https://api.zotero.org/' + libraryType + '/' + libraryId + '/items/top' +
      '?format=json&include=data,bib,citation&style=' + CITATION_STYLE +
      '&limit=' + pageSize + '&start=' + start;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'Zotero-API-Key': apiKey },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      throw new Error('Zotero API error ' + res.getResponseCode() + ': ' + res.getContentText());
    }
    const batch = JSON.parse(res.getContentText());
    items.push.apply(items, batch);
    if (batch.length < pageSize) break;
    start += pageSize;
  }
  return items;
}

function buildRowFromZoteroItem(item, apiKey, libraryType, libraryId) {
  const d = item.data;
  const authors = (d.creators || [])
    .filter(function (c) { return c.creatorType === 'author'; })
    .map(formatCreatorName)
    .join('; ');
  const editors = (d.creators || [])
    .filter(function (c) { return c.creatorType === 'editor'; })
    .map(formatCreatorName)
    .join('; ');

  const year = (d.date || '').match(/\d{4}/);
  const citationKey = extractCitationKeyFromExtra(d.extra);
  const bib = item.bib ? stripHtml(item.bib) : '';

  return {
    'Zotero Key': item.key,
    'Title': d.title || '',
    'URL': d.url || '',
    'Authors': authors,
    'Editors': editors,
    'Year': year ? year[0] : '',
    'Date': d.date || '',
    'Date Added': d.dateAdded || '',
    'Date Modified': d.dateModified || '',
    'Publication': d.publicationTitle || '',
    'Proceedings Title': d.proceedingsTitle || '',
    'Series Title': d.seriesTitle || '',
    'Short Title': d.shortTitle || '',
    'Place': d.place || '',
    'DOI': d.DOI || '',
    'Abstract': d.abstractNote || '',
    'Citation Key': citationKey,
    'Extra': d.extra || '',
    'Full Citation': bib,
    'Tags': (d.tags || []).map(function (t) { return t.tag; }).join(', '),
    'Collections': (d.collections || []).join(', '),
    'Zotero URI': libraryType && libraryId
      ? 'https://www.zotero.org/' + libraryType + '/' + libraryId + '/items/' + item.key
      : '',
  };
}

function formatCreatorName(c) {
  if (c.name) return c.name;
  return [c.firstName, c.lastName].filter(Boolean).join(' ');
}

function extractCitationKeyFromExtra(extra) {
  if (!extra) return '';
  const m = extra.match(/Citation Key:\s*(\S+)/i);
  return m ? m[1] : '';
}

function stripHtml(html) {
  return html.replace(/<[^>]+>/g, '').trim();
}

// ---------------------------------------------------------------------------
// STEP 2: CREATE A GOOGLE DOC FOR EVERY ROW THAT DOESN'T HAVE ONE YET
// ---------------------------------------------------------------------------

function createDocsForNewRows(optSheet) {
  const sheet = optSheet || getSheet();
  const colIndex = getColumnIndexMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const folderId = getProp('DOCS_FOLDER_ID', true);
  var folder = null;
  if (folderId) {
    var rootFolder = DriveApp.getFolderById(folderId);
    var sheetName = sheet.getName();
    var subfolders = rootFolder.getFoldersByName(sheetName);
    folder = subfolders.hasNext() ? subfolders.next() : rootFolder.createFolder(sheetName);
  }
  const docLinkCol = colIndex['Doc Link'] + 1;

  // Collect {rowNum, url} for every Doc created this run, then write them all
  // as smart chips in one batched Sheets API call at the end (falls back to
  // plain URLs per-row if the advanced service isn't enabled/available).
  const newDocLinks = [];

  data.forEach(function (row, i) {
    const rowNum = i + 2;
    const docLink = row[colIndex['Doc Link']];
    if (docLink) return; // already has a doc

    const title = row[colIndex['Title']] || '(untitled reference)';
    if (!row[colIndex['Title']]) return; // no title, nothing to create
    const doc = DocumentApp.create(title);
    if (folder) {
      const file = DriveApp.getFileById(doc.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file); // move out of My Drive root
    }
    populateDocTemplate(doc, row, colIndex);

    newDocLinks.push({ rowNum: rowNum, url: doc.getUrl() });
  });

  writeDocLinkChips(sheet, docLinkCol, newDocLinks);
}

/**
 * Writes Doc Link URLs as Google Sheets smart chips (the pill-shaped preview
 * with the file's icon/title) instead of raw URL text. Requires the "Google
 * Sheets API" advanced service; falls back to plain setValue() per row if
 * that service isn't enabled, so Doc creation never breaks because of this.
 */
function writeDocLinkChips(sheet, docLinkCol, entries) {
  if (!entries || entries.length === 0) return;
  const sheetId = sheet.getSheetId();

  try {
    const requests = entries.map(function (entry) {
      return {
        updateCells: {
          rows: [{
            values: [{
              userEnteredValue: { stringValue: '@' },
              chipRuns: [{ chip: { richLinkProperties: { uri: entry.url } } }],
            }],
          }],
          fields: 'userEnteredValue,chipRuns',
          start: { sheetId: sheetId, rowIndex: entry.rowNum - 1, columnIndex: docLinkCol - 1 },
        },
      };
    });
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, SpreadsheetApp.getActiveSpreadsheet().getId());
  } catch (e) {
    Logger.log('Falling back to plain Doc Link URLs (Sheets advanced service not enabled?): ' + e);
    entries.forEach(function (entry) {
      sheet.getRange(entry.rowNum, docLinkCol).setValue(entry.url);
    });
  }
}

function populateDocTemplate(doc, row, colIndex) {
  const body = doc.getBody();
  body.appendParagraph(row[colIndex['Title']] || '(untitled)').setHeading(DocumentApp.ParagraphHeading.TITLE);

  const meta = body.appendParagraph('');
  meta.appendText('Authors: ' + (row[colIndex['Authors']] || '') + '\n');
  meta.appendText('Year: ' + (row[colIndex['Year']] || '') + '\n');
  meta.appendText('URL: ' + (row[colIndex['URL']] || '') + '\n');
  meta.appendText('DOI: ' + (row[colIndex['DOI']] || ''));

  if (row[colIndex['Abstract']]) {
    body.appendParagraph('Abstract').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(row[colIndex['Abstract']]);
  }

  // One editable section per manually-synced field, so syncDocsToSheet() can find them.
  DOC_SYNCED_FIELDS.forEach(function (field) {
    body.appendParagraph(field).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(row[colIndex[field]] || '');
  });

  doc.saveAndClose();
}

// ---------------------------------------------------------------------------
// STEP 3: READ THE MANUAL SECTIONS BACK OUT OF EACH DOC INTO THE SHEET
// ---------------------------------------------------------------------------

function syncDocsToSheet(optSheet) {
  const sheet = optSheet || getSheet();
  const colIndex = getColumnIndexMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  data.forEach(function (row, i) {
    const rowNum = i + 2;
    const docUrl = row[colIndex['Doc Link']];
    if (!docUrl) return;

    const docId = extractDocIdFromUrl(docUrl);
    if (!docId) return;

    let doc;
    try {
      doc = DocumentApp.openById(docId);
    } catch (e) {
      Logger.log('Could not open doc for row ' + rowNum + ': ' + e);
      return;
    }

    const sections = extractHeadingSections(doc.getBody(), DOC_SYNCED_FIELDS);
    DOC_SYNCED_FIELDS.forEach(function (field) {
      if (sections[field] === undefined) return;
      const col = colIndex[field] + 1;
      const current = sheet.getRange(rowNum, col).getValue();
      if (sections[field] !== current) {
        sheet.getRange(rowNum, col).setValue(sections[field]);
      }
    });
  });
}

function extractDocIdFromUrl(url) {
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

/**
 * Walks the doc body and, for each Heading2 whose text matches one of
 * `fieldNames`, collects all following paragraph text until the next
 * Heading2 (or Heading1/TITLE), and returns { fieldName: text }.
 */
function extractHeadingSections(body, fieldNames) {
  const result = {};
  const numChildren = body.getNumChildren();
  let currentField = null;
  let buffer = [];

  function flush() {
    if (currentField) {
      result[currentField] = buffer.join('\n').trim();
    }
    buffer = [];
  }

  for (let i = 0; i < numChildren; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = el.asParagraph();
    const heading = para.getHeading();
    const text = para.getText();

    if (heading === DocumentApp.ParagraphHeading.HEADING2) {
      flush();
      currentField = fieldNames.indexOf(text.trim()) !== -1 ? text.trim() : null;
      continue;
    }
    if (heading === DocumentApp.ParagraphHeading.HEADING1 || heading === DocumentApp.ParagraphHeading.TITLE) {
      flush();
      currentField = null;
      continue;
    }
    if (currentField) buffer.push(text);
  }
  flush();
  return result;
}

// ---------------------------------------------------------------------------
// SCRIPT STORAGE: INSPECT / WIPE SCRIPT PROPERTIES + CACHE
// (This is the ONLY persistent storage this script uses outside the Sheet
// itself and the Google Docs it creates -- no database, no other backend.)
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// SHARED HELPERS
// ---------------------------------------------------------------------------

function getSheet() {
  return SpreadsheetApp.getActiveSheet();
}

function getProp(name, optional) {
  const val = PropertiesService.getScriptProperties().getProperty(name);
  if (!val && !optional) {
    throw new Error('Missing Script Property "' + name + '". Set it under Project Settings > Script Properties.');
  }
  return val;
}

function getColumnIndexMap(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn() || HEADERS.length).getValues()[0];
  const map = {};
  headerRow.forEach(function (h, i) {
    if (h) map[h] = i;
  });
  HEADERS.forEach(function (h) {
    if (map[h] === undefined) {
      throw new Error('Sheet is missing expected column "' + h + '". Run setupSheet() or add it manually.');
    }
  });
  return map;
}

function getExistingRowsByZoteroKey(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < 2) return result;
  const keys = sheet.getRange(2, colIndex['Zotero Key'] + 1, lastRow - 1, 1).getValues();
  keys.forEach(function (r, i) {
    if (r[0]) result[r[0]] = { rowNum: i + 2 };
  });
  return result;
}

/** Writes only the Zotero-sourced columns for an existing row (never touches MANUAL_COLUMNS or Doc Link). */
function writeZoteroColumnsToRow(sheet, rowNum, colIndex, values) {
  Object.keys(values).forEach(function (field) {
    if (field === 'Zotero Key') return; // key itself doesn't need rewriting
    if (MANUAL_COLUMNS.indexOf(field) !== -1) return; // never touch user-owned columns
    const col = colIndex[field];
    if (col === undefined) return;
    sheet.getRange(rowNum, col + 1).setValue(values[field]);
  });
}

/** Builds a full row array (in HEADERS order) for a brand-new row, manual columns left blank. */
function buildFullRowArray(colIndex, values) {
  const row = new Array(HEADERS.length).fill('');
  Object.keys(values).forEach(function (field) {
    const col = colIndex[field];
    if (col !== undefined) row[col] = values[field];
  });
  return row;
}
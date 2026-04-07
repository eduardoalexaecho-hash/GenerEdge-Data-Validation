/**
 * =============================================================================
 * CRM CALLERS READY GENERATOR - V5.3
 * =============================================================================
 *
 * Creates a clean calling list directly from the active validated sheet,
 * plus an "Excluded Contacts" report sheet for every contact that didn't qualify.
 *
 * FILTERING CRITERIA (both required):
 * - At least one valid email (any of: Email 1, Email 2, Personal Email, Primary Email)
 * - At least one phone that is VALID_CONFIRMED + MOBILE line type
 *
 * EMAIL PRIORITY (first valid one wins):
 *   1. Email 1
 *   2. Email 2
 *   3. Personal Email
 *   4. Primary Email
 *   Note: Only email columns that have a _Status column are eligible.
 *         Status must be exactly "valid" — anything else is rejected.
 *
 * PHONE FALLBACK CHAIN (first passing one wins → single "phone" column):
 *   1. Contact Mobile Phone
 *   2. Contact Phone 1
 *   3. Company Phone 1
 *   4. Company Phone 2
 *
 * STATE RESOLUTION PRIORITY (per row):
 *   1. "Company State" column (full name → abbreviation, e.g. Arizona → AZ)
 *   2. "Company City" already contains an embedded abbreviation (e.g. "Phoenix, AZ")
 *   3. City database lookup in 08-Citystagedatabase.gs
 *      - Unique   → use that abbreviation
 *      - Ambiguous → write "AMBIGUOUS: AZ, CA" for manual review
 *      - Not found → leave blank
 *
 * OUTPUT — CRM Callers Ready (7 columns):
 *   first_name, last_name, organization, email, phone, city, state
 *
 * OUTPUT — Included Contacts (9 columns):
 *   first_name, last_name, organization, email, email_source,
 *   phone, phone_source, city, state
 *   email_source → e.g. "Email 1 — valid"
 *   phone_source → e.g. "Contact Mobile Phone — VALID_CONFIRMED / MOBILE"
 *
 * OUTPUT — Excluded Contacts (dynamic columns):
 *   first_name, last_name, organization,
 *   [each validated email col + _Status],
 *   [each validated phone col + _Status + _Line Type],
 *   exclusion_reason
 *
 * REQUIRES: 08-Citystagedatabase.gs in the same Apps Script project
 * MENU: 🔧 Data Tools → 🎯 CRM Ready → 📞 Generate CRM Callers Ready
 *
 * @author: Claude
 * @version: 5.3 - Included Contacts report added
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CALLERS_CONFIG = {
  OUTPUT_SHEET:   'CRM Callers Ready',
  EXCLUDED_SHEET: 'Excluded Contacts',
  INCLUDED_SHEET: 'Included Contacts',

  // Email columns in priority order (first valid one is used)
  EMAIL_PRIORITY: [
    'Email 1',
    'Email 2',
    'Personal Email'
  ],

  // Phone fallback chain in priority order (first passing one is used)
  PHONE_CHAIN: [
    'Contact Mobile Phone',
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2'
  ],

  // Source column names → output column names for base fields
  BASE_COLUMNS: {
    'Contact Full Name': 'contact_full_name',
    'First Name':        'first_name',
    'Last Name':         'last_name',
    'Organization':      'organization',
    'Company City':  'city',
    'Address 1':     'address_1',
    'Address 2':     'address_2',
    'County':        'county',
    'Zip':           'zip',
    'Gender':        'gender',
    'Race/Ethnicity':'race_ethnicity',
    'Industry':      'industry'
  },

  // State source columns
  STATE_COLUMN:  'Company State',
  CITY_COLUMN:   'Company City',

  // Validation column suffixes
  STATUS_SUFFIX:    '_Status',
  LINE_TYPE_SUFFIX: '_Line Type',

  // Valid email status (ZeroBounce)
  VALID_EMAIL_STATUS: 'valid',

  // Valid phone criteria (Byteplant)
  VALID_PHONE_STATUS:    'VALID_CONFIRMED',
  VALID_PHONE_LINE_TYPE: 'MOBILE',

  // SKIP marker written by phone validation — treat as no result
  SKIP_MARKER: 'SKIP',

  // Final output column order for CRM Callers Ready
  OUTPUT_COLUMNS: [
    'first_name', 'last_name', 'contact_full_name', 'organization',
    'email', 'phone',
    'city', 'state',
    'address_1', 'address_2', 'county', 'zip',
    'gender', 'race_ethnicity', 'industry'
  ],

  // Header colors
  CALLERS_HEADER_BG:   '#4285f4',  // Blue       — CRM Callers Ready
  INCLUDED_HEADER_BG:  '#0f9d58',  // Green      — Included Contacts
  SOURCE_COL_BG:       '#e6f4ea',  // Light green — email_source / phone_source cells
  SOURCE_HEADER_BG:    '#137333',  // Dark green  — email_source / phone_source headers
  EXCLUDED_HEADER_BG:  '#ea4335',  // Red        — Excluded Contacts
  REASON_COL_BG:       '#fce8e6',  // Light red   — exclusion_reason cells
  REASON_HEADER_BG:    '#c62828'   // Dark red    — exclusion_reason header
};

// =============================================================================
// MAIN FUNCTION
// =============================================================================

/**
 * Entry point. Reads the active sheet, builds both output sheets.
 */
function generateCRMCallersReady() {
  const startTime = new Date();

  try {
    Logger.log('=== CRM Callers Ready V5.3 ===');

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getActiveSheet();

    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert('Error', 'No active sheet found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    Logger.log(`Source: "${sourceSheet.getName()}"`);

    const lastRow = sourceSheet.getLastRow();
    const lastCol = sourceSheet.getLastColumn();

    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('Error', 'The active sheet has no data rows.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Read everything at once for performance
    const allData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = allData[0].map(h => h.toString().trim());

    // Build column index map
    const colMap = buildColumnMap_(headers);

    if (!colMap) {
      SpreadsheetApp.getUi().alert(
        'Missing Columns',
        'Could not find the required columns in the active sheet.\n\n' +
        'Required:\n' +
        '• At least one validated email column (must have a _Status column)\n' +
        '• At least one validated phone column (must have _Status and _Line Type columns)\n' +
        '• First Name, Last Name, Organization',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }

    // Process all rows — splits into qualified and excluded
    const result = processRows_(allData, colMap);

    // Always write both reports (even if 0 rows, so sheets always exist)
    writeOutputSheet_(ss, result.rows);
    writeIncludedSheet_(ss, result.rows, colMap);
    writeExcludedSheet_(ss, result.excluded, colMap);

    const duration = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log(`=== Done — ${result.rows.length} included, ${result.excluded.length} excluded (${duration}s) ===`);

    const s = result.stats;
    SpreadsheetApp.getUi().alert(
      '✅ CRM Callers Ready — Complete',
      `Processing finished!\n\n` +
      `📊 Results:\n` +
      `  Total rows scanned:             ${s.total}\n` +
      `  ✅ Included in Callers Ready:   ${result.rows.length}\n` +
      `  ❌ Excluded (see report sheet): ${result.excluded.length}\n\n` +
      `❌ Exclusion breakdown:\n` +
      `  No valid email:                 ${s.noEmail}\n` +
      `  No valid mobile phone:          ${s.noPhone}\n` +
      `  Missing both:                   ${s.noBoth}\n\n` +
      `🗺️  State resolution:\n` +
      `  From State column:              ${s.stateFromCol}\n` +
      `  From City (embedded abbr):      ${s.stateFromCity}\n` +
      `  From City (DB lookup):          ${s.stateFromDb}\n` +
      `  Ambiguous city:                 ${s.stateAmbiguous}\n` +
      `  Not found:                      ${s.stateNotFound}\n\n` +
      `  Duration: ${duration}s\n\n` +
      `Sheets created:\n` +
      `  • "${CALLERS_CONFIG.OUTPUT_SHEET}"\n` +
      `  • "${CALLERS_CONFIG.INCLUDED_SHEET}"\n` +
      `  • "${CALLERS_CONFIG.EXCLUDED_SHEET}"`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    // Switch to callers ready if it exists, otherwise show excluded
    const callersSheet = ss.getSheetByName(CALLERS_CONFIG.OUTPUT_SHEET);
    ss.setActiveSheet(callersSheet || ss.getSheetByName(CALLERS_CONFIG.EXCLUDED_SHEET));

  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'Failed to generate CRM Callers Ready:\n\n' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

// =============================================================================
// COLUMN MAP BUILDER
// =============================================================================

/**
 * Builds a map of all relevant column indices from the header row.
 * Returns null if minimum required columns are missing.
 *
 * @param {string[]} headers
 * @returns {object|null}
 */
function buildColumnMap_(headers) {
  const map = {
    base:     {},   // { outputName: colIndex }
    stateIdx: -1,
    cityIdx:  -1,
    emails:   [],   // [{ name, dataIdx, statusIdx }] — only validated columns
    phones:   []    // [{ name, dataIdx, statusIdx, lineTypeIdx }] — only validated columns
  };

  // Base columns
  for (const [sourceName, outputName] of Object.entries(CALLERS_CONFIG.BASE_COLUMNS)) {
    const idx = headers.findIndex(h => h.toLowerCase() === sourceName.toLowerCase());
    if (idx !== -1) map.base[outputName] = idx;
  }

  // State column (optional)
  map.stateIdx = headers.findIndex(h => h.toLowerCase() === CALLERS_CONFIG.STATE_COLUMN.toLowerCase());

  // City column
  map.cityIdx = headers.findIndex(h => h.toLowerCase() === CALLERS_CONFIG.CITY_COLUMN.toLowerCase());

  // Email columns — only include if a _Status column exists (i.e. was validated)
  for (const colName of CALLERS_CONFIG.EMAIL_PRIORITY) {
    const dataIdx = headers.findIndex(h => h.toLowerCase() === colName.toLowerCase());
    if (dataIdx === -1) continue;

    const statusIdx = headers.findIndex(h =>
      h.toLowerCase() === (colName + CALLERS_CONFIG.STATUS_SUFFIX).toLowerCase()
    );

    if (statusIdx === -1) {
      Logger.log(`Email "${colName}" has no _Status column — skipping`);
      continue;
    }

    map.emails.push({ name: colName, dataIdx, statusIdx });
    Logger.log(`Email "${colName}" included (statusIdx: ${statusIdx})`);
  }

  // Phone columns — only include if both _Status and _Line Type columns exist
  for (const colName of CALLERS_CONFIG.PHONE_CHAIN) {
    const dataIdx     = headers.findIndex(h => h.toLowerCase() === colName.toLowerCase());
    const statusIdx   = headers.findIndex(h => h.toLowerCase() === (colName + CALLERS_CONFIG.STATUS_SUFFIX).toLowerCase());
    const lineTypeIdx = headers.findIndex(h => h.toLowerCase() === (colName + CALLERS_CONFIG.LINE_TYPE_SUFFIX).toLowerCase());

    if (dataIdx !== -1 && statusIdx !== -1 && lineTypeIdx !== -1) {
      map.phones.push({ name: colName, dataIdx, statusIdx, lineTypeIdx });
      Logger.log(`Phone "${colName}" included`);
    } else {
      Logger.log(`Phone "${colName}" skipped — missing validation columns`);
    }
  }

  if (map.emails.length === 0 || map.phones.length === 0) {
    Logger.log('Column map failed: no validated email or phone columns found');
    return null;
  }

  return map;
}

// =============================================================================
// ROW PROCESSING
// =============================================================================

/**
 * Iterates all data rows. Returns qualified rows and excluded rows separately.
 *
 * @param {any[][]} allData
 * @param {object}  colMap
 * @returns {{ rows: object[], excluded: object[], stats: object }}
 */
function processRows_(allData, colMap) {
  const outputRows = [];
  const excludedRows = [];
  let noEmail = 0, noPhone = 0, noBoth = 0;
  let stateFromCol = 0, stateNotFound = 0;
  const total = allData.length - 1;

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];

    // Skip completely empty rows
    if (row.every(cell => cell === '' || cell === null || cell === undefined)) continue;

    const emailResult = pickEmail_(row, colMap.emails);   // { value, detail } or null
    const phoneResult = pickPhone_(row, colMap.phones);   // { value, detail } or null

    const hasEmail = emailResult !== null;
    const hasPhone = phoneResult !== null;

    if (!hasEmail || !hasPhone) {
      // Build exclusion reason with detail
      const reason = buildExclusionReason_(row, colMap, hasEmail, hasPhone);

      // Build excluded row object
      const exc = buildBaseFields_(row, colMap);
      exc['exclusion_reason'] = reason.summary;
      exc['_emailDetail']     = reason.emailDetail;   // used by writeExcludedSheet_
      exc['_phoneDetail']     = reason.phoneDetail;   // used by writeExcludedSheet_

      // Store raw per-column values so writeExcludedSheet_ can build individual cells
      exc['_emailCols'] = colMap.emails.map(col => ({
        name:   col.name,
        value:  allData[i][col.dataIdx]   !== undefined ? allData[i][col.dataIdx].toString().trim()   : '',
        status: allData[i][col.statusIdx] !== undefined ? allData[i][col.statusIdx].toString().trim() : ''
      }));
      exc['_phoneCols'] = colMap.phones.map(col => ({
        name:     col.name,
        value:    allData[i][col.dataIdx]     !== undefined ? allData[i][col.dataIdx].toString().trim()     : '',
        status:   allData[i][col.statusIdx]   !== undefined ? allData[i][col.statusIdx].toString().trim()   : '',
        lineType: allData[i][col.lineTypeIdx] !== undefined ? allData[i][col.lineTypeIdx].toString().trim() : ''
      }));

      excludedRows.push(exc);

      if (!hasEmail && !hasPhone) noBoth++;
      else if (!hasEmail)         noEmail++;
      else                        noPhone++;
      continue;
    }

    // ── Qualified row ────────────────────────────────────────────────────────
    const out = buildBaseFields_(row, colMap);
    out['email'] = emailResult.value;
    out['phone'] = phoneResult.value;

    const stateResult = resolveState_(row, colMap);
    out['state'] = stateResult.value;

    if (stateResult.source === 'col') stateFromCol++;
    else                              stateNotFound++;

    outputRows.push(out);
  }

  Logger.log(`Processed ${total} rows — included: ${outputRows.length}, excluded: ${excludedRows.length}`);

  return {
    rows:     outputRows,
    excluded: excludedRows,
    stats:    { total, noEmail, noPhone, noBoth, stateFromCol, stateNotFound }
  };
}

/**
 * Builds the base field object (first_name, last_name, organization, city) from a row.
 *
 * @param {any[]}  row
 * @param {object} colMap
 * @returns {object}
 */
function buildBaseFields_(row, colMap) {
  const out = {};
  for (const [outputName, idx] of Object.entries(colMap.base)) {
    out[outputName] = row[idx] !== undefined ? row[idx].toString().trim() : '';
  }

  // If contact_full_name is empty (column missing or blank cell),
  // build it from first_name + last_name
  if (!out['contact_full_name']) {
    const first = (out['first_name'] || '').trim();
    const last  = (out['last_name']  || '').trim();
    out['contact_full_name'] = [first, last].filter(Boolean).join(' ');
  }
  return out;
}

/**
 * Builds the exclusion reason and per-column detail for a rejected row.
 *
 * exclusion_reason values (simple, filterable):
 *   "No valid email"                      — has phone but no valid email
 *   "No valid mobile phone"               — has email but no valid phone
 *   "No valid email | No valid mobile phone" — missing both
 *
 * emailDetail / phoneDetail strings are still built for the individual
 * detail columns in the Excluded Contacts sheet.
 *
 * @param {any[]}   row
 * @param {object}  colMap
 * @param {boolean} hasEmail
 * @param {boolean} hasPhone
 * @returns {{ summary: string, emailDetail: string, phoneDetail: string }}
 */
function buildExclusionReason_(row, colMap, hasEmail, hasPhone) {
  // Email detail — for the individual _Status columns in the report
  const emailParts = colMap.emails.map(col => {
    const value  = row[col.dataIdx]   !== undefined ? row[col.dataIdx].toString().trim()                 : '';
    const status = row[col.statusIdx] !== undefined ? row[col.statusIdx].toString().trim().toLowerCase() : '';
    if (!value)  return `${col.name}: (empty)`;
    if (!status) return `${col.name}: (not validated)`;
    return `${col.name}: ${status}`;
  });

  // Phone detail — for the individual _Status / _Line Type columns in the report
  const phoneParts = colMap.phones.map(col => {
    const value    = row[col.dataIdx]     !== undefined ? row[col.dataIdx].toString().trim()   : '';
    const status   = row[col.statusIdx]   !== undefined ? row[col.statusIdx].toString().trim() : '';
    const lineType = row[col.lineTypeIdx] !== undefined ? row[col.lineTypeIdx].toString().trim() : '';
    if (!value)                                return `${col.name}: (empty)`;
    if (status === CALLERS_CONFIG.SKIP_MARKER) return `${col.name}: SKIP`;
    if (!status)                               return `${col.name}: (not validated)`;
    return `${col.name}: ${status}${lineType ? ' / ' + lineType : ''}`;
  });

  const emailDetail = emailParts.join(' | ');
  const phoneDetail = phoneParts.join(' | ');

  // Simple filterable summary — no detail, just the category
  let summary = '';
  if (!hasEmail && !hasPhone) {
    summary = 'No valid email | No valid mobile phone';
  } else if (!hasEmail) {
    summary = 'No valid email';
  } else {
    summary = 'No valid mobile phone';
  }

  return { summary, emailDetail, phoneDetail };
}

// =============================================================================
// EMAIL & PHONE PICKERS
// =============================================================================

/**
 * Returns { value, detail } for the first email column that has status "valid",
 * or null if none qualify.
 *
 * @param {any[]}    row
 * @param {object[]} emails
 * @returns {{ value: string, detail: string }|null}
 */
function pickEmail_(row, emails) {
  for (const col of emails) {
    const value  = row[col.dataIdx]   !== undefined ? row[col.dataIdx].toString().trim()                 : '';
    const status = row[col.statusIdx] !== undefined ? row[col.statusIdx].toString().trim().toLowerCase() : '';

    if (!value)  continue;
    if (!status) continue; // Not yet validated — skip

    if (status === CALLERS_CONFIG.VALID_EMAIL_STATUS) {
      return { value, detail: `${col.name} (${status})` };
    }
  }
  return null;
}

/**
 * Returns { value, detail } for the first phone column that passes
 * (VALID_CONFIRMED + MOBILE), or null if none qualify.
 * SKIP marker is treated as no result and the chain continues.
 *
 * @param {any[]}    row
 * @param {object[]} phones
 * @returns {{ value: string, detail: string }|null}
 */
function pickPhone_(row, phones) {
  for (const col of phones) {
    const status   = row[col.statusIdx]   !== undefined ? row[col.statusIdx].toString().trim().toUpperCase()   : '';
    const lineType = row[col.lineTypeIdx] !== undefined ? row[col.lineTypeIdx].toString().trim().toUpperCase() : '';

    if (status === CALLERS_CONFIG.SKIP_MARKER) continue;

    if (status === CALLERS_CONFIG.VALID_PHONE_STATUS && lineType === CALLERS_CONFIG.VALID_PHONE_LINE_TYPE) {
      const number = row[col.dataIdx] !== undefined ? row[col.dataIdx].toString().trim() : '';
      if (number) return { value: number, detail: `${col.name} (${status}/${lineType})` };
    }
  }
  return null;
}

// =============================================================================
// STATE RESOLUTION
// =============================================================================

/**
 * Resolves the state for a qualified row from the Company State column.
 * The value is used as-is (already the correct format from your data).
 * No city database lookup needed since Company State is always present.
 *
 * @param {any[]}  row
 * @param {object} colMap
 * @returns {{ value: string, source: string }}
 */
function resolveState_(row, colMap) {
  if (colMap.stateIdx !== -1) {
    const raw = row[colMap.stateIdx] !== undefined ? row[colMap.stateIdx].toString().trim() : '';
    if (raw) return { value: raw, source: 'col' };
  }
  return { value: '', source: 'notfound' };
}

// =============================================================================
// OUTPUT SHEET — CRM CALLERS READY
// =============================================================================

/**
 * Creates or overwrites the CRM Callers Ready sheet.
 *
 * @param {Spreadsheet} ss
 * @param {object[]}    rows
 */
function writeOutputSheet_(ss, rows) {
  const existing = ss.getSheetByName(CALLERS_CONFIG.OUTPUT_SHEET);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(CALLERS_CONFIG.OUTPUT_SHEET);
  const cols  = CALLERS_CONFIG.OUTPUT_COLUMNS;

  // Header
  sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  sheet.getRange(1, 1, 1, cols.length)
    .setBackground(CALLERS_CONFIG.CALLERS_HEADER_BG)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Data
  if (rows.length > 0) {
    const data = rows.map(r => cols.map(col => r[col] || ''));
    sheet.getRange(2, 1, data.length, cols.length).setValues(data);
  }

  sheet.setFrozenRows(1);
  for (let c = 1; c <= cols.length; c++) sheet.autoResizeColumn(c);

  Logger.log(`CRM Callers Ready: ${rows.length} rows written`);
}


// =============================================================================
// OUTPUT SHEET — INCLUDED CONTACTS REPORT
// =============================================================================

/**
 * Creates or overwrites the Included Contacts report sheet.
 * Uses the same columns as CRM Callers Ready.
 *
 * @param {Spreadsheet} ss
 * @param {object[]}    rows  - Qualified row objects
 * @param {object}      colMap
 */
function writeIncludedSheet_(ss, rows, colMap) {
  const existing = ss.getSheetByName(CALLERS_CONFIG.INCLUDED_SHEET);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(CALLERS_CONFIG.INCLUDED_SHEET);

  // Same columns as CRM Callers Ready — no source columns, just the clean data
  const headers = CALLERS_CONFIG.OUTPUT_COLUMNS;

  // ── Header row ────────────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(CALLERS_CONFIG.INCLUDED_HEADER_BG)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // ── Data rows ─────────────────────────────────────────────────────────────
  if (rows.length > 0) {
    const data = rows.map(r => headers.map(col => r[col] || ''));
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }

  sheet.setFrozenRows(1);
  for (let c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);

  Logger.log(`Included Contacts: ${rows.length} rows written`);
}

// =============================================================================
// OUTPUT SHEET — EXCLUDED CONTACTS
// =============================================================================

/**
 * Creates or overwrites the Excluded Contacts report sheet.
 *
 * Columns (dynamic):
 *   first_name | last_name | organization |
 *   [Email col] | [Email col _Status] | ... (one pair per validated email column) |
 *   [Phone col] | [Phone col _Status] | [Phone col _Line Type] | ... (one trio per validated phone column) |
 *   exclusion_reason
 *
 * @param {Spreadsheet} ss
 * @param {object[]}    excluded  - Excluded row objects with _emailDetail and _phoneDetail
 * @param {object}      colMap
 */
function writeExcludedSheet_(ss, excluded, colMap) {
  const existing = ss.getSheetByName(CALLERS_CONFIG.EXCLUDED_SHEET);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(CALLERS_CONFIG.EXCLUDED_SHEET);

  // ── Build dynamic header list ─────────────────────────────────────────────
  const headers = ['first_name', 'last_name', 'organization'];

  // One column per email column value, one per email _Status
  for (const col of colMap.emails) {
    headers.push(col.name);
    headers.push(col.name + CALLERS_CONFIG.STATUS_SUFFIX);
  }

  // One column per phone column value, _Status, _Line Type
  for (const col of colMap.phones) {
    headers.push(col.name);
    headers.push(col.name + CALLERS_CONFIG.STATUS_SUFFIX);
    headers.push(col.name + CALLERS_CONFIG.LINE_TYPE_SUFFIX);
  }

  headers.push('exclusion_reason');
  const reasonColIndex = headers.length; // 1-based column index for the reason column

  // ── Write header row ──────────────────────────────────────────────────────
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style all headers red
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(CALLERS_CONFIG.EXCLUDED_HEADER_BG)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Style the exclusion_reason header darker red
  sheet.getRange(1, reasonColIndex)
    .setBackground(CALLERS_CONFIG.REASON_HEADER_BG)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // ── Write data rows ───────────────────────────────────────────────────────
  if (excluded.length > 0) {
    // _emailCols and _phoneCols arrays are stored on each excluded object
    // by processRows_() — they hold the raw value, status, and line type
    // for every validated email/phone column for that row.

    const dataRows = excluded.map(exc => {
      const row = [];

      // Base fields
      row.push(exc['first_name']   || '');
      row.push(exc['last_name']    || '');
      row.push(exc['organization'] || '');

      // Email columns: value + status (stored as _emailCols array on exc)
      for (const colData of (exc['_emailCols'] || [])) {
        row.push(colData.value);
        row.push(colData.status);
      }

      // Phone columns: value + status + line type (stored as _phoneCols array on exc)
      for (const colData of (exc['_phoneCols'] || [])) {
        row.push(colData.value);
        row.push(colData.status);
        row.push(colData.lineType);
      }

      // Exclusion reason
      row.push(exc['exclusion_reason'] || '');

      return row;
    });

    sheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

    // Light red background for the exclusion_reason column data cells
    sheet.getRange(2, reasonColIndex, dataRows.length, 1)
      .setBackground(CALLERS_CONFIG.REASON_COL_BG)
      .setFontWeight('bold');
  }

  // ── Format ────────────────────────────────────────────────────────────────
  sheet.setFrozenRows(1);
  for (let c = 1; c <= headers.length; c++) sheet.autoResizeColumn(c);

  Logger.log(`Excluded Contacts: ${excluded.length} rows written`);
}

// =============================================================================
// END OF SCRIPT
// =============================================================================

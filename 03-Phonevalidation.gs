/**
 * =============================================================================
 * BYTEPLANT PHONE VALIDATION SCRIPT
 * =============================================================================
 *
 * Validates phone numbers using Byteplant Real-Time Phone Validation API
 *
 * IMPORTANT: This file should be in the same Apps Script project as:
 *            DataCleaningPipeline_V3.gs and EmailValidation.gs
 *
 * API SPECIFICATION:
 * - Endpoint: https://api.phone-validator.net/api/v2/verify
 * - Method: GET
 * - Required Parameters: PhoneNumber, CountryCode, APIKey
 * - Phone Format: National format (e.g., "2109791958") with CountryCode="us"
 *
 * STATUS VALUES:
 * - VALID_CONFIRMED, VALID_UNCONFIRMED, INVALID, DELAYED,
 *   RATE_LIMIT_EXCEEDED, API_KEY_INVALID_OR_DEPLETED
 *
 * LINE TYPES:
 * - MOBILE, FIXED_LINE, VOIP, TOLL_FREE, PREMIUM_RATE, etc.
 *
 * FALLBACK CHAIN LOGIC:
 * - Contact Mobile Phone: no predecessors — validates every eligible row
 * - Contact Phone 1:  skips row if Contact Mobile Phone passes
 * - Company Phone 1:  skips row if Contact Mobile Phone OR Contact Phone 1 passes
 * - Company Phone 2:  skips row if any of the above 3 pass
 * - "Passes" = VALID_CONFIRMED AND MOBILE
 * - Skipped rows get "SKIP" written in their _Status cell
 *
 * EMAIL PRE-CHECK (all buttons including individual ones):
 * - Before calling the phone API for any row, checks whether Email 1 has
 *   status "valid" for that row
 * - If Email 1 is not valid → writes "NO_EMAIL" to the _Status cell, skips API call
 * - Saves phone validation credits — no point validating phones for invalid emails
 *
 * FULL CHAIN (▶️ Validate All):
 * - Step 1: Email 1       (via EmailValidation.gs — startEmailValidation_)
 * - Step 2: Contact Mobile Phone
 * - Step 3: Contact Phone 1
 * - Step 4: Company Phone 1
 * - Step 5: Company Phone 2
 * - Each step hands off to the next automatically when done
 * - Time-aware continuation works within each step
 *
 * SMART RATE LIMITING:
 * - No fixed delay between API calls
 * - Every Byteplant response includes ratelimit_remain and ratelimit_seconds
 * - Script only waits when ratelimit_remain <= RATE_LIMIT_BUFFER (5 calls left)
 * - If RATE_LIMIT_EXCEEDED hit anyway, automatically waits and retries
 * - Goes as fast as Byteplant allows — ~100 calls per 300 seconds
 *
 * TIME-AWARE AUTO-CONTINUATION:
 * - All sheet data read in ONE batch at start of each run (fast)
 * - Results written immediately after each API call (progress always saved)
 * - When elapsed time approaches 240s, saves position, schedules 1-min trigger
 * - continuePhoneValidation() resumes from saved row automatically
 * - Use "Cancel Running Validation" to stop at any time
 *
 * @author: Claude
 * @version: 1.8 - Smart rate limit throttling (no fixed delay)
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const PHONE_VALIDATION_CONFIG = {
  // Byteplant API
  API_KEY:      'pv-0e4b530a9be44e529631bd5cc323279d',
  API_ENDPOINT: 'https://api.phone-validator.net/api/v2/verify',
  COUNTRY_CODE: 'us',

  // Test Mode
  TEST_MODE:       false,
  TEST_ROWS_START: 2,
  TEST_ROWS_END:   6,

  // Full chain order — Email 1 is handled by EmailValidation.gs as step 1
  // This array defines the phone steps only
  CHAIN_STEPS: [
    'Contact Mobile Phone',
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2'
  ],

  // Fallback chain predecessors per phone column
  FALLBACK_PREDECESSORS: {
    'Contact Phone 1': ['Contact Mobile Phone'],
    'Company Phone 1': ['Contact Mobile Phone', 'Contact Phone 1'],
    'Company Phone 2': ['Contact Mobile Phone', 'Contact Phone 1', 'Company Phone 1']
  },

  // Email column to check before validating any phone (all buttons)
  EMAIL_CHECK_COLUMN: 'Email 1',
  EMAIL_STATUS_SUFFIX: '_Status',
  VALID_EMAIL_STATUS:  'valid',

  // Result column suffixes
  RESULT_COLUMNS: {
    STATUS:    '_Status',
    LINE_TYPE: '_Line Type',
    LOCATION:  '_Location',
    FORMAT:    '_International'
  },

  // Passing criteria
  VALID_STATUS:    'VALID_CONFIRMED',
  VALID_LINE_TYPE: 'MOBILE',

  // Marker values written to _Status
  SKIP_MARKER:     'SKIP',      // Higher-priority phone already passed
  NO_EMAIL_MARKER: 'NO_EMAIL',  // No valid email — phone skipped

  // API settings
  // No fixed delay — script reads ratelimit_remain and ratelimit_seconds from
  // every Byteplant response and only waits when actually near the rate limit.
  // Byteplant allows 100 calls per 300s. We pause when fewer than RATE_LIMIT_BUFFER
  // calls remain in the current window.
  RATE_LIMIT_BUFFER:   5,     // pause when this many calls remain in current window
  REQUEST_TIMEOUT:     10000,
  MAX_RETRIES:         2,
  RETRY_DELAY:         2000,  // ms to wait before retrying after a non-rate-limit error

  // Time-aware
  TIME_LIMIT_SECONDS:         240,
  CONTINUATION_DELAY_MINUTES: 1
};

// PropertiesService keys
const PROPS = {
  COLUMN:     'phoneVal_column',
  NEXT_ROW:   'phoneVal_nextRow',
  END_ROW:    'phoneVal_endRow',
  SHEET_NAME: 'phoneVal_sheetName',
  TRIGGER_ID: 'phoneVal_triggerId',
  IS_RUNNING: 'phoneVal_isRunning',
  CHAIN_NEXT: 'phoneVal_chainNext'  // next phone column in chain after current one finishes
};

// =============================================================================
// PUBLIC ENTRY POINTS (individual menu buttons)
// =============================================================================

function validateContactMobilePhone() { startValidation_('Contact Mobile Phone', null); }
function validateContactPhone1()      { startValidation_('Contact Phone 1',       null); }
function validateCompanyPhone1()      { startValidation_('Company Phone 1',        null); }
function validateCompanyPhone2()      { startValidation_('Company Phone 2',        null); }

/**
 * Full chain entry point — called from the menu.
 * Starts with Email 1 validation, which automatically hands off to the phone chain.
 */
function validateAll() {
  // Email 1 is step 1 — when it finishes it calls startPhoneChainFromStep_('Contact Mobile Phone')
  startEmailValidation_('Email 1', 'Contact Mobile Phone');
}

/**
 * Called by EmailValidation.gs when Email 1 finishes in full chain mode.
 * Starts the phone chain from the given step.
 *
 * @param {string} firstPhoneStep - First phone column to validate (e.g. 'Contact Mobile Phone')
 */
function startPhoneChainFromStep_(firstPhoneStep) {
  // Find what comes after this step in the chain
  const chainSteps = PHONE_VALIDATION_CONFIG.CHAIN_STEPS;
  const currentIdx = chainSteps.indexOf(firstPhoneStep);
  const chainNext  = currentIdx !== -1 && currentIdx < chainSteps.length - 1
    ? chainSteps[currentIdx + 1]
    : null;

  startValidation_(firstPhoneStep, chainNext);
}

/**
 * Cancels any in-progress phone validation.
 */
function cancelPhoneValidation() {
  deleteContinuationTrigger_();
  clearSavedState_();
  SpreadsheetApp.getUi().alert(
    '🛑 Validation Cancelled',
    'The running validation has been cancelled.\n\n' +
    'All progress already written to the sheet is preserved.\n\n' +
    'You can click the same column button again to resume —\n' +
    'already-validated rows will be skipped automatically.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// CORE ORCHESTRATION
// =============================================================================

/**
 * Prepares result columns, saves state, runs first batch.
 *
 * @param {string}      columnName - Phone column to validate
 * @param {string|null} chainNext  - Next phone column in chain, or null if last/individual
 */
function startValidation_(columnName, chainNext) {
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty(PROPS.IS_RUNNING) === 'true') {
    const runningCol = props.getProperty(PROPS.COLUMN) || '(unknown)';
    SpreadsheetApp.getUi().alert(
      '⚠️ Already Running',
      `A phone validation is already in progress for:\n"${runningCol}"\n\n` +
      'Wait for it to finish, or cancel it first:\n' +
      '📞 Phone Validation → 🛑 Cancel Running Validation',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  const startRow = PHONE_VALIDATION_CONFIG.TEST_ROWS_START;
  const endRow   = PHONE_VALIDATION_CONFIG.TEST_MODE
    ? Math.min(PHONE_VALIDATION_CONFIG.TEST_ROWS_END, lastRow)
    : lastRow;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  prepareResultColumnsForPhone_(sheet, headers, columnName);

  props.setProperties({
    [PROPS.COLUMN]:     columnName,
    [PROPS.NEXT_ROW]:   String(startRow),
    [PROPS.END_ROW]:    String(endRow),
    [PROPS.SHEET_NAME]: sheet.getName(),
    [PROPS.IS_RUNNING]: 'true',
    [PROPS.CHAIN_NEXT]: chainNext || ''
  });

  Logger.log(`=== Starting phone validation: "${columnName}", rows ${startRow}–${endRow} ===`);
  Logger.log(`Chain next: ${chainNext || 'none'}`);

  ss.toast(
    `Validating "${columnName}" — rows ${startRow} to ${endRow}. Running in background.`,
    '📞 Phone Validation Started', 8
  );

  runBatch_();
}

/**
 * Called by the time-based trigger to resume a paused phone validation.
 */
function continuePhoneValidation() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROPS.IS_RUNNING) !== 'true') {
    deleteContinuationTrigger_();
    Logger.log('continuePhoneValidation: no active run. Cleaned up stale trigger.');
    return;
  }
  Logger.log('=== Continuing phone validation (trigger fired) ===');
  runBatch_();
}

/**
 * Core batch loop.
 * Reads all sheet data upfront, processes rows in memory, flushes writes at end.
 */
function runBatch_() {
  const batchStart = new Date();
  const props      = PropertiesService.getScriptProperties();

  const columnName = props.getProperty(PROPS.COLUMN);
  const nextRow    = parseInt(props.getProperty(PROPS.NEXT_ROW), 10);
  const endRow     = parseInt(props.getProperty(PROPS.END_ROW),  10);
  const sheetName  = props.getProperty(PROPS.SHEET_NAME);
  const chainNext  = props.getProperty(PROPS.CHAIN_NEXT) || '';

  if (!columnName || isNaN(nextRow) || isNaN(endRow)) {
    Logger.log('ERROR: Incomplete saved state. Aborting.');
    clearSavedState_();
    deleteContinuationTrigger_();
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();

  // ── Read ALL data in one batch ────────────────────────────────────────────
  const lastCol   = sheet.getLastColumn();
  const totalRows = sheet.getLastRow();
  const allData   = sheet.getRange(1, 1, totalRows, lastCol).getValues();
  const headers   = allData[0].map(h => h.toString().trim());
  // ─────────────────────────────────────────────────────────────────────────

  // Resolve column indices
  const predecessors     = PHONE_VALIDATION_CONFIG.FALLBACK_PREDECESSORS[columnName] || [];
  const phoneColIndex    = headers.indexOf(columnName);
  const statusColIndex   = headers.indexOf(columnName + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS);
  const lineTypeColIndex = headers.indexOf(columnName + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LINE_TYPE);
  const locationColIndex = headers.indexOf(columnName + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LOCATION);
  const formatColIndex   = headers.indexOf(columnName + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.FORMAT);

  if (phoneColIndex === -1 || statusColIndex === -1) {
    Logger.log(`ERROR: Required columns not found for "${columnName}". Aborting.`);
    clearSavedState_();
    deleteContinuationTrigger_();
    return;
  }

  // Predecessor index pairs for in-memory fallback chain check
  const predecessorIndices = predecessors.map(pred => ({
    statusIdx:   headers.indexOf(pred + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS),
    lineTypeIdx: headers.indexOf(pred + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LINE_TYPE)
  }));

  // Email 1 status column index for pre-check
  const emailCheckCol    = PHONE_VALIDATION_CONFIG.EMAIL_CHECK_COLUMN;
  const emailStatusIdx   = headers.indexOf(emailCheckCol + PHONE_VALIDATION_CONFIG.EMAIL_STATUS_SUFFIX);
  const hasEmailCheck    = emailStatusIdx !== -1;

  if (!hasEmailCheck) {
    Logger.log(`Warning: "${emailCheckCol}_Status" column not found — email pre-check disabled for this run`);
  }

  let processed      = 0;
  let validated      = 0;
  let skippedByChain = 0;
  let skippedAlready = 0;
  let skippedNoEmail = 0;
  let errors         = 0;
  let lastRowReached = nextRow - 1;
  let timedOut       = false;

  for (let row = nextRow; row <= endRow; row++) {
    const rowData = allData[row - 1];
    const dataIdx = row - 1;

    // ── TIME CHECK ────────────────────────────────────────────────────────────
    const elapsed = (new Date() - batchStart) / 1000;
    if (elapsed >= PHONE_VALIDATION_CONFIG.TIME_LIMIT_SECONDS) {
      Logger.log(`Time limit hit at row ${row} (${elapsed.toFixed(1)}s). Pausing.`);
      props.setProperty(PROPS.NEXT_ROW, String(row));
      timedOut = true;
      break;
    }
    // ─────────────────────────────────────────────────────────────────────────

    lastRowReached = row;

    // ── FALLBACK CHAIN CHECK (in-memory) ─────────────────────────────────────
    if (predecessorIndices.length > 0) {
      let predecessorPassed = false;
      for (const pred of predecessorIndices) {
        if (pred.statusIdx === -1 || pred.lineTypeIdx === -1) continue;
        const predStatus   = (rowData[pred.statusIdx]   || '').toString().trim().toUpperCase();
        const predLineType = (rowData[pred.lineTypeIdx] || '').toString().trim().toUpperCase();
        if (predStatus   === PHONE_VALIDATION_CONFIG.VALID_STATUS &&
            predLineType === PHONE_VALIDATION_CONFIG.VALID_LINE_TYPE) {
          predecessorPassed = true;
          break;
        }
      }
      if (predecessorPassed) {
        const current = (rowData[statusColIndex] || '').toString().trim();
        if (current === '' || current === PHONE_VALIDATION_CONFIG.SKIP_MARKER) {
          SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(row, statusColIndex + 1).setValue(PHONE_VALIDATION_CONFIG.SKIP_MARKER);
        }
        skippedByChain++;
        continue;
      }
    }
    // ─────────────────────────────────────────────────────────────────────────

    const phoneNumber = (rowData[phoneColIndex] || '').toString().trim();
    if (!phoneNumber) continue;

    // Skip already-validated rows (re-validate DELAYED)
    const existingStatus = (rowData[statusColIndex] || '').toString().trim();
    if (existingStatus !== '' &&
        existingStatus !== PHONE_VALIDATION_CONFIG.SKIP_MARKER &&
        existingStatus !== PHONE_VALIDATION_CONFIG.NO_EMAIL_MARKER) {
      if (existingStatus.toUpperCase() !== 'DELAYED') {
        skippedAlready++;
        continue;
      }
      Logger.log(`Row ${row}: DELAYED — re-validating`);
    }

    // ── EMAIL PRE-CHECK (in-memory) ───────────────────────────────────────────
    if (hasEmailCheck) {
      const emailStatus = (rowData[emailStatusIdx] || '').toString().trim().toLowerCase();
      if (emailStatus !== PHONE_VALIDATION_CONFIG.VALID_EMAIL_STATUS) {
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(row, statusColIndex + 1).setValue(PHONE_VALIDATION_CONFIG.NO_EMAIL_MARKER);
        skippedNoEmail++;
        Logger.log(`Row ${row}: Email 1 not valid ("${emailStatus}") — skipping phone API call`);
        continue;
      }
    }
    // ─────────────────────────────────────────────────────────────────────────

    // ── API CALL ─────────────────────────────────────────────────────────────
    processed++;
    Logger.log(`Row ${row}: Validating ${phoneNumber}`);

    const result = validatePhoneWithAPI_(phoneNumber);

    // Write result immediately so progress is never lost if script times out
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (result.success) {
      activeSheet.getRange(row, statusColIndex   + 1).setValue(result.phone_status        || '');
      activeSheet.getRange(row, lineTypeColIndex + 1).setValue(result.line_type            || '');
      activeSheet.getRange(row, locationColIndex + 1).setValue(result.location             || '');
      activeSheet.getRange(row, formatColIndex   + 1).setValue(result.format_international || '');
      validated++;
      Logger.log(`Row ${row}: ✓ ${result.phone_status} / ${result.line_type}`);
    } else {
      activeSheet.getRange(row, statusColIndex   + 1).setValue('ERROR: ' + result.error);
      activeSheet.getRange(row, lineTypeColIndex + 1).setValue('');
      activeSheet.getRange(row, locationColIndex + 1).setValue('');
      activeSheet.getRange(row, formatColIndex   + 1).setValue('');
      errors++;
      Logger.log(`Row ${row}: ✗ ${result.error}`);
    }
    SpreadsheetApp.flush();

    // Smart throttle: if Byteplant says we are near the rate limit,
    // wait for the current window to reset before the next call.
    // result.ratelimit_remain and result.ratelimit_seconds come from the API response.
    if (result.success && result.ratelimit_remain !== undefined) {
      if (result.ratelimit_remain <= PHONE_VALIDATION_CONFIG.RATE_LIMIT_BUFFER) {
        const waitMs = Math.max(0, (result.ratelimit_seconds || 10)) * 1000;
        Logger.log(`Rate limit buffer hit (${result.ratelimit_remain} calls left). Waiting ${waitMs}ms...`);
        Utilities.sleep(waitMs + 500); // extra 500ms buffer
      }
    }
    // ─────────────────────────────────────────────────────────────────────────
  }

  // Flush any pending Sheets writes
  SpreadsheetApp.flush();

  // ── CONTINUATION OR COMPLETION ────────────────────────────────────────────
  if (timedOut) {
    const resumeAt = parseInt(props.getProperty(PROPS.NEXT_ROW), 10);
    scheduleContinuationTrigger_();

    Logger.log(`Pausing. Resuming at row ${resumeAt} in ${PHONE_VALIDATION_CONFIG.CONTINUATION_DELAY_MINUTES} min.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Phone validation processed up to row ${lastRowReached}. ` +
      `Resuming at row ${resumeAt} in ${PHONE_VALIDATION_CONFIG.CONTINUATION_DELAY_MINUTES} minute(s)…`,
      '⏱️ Auto-Continuing in Background', 30
    );

  } else {
    // This column is fully done
    deleteContinuationTrigger_();
    clearSavedState_();

    const duration = ((new Date() - batchStart) / 1000).toFixed(1);
    Logger.log(`=== Phone validation complete for "${columnName}" ===`);

    if (chainNext) {
      // Auto-start next phone column in chain
      Logger.log(`Full chain: "${columnName}" done. Starting next: "${chainNext}"`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `"${columnName}" done. Starting "${chainNext}"…`,
        `✅ Moving to Next Step`, 8
      );

      // Find what comes after chainNext in the chain
      const chainSteps   = PHONE_VALIDATION_CONFIG.CHAIN_STEPS;
      const nextIdx      = chainSteps.indexOf(chainNext);
      const afterNext    = nextIdx !== -1 && nextIdx < chainSteps.length - 1
        ? chainSteps[nextIdx + 1]
        : null;

      startValidation_(chainNext, afterNext);

    } else {
      // Last step or individual run — show completion alert
      const isChainEnd = PHONE_VALIDATION_CONFIG.CHAIN_STEPS.includes(columnName);
      SpreadsheetApp.getUi().alert(
        `✅ ${columnName} — Complete`,
        `Phone validation finished for "${columnName}"!\n\n` +
        `📊 Summary:\n` +
        `  Rows checked:                       ${PHONE_VALIDATION_CONFIG.TEST_ROWS_START}–${endRow}\n` +
        `  API calls made:                     ${processed}\n` +
        `  Successfully validated:             ${validated}\n` +
        `  Skipped — chain (SKIP):             ${skippedByChain}\n` +
        `  Skipped — no valid email (NO_EMAIL):${skippedNoEmail}\n` +
        `  Skipped — already validated:        ${skippedAlready}\n` +
        `  Errors:                             ${errors}\n` +
        `  Duration (this batch):              ${duration}s\n\n` +
        `${isChainEnd && !chainNext ? '🏁 Full chain complete — all columns done!' : ''}\n` +
        `${PHONE_VALIDATION_CONFIG.TEST_MODE ? '⚠️ TEST MODE — rows 2–6 only' : '✓ FULL MODE — all rows'}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

// =============================================================================
// COLUMN PREPARATION
// =============================================================================

function prepareResultColumnsForPhone_(sheet, headers, phoneColumn) {
  const phoneColIndex = headers.indexOf(phoneColumn);
  if (phoneColIndex === -1) {
    Logger.log(`Warning: "${phoneColumn}" not found — cannot prepare result columns`);
    return;
  }

  const statusColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  if (headers.includes(statusColName)) {
    Logger.log(`Result columns already exist for "${phoneColumn}"`);
    return;
  }

  const insertAfterCol = phoneColIndex + 1;
  sheet.insertColumnsAfter(insertAfterCol, 4);

  const firstNewCol = insertAfterCol + 1;
  sheet.getRange(1, firstNewCol    ).setValue(phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS);
  sheet.getRange(1, firstNewCol + 1).setValue(phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LINE_TYPE);
  sheet.getRange(1, firstNewCol + 2).setValue(phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LOCATION);
  sheet.getRange(1, firstNewCol + 3).setValue(phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.FORMAT);

  sheet.getRange(1, firstNewCol, 1, 4)
    .setBackground('#4285f4')
    .setFontColor('#000000')
    .setFontWeight('bold');

  Logger.log(`Created 4 result columns after "${phoneColumn}"`);
}

// =============================================================================
// TRIGGER MANAGEMENT
// =============================================================================

function scheduleContinuationTrigger_() {
  deleteContinuationTrigger_();

  const trigger = ScriptApp.newTrigger('continuePhoneValidation')
    .timeBased()
    .after(PHONE_VALIDATION_CONFIG.CONTINUATION_DELAY_MINUTES * 60 * 1000)
    .create();

  PropertiesService.getScriptProperties()
    .setProperty(PROPS.TRIGGER_ID, trigger.getUniqueId());

  Logger.log(`Scheduled phone continuation trigger: ${trigger.getUniqueId()}`);
}

function deleteContinuationTrigger_() {
  const props     = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(PROPS.TRIGGER_ID);
  if (!triggerId) return;

  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(t);
      Logger.log(`Deleted phone trigger: ${triggerId}`);
      break;
    }
  }
  props.deleteProperty(PROPS.TRIGGER_ID);
}

function clearSavedState_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROPS.COLUMN);
  props.deleteProperty(PROPS.NEXT_ROW);
  props.deleteProperty(PROPS.END_ROW);
  props.deleteProperty(PROPS.SHEET_NAME);
  props.deleteProperty(PROPS.IS_RUNNING);
  props.deleteProperty(PROPS.CHAIN_NEXT);
  Logger.log('Cleared phone saved state');
}

// =============================================================================
// BYTEPLANT API
// =============================================================================

/**
 * Calls the Byteplant API to validate a single phone number.
 *
 * SMART RATE LIMITING:
 * - Returns ratelimit_remain and ratelimit_seconds from every response
 * - The caller (runBatch_) uses those to decide whether to wait before the next call
 * - If RATE_LIMIT_EXCEEDED is returned (hit limit mid-run), this function
 *   automatically waits ratelimit_seconds then retries — transparent to the caller
 *
 * @param {string} phoneNumber
 * @returns {{
 *   success: boolean,
 *   phone_status?: string,
 *   line_type?: string,
 *   location?: string,
 *   format_international?: string,
 *   ratelimit_remain?: number,
 *   ratelimit_seconds?: number,
 *   error?: string
 * }}
 */
function validatePhoneWithAPI_(phoneNumber) {
  const { API_KEY, API_ENDPOINT, COUNTRY_CODE, MAX_RETRIES, RETRY_DELAY, REQUEST_TIMEOUT } = PHONE_VALIDATION_CONFIG;
  const formatted = phoneNumber.toString().trim();

  for (let attempt = 1; attempt <= MAX_RETRIES + 1; attempt++) {
    try {
      const url      = `${API_ENDPOINT}?PhoneNumber=${encodeURIComponent(formatted)}&CountryCode=${COUNTRY_CODE}&APIKey=${API_KEY}`;
      const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, timeout: REQUEST_TIMEOUT });
      const code     = response.getResponseCode();
      const text     = response.getContentText();

      Logger.log(`Attempt ${attempt} — HTTP ${code}: ${text.substring(0, 150)}`);

      if (code === 200) {
        const data = JSON.parse(text);
        if (!data.status) return { success: false, error: 'No status in API response' };

        const status = data.status.toUpperCase();

        // ── Rate limit hit — wait and retry automatically ─────────────────────
        if (status === 'RATE_LIMIT_EXCEEDED') {
          const waitSeconds = parseInt(data.ratelimit_seconds || 10, 10);
          const waitMs      = (waitSeconds + 1) * 1000; // +1s buffer
          Logger.log(`RATE_LIMIT_EXCEEDED on attempt ${attempt}. Waiting ${waitMs}ms then retrying...`);
          Utilities.sleep(waitMs);
          continue; // retry same attempt after waiting
        }
        // ─────────────────────────────────────────────────────────────────────

        if (status === 'API_KEY_INVALID_OR_DEPLETED') {
          return { success: false, error: 'API key invalid or depleted' };
        }

        const isInvalid = status === 'INVALID';
        return {
          success:              true,
          phone_status:         data.status,
          line_type:            isInvalid ? '' : (data.linetype             || ''),
          location:             isInvalid ? '' : (data.location              || ''),
          format_international: isInvalid ? '' : (data.formatinternational   || formatted),
          // Pass rate limit info back so the loop can throttle proactively
          ratelimit_remain:  parseInt(data.ratelimit_remain  || 100, 10),
          ratelimit_seconds: parseInt(data.ratelimit_seconds || 0,   10)
        };

      } else {
        if (attempt <= MAX_RETRIES) { Utilities.sleep(RETRY_DELAY); continue; }
        return { success: false, error: `API error ${code}` };
      }

    } catch (err) {
      if (attempt <= MAX_RETRIES) { Utilities.sleep(RETRY_DELAY); continue; }
      return { success: false, error: err.toString().substring(0, 100) };
    }
  }

  return { success: false, error: 'Max retries exceeded' };
}

// =============================================================================
// UTILITY FUNCTIONS (menu items)
// =============================================================================

function showConfigDialog() {
  const cfg = PHONE_VALIDATION_CONFIG;
  SpreadsheetApp.getUi().alert(
    '⚙️ Phone Validation — Configuration',
    'Current Settings:\n\n' +
    `Test Mode:            ${cfg.TEST_MODE ? 'ENABLED (rows 2–6 only)' : 'DISABLED (all rows)'}\n` +
    `Email pre-check:      Email 1 must be "valid" before any phone API call\n` +
    `Rate limit buffer:    Waits when ${cfg.RATE_LIMIT_BUFFER} calls remain in window\n` +
    `Max Retries:          ${cfg.MAX_RETRIES}\n` +
    `Time Limit Per Run:   ${cfg.TIME_LIMIT_SECONDS}s\n` +
    `Continuation Delay:   ${cfg.CONTINUATION_DELAY_MINUTES} minute(s)\n\n` +
    'Full Chain Order:\n' +
    '  1. Email 1               (ZeroBounce)\n' +
    '  2. Contact Mobile Phone  (Byteplant — no predecessors)\n' +
    '  3. Contact Phone 1       (skips row if Mobile passes)\n' +
    '  4. Company Phone 1       (skips row if Mobile or Phone 1 passes)\n' +
    '  5. Company Phone 2       (skips row if any above passes)\n\n' +
    'Status markers:\n' +
    '  SKIP     = higher-priority phone already passed\n' +
    '  NO_EMAIL = Email 1 not valid — phone validation skipped\n\n' +
    'Edit PHONE_VALIDATION_CONFIG in the script editor to change any setting.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showValidationLogs() {
  const logs = Logger.getLog();
  SpreadsheetApp.getUi().alert(
    '📋 Phone Validation Logs',
    logs || 'No logs available. Run a validation first.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// END OF SCRIPT
// =============================================================================

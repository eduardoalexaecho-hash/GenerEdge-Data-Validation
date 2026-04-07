/**
 * =============================================================================
 * ZEROBOUNCE EMAIL VALIDATION SCRIPT
 * =============================================================================
 *
 * Validates email addresses using ZeroBounce Batch Validation API
 *
 * IMPORTANT: This file should be in the same Apps Script project as:
 *            DataCleaningPipeline_V3.gs and PhoneValidation.gs
 *
 * BATCH API SPECIFICATION:
 * - Endpoint: POST https://api.zerobounce.net/v2/validatebatch
 * - Sends up to 100 emails per call
 * - Rate limited to 5 calls per minute (12s minimum between calls)
 * - Can take up to 70 seconds to return results
 * - Request body: JSON with api_key and email_batch array
 * - email_batch format: [{ email_address: "...", ip_address: null }, ...]
 *
 * STATUS VALUES (same as single endpoint):
 * - valid:       Email is valid and safe to send
 * - invalid:     Email is invalid and will bounce
 * - catch-all:   Domain accepts all emails (risky)
 * - unknown:     Cannot determine validity
 * - spamtrap:    Email is a spam trap
 * - abuse:       Email is known for abuse complaints
 * - do_not_mail: Email should not be mailed
 *
 * PERFORMANCE:
 * - 100 emails per batch call
 * - 5 calls per minute max → 500 emails per minute
 * - 5,000 rows ≈ 10 minutes total
 * - vs single-email approach ≈ 33+ minutes
 *
 * TIME-AWARE AUTO-CONTINUATION:
 * - Processes rows in chunks of BATCH_SIZE (100)
 * - After each batch call, writes all results immediately to the sheet
 * - Checks elapsed time before each batch — if near 240s limit, saves
 *   position and schedules a continuation trigger for 1 minute later
 * - continueEmailValidation() resumes from saved position automatically
 * - After completing, if running as part of full chain, schedules a trigger
 *   to start the phone validation chain in a fresh execution
 *
 * @author: Claude
 * @version: 3.0 - ZeroBounce batch endpoint, 100 emails per call
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const EMAIL_VALIDATION_CONFIG = {
  // ZeroBounce API
  API_KEY:           '48891e1db6e34f00a0d4062edc9fd5a4',
  API_ENDPOINT:      'https://api.zerobounce.net/v2/validate',       // single (kept for reference)
  BATCH_ENDPOINT:    'https://api.zerobounce.net/v2/validatebatch',   // batch (used by this script)

  // Test Mode
  TEST_MODE:       false,
  TEST_ROWS_START: 2,
  TEST_ROWS_END:   6,

  // Email columns available for individual validation
  EMAIL_COLUMNS: ['Email 1', 'Email 2', 'Personal Email'],

  // Result column suffix
  RESULT_COLUMNS: {
    STATUS: '_Status'
  },

  // Batch settings
  BATCH_SIZE:            100,   // Emails per API call (ZeroBounce max is 100)
  DELAY_BETWEEN_BATCHES: 12000, // ms between batch calls (5 calls/min = 12s minimum)
  REQUEST_TIMEOUT:       75000, // ms — batch can take up to 70s, give 75s buffer

  // Error handling
  MAX_RETRIES:  2,
  RETRY_DELAY:  3000,  // ms before retrying a failed batch

  // Time-aware: stop processing when elapsed seconds exceeds this value
  // Each batch can take up to 70s + 12s delay = 82s, so we need a generous buffer
  // With 240s limit: we can safely fit ~2-3 batches (200-300 emails) per execution
  TIME_LIMIT_SECONDS:         200,  // More conservative — each batch takes up to 82s
  CONTINUATION_DELAY_MINUTES: 1
};

// PropertiesService keys
const EMAIL_PROPS = {
  COLUMN:             'emailVal_column',
  NEXT_ROW:           'emailVal_nextRow',
  END_ROW:            'emailVal_endRow',
  SHEET_NAME:         'emailVal_sheetName',
  TRIGGER_ID:         'emailVal_triggerId',
  IS_RUNNING:         'emailVal_isRunning',
  CHAIN_NEXT:         'emailVal_chainNext',       // next step after this column finishes
  HANDOFF_TRIGGER_ID: 'emailVal_handoffTrigger'   // trigger that fires to start phone chain
};

// =============================================================================
// PUBLIC ENTRY POINTS (menu buttons)
// =============================================================================

function validateEmail1()        { startEmailValidation_('Email 1',        null); }
function validateEmail2()        { startEmailValidation_('Email 2',        null); }
function validatePersonalEmail() { startEmailValidation_('Personal Email', null); }

/**
 * Cancels any in-progress email validation.
 */
function cancelEmailValidation() {
  deleteEmailContinuationTrigger_();
  deletePhoneChainHandoffTrigger_();
  clearEmailSavedState_();
  SpreadsheetApp.getUi().alert(
    '🛑 Email Validation Cancelled',
    'The running email validation has been cancelled.\n\n' +
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
 * Called by a menu button or by validateAll() in the phone file.
 * Prepares result column, saves state, runs first batch.
 *
 * @param {string}      columnName - Email column to validate (e.g. 'Email 1')
 * @param {string|null} chainNext  - First phone step after email finishes (full chain only)
 */
function startEmailValidation_(columnName, chainNext) {
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty(EMAIL_PROPS.IS_RUNNING) === 'true') {
    const runningCol = props.getProperty(EMAIL_PROPS.COLUMN) || '(unknown)';
    SpreadsheetApp.getUi().alert(
      '⚠️ Already Running',
      `An email validation is already in progress for:\n"${runningCol}"\n\n` +
      'Wait for it to finish, or cancel it first:\n' +
      '📧 Email Validation → 🛑 Cancel Running Validation',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();

  const startRow = EMAIL_VALIDATION_CONFIG.TEST_ROWS_START;
  const endRow   = EMAIL_VALIDATION_CONFIG.TEST_MODE
    ? Math.min(EMAIL_VALIDATION_CONFIG.TEST_ROWS_END, lastRow)
    : lastRow;

  // Prepare _Status column if it doesn't exist yet
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  prepareResultColumnsForEmail_(sheet, headers, columnName);

  props.setProperties({
    [EMAIL_PROPS.COLUMN]:     columnName,
    [EMAIL_PROPS.NEXT_ROW]:   String(startRow),
    [EMAIL_PROPS.END_ROW]:    String(endRow),
    [EMAIL_PROPS.SHEET_NAME]: sheet.getName(),
    [EMAIL_PROPS.IS_RUNNING]: 'true',
    [EMAIL_PROPS.CHAIN_NEXT]: chainNext || ''
  });

  Logger.log(`=== Starting email validation: "${columnName}", rows ${startRow}–${endRow} ===`);
  Logger.log(`Batch size: ${EMAIL_VALIDATION_CONFIG.BATCH_SIZE} | Chain next: ${chainNext || 'none'}`);

  ss.toast(
    `Validating "${columnName}" in batches of ${EMAIL_VALIDATION_CONFIG.BATCH_SIZE}. Running in background.`,
    '📧 Email Validation Started', 8
  );

  runEmailBatch_();
}

/**
 * Called by the time-based trigger to resume a paused email validation.
 * Function name must stay stable — it is the registered trigger handler.
 */
function continueEmailValidation() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(EMAIL_PROPS.IS_RUNNING) !== 'true') {
    deleteEmailContinuationTrigger_();
    Logger.log('continueEmailValidation: no active run. Cleaned up stale trigger.');
    return;
  }
  Logger.log('=== Continuing email validation (trigger fired) ===');
  runEmailBatch_();
}

/**
 * Core batch loop.
 *
 * Instead of one API call per row, this collects up to BATCH_SIZE rows at a time,
 * sends them all in one POST request to /v2/validatebatch, gets back all results,
 * writes them to the sheet, then moves to the next batch.
 *
 * Time check happens BEFORE each batch so we never start a batch we can't finish.
 */
function runEmailBatch_() {
  const batchStart = new Date();
  const props      = PropertiesService.getScriptProperties();

  const columnName = props.getProperty(EMAIL_PROPS.COLUMN);
  const nextRow    = parseInt(props.getProperty(EMAIL_PROPS.NEXT_ROW), 10);
  const endRow     = parseInt(props.getProperty(EMAIL_PROPS.END_ROW),  10);
  const sheetName  = props.getProperty(EMAIL_PROPS.SHEET_NAME);
  const chainNext  = props.getProperty(EMAIL_PROPS.CHAIN_NEXT) || '';

  if (!columnName || isNaN(nextRow) || isNaN(endRow)) {
    Logger.log('ERROR: Incomplete saved state. Aborting.');
    clearEmailSavedState_();
    deleteEmailContinuationTrigger_();
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();

  // ── Read ALL sheet data in one call ──────────────────────────────────────
  const lastCol   = sheet.getLastColumn();
  const totalRows = sheet.getLastRow();
  const allData   = sheet.getRange(1, 1, totalRows, lastCol).getValues();
  const headers   = allData[0].map(h => h.toString().trim());
  // ─────────────────────────────────────────────────────────────────────────

  const emailColIndex  = headers.indexOf(columnName);
  const statusColName  = columnName + EMAIL_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  const statusColIndex = headers.indexOf(statusColName);

  if (emailColIndex === -1 || statusColIndex === -1) {
    Logger.log(`ERROR: Required columns not found for "${columnName}". Aborting.`);
    clearEmailSavedState_();
    deleteEmailContinuationTrigger_();
    return;
  }

  let totalProcessed  = 0;
  let totalValidated  = 0;
  let totalSkipped    = 0;
  let totalErrors     = 0;
  let currentRow      = nextRow;
  let timedOut        = false;

  // ── BATCH PROCESSING LOOP ─────────────────────────────────────────────────
  while (currentRow <= endRow) {

    // Time check BEFORE starting a new batch
    // Each batch can take up to 70s + 12s delay = 82s
    // Stop if we don't have enough time for at least one more full batch
    const elapsed = (new Date() - batchStart) / 1000;
    if (elapsed >= EMAIL_VALIDATION_CONFIG.TIME_LIMIT_SECONDS) {
      Logger.log(`Time limit hit before batch at row ${currentRow} (${elapsed.toFixed(1)}s). Pausing.`);
      props.setProperty(EMAIL_PROPS.NEXT_ROW, String(currentRow));
      timedOut = true;
      break;
    }

    // ── Collect up to BATCH_SIZE rows that need validation ────────────────
    const batch = [];  // [{ row (1-based), email }]

    for (let row = currentRow; row <= endRow && batch.length < EMAIL_VALIDATION_CONFIG.BATCH_SIZE; row++) {
      const rowData       = allData[row - 1];
      const emailAddress  = (rowData[emailColIndex] || '').toString().trim();
      const existingStatus = (rowData[statusColIndex] || '').toString().trim();

      if (!emailAddress) {
        // Empty — skip silently, advance currentRow
        currentRow = row + 1;
        continue;
      }

      if (existingStatus !== '') {
        // Already validated — skip, advance currentRow
        totalSkipped++;
        currentRow = row + 1;
        continue;
      }

      batch.push({ row, email: emailAddress });
      currentRow = row + 1;
    }

    if (batch.length === 0) {
      // No rows left to validate in this segment
      break;
    }

    Logger.log(`Sending batch of ${batch.length} emails (rows ${batch[0].row}–${batch[batch.length-1].row})`);

    // ── API CALL — send entire batch ──────────────────────────────────────
    const results = validateEmailBatchWithAPI_(batch.map(b => b.email));
    totalProcessed += batch.length;

    // ── Write results immediately to sheet ────────────────────────────────
    if (results.success) {
      // Build a map of email → result for fast lookup
      // ZeroBounce returns results in the same order but we match by email to be safe
      const resultMap = {};
      for (const r of results.data) {
        resultMap[(r.address || '').toLowerCase().trim()] = r.status || 'unknown';
      }

      for (const item of batch) {
        const status = resultMap[item.email.toLowerCase().trim()] || 'unknown';
        sheet.getRange(item.row, statusColIndex + 1).setValue(status);
        if (status === 'valid') totalValidated++;
        Logger.log(`Row ${item.row}: ${item.email} → ${status}`);
      }

    } else {
      // Entire batch failed — write error to all rows in batch
      for (const item of batch) {
        sheet.getRange(item.row, statusColIndex + 1).setValue('ERROR: ' + results.error);
      }
      totalErrors += batch.length;
      Logger.log(`Batch error: ${results.error}`);
    }

    SpreadsheetApp.flush();
    // ─────────────────────────────────────────────────────────────────────

    // Wait between batches to respect rate limit (5 calls/min)
    // Only wait if there are more rows to process
    if (currentRow <= endRow) {
      Logger.log(`Waiting ${EMAIL_VALIDATION_CONFIG.DELAY_BETWEEN_BATCHES}ms before next batch...`);
      Utilities.sleep(EMAIL_VALIDATION_CONFIG.DELAY_BETWEEN_BATCHES);
    }
  }
  // ── END BATCH LOOP ────────────────────────────────────────────────────────

  // ── CONTINUATION OR COMPLETION ────────────────────────────────────────────
  if (timedOut) {
    scheduleEmailContinuationTrigger_();
    const resumeAt = parseInt(props.getProperty(EMAIL_PROPS.NEXT_ROW), 10);

    Logger.log(`Pausing at row ${resumeAt}. Resuming in ${EMAIL_VALIDATION_CONFIG.CONTINUATION_DELAY_MINUTES} min.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Email validation paused at row ${resumeAt}. ` +
      `Resuming automatically in ${EMAIL_VALIDATION_CONFIG.CONTINUATION_DELAY_MINUTES} minute(s)…`,
      '⏱️ Auto-Continuing in Background', 30
    );

  } else {
    // All rows done
    deleteEmailContinuationTrigger_();
    clearEmailSavedState_();

    const duration = ((new Date() - batchStart) / 1000).toFixed(1);
    Logger.log(`=== Email validation complete: "${columnName}" ===`);
    Logger.log(`Processed: ${totalProcessed}, Valid: ${totalValidated}, Skipped: ${totalSkipped}, Errors: ${totalErrors}`);

    if (chainNext) {
      // Schedule phone chain to start in a fresh execution
      Logger.log(`Scheduling phone chain handoff → "${chainNext}"`);
      schedulePhoneChainHandoffTrigger_(chainNext);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `"${columnName}" validation done (${totalProcessed} emails). ` +
        `Phone validation will start automatically in ~1 minute…`,
        '✅ Email Done — Phones Starting Soon', 15
      );

    } else {
      SpreadsheetApp.getUi().alert(
        `✅ ${columnName} — Complete`,
        `Email validation finished for "${columnName}"!\n\n` +
        `📊 Summary:\n` +
        `  Rows checked:           ${EMAIL_VALIDATION_CONFIG.TEST_ROWS_START}–${endRow}\n` +
        `  Emails sent to API:     ${totalProcessed}\n` +
        `  Valid emails found:     ${totalValidated}\n` +
        `  Skipped (already done): ${totalSkipped}\n` +
        `  Errors:                 ${totalErrors}\n` +
        `  Batch size used:        ${EMAIL_VALIDATION_CONFIG.BATCH_SIZE} emails/call\n` +
        `  Duration (this run):    ${duration}s\n\n` +
        `${EMAIL_VALIDATION_CONFIG.TEST_MODE ? '⚠️ TEST MODE — rows 2–6 only' : '✓ FULL MODE — all rows'}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

// =============================================================================
// COLUMN PREPARATION
// =============================================================================

/**
 * Inserts 1 _Status column immediately after the email column if it doesn't exist.
 */
function prepareResultColumnsForEmail_(sheet, headers, emailColumn) {
  const emailColIndex = headers.indexOf(emailColumn);
  if (emailColIndex === -1) {
    Logger.log(`Warning: "${emailColumn}" not found — cannot prepare result column`);
    return;
  }

  const statusColName = emailColumn + EMAIL_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  if (headers.includes(statusColName)) {
    Logger.log(`Result column already exists for "${emailColumn}"`);
    return;
  }

  const insertAfterCol = emailColIndex + 1;
  sheet.insertColumnsAfter(insertAfterCol, 1);

  const firstNewCol = insertAfterCol + 1;
  sheet.getRange(1, firstNewCol).setValue(statusColName);
  sheet.getRange(1, firstNewCol, 1, 1)
    .setBackground('#4285f4')
    .setFontColor('#000000')
    .setFontWeight('bold');

  Logger.log(`Created _Status column after "${emailColumn}"`);
}

// =============================================================================
// TRIGGER MANAGEMENT — EMAIL CONTINUATION
// =============================================================================

function scheduleEmailContinuationTrigger_() {
  deleteEmailContinuationTrigger_();

  const trigger = ScriptApp.newTrigger('continueEmailValidation')
    .timeBased()
    .after(EMAIL_VALIDATION_CONFIG.CONTINUATION_DELAY_MINUTES * 60 * 1000)
    .create();

  PropertiesService.getScriptProperties()
    .setProperty(EMAIL_PROPS.TRIGGER_ID, trigger.getUniqueId());

  Logger.log(`Scheduled email continuation trigger: ${trigger.getUniqueId()}`);
}

function deleteEmailContinuationTrigger_() {
  const props     = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(EMAIL_PROPS.TRIGGER_ID);
  if (!triggerId) return;

  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(t);
      Logger.log(`Deleted email continuation trigger: ${triggerId}`);
      break;
    }
  }
  props.deleteProperty(EMAIL_PROPS.TRIGGER_ID);
}

function clearEmailSavedState_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(EMAIL_PROPS.COLUMN);
  props.deleteProperty(EMAIL_PROPS.NEXT_ROW);
  props.deleteProperty(EMAIL_PROPS.END_ROW);
  props.deleteProperty(EMAIL_PROPS.SHEET_NAME);
  props.deleteProperty(EMAIL_PROPS.IS_RUNNING);
  props.deleteProperty(EMAIL_PROPS.CHAIN_NEXT);
  props.deleteProperty(EMAIL_PROPS.HANDOFF_TRIGGER_ID);
  Logger.log('Cleared email saved state');
}

// =============================================================================
// TRIGGER MANAGEMENT — PHONE CHAIN HANDOFF
// =============================================================================

/**
 * Schedules a trigger to start the phone chain in a fresh execution.
 * This ensures the phone chain has a full time budget and doesn't run
 * inside the tail end of the email execution.
 */
function schedulePhoneChainHandoffTrigger_(firstPhoneStep) {
  deletePhoneChainHandoffTrigger_();

  PropertiesService.getScriptProperties()
    .setProperty(EMAIL_PROPS.CHAIN_NEXT, firstPhoneStep);

  const trigger = ScriptApp.newTrigger('startPhoneChainFromEmail_')
    .timeBased()
    .after(1 * 60 * 1000)
    .create();

  PropertiesService.getScriptProperties()
    .setProperty(EMAIL_PROPS.HANDOFF_TRIGGER_ID, trigger.getUniqueId());

  Logger.log(`Scheduled phone chain handoff trigger: ${trigger.getUniqueId()} → "${firstPhoneStep}"`);
}

function deletePhoneChainHandoffTrigger_() {
  const props     = PropertiesService.getScriptProperties();
  const triggerId = props.getProperty(EMAIL_PROPS.HANDOFF_TRIGGER_ID);
  if (!triggerId) return;

  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(t);
      Logger.log(`Deleted phone chain handoff trigger: ${triggerId}`);
      break;
    }
  }
  props.deleteProperty(EMAIL_PROPS.HANDOFF_TRIGGER_ID);
}

/**
 * Fired by the handoff trigger ~1 minute after email finishes.
 * Starts the phone chain in a fresh execution with full time budget.
 * Function name must stay stable — it is the registered trigger handler.
 */
function startPhoneChainFromEmail_() {
  const props          = PropertiesService.getScriptProperties();
  const firstPhoneStep = props.getProperty(EMAIL_PROPS.CHAIN_NEXT) || '';

  deletePhoneChainHandoffTrigger_();
  props.deleteProperty(EMAIL_PROPS.CHAIN_NEXT);

  if (!firstPhoneStep) {
    Logger.log('startPhoneChainFromEmail_: no chain step saved. Nothing to do.');
    return;
  }

  Logger.log(`=== Phone chain handoff fired — starting "${firstPhoneStep}" ===`);
  startPhoneChainFromStep_(firstPhoneStep);
}

// =============================================================================
// ZEROBOUNCE BATCH API
// =============================================================================

/**
 * Sends a batch of up to 100 email addresses to ZeroBounce's /v2/validatebatch endpoint.
 * Returns all results in one response.
 *
 * @param {string[]} emails - Array of email address strings (max 100)
 * @returns {{ success: boolean, data?: object[], error?: string }}
 */
function validateEmailBatchWithAPI_(emails) {
  const { API_KEY, BATCH_ENDPOINT, MAX_RETRIES, RETRY_DELAY, REQUEST_TIMEOUT } = EMAIL_VALIDATION_CONFIG;

  // Build the email_batch array — ip_address is optional, set to null
  const emailBatch = emails.map(email => ({
    email_address: email,
    ip_address:    null
  }));

  const payload = JSON.stringify({
    api_key:     API_KEY,
    email_batch: emailBatch
  });

  for (let attempt = 1; attempt <= MAX_RETRIES + 1; attempt++) {
    try {
      Logger.log(`Batch API attempt ${attempt}: ${emails.length} emails`);

      const response = UrlFetchApp.fetch(BATCH_ENDPOINT, {
        method:             'post',
        contentType:        'application/json',
        payload:            payload,
        muteHttpExceptions: true,
        timeout:            REQUEST_TIMEOUT
      });

      const code = response.getResponseCode();
      const text = response.getContentText();

      Logger.log(`Batch response HTTP ${code}: ${text.substring(0, 200)}`);

      if (code === 200) {
        const data = JSON.parse(text);

        // ZeroBounce batch response: { "email_batch": [...], "errors": [...] }
        if (!data.email_batch) {
          return { success: false, error: 'No email_batch in response' };
        }

        return { success: true, data: data.email_batch };

      } else if (code === 429) {
        // Rate limited — wait longer before retrying
        Logger.log('Rate limited (429). Waiting 60s before retry...');
        Utilities.sleep(60000);
        continue;

      } else {
        if (attempt <= MAX_RETRIES) {
          Logger.log(`HTTP ${code} — retrying in ${RETRY_DELAY}ms...`);
          Utilities.sleep(RETRY_DELAY);
          continue;
        }
        return { success: false, error: `API error ${code}: ${text.substring(0, 100)}` };
      }

    } catch (err) {
      if (attempt <= MAX_RETRIES) {
        Logger.log(`Attempt ${attempt} failed: ${err}. Retrying...`);
        Utilities.sleep(RETRY_DELAY);
        continue;
      }
      return { success: false, error: err.toString().substring(0, 100) };
    }
  }

  return { success: false, error: 'Max retries exceeded' };
}

// =============================================================================
// UTILITY FUNCTIONS (menu items)
// =============================================================================

function showEmailConfigDialog() {
  const cfg = EMAIL_VALIDATION_CONFIG;
  SpreadsheetApp.getUi().alert(
    '⚙️ Email Validation — Configuration',
    'Current Settings:\n\n' +
    `Test Mode:              ${cfg.TEST_MODE ? 'ENABLED (rows 2–6 only)' : 'DISABLED (all rows)'}\n` +
    `Batch Size:             ${cfg.BATCH_SIZE} emails per API call\n` +
    `Delay Between Batches:  ${cfg.DELAY_BETWEEN_BATCHES / 1000}s (rate limit: 5 calls/min)\n` +
    `Request Timeout:        ${cfg.REQUEST_TIMEOUT / 1000}s per batch\n` +
    `Max Retries:            ${cfg.MAX_RETRIES}\n` +
    `Time Limit Per Run:     ${cfg.TIME_LIMIT_SECONDS}s\n` +
    `Continuation Delay:     ${cfg.CONTINUATION_DELAY_MINUTES} minute(s)\n\n` +
    'Performance estimate:\n' +
    '  500 emails/min → 5,000 rows ≈ 10 minutes\n\n' +
    'Edit EMAIL_VALIDATION_CONFIG in the script editor to change any setting.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showEmailValidationLogs() {
  const logs = Logger.getLog();
  SpreadsheetApp.getUi().alert(
    '📋 Email Validation Logs',
    logs || 'No logs available. Run a validation first.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// END OF SCRIPT
// =============================================================================

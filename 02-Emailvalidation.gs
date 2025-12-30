/**
 * =============================================================================
 * ZEROBOUNCE EMAIL VALIDATION SCRIPT
 * =============================================================================
 * 
 * Validates email addresses using ZeroBounce Email Validation API
 * 
 * IMPORTANT: This file should be in the same Apps Script project as:
 *            DataCleaningPipeline_V3.gs
 *            PhoneValidation_Byteplant.gs
 * 
 * The onOpen() function is located in DataCleaningPipeline_V3.gs and creates
 * menus for all scripts.
 * 
 * API SPECIFICATION:
 * - Endpoint: https://api.zerobounce.net/v2/validate
 * - Method: GET (with query parameters)
 * - Required Parameters: api_key, email
 * - Optional Parameters: ip_address
 * 
 * STATUS VALUES:
 * - valid: Email is valid and safe to send
 * - invalid: Email is invalid and will bounce
 * - catch-all: Domain accepts all emails (risky)
 * - unknown: Cannot determine validity
 * - spamtrap: Email is a spam trap (dangerous!)
 * - abuse: Email is known for abuse complaints
 * - do_not_mail: Email should not be mailed
 * 
 * SUB-STATUS VALUES:
 * - antispam_system, greylisted, mail_server_temporary_error
 * - mailbox_quota_exceeded, mailbox_not_found, no_dns_entries
 * - possible_typo, disposable, toxic, role_based, etc.
 * 
 * TEST MODE: Processes only the first 5 data rows (rows 2-6)
 * 
 * Features:
 * - Validates Primary Email, Email 1, Email 2, Personal Email
 * - Processes ONE email column at a time (sequential)
 * - Creates 1 result column immediately after each email column:
 *   • _Status (valid, invalid, catch-all, unknown, spamtrap, abuse, do_not_mail)
 * - Result column has BLUE background with BLACK text
 * - Includes delay between API calls to avoid rate limits
 * - Handles errors gracefully with retry logic
 * - Easy to extend to more rows later
 * 
 * @author: Claude
 * @version: 1.1 - Simplified to Status column only
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const EMAIL_VALIDATION_CONFIG = {
  // ZeroBounce API Configuration
  API_KEY: '48891e1db6e34f00a0d4062edc9fd5a4',
  API_ENDPOINT: 'https://api.zerobounce.net/v2/validate',
  
  // Test Mode Settings
  TEST_MODE: false,          // Set to true to only process 5 rows for testing
  TEST_ROWS_START: 2,        // First data row (row 2 = first after header)
  TEST_ROWS_END: 6,          // Last data row for test (row 6 = 5 data rows)
  
  // Column Names (must match your sheet headers)
  EMAIL_COLUMNS: [
    'Primary Email',
    'Email 1',
    'Email 2',
    'Personal Email'
  ],
  
  // Result Column Suffixes (will be added next to each email column)
  RESULT_COLUMNS: {
    STATUS: '_Status'           // e.g., "Primary Email_Status"
  },
  
  // API Settings
  DELAY_BETWEEN_CALLS: 500,  // Milliseconds between API calls (0.5 seconds)
  REQUEST_TIMEOUT: 10000,    // Request timeout in milliseconds (10 seconds)
  
  // Error Handling
  MAX_RETRIES: 2,            // Number of retry attempts for failed requests
  RETRY_DELAY: 1000          // Delay between retries in milliseconds
};

// =============================================================================
// MAIN VALIDATION FUNCTIONS
// =============================================================================

/**
 * Main validation function for a single email column
 * @param {string} columnName - The email column to validate
 */
function validateEmailColumn_Single(columnName) {
  try {
    const startTime = new Date();
    Logger.log(`=== Starting Email Validation for ${columnName} ===`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Get sheet data
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Determine rows to process
    const startRow = EMAIL_VALIDATION_CONFIG.TEST_ROWS_START;
    const endRow = EMAIL_VALIDATION_CONFIG.TEST_MODE 
      ? Math.min(EMAIL_VALIDATION_CONFIG.TEST_ROWS_END, lastRow)
      : lastRow;
    
    Logger.log(`Processing rows ${startRow} to ${endRow}`);
    Logger.log(`Test mode: ${EMAIL_VALIDATION_CONFIG.TEST_MODE}`);
    
    // Step 1: Prepare result columns for this email column
    Logger.log('Step 1: Preparing result columns...');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    prepareResultColumnsForEmail(sheet, headers, columnName);
    
    // Step 2: Get updated headers after adding columns
    const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Step 3: Validate all rows for this email column
    Logger.log('Step 2: Validating email addresses...');
    const result = validateEmailColumn(
      sheet,
      updatedHeaders,
      columnName,
      startRow,
      endRow
    );
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log('\n=== Validation Complete ===');
    Logger.log(`Column: ${columnName}`);
    Logger.log(`Rows processed: ${startRow}-${endRow} (${endRow - startRow + 1} rows)`);
    Logger.log(`Email addresses processed: ${result.processed}`);
    Logger.log(`Successfully validated: ${result.validated}`);
    Logger.log(`Errors: ${result.errors}`);
    Logger.log(`Duration: ${duration.toFixed(2)}s`);
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      `${columnName} - Validation Complete`,
      `Email validation complete for ${columnName}!\n\n` +
      `Rows: ${startRow}-${endRow}\n` +
      `Processed: ${result.processed} email addresses\n` +
      `Validated: ${result.validated}\n` +
      `Errors: ${result.errors}\n` +
      `Duration: ${duration.toFixed(2)}s\n\n` +
      `${EMAIL_VALIDATION_CONFIG.TEST_MODE ? '✓ TEST MODE (5 rows only)' : '✓ FULL MODE'}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      `Email validation failed for ${columnName}: ` + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Validate Primary Email column (5 rows in test mode)
 */
function validatePrimaryEmail() {
  validateEmailColumn_Single('Primary Email');
}

/**
 * Validate Email 1 column (5 rows in test mode)
 */
function validateEmail1() {
  validateEmailColumn_Single('Email 1');
}

/**
 * Validate Email 2 column (5 rows in test mode)
 */
function validateEmail2() {
  validateEmailColumn_Single('Email 2');
}

/**
 * Validate Personal Email column (5 rows in test mode)
 */
function validatePersonalEmail() {
  validateEmailColumn_Single('Personal Email');
}

// =============================================================================
// COLUMN PREPARATION
// =============================================================================

/**
 * Prepares result columns immediately after a specific email column
 * Inserts 1 column: Status
 * Applies blue background with black text formatting
 */
function prepareResultColumnsForEmail(sheet, headers, emailColumn) {
  Logger.log(`Preparing result column for: ${emailColumn}`);
  
  const emailColIndex = headers.indexOf(emailColumn);
  
  if (emailColIndex === -1) {
    Logger.log(`Warning: Column "${emailColumn}" not found`);
    return;
  }
  
  // Define result column name
  const statusColName = emailColumn + EMAIL_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  
  // Check if column already exists
  if (headers.includes(statusColName)) {
    Logger.log(`Result column already exists for ${emailColumn}`);
    return;
  }
  
  // Insert 1 new column immediately after the email column
  // emailColIndex is 0-based, but insertColumnsAfter uses 1-based
  const insertAfterCol = emailColIndex + 1;
  sheet.insertColumnsAfter(insertAfterCol, 1);
  
  // Set header for the new column
  const headerRow = 1;
  const firstNewCol = insertAfterCol + 1; // First column after email column
  
  sheet.getRange(headerRow, firstNewCol).setValue(statusColName);
  
  // Format header: Blue background (#4285f4), Black text, Bold
  const headerRange = sheet.getRange(headerRow, firstNewCol, 1, 1);
  headerRange
    .setBackground('#4285f4')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  Logger.log(`Created 1 result column after ${emailColumn} (column ${insertAfterCol})`);
}

// =============================================================================
// EMAIL VALIDATION
// =============================================================================

/**
 * Validates all email addresses in a specific column
 */
function validateEmailColumn(sheet, headers, emailColumn, startRow, endRow) {
  const emailColIndex = headers.indexOf(emailColumn);
  
  if (emailColIndex === -1) {
    Logger.log(`Column "${emailColumn}" not found, skipping`);
    return { processed: 0, validated: 0, errors: 0 };
  }
  
  // Find result column index
  const statusColName = emailColumn + EMAIL_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  const statusColIndex = headers.indexOf(statusColName);
  
  if (statusColIndex === -1) {
    Logger.log(`Result column not found for ${emailColumn}, skipping`);
    return { processed: 0, validated: 0, errors: 0 };
  }
  
  let processed = 0;
  let validated = 0;
  let errors = 0;
  
  // Process each row
  for (let row = startRow; row <= endRow; row++) {
    const emailAddress = sheet.getRange(row, emailColIndex + 1).getValue();
    
    // Skip empty cells
    if (!emailAddress || emailAddress.toString().trim() === '') {
      Logger.log(`Row ${row}: Empty email address, skipping`);
      continue;
    }
    
    // ✅ SKIP if already validated (check if status column has data)
    const existingStatus = sheet.getRange(row, statusColIndex + 1).getValue();
    if (existingStatus && existingStatus.toString().trim() !== '') {
      Logger.log(`Row ${row}: Already validated (status: ${existingStatus}), skipping`);
      continue;
    }
    
    processed++;
    Logger.log(`Row ${row}: Validating ${emailAddress}`);
    
    // Validate email address
    const result = validateEmailWithAPI(emailAddress.toString().trim());
    
    if (result.success) {
      // Write only status to sheet
      sheet.getRange(row, statusColIndex + 1).setValue(result.status || '');
      
      validated++;
      Logger.log(`Row ${row}: ✓ Status: ${result.status}`);
    } else {
      // Write error to status column
      sheet.getRange(row, statusColIndex + 1).setValue('ERROR: ' + result.error);
      
      errors++;
      Logger.log(`Row ${row}: ✗ Error - ${result.error}`);
    }
    
    // Add delay between API calls (except for last row)
    if (row < endRow) {
      Utilities.sleep(EMAIL_VALIDATION_CONFIG.DELAY_BETWEEN_CALLS);
    }
  }
  
  return { processed, validated, errors };
}

/**
 * Calls the ZeroBounce API to validate an email address
 */
function validateEmailWithAPI(emailAddress) {
  const apiKey = EMAIL_VALIDATION_CONFIG.API_KEY;
  const endpoint = EMAIL_VALIDATION_CONFIG.API_ENDPOINT;
  
  // Try with retries
  for (let attempt = 1; attempt <= EMAIL_VALIDATION_CONFIG.MAX_RETRIES + 1; attempt++) {
    try {
      Logger.log(`API call attempt ${attempt}: ${emailAddress}`);
      
      // Build URL with query parameters (GET method)
      const encodedEmail = encodeURIComponent(emailAddress);
      const url = `${endpoint}?api_key=${apiKey}&email=${encodedEmail}&ip_address=`;
      
      const options = {
        method: 'get',
        muteHttpExceptions: true,
        timeout: EMAIL_VALIDATION_CONFIG.REQUEST_TIMEOUT
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log(`Response code: ${responseCode}`);
      Logger.log(`Response: ${responseText.substring(0, 200)}`); // Log first 200 chars
      
      if (responseCode === 200) {
        const data = JSON.parse(responseText);
        
        // Check if we have a status field
        if (!data.status) {
          return {
            success: false,
            error: 'No status in API response'
          };
        }
        
        // Return all relevant data
        return {
          success: true,
          status: data.status || '',
          sub_status: data.sub_status || '',
          did_you_mean: data.did_you_mean || '',
          free_email: data.free_email ? 'Yes' : 'No',
          smtp_provider: data.smtp_provider || '',
          domain_age_days: data.domain_age_days || ''
        };
      } else {
        // Non-200 response
        Logger.log(`API returned ${responseCode}: ${responseText}`);
        
        if (attempt <= EMAIL_VALIDATION_CONFIG.MAX_RETRIES) {
          Logger.log(`Retrying in ${EMAIL_VALIDATION_CONFIG.RETRY_DELAY}ms...`);
          Utilities.sleep(EMAIL_VALIDATION_CONFIG.RETRY_DELAY);
          continue;
        }
        
        return {
          success: false,
          error: `API error ${responseCode}`
        };
      }
      
    } catch (error) {
      Logger.log(`Attempt ${attempt} failed: ${error.toString()}`);
      
      if (attempt <= EMAIL_VALIDATION_CONFIG.MAX_RETRIES) {
        Logger.log(`Retrying in ${EMAIL_VALIDATION_CONFIG.RETRY_DELAY}ms...`);
        Utilities.sleep(EMAIL_VALIDATION_CONFIG.RETRY_DELAY);
        continue;
      }
      
      return {
        success: false,
        error: error.toString().substring(0, 100)
      };
    }
  }
  
  return {
    success: false,
    error: 'Max retries exceeded'
  };
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Shows configuration dialog
 */
function showEmailConfigDialog() {
  const html = HtmlService.createHtmlOutput(`
    <h3>Email Validation Configuration</h3>
    <p><strong>API Endpoint:</strong> ${EMAIL_VALIDATION_CONFIG.API_ENDPOINT}</p>
    <p><strong>Test Mode:</strong> ${EMAIL_VALIDATION_CONFIG.TEST_MODE ? 'ON' : 'OFF'}</p>
    <p><strong>Test Rows:</strong> ${EMAIL_VALIDATION_CONFIG.TEST_ROWS_START} to ${EMAIL_VALIDATION_CONFIG.TEST_ROWS_END}</p>
    <p><strong>Email Columns:</strong></p>
    <ul>
      ${EMAIL_VALIDATION_CONFIG.EMAIL_COLUMNS.map(col => '<li>' + col + '</li>').join('')}
    </ul>
    <p><strong>Result Columns:</strong> 1 column per email (_Status only)</p>
    <p><strong>Delay Between Calls:</strong> ${EMAIL_VALIDATION_CONFIG.DELAY_BETWEEN_CALLS}ms</p>
    <p><strong>Max Retries:</strong> ${EMAIL_VALIDATION_CONFIG.MAX_RETRIES}</p>
    <br>
    <p><em>To change these settings, edit the EMAIL_VALIDATION_CONFIG in the script.</em></p>
  `)
    .setWidth(400)
    .setHeight(450);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Email Validation Settings');
}

/**
 * Shows execution logs for email validation
 */
function showEmailValidationLogs() {
  const logs = Logger.getLog();
  const ui = SpreadsheetApp.getUi();
  
  if (logs) {
    ui.alert('Email Validation Logs', logs, ui.ButtonSet.OK);
  } else {
    ui.alert('Email Validation Logs', 'No logs available. Run the validation first.', ui.ButtonSet.OK);
  }
}

/**
 * =============================================================================
 * BYTEPLANT PHONE VALIDATION SCRIPT
 * =============================================================================
 * 
 * Validates phone numbers using Byteplant Real-Time Phone Validation API
 * 
 * IMPORTANT: This file should be in the same Apps Script project as:
 *            DataCleaningPipeline_V3.gs
 * 
 * The onOpen() function is located in DataCleaningPipeline_V3.gs and creates
 * menus for both scripts.
 * 
 * API SPECIFICATION:
 * - Endpoint: https://api.phone-validator.net/api/v2/verify
 * - Method: GET (with query parameters)
 * - Required Parameters: PhoneNumber, CountryCode, APIKey
 * - Phone Format: National format (e.g., "2109791958") with CountryCode="us"
 *                 OR International format (e.g., "+12109791958") without CountryCode
 * 
 * STATUS VALUES:
 * - VALID_CONFIRMED: Phone is valid and confirmed active
 * - VALID_UNCONFIRMED: Phone is valid but cannot confirm if active
 * - INVALID: Phone number is invalid
 * - DELAYED: Validation is delayed (async processing)
 * - RATE_LIMIT_EXCEEDED: API rate limit reached
 * - API_KEY_INVALID_OR_DEPLETED: API key issue
 * 
 * LINE TYPES:
 * - MOBILE, FIXED_LINE (landline), VOIP, TOLL_FREE, PREMIUM_RATE,
 *   SHARED_COST, PERSONAL_NUMBER, PAGER, UAN, VOICEMAIL
 * 
 * TEST MODE: Processes only the first 5 data rows (rows 2-6)
 * 
 * Features:
 * - Validates 4 phone columns individually:
 *   • Contact Phone 1
 *   • Company Phone 1
 *   • Company Phone 2
 *   • Contact Mobile Phone
 * - Each column validated separately via menu
 * - Creates 4 result columns immediately after each phone column:
 *   • _Status (valid confirmed, invalid, etc.)
 *   • _Line Type (mobile, landline, voip, etc.)
 *   • _Location (city, state)
 *   • _International (formatted phone number)
 * - Result columns have BLUE background with BLACK text
 * - Includes delay between API calls to avoid rate limits
 * - Handles errors gracefully with retry logic
 * - TEST MODE: Validates only rows 2-6 (5 rows) per column
 * 
 * @author: Claude
 * @version: 1.3 - Individual column validation with 4 phone columns
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const PHONE_VALIDATION_CONFIG = {
  // Byteplant API Configuration
  API_KEY: 'pv-0e4b530a9be44e529631bd5cc323279d',
  API_ENDPOINT: 'https://api.phone-validator.net/api/v2/verify',
  COUNTRY_CODE: 'us',        // Two-letter ISO 3166-1 country code (use with national format numbers)
  
  // Test Mode Settings
  TEST_MODE: false,          // Set to true to only process 5 rows for testing
  TEST_ROWS_START: 2,        // First data row (row 2 = first after header)
  TEST_ROWS_END: 6,          // Last data row for test (row 6 = 5 data rows)
  
  // Column Names (must match your sheet headers)
  PHONE_COLUMNS: [
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2',
    'Contact Mobile Phone'
  ],
  
  // Result Column Suffixes (will be added next to each phone column)
  RESULT_COLUMNS: {
    STATUS: '_Status',           // e.g., "Company Phone 1_Status"
    LINE_TYPE: '_Line Type',     // e.g., "Company Phone 1_Line Type"
    LOCATION: '_Location',       // e.g., "Company Phone 1_Location"
    FORMAT: '_International'     // e.g., "Company Phone 1_International"
  },
  
  // API Settings
  DELAY_BETWEEN_CALLS: 500,  // Milliseconds between API calls (0.5 seconds)
  REQUEST_TIMEOUT: 10000,    // Request timeout in milliseconds (10 seconds)
  
  // Error Handling
  MAX_RETRIES: 2,            // Number of retries on API failure
  RETRY_DELAY: 1000          // Delay before retry (1 second)
};

// =============================================================================
// MAIN EXECUTION
// =============================================================================

/**
 * Main validation function for a single column
 * @param {string} columnName - The phone column to validate
 */
function validatePhoneColumn_Single(columnName) {
  try {
    const startTime = new Date();
    Logger.log(`=== Starting Phone Validation for ${columnName} ===`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Get sheet data
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // Determine rows to process
    const startRow = PHONE_VALIDATION_CONFIG.TEST_ROWS_START;
    const endRow = PHONE_VALIDATION_CONFIG.TEST_MODE 
      ? Math.min(PHONE_VALIDATION_CONFIG.TEST_ROWS_END, lastRow)
      : lastRow;
    
    Logger.log(`Processing rows ${startRow} to ${endRow}`);
    Logger.log(`Test mode: ${PHONE_VALIDATION_CONFIG.TEST_MODE}`);
    
    // Step 1: Prepare result columns for this phone column
    Logger.log('Step 1: Preparing result columns...');
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    prepareResultColumnsForPhone(sheet, headers, columnName);
    
    // Step 2: Get updated headers after adding columns
    const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Step 3: Validate all rows for this phone column
    Logger.log('Step 2: Validating phone numbers...');
    const result = validatePhoneColumn(
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
    Logger.log(`Phone numbers processed: ${result.processed}`);
    Logger.log(`Successfully validated: ${result.validated}`);
    Logger.log(`Errors: ${result.errors}`);
    Logger.log(`Duration: ${duration.toFixed(2)}s`);
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      `${columnName} - Validation Complete`,
      `Phone validation complete for ${columnName}!\n\n` +
      `Rows: ${startRow}-${endRow}\n` +
      `Processed: ${result.processed} phone numbers\n` +
      `Validated: ${result.validated}\n` +
      `Errors: ${result.errors}\n` +
      `Duration: ${duration.toFixed(2)}s\n\n` +
      `${PHONE_VALIDATION_CONFIG.TEST_MODE ? '✓ TEST MODE (5 rows only)' : '✓ FULL MODE'}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      `Phone validation failed for ${columnName}: ` + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Validate Contact Phone 1 column (5 rows in test mode)
 */
function validateContactPhone1() {
  validatePhoneColumn_Single('Contact Phone 1');
}

/**
 * Validate Company Phone 1 column (5 rows in test mode)
 */
function validateCompanyPhone1() {
  validatePhoneColumn_Single('Company Phone 1');
}

/**
 * Validate Company Phone 2 column (5 rows in test mode)
 */
function validateCompanyPhone2() {
  validatePhoneColumn_Single('Company Phone 2');
}

/**
 * Validate Contact Mobile Phone column (5 rows in test mode)
 */
function validateContactMobilePhone() {
  validatePhoneColumn_Single('Contact Mobile Phone');
}

// =============================================================================
// COLUMN PREPARATION
// =============================================================================

/**
 * Prepares result columns immediately after a specific phone column
 * Inserts 4 columns: Status, Line Type, Location, International
 * Applies blue background with black text formatting
 */
function prepareResultColumnsForPhone(sheet, headers, phoneColumn) {
  Logger.log(`Preparing result columns for: ${phoneColumn}`);
  
  const phoneColIndex = headers.indexOf(phoneColumn);
  
  if (phoneColIndex === -1) {
    Logger.log(`Warning: Column "${phoneColumn}" not found`);
    return;
  }
  
  // Define result column names
  const statusColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  const lineTypeColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LINE_TYPE;
  const locationColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LOCATION;
  const formatColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.FORMAT;
  
  // Check if columns already exist
  if (headers.includes(statusColName)) {
    Logger.log(`Result columns already exist for ${phoneColumn}`);
    return;
  }
  
  // Insert 4 new columns immediately after the phone column
  // phoneColIndex is 0-based, but insertColumnsAfter uses 1-based
  const insertAfterCol = phoneColIndex + 1;
  sheet.insertColumnsAfter(insertAfterCol, 4);
  
  // Set headers for the 4 new columns
  const headerRow = 1;
  const firstNewCol = insertAfterCol + 1; // First column after phone column
  
  sheet.getRange(headerRow, firstNewCol).setValue(statusColName);
  sheet.getRange(headerRow, firstNewCol + 1).setValue(lineTypeColName);
  sheet.getRange(headerRow, firstNewCol + 2).setValue(locationColName);
  sheet.getRange(headerRow, firstNewCol + 3).setValue(formatColName);
  
  // Format headers: Blue background (#4285f4), Black text, Bold
  const headerRange = sheet.getRange(headerRow, firstNewCol, 1, 4);
  headerRange
    .setBackground('#4285f4')
    .setFontColor('#000000')
    .setFontWeight('bold');
  
  Logger.log(`Created 4 result columns after ${phoneColumn} (column ${insertAfterCol})`);
}

// =============================================================================
// PHONE VALIDATION
// =============================================================================

/**
 * Validates all phone numbers in a specific column
 */
function validatePhoneColumn(sheet, headers, phoneColumn, startRow, endRow) {
  const phoneColIndex = headers.indexOf(phoneColumn);
  
  if (phoneColIndex === -1) {
    Logger.log(`Column "${phoneColumn}" not found, skipping`);
    return { processed: 0, validated: 0, errors: 0 };
  }
  
  // Find result column indices
  const statusColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.STATUS;
  const lineTypeColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LINE_TYPE;
  const locationColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.LOCATION;
  const formatColName = phoneColumn + PHONE_VALIDATION_CONFIG.RESULT_COLUMNS.FORMAT;
  
  const statusColIndex = headers.indexOf(statusColName);
  const lineTypeColIndex = headers.indexOf(lineTypeColName);
  const locationColIndex = headers.indexOf(locationColName);
  const formatColIndex = headers.indexOf(formatColName);
  
  if (statusColIndex === -1) {
    Logger.log(`Result columns not found for ${phoneColumn}, skipping`);
    return { processed: 0, validated: 0, errors: 0 };
  }
  
  let processed = 0;
  let validated = 0;
  let errors = 0;
  
  // Process each row
  for (let row = startRow; row <= endRow; row++) {
    const phoneNumber = sheet.getRange(row, phoneColIndex + 1).getValue();
    
    // Skip empty cells
    if (!phoneNumber || phoneNumber.toString().trim() === '') {
      Logger.log(`Row ${row}: Empty phone number, skipping`);
      continue;
    }
    
    // ✅ SKIP if already validated (but RE-VALIDATE if DELAYED)
    const existingStatus = sheet.getRange(row, statusColIndex + 1).getValue();
    if (existingStatus && existingStatus.toString().trim() !== '') {
      const statusUpper = existingStatus.toString().trim().toUpperCase();
      
      // Re-validate DELAYED (it's temporary, needs another check)
      if (statusUpper === 'DELAYED') {
        Logger.log(`Row ${row}: Status is DELAYED, re-validating...`);
        // Don't skip - continue to validation
      } else {
        // Skip all other statuses (VALID_CONFIRMED, INVALID, ERROR, etc.)
        Logger.log(`Row ${row}: Already validated (status: ${existingStatus}), skipping`);
        continue;
      }
    }
    
    processed++;
    Logger.log(`Row ${row}: Validating ${phoneNumber}`);
    
    // Validate phone number
    const result = validatePhoneWithAPI(phoneNumber.toString().trim());
    
    if (result.success) {
      // Write results to sheet (4 columns)
      sheet.getRange(row, statusColIndex + 1).setValue(result.phone_status || '');
      sheet.getRange(row, lineTypeColIndex + 1).setValue(result.line_type || '');
      sheet.getRange(row, locationColIndex + 1).setValue(result.location || '');
      sheet.getRange(row, formatColIndex + 1).setValue(result.format_international || '');
      
      validated++;
      Logger.log(`Row ${row}: ✓ Valid - Status: ${result.phone_status}, Type: ${result.line_type}, Location: ${result.location}`);
    } else {
      // Write error to status column, clear others
      sheet.getRange(row, statusColIndex + 1).setValue('ERROR: ' + result.error);
      sheet.getRange(row, lineTypeColIndex + 1).setValue('');
      sheet.getRange(row, locationColIndex + 1).setValue('');
      sheet.getRange(row, formatColIndex + 1).setValue('');
      
      errors++;
      Logger.log(`Row ${row}: ✗ Error - ${result.error}`);
    }
    
    // Add delay between API calls (except for last row)
    if (row < endRow) {
      Utilities.sleep(PHONE_VALIDATION_CONFIG.DELAY_BETWEEN_CALLS);
    }
  }
  
  return { processed, validated, errors };
}

/**
 * Calls the Byteplant API to validate a phone number
 */
function validatePhoneWithAPI(phoneNumber) {
  const apiKey = PHONE_VALIDATION_CONFIG.API_KEY;
  const endpoint = PHONE_VALIDATION_CONFIG.API_ENDPOINT;
  const countryCode = PHONE_VALIDATION_CONFIG.COUNTRY_CODE;
  
  // Phone number in national format (no need to add +1)
  const formattedPhone = phoneNumber.toString().trim();
  
  // Try with retries
  for (let attempt = 1; attempt <= PHONE_VALIDATION_CONFIG.MAX_RETRIES + 1; attempt++) {
    try {
      Logger.log(`API call attempt ${attempt}: ${formattedPhone}`);
      
      // Build URL with query parameters (GET method)
      // Byteplant accepts both GET and POST, GET is simpler
      const encodedPhone = encodeURIComponent(formattedPhone);
      const url = `${endpoint}?PhoneNumber=${encodedPhone}&CountryCode=${countryCode}&APIKey=${apiKey}`;
      
      const options = {
        method: 'get',
        muteHttpExceptions: true,
        timeout: PHONE_VALIDATION_CONFIG.REQUEST_TIMEOUT
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      Logger.log(`Response code: ${responseCode}`);
      Logger.log(`Response: ${responseText.substring(0, 200)}`); // Log first 200 chars
      
      if (responseCode === 200) {
        const data = JSON.parse(responseText);
        
        // Check status field
        if (!data.status) {
          return {
            success: false,
            error: 'No status in API response'
          };
        }
        
        const status = data.status.toUpperCase();
        
        // Handle different status codes
        if (status === 'VALID_CONFIRMED') {
          return {
            success: true,
            phone_status: data.status,  // Keep EXACT API value: VALID_CONFIRMED
            line_type: data.linetype || '',  // Keep EXACT API value: FIXED_LINE, MOBILE, etc.
            location: data.location || '',
            format_international: data.formatinternational || formattedPhone
          };
        } else if (status === 'VALID_UNCONFIRMED') {
          return {
            success: true,
            phone_status: data.status,  // Keep EXACT API value: VALID_UNCONFIRMED
            line_type: data.linetype || '',  // Keep EXACT API value
            location: data.location || '',
            format_international: data.formatinternational || formattedPhone
          };
        } else if (status === 'INVALID') {
          return {
            success: true,
            phone_status: data.status,  // Keep EXACT API value: INVALID
            line_type: '',
            location: '',
            format_international: ''
          };
        } else if (status === 'RATE_LIMIT_EXCEEDED') {
          return {
            success: false,
            error: 'Rate limit exceeded. Wait and try again.'
          };
        } else if (status === 'API_KEY_INVALID_OR_DEPLETED') {
          return {
            success: false,
            error: 'API key invalid or depleted'
          };
        } else if (status === 'DELAYED') {
          return {
            success: true,
            phone_status: data.status,  // Keep EXACT API value: DELAYED
            line_type: data.linetype || '',  // Keep EXACT API value
            location: data.location || '',
            format_international: data.formatinternational || formattedPhone
          };
        } else {
          // Unknown status - keep exact API value
          return {
            success: true,
            phone_status: data.status,  // Keep whatever API returns (uppercase)
            line_type: data.linetype || '',  // Keep EXACT API value
            location: data.location || '',
            format_international: data.formatinternational || formattedPhone
          };
        }
      } else {
        // Non-200 response
        Logger.log(`API returned ${responseCode}: ${responseText}`);
        
        if (attempt <= PHONE_VALIDATION_CONFIG.MAX_RETRIES) {
          Logger.log(`Retrying in ${PHONE_VALIDATION_CONFIG.RETRY_DELAY}ms...`);
          Utilities.sleep(PHONE_VALIDATION_CONFIG.RETRY_DELAY);
          continue;
        }
        
        return {
          success: false,
          error: `API error ${responseCode}`
        };
      }
      
    } catch (error) {
      Logger.log(`Attempt ${attempt} failed: ${error.toString()}`);
      
      if (attempt <= PHONE_VALIDATION_CONFIG.MAX_RETRIES) {
        Logger.log(`Retrying in ${PHONE_VALIDATION_CONFIG.RETRY_DELAY}ms...`);
        Utilities.sleep(PHONE_VALIDATION_CONFIG.RETRY_DELAY);
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

/**
 * Formats line type from API response to readable format
 */
// NO LONGER USED - Now using exact API response values
// function formatLineType(linetype) {
//   if (!linetype) return '';
//   
//   // Convert from MOBILE, FIXED_LINE, etc. to readable format
//   const typeMap = {
//     'MOBILE': 'mobile',
//     'FIXED_LINE': 'landline',
//     'VOIP': 'voip',
//     'TOLL_FREE': 'toll-free',
//     'PREMIUM_RATE': 'premium',
//     'SHARED_COST': 'shared cost',
//     'PERSONAL_NUMBER': 'personal',
//     'PAGER': 'pager',
//     'UAN': 'uan',
//     'VOICEMAIL': 'voicemail'
//   };
//   
//   const upper = linetype.toUpperCase();
//   return typeMap[upper] || linetype.toLowerCase();
// }

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Shows configuration dialog
 */
function showConfigDialog() {
  const config = PHONE_VALIDATION_CONFIG;
  const message = 
    'Current Configuration:\n\n' +
    `Test Mode: ${config.TEST_MODE ? 'ENABLED' : 'DISABLED'}\n` +
    `Test Rows: ${config.TEST_ROWS_START}-${config.TEST_ROWS_END}\n` +
    `Delay: ${config.DELAY_BETWEEN_CALLS}ms\n` +
    `Max Retries: ${config.MAX_RETRIES}\n\n` +
    'To change settings, edit the PHONE_VALIDATION_CONFIG object in the script editor.\n\n' +
    'To process ALL rows:\n' +
    '1. Set TEST_MODE = false\n' +
    '2. Save and refresh the sheet\n' +
    '3. Run validation again';
  
  SpreadsheetApp.getUi().alert('Configuration', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Shows execution logs for phone validation
 */
function showValidationLogs() {
  const logs = Logger.getLog();
  const ui = SpreadsheetApp.getUi();
  
  if (logs) {
    ui.alert('Phone Validation Logs', logs, ui.ButtonSet.OK);
  } else {
    ui.alert('Phone Validation Logs', 'No logs available. Run the validation first.', ui.ButtonSet.OK);
  }
}

// =============================================================================
// INDIVIDUAL COLUMN VALIDATION FUNCTIONS (Menu Items)
// =============================================================================

/**
 * Validates Contact Phone 1 column (rows 2-6 in test mode)
 */
function validateContactPhone1() {
  validatePhoneColumn_Single('Contact Phone 1');
}

/**
 * Validates Company Phone 1 column (rows 2-6 in test mode)
 */
function validateCompanyPhone1() {
  validatePhoneColumn_Single('Company Phone 1');
}

/**
 * Validates Company Phone 2 column (rows 2-6 in test mode)
 */
function validateCompanyPhone2() {
  validatePhoneColumn_Single('Company Phone 2');
}

/**
 * Validates Contact Mobile Phone column (rows 2-6 in test mode)
 */
function validateContactMobilePhone() {
  validatePhoneColumn_Single('Contact Mobile Phone');
}

// =============================================================================
// END OF SCRIPT
// =============================================================================

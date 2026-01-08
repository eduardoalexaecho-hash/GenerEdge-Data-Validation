/**
 * =============================================================================
 * CRM READY FILTER - V2.2 (ENHANCED LOGGING)
 * =============================================================================
 * 
 * Creates a single "CRM Ready" sheet with cleaned, validated data
 * 
 * V2.2 CHANGES:
 * - Enhanced logging to show which columns are validated vs skipped
 * - Clear visibility into which emails/phones are considered for inclusion
 * - Transparent reporting of validation status
 * 
 * V2.1 CHANGES:
 * - CRITICAL FIX: Initialize ALL email and phone columns to empty first
 * - Only fill in columns with explicit validation status
 * - Prevents unvalidated data from appearing in output
 * - Ensures only validated emails (status='valid') appear
 * - Ensures only validated mobiles (status='valid confirmed' + type='mobile') appear
 * 
 * CRITERIA:
 * - At least one valid email (status='valid')
 * - At least one valid mobile phone (status='valid confirmed' + type='mobile')
 * 
 * VALIDATION RULES:
 * - Emails: ONLY included if status column exists AND status='valid'
 * - Phones: ONLY included if status AND line type columns exist AND status='valid confirmed' AND type='mobile'
 * - Columns without validation columns: NEVER considered for inclusion or output
 * 
 * DATA CLEANING:
 * - Only keeps valid emails (invalid/unvalidated emails are cleared/empty)
 * - Only keeps valid mobile phones (invalid/unvalidated/landline phones are cleared/empty)
 * - Extracts state from first valid phone location
 * 
 * OUTPUT COLUMNS (FIXED ORDER):
 * organization, first_name, last_name, Primary Email, Personal Email, 
 * Email 1, Email 2, Contact Phone 1, Company Phone 1, Company Phone 2, 
 * Contact Mobile Phone, description, website, city, state
 * 
 * @author: Claude
 * @version: 2.2 - Enhanced logging for transparency
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CRM_CONFIG = {
  // Email column names (must match exactly)
  EMAIL_COLUMNS: {
    'Primary Email': 'Primary Email',
    'Personal Email': 'Personal Email',
    'Email 1': 'Email 1',
    'Email 2': 'Email 2'
  },
  
  // Phone column names (must match exactly)
  PHONE_COLUMNS: {
    'Contact Phone 1': 'Contact Phone 1',
    'Company Phone 1': 'Company Phone 1',
    'Company Phone 2': 'Company Phone 2',
    'Contact Mobile Phone': 'Contact Mobile Phone'
  },
  
  // Base column names (original data)
  BASE_COLUMNS: {
    'organization': 'Organization',
    'first_name': 'First Name',
    'last_name': 'Last Name',
    'description': 'Company Description',
    'website': 'Website',
    'city': 'Company City'
  },
  
  // Validation column suffixes
  STATUS_SUFFIX: '_Status',
  LINE_TYPE_SUFFIX: '_Line Type',
  LOCATION_SUFFIX: '_Location',
  
  // Valid criteria
  VALID_EMAIL_STATUS: 'valid',
  VALID_PHONE_STATUS: 'valid confirmed',
  VALID_LINE_TYPE: 'mobile',
  
  // Output sheet name
  OUTPUT_SHEET: 'CRM Ready',
  
  // Output column order (EXACT ORDER - DO NOT CHANGE)
  OUTPUT_COLUMNS: [
    'organization',
    'first_name',
    'last_name',
    'Primary Email',
    'Personal Email',
    'Email 1',
    'Email 2',
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2',
    'Contact Mobile Phone',
    'description',
    'website',
    'city',
    'state'
  ],
  
  FIRST_DATA_ROW: 2
};

// =============================================================================
// MAIN FUNCTION
// =============================================================================

/**
 * Creates CRM Ready sheet with cleaned, validated data
 */
function createCRMReady() {
  const startTime = new Date();
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getActiveSheet();
  
  if (!spreadsheet || !sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'No active spreadsheet or sheet found.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log('=== CREATING CRM READY SHEET ===');
  Logger.log(`Spreadsheet: ${spreadsheet.getName()}`);
  Logger.log(`Active Sheet: ${sourceSheet.getName()}`);
  
  // Find all required columns
  const columnMap = findAllColumns(sourceSheet);
  
  if (!columnMap) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Could not find required columns!\n\n' +
      'Please ensure you have:\n' +
      '1. Run email validation (at least one email column)\n' +
      '2. Run phone validation (at least one phone column)\n\n' +
      'Then try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log('Found all required columns');
  
  // Process data
  const crmData = processCRMData(sourceSheet, columnMap);
  
  if (crmData.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Data Found',
      'No contacts meet the criteria:\n' +
      '- At least one valid email\n' +
      '- At least one valid mobile phone\n\n' +
      'Please validate your data and try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Create CRM Ready sheet
  createCRMSheet(spreadsheet, crmData);
  
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  
  Logger.log('=== CRM READY COMPLETE ===');
  Logger.log(`Total contacts: ${crmData.length}`);
  Logger.log(`Duration: ${duration}s`);
  
  SpreadsheetApp.getUi().alert(
    'CRM Ready Created!',
    `Successfully created CRM Ready sheet with ${crmData.length} contacts in ${duration} seconds.\n\n` +
    `✅ All contacts have:\n` +
    `   - At least one valid email\n` +
    `   - At least one valid mobile phone\n\n` +
    `✅ Data cleaned:\n` +
    `   - Invalid emails removed\n` +
    `   - Invalid/landline phones removed\n` +
    `   - State extracted from phone location\n\n` +
    `Sheet: "${CRM_CONFIG.OUTPUT_SHEET}"`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// COLUMN DETECTION
// =============================================================================

/**
 * Finds all required columns in the source sheet
 */
function findAllColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const columnMap = {
    base: {},
    emails: {},
    phones: {}
  };
  
  // Find base columns
  for (const [key, colName] of Object.entries(CRM_CONFIG.BASE_COLUMNS)) {
    const idx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === colName.toLowerCase()
    );
    if (idx !== -1) {
      columnMap.base[key] = idx;
    }
  }
  
  // Find email columns and their status columns
  const validatedEmails = [];
  const skippedEmails = [];
  
  for (const [key, colName] of Object.entries(CRM_CONFIG.EMAIL_COLUMNS)) {
    const emailIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === colName.toLowerCase()
    );
    const statusIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === (colName + CRM_CONFIG.STATUS_SUFFIX).toLowerCase()
    );
    
    if (emailIdx !== -1 && statusIdx !== -1) {
      columnMap.emails[key] = {
        dataIndex: emailIdx,
        statusIndex: statusIdx
      };
      validatedEmails.push(key);
    } else {
      const reason = emailIdx === -1 ? 'data column missing' : 'status column missing';
      skippedEmails.push(`${key} (${reason})`);
    }
  }
  
  // Find phone columns and their validation columns
  const validatedPhones = [];
  const skippedPhones = [];
  
  for (const [key, colName] of Object.entries(CRM_CONFIG.PHONE_COLUMNS)) {
    const phoneIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === colName.toLowerCase()
    );
    const statusIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === (colName + CRM_CONFIG.STATUS_SUFFIX).toLowerCase()
    );
    const lineTypeIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === (colName + CRM_CONFIG.LINE_TYPE_SUFFIX).toLowerCase()
    );
    const locationIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === (colName + CRM_CONFIG.LOCATION_SUFFIX).toLowerCase()
    );
    
    if (phoneIdx !== -1 && statusIdx !== -1 && lineTypeIdx !== -1) {
      columnMap.phones[key] = {
        dataIndex: phoneIdx,
        statusIndex: statusIdx,
        lineTypeIndex: lineTypeIdx,
        locationIndex: locationIdx !== -1 ? locationIdx : null
      };
      validatedPhones.push(key);
    } else {
      const reasons = [];
      if (phoneIdx === -1) reasons.push('data column missing');
      if (statusIdx === -1) reasons.push('status column missing');
      if (lineTypeIdx === -1) reasons.push('line type column missing');
      skippedPhones.push(`${key} (${reasons.join(', ')})`);
    }
  }
  
  // Check if we have at least one email and one phone
  if (Object.keys(columnMap.emails).length === 0 || 
      Object.keys(columnMap.phones).length === 0) {
    return null;
  }
  
  Logger.log(`Found ${Object.keys(columnMap.base).length} base columns`);
  Logger.log(`\n✅ VALIDATED EMAILS (will be checked for inclusion):`);
  validatedEmails.forEach(email => Logger.log(`  - ${email}`));
  if (skippedEmails.length > 0) {
    Logger.log(`\n❌ SKIPPED EMAILS (no validation columns, will NOT be checked):`);
    skippedEmails.forEach(email => Logger.log(`  - ${email}`));
  }
  Logger.log(`\n✅ VALIDATED PHONES (will be checked for inclusion):`);
  validatedPhones.forEach(phone => Logger.log(`  - ${phone}`));
  if (skippedPhones.length > 0) {
    Logger.log(`\n❌ SKIPPED PHONES (no validation columns, will NOT be checked):`);
    skippedPhones.forEach(phone => Logger.log(`  - ${phone}`));
  }
  Logger.log('');
  
  return columnMap;
}

// =============================================================================
// DATA PROCESSING
// =============================================================================

/**
 * Processes all data and creates CRM-ready rows
 */
function processCRMData(sourceSheet, columnMap) {
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  
  const allData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const crmData = [];
  
  Logger.log(`Processing ${lastRow - CRM_CONFIG.FIRST_DATA_ROW + 1} rows...`);
  
  for (let i = CRM_CONFIG.FIRST_DATA_ROW - 1; i < lastRow; i++) {
    const row = allData[i];
    
    // Check if row meets criteria
    const hasValidEmail = checkEmailCriteriaCRM(row, columnMap.emails);
    const hasValidMobile = checkPhoneCriteriaCRM(row, columnMap.phones);
    
    if (hasValidEmail && hasValidMobile) {
      // Build CRM row
      const crmRow = buildCRMRow(row, columnMap);
      crmData.push(crmRow);
    }
    
    if ((i + 1) % 1000 === 0) {
      Logger.log(`  Processed ${i + 1 - CRM_CONFIG.FIRST_DATA_ROW + 1} rows...`);
    }
  }
  
  Logger.log(`Found ${crmData.length} contacts that meet criteria`);
  
  return crmData;
}

/**
 * Checks if row has at least one valid email
 */
function checkEmailCriteriaCRM(row, emailCols) {
  for (const col of Object.values(emailCols)) {
    const status = String(row[col.statusIndex] || '').toLowerCase().trim();
    if (status === CRM_CONFIG.VALID_EMAIL_STATUS) {
      return true;
    }
  }
  return false;
}

/**
 * Checks if row has at least one valid mobile phone
 */
function checkPhoneCriteriaCRM(row, phoneCols) {
  for (const col of Object.values(phoneCols)) {
    const status = String(row[col.statusIndex] || '').toLowerCase().trim();
    const lineType = String(row[col.lineTypeIndex] || '').toLowerCase().trim();
    if (status === CRM_CONFIG.VALID_PHONE_STATUS && lineType === CRM_CONFIG.VALID_LINE_TYPE) {
      return true;
    }
  }
  return false;
}

/**
 * Builds a CRM-ready row with cleaned data
 */
function buildCRMRow(row, columnMap) {
  const crmRow = {};
  
  // 1. Get base columns (original data)
  for (const [key, idx] of Object.entries(columnMap.base)) {
    crmRow[key] = idx !== undefined ? String(row[idx] || '').trim() : '';
  }
  
  // 2. Initialize ALL email columns to empty FIRST
  //    (This ensures unvalidated emails don't appear in output)
  for (const key of Object.keys(CRM_CONFIG.EMAIL_COLUMNS)) {
    crmRow[key] = '';
  }
  
  // 3. Fill in ONLY emails with status='valid'
  //    (Only validated emails with 'valid' status will appear)
  for (const [key, col] of Object.entries(columnMap.emails)) {
    const status = String(row[col.statusIndex] || '').toLowerCase().trim();
    if (status === CRM_CONFIG.VALID_EMAIL_STATUS) {
      crmRow[key] = String(row[col.dataIndex] || '').trim();
    }
    // else stays empty (already initialized above)
  }
  
  // 4. Initialize ALL phone columns to empty FIRST
  //    (This ensures unvalidated phones don't appear in output)
  for (const key of Object.keys(CRM_CONFIG.PHONE_COLUMNS)) {
    crmRow[key] = '';
  }
  
  // 5. Fill in ONLY valid mobile phones
  //    (Only phones with 'valid confirmed' status AND 'mobile' type will appear)
  let stateExtracted = false;
  for (const [key, col] of Object.entries(columnMap.phones)) {
    const status = String(row[col.statusIndex] || '').toLowerCase().trim();
    const lineType = String(row[col.lineTypeIndex] || '').toLowerCase().trim();
    
    if (status === CRM_CONFIG.VALID_PHONE_STATUS && lineType === CRM_CONFIG.VALID_LINE_TYPE) {
      crmRow[key] = String(row[col.dataIndex] || '').trim();
      
      // Extract state from first valid phone (if not already extracted)
      if (!stateExtracted && col.locationIndex !== null) {
        const location = String(row[col.locationIndex] || '').trim();
        crmRow['state'] = extractState(location);
        stateExtracted = true;
      }
    }
    // else stays empty (already initialized above)
  }
  
  // If no state extracted yet, leave empty
  if (!stateExtracted) {
    crmRow['state'] = '';
  }
  
  return crmRow;
}

/**
 * Extracts state abbreviation from location string
 */
function extractState(location) {
  if (!location) return '';
  
  // State abbreviations map
  const stateAbbreviations = {
    'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR', 'california': 'CA',
    'colorado': 'CO', 'connecticut': 'CT', 'delaware': 'DE', 'florida': 'FL', 'georgia': 'GA',
    'hawaii': 'HI', 'idaho': 'ID', 'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA',
    'kansas': 'KS', 'kentucky': 'KY', 'louisiana': 'LA', 'maine': 'ME', 'maryland': 'MD',
    'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN', 'mississippi': 'MS', 'missouri': 'MO',
    'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV', 'new hampshire': 'NH', 'new jersey': 'NJ',
    'new mexico': 'NM', 'new york': 'NY', 'north carolina': 'NC', 'north dakota': 'ND', 'ohio': 'OH',
    'oklahoma': 'OK', 'oregon': 'OR', 'pennsylvania': 'PA', 'rhode island': 'RI', 'south carolina': 'SC',
    'south dakota': 'SD', 'tennessee': 'TN', 'texas': 'TX', 'utah': 'UT', 'vermont': 'VT',
    'virginia': 'VA', 'washington': 'WA', 'west virginia': 'WV', 'wisconsin': 'WI', 'wyoming': 'WY'
  };
  
  const locationLower = location.toLowerCase();
  
  // Check if already has state abbreviation (e.g., "San Antonio, TX")
  const abbrevMatch = location.match(/\b([A-Z]{2})\b/);
  if (abbrevMatch && Object.values(stateAbbreviations).includes(abbrevMatch[1])) {
    return abbrevMatch[1];
  }
  
  // Check for full state name
  for (const [stateName, abbrev] of Object.entries(stateAbbreviations)) {
    if (locationLower.includes(stateName)) {
      return abbrev;
    }
  }
  
  // If no state found, check for common city patterns
  // "Austin, TX" or "Austin TX"
  const cityStateMatch = location.match(/,\s*([A-Z]{2})\s*$/);
  if (cityStateMatch) {
    return cityStateMatch[1];
  }
  
  return '';
}

// =============================================================================
// SHEET CREATION
// =============================================================================

/**
 * Creates the CRM Ready sheet
 */
function createCRMSheet(spreadsheet, crmData) {
  // Delete existing sheet if it exists
  const existingSheet = spreadsheet.getSheetByName(CRM_CONFIG.OUTPUT_SHEET);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // Create new sheet
  const newSheet = spreadsheet.insertSheet(CRM_CONFIG.OUTPUT_SHEET);
  
  // Set headers (EXACT ORDER from config)
  const headers = CRM_CONFIG.OUTPUT_COLUMNS;
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  newSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // Build data rows in correct order
  const dataRows = crmData.map(crmRow => {
    return headers.map(colName => crmRow[colName] || '');
  });
  
  // Add data
  if (dataRows.length > 0) {
    newSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);
  }
  
  // Format
  newSheet.setFrozenRows(1);
  newSheet.autoResizeColumns(1, headers.length);
  
  Logger.log(`Created CRM Ready sheet with ${dataRows.length} rows`);
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Shows configuration information
 */
function showCRMConfig() {
  const html = HtmlService.createHtmlOutput(`
    <h3>CRM Ready Configuration</h3>
    
    <h4>✅ Criteria:</h4>
    <ul>
      <li><strong>Email:</strong> At least one email with status = "valid"</li>
      <li><strong>Phone:</strong> At least one phone with:
        <ul>
          <li>Status = "valid confirmed"</li>
          <li>Line Type = "mobile"</li>
        </ul>
      </li>
    </ul>
    
    <h4>🧹 Data Cleaning:</h4>
    <ul>
      <li>Invalid emails → Cleared (empty cell)</li>
      <li>Invalid phones → Cleared (empty cell)</li>
      <li>Landline phones → Cleared (empty cell)</li>
      <li>State → Extracted from first valid phone location</li>
    </ul>
    
    <h4>📊 Output Columns (Fixed Order):</h4>
    <ol>
      ${CRM_CONFIG.OUTPUT_COLUMNS.map(col => '<li>' + col + '</li>').join('')}
    </ol>
    
    <h4>📋 Output Sheet:</h4>
    <p><strong>${CRM_CONFIG.OUTPUT_SHEET}</strong> (Blue header)</p>
    
    <br>
    <p><em>Only contacts meeting both criteria are included.</em></p>
  `)
    .setWidth(500)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'CRM Ready Settings');
}

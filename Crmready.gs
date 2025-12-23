/**
 * =============================================================================
 * CRM READY FILTER - V3.0
 * =============================================================================
 * 
 * Creates a single "CRM Ready" sheet with cleaned, validated data
 * 
 * CRITERIA:
 * - At least one valid email (only ONE kept with priority)
 * - At least one valid mobile phone
 * 
 * EMAIL PRIORITY (only keeps ONE in single "email" column):
 * 1. Email 1 (highest priority)
 * 2. Email 2
 * 3. Personal Email
 * 4. Primary Email (lowest priority)
 * 
 * VALIDATION:
 * - Email validation: OPTIONAL (uses _Status columns if they exist)
 * - Phone validation: REQUIRED (_Status and _Line Type columns must exist)
 * - Only phone columns with BOTH validation columns will be included
 * - Example: If you only validated Contact Phone 1 and Contact Mobile Phone,
 *   only those two phone columns will appear in CRM Ready output
 * 
 * DATA CLEANING (when validation exists):
 * - Only keeps ONE valid email in single "email" column (highest priority)
 * - Invalid emails are cleared/empty
 * - Only keeps valid phones with line type mobile
 * - Invalid phones are cleared/empty
 * - Extracts state from first valid phone location
 * 
 * OUTPUT COLUMNS (DYNAMIC):
 * Base columns (always included):
 *   organization, first_name, last_name, email, description, website, city, state
 * 
 * Phone columns (only if they exist in source):
 *   contact_phone1, company_phone1, company_phone2, contact_mobile_phone
 * 
 * Example output (all phones exist):
 *   organization, first_name, last_name, email, contact_phone1, company_phone1,
 *   company_phone2, contact_mobile_phone, description, website, city, state
 * 
 * Example output (only 2 phones exist):
 *   organization, first_name, last_name, email, contact_phone1, contact_mobile_phone,
 *   description, website, city, state
 * 
 * @author: Claude
 * @version: 3.0 - Single email column, dynamic phone columns, underscore naming
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CRM_CONFIG = {
  // Email column names (must match exactly)
  EMAIL_COLUMNS: {
    'Primary Email': 'Primary Email',
    'Email 1': 'Email 1',
    'Email 2': 'Email 2',
    'Personal Email': 'Personal Email'
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
  
  // Valid criteria (case-insensitive matching)
  VALID_EMAIL_STATUS: 'valid',
  VALID_PHONE_STATUS: 'valid_confirmed',  // Matches both "VALID_CONFIRMED" and "valid confirmed"
  VALID_LINE_TYPES: [
    'mobile', 'MOBILE'  // ONLY mobile phones (removed fixed_line, landline)
  ],
  
  // Output sheet name
  OUTPUT_SHEET: 'CRM Ready',
  
  // Base output columns (always included)
  BASE_OUTPUT_COLUMNS: [
    'organization',
    'first_name',
    'last_name',
    'email',  // Single email column (highest priority valid email)
    'description',
    'website',
    'city',
    'state'
  ],
  
  // Phone column mapping (output name → source name)
  // Only phone columns that exist in source will be included in output
  PHONE_OUTPUT_MAP: {
    'contact_phone1': 'Contact Phone 1',
    'company_phone1': 'Company Phone 1',
    'company_phone2': 'Company Phone 2',
    'contact_mobile_phone': 'Contact Mobile Phone'
  },
  
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
      'Please ensure you have these columns in your active sheet:\n' +
      '• At least one email column (Primary Email, Email 1, Email 2, or Personal Email)\n' +
      '• At least one phone column (Contact Phone 1, Company Phone 1, Company Phone 2, or Contact Mobile Phone)\n\n' +
      'Note: Validation status columns are optional. If present, they will be used to filter valid contacts.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log('Found all required columns');
  
  // Process data
  const result = processCRMData(sourceSheet, columnMap);
  const crmData = result.data;
  const stats = result.stats;
  
  if (crmData.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Data Found',
      'No contacts meet the STRICT criteria:\n\n' +
      `Total rows: ${stats.total}\n` +
      `Excluded:\n` +
      `  - Missing valid email: ${stats.excludedNoEmail}\n` +
      `  - Missing mobile phone: ${stats.excludedNoPhone}\n` +
      `  - Missing BOTH: ${stats.excludedBoth}\n\n` +
      'Requirements (BOTH needed):\n' +
      '✓ At least one valid email (Primary Email, Email 1, Email 2, or Personal Email)\n' +
      '✓ At least 1 mobile phone = valid\n\n' +
      'Please validate your data and try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Create CRM Ready sheet
  createCRMSheet(spreadsheet, crmData, columnMap);
  
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  
  Logger.log('=== CRM READY COMPLETE ===');
  Logger.log(`Total contacts: ${crmData.length}`);
  Logger.log(`Duration: ${duration}s`);
  
  SpreadsheetApp.getUi().alert(
    'CRM Ready Created!',
    `✅ Successfully created CRM Ready sheet!\n\n` +
    `📊 Summary:\n` +
    `Total rows: ${stats.total}\n` +
    `Included: ${stats.included} (${(stats.included / stats.total * 100).toFixed(1)}%)\n` +
    `Excluded: ${stats.excluded} (${(stats.excluded / stats.total * 100).toFixed(1)}%)\n\n` +
    `❌ Exclusions:\n` +
    `  - Missing valid email: ${stats.excludedNoEmail}\n` +
    `  - Missing mobile phone: ${stats.excludedNoPhone}\n` +
    `  - Missing BOTH: ${stats.excludedBoth}\n\n` +
    `✅ Included contacts have:\n` +
    `  - Valid email (any email column) ✓\n` +
    `  - At least 1 valid mobile phone ✓\n\n` +
    `Sheet: "${CRM_CONFIG.OUTPUT_SHEET}"\n` +
    `Duration: ${duration}s`,
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
  
  // Find email columns (with OR without status columns)
  for (const [key, colName] of Object.entries(CRM_CONFIG.EMAIL_COLUMNS)) {
    const emailIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === colName.toLowerCase()
    );
    const statusIdx = headers.findIndex(h => 
      String(h).trim().toLowerCase() === (colName + CRM_CONFIG.STATUS_SUFFIX).toLowerCase()
    );
    
    // Include email column even without status (just check if email exists)
    if (emailIdx !== -1) {
      columnMap.emails[key] = {
        dataIndex: emailIdx,
        statusIndex: statusIdx !== -1 ? statusIdx : null  // null if no status column
      };
    }
  }
  
  // Find phone columns (ONLY with validation columns - must have _Status and _Line Type)
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
    
    // ONLY include phone column if it has BOTH status AND line type validation columns
    // This ensures we only include phones that were actually validated
    if (phoneIdx !== -1 && statusIdx !== -1 && lineTypeIdx !== -1) {
      columnMap.phones[key] = {
        dataIndex: phoneIdx,
        statusIndex: statusIdx,
        lineTypeIndex: lineTypeIdx,
        locationIndex: locationIdx !== -1 ? locationIdx : null  // Location is optional
      };
    }
  }
  
  // Check if we have at least one email and one phone column
  if (Object.keys(columnMap.emails).length === 0 || 
      Object.keys(columnMap.phones).length === 0) {
    return null;
  }
  
  Logger.log(`Found ${Object.keys(columnMap.base).length} base columns`);
  Logger.log(`Found ${Object.keys(columnMap.emails).length} email columns`);
  Logger.log(`Found ${Object.keys(columnMap.phones).length} phone columns`);
  
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
  
  // Counters for tracking inclusions/exclusions
  let includedCount = 0;
  let excludedNoEmail = 0;
  let excludedNoPhone = 0;
  let excludedBoth = 0;
  
  Logger.log(`Processing ${lastRow - CRM_CONFIG.FIRST_DATA_ROW + 1} rows...`);
  
  for (let i = CRM_CONFIG.FIRST_DATA_ROW - 1; i < lastRow; i++) {
    const row = allData[i];
    
    // Check if row meets criteria (BOTH required)
    const hasValidEmail = checkEmailCriteriaCRM(row, columnMap.emails);
    const hasValidPhone = checkPhoneCriteriaCRM(row, columnMap.phones);
    
    // BOTH email AND phone required - if either is missing, exclude
    if (hasValidEmail && hasValidPhone) {
      // ✅ Contact meets BOTH criteria - include in CRM Ready
      const crmRow = buildCRMRow(row, columnMap);
      crmData.push(crmRow);
      includedCount++;
    } else {
      // ❌ Contact missing email OR phone - exclude from CRM Ready
      if (!hasValidEmail && !hasValidPhone) {
        excludedBoth++;
        Logger.log(`  Row ${i + 1}: Excluded (no valid email AND no valid mobile phone)`);
      } else if (!hasValidEmail) {
        excludedNoEmail++;
        Logger.log(`  Row ${i + 1}: Excluded (no valid email)`);
      } else if (!hasValidPhone) {
        excludedNoPhone++;
        Logger.log(`  Row ${i + 1}: Excluded (no valid mobile phone)`);
      }
    }
    
    if ((i + 1) % 1000 === 0) {
      Logger.log(`  Processed ${i + 1 - CRM_CONFIG.FIRST_DATA_ROW + 1} rows...`);
    }
  }
  
  // Summary
  const totalProcessed = lastRow - CRM_CONFIG.FIRST_DATA_ROW + 1;
  const totalExcluded = excludedNoEmail + excludedNoPhone + excludedBoth;
  
  Logger.log('\n=== CRM Ready Summary ===');
  Logger.log(`Total rows processed: ${totalProcessed}`);
  Logger.log(`✅ Included (meet BOTH criteria): ${includedCount}`);
  Logger.log(`❌ Excluded (total): ${totalExcluded}`);
  Logger.log(`   - Missing valid email only: ${excludedNoEmail}`);
  Logger.log(`   - Missing mobile phone only: ${excludedNoPhone}`);
  Logger.log(`   - Missing BOTH: ${excludedBoth}`);
  Logger.log(`\nInclusion rate: ${(includedCount / totalProcessed * 100).toFixed(1)}%`);
  
  return {
    data: crmData,
    stats: {
      total: totalProcessed,
      included: includedCount,
      excluded: totalExcluded,
      excludedNoEmail: excludedNoEmail,
      excludedNoPhone: excludedNoPhone,
      excludedBoth: excludedBoth
    }
  };
}

/**
 * Checks if row has at least one valid email
 * If validation status exists, use it. Otherwise, just check if email has a value.
 */
function checkEmailCriteriaCRM(row, emailCols) {
  for (const col of Object.values(emailCols)) {
    // If status column exists, check validation status
    if (col.statusIndex !== null) {
      const status = String(row[col.statusIndex] || '').toLowerCase().trim();
      if (status === CRM_CONFIG.VALID_EMAIL_STATUS) {
        return true;
      }
    } else {
      // No status column - just check if email has a value
      const email = String(row[col.dataIndex] || '').trim();
      if (email && email.length > 0) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Checks if row has at least one valid mobile phone
 * All phones in phoneCols are guaranteed to have validation columns
 */
function checkPhoneCriteriaCRM(row, phoneCols) {
  for (const col of Object.values(phoneCols)) {
    // Check validation status (guaranteed to exist)
    const status = String(row[col.statusIndex] || '').toLowerCase().trim();
    const lineType = String(row[col.lineTypeIndex] || '').toLowerCase().trim();
    
    // Check if status is valid (handle both "valid_confirmed" and "valid confirmed")
    const isValidStatus = (status === 'valid_confirmed' || status === 'valid confirmed');
    
    // Check if line type is valid (case-insensitive)
    const isValidLineType = CRM_CONFIG.VALID_LINE_TYPES.some(type => 
      type.toLowerCase() === lineType
    );
    
    if (isValidStatus && isValidLineType) {
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
  
  // 2. Get valid emails - ONLY keep ONE (highest priority) in single "email" column
  // Priority: Email 1 > Email 2 > Personal Email > Primary Email
  let validEmailFound = false;
  const emailPriority = ['Email 1', 'Email 2', 'Personal Email', 'Primary Email'];
  
  // Initialize email as empty
  crmRow['email'] = '';
  
  // Find first valid email in priority order
  for (const priorityKey of emailPriority) {
    if (validEmailFound) break;  // Already found valid email, stop
    
    const col = columnMap.emails[priorityKey];
    if (!col) continue;  // Column doesn't exist in source
    
    // Check if email is valid
    let isValid = false;
    let emailValue = '';
    
    if (col.statusIndex !== null) {
      // Status column exists - check validation status
      const status = String(row[col.statusIndex] || '').toLowerCase().trim();
      if (status === CRM_CONFIG.VALID_EMAIL_STATUS) {
        isValid = true;
        emailValue = String(row[col.dataIndex] || '').trim();
      }
    } else {
      // No status column - check if email has a value
      emailValue = String(row[col.dataIndex] || '').trim();
      if (emailValue && emailValue.length > 0) {
        isValid = true;
      }
    }
    
    // If valid email found, put in single "email" column
    if (isValid && emailValue) {
      crmRow['email'] = emailValue;
      validEmailFound = true;
    }
  }
  
  // 3. Get valid phones only (mobile phones) - use output column names (underscore format)
  // Note: All phones in columnMap.phones are guaranteed to have validation columns
  let stateExtracted = false;
  for (const [sourceKey, col] of Object.entries(columnMap.phones)) {
    // Map source column name to output column name (with underscores)
    const outputKey = Object.keys(CRM_CONFIG.PHONE_OUTPUT_MAP).find(
      k => CRM_CONFIG.PHONE_OUTPUT_MAP[k] === sourceKey
    );
    
    if (!outputKey) continue;  // Column not in output map
    
    // Check validation status (guaranteed to exist)
    const status = String(row[col.statusIndex] || '').toLowerCase().trim();
    const lineType = String(row[col.lineTypeIndex] || '').toLowerCase().trim();
    
    // Check if status is valid (handle both formats)
    const isValidStatus = (status === 'valid_confirmed' || status === 'valid confirmed');
    
    // Check if line type is valid (case-insensitive)
    const isValidLineType = CRM_CONFIG.VALID_LINE_TYPES.some(type => 
      type.toLowerCase() === lineType
    );
    
    if (isValidStatus && isValidLineType) {
      crmRow[outputKey] = String(row[col.dataIndex] || '').trim();
      
      // Extract state from first valid phone (if not already extracted)
      if (!stateExtracted && col.locationIndex !== null) {
        const location = String(row[col.locationIndex] || '').trim();
        crmRow['state'] = extractState(location);
        stateExtracted = true;
      }
    } else {
      crmRow[outputKey] = ''; // Clear invalid phone
    }
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
function createCRMSheet(spreadsheet, crmData, columnMap) {
  // Delete existing sheet if it exists
  const existingSheet = spreadsheet.getSheetByName(CRM_CONFIG.OUTPUT_SHEET);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
  }
  
  // Create new sheet
  const newSheet = spreadsheet.insertSheet(CRM_CONFIG.OUTPUT_SHEET);
  
  // Build dynamic headers based on what phone columns exist in source
  const headers = [...CRM_CONFIG.BASE_OUTPUT_COLUMNS];
  
  // Insert phone columns in specific order (only if they exist in source)
  const phoneHeaders = [];
  for (const [outputName, sourceName] of Object.entries(CRM_CONFIG.PHONE_OUTPUT_MAP)) {
    if (columnMap.phones[sourceName]) {
      phoneHeaders.push(outputName);
    }
  }
  
  // Insert phone columns after 'email' (index 3)
  headers.splice(4, 0, ...phoneHeaders);
  
  Logger.log(`Output columns: ${headers.join(', ')}`);
  
  // Set headers
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
  
  Logger.log(`Created CRM Ready sheet with ${dataRows.length} rows and ${headers.length} columns`);
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
    
    <h4>✅ Criteria (STRICT):</h4>
    <ul>
      <li><strong>Email:</strong> Primary Email must have status = "valid"</li>
      <li><strong>Phone:</strong> At least one phone with:
        <ul>
          <li>Status = "valid confirmed" (or "VALID_CONFIRMED")</li>
          <li>Line Type = "mobile" (or "MOBILE") ONLY</li>
        </ul>
      </li>
    </ul>
    
    <p><strong>⚠️ Note:</strong> Fixed line/landline phones are NOT included (mobile only).</p>
    
    <h4>🧹 Data Cleaning:</h4>
    <ul>
      <li>Invalid emails → Cleared (empty cell)</li>
      <li>Other emails (Personal, Email 1, Email 2) → Not included</li>
      <li>Invalid phones → Cleared (empty cell)</li>
      <li>Valid mobile phones → Kept ✓</li>
      <li>Fixed line phones → Cleared (mobile only) ❌</li>
      <li>State → Extracted from first valid mobile phone location</li>
    </ul>
    
    <h4>📊 Output Columns (12 Total):</h4>
    <ol>
      ${CRM_CONFIG.OUTPUT_COLUMNS.map(col => '<li>' + col + '</li>').join('')}
    </ol>
    
    <h4>📋 Output Sheet:</h4>
    <p><strong>${CRM_CONFIG.OUTPUT_SHEET}</strong> (Blue header)</p>
    
    <br>
    <p><em>Only contacts with valid Primary Email AND at least one valid mobile phone are included.</em></p>
  `)
    .setWidth(550)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'CRM Ready Settings (STRICT)');
}

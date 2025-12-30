/**
 * ========================================================================
 * CRM CALLERS READY GENERATOR
 * ========================================================================
 * 
 * Creates "CRM Callers Ready" sheet from "CRM Ready" with:
 * - 5 columns: organization, first_name, last_name, email, contact_mobile_phone
 * - Phone Priority: contact_mobile_phone > contact_phone1 > company_phone1 > company_phone2
 * - Only includes contacts with at least one phone number
 * 
 * Phone Selection Logic:
 * 1. If contact_mobile_phone exists → Use it
 * 2. Else if contact_phone1 exists → Use it
 * 3. Else if company_phone1 exists → Use it
 * 4. Else if company_phone2 exists → Use it
 * 5. Else → Skip contact (no phone number)
 * 
 * MENU: Added to DataCleaningPipeline_V3.gs onOpen()
 *       🔧 Data Tools → 🎯 CRM Ready → 📞 Generate CRM Callers Ready
 * 
 * ========================================================================
 */

// Configuration
const CALLERS_CONFIG = {
  SOURCE_SHEET: 'CRM Ready',
  OUTPUT_SHEET: 'CRM Callers Ready',
  
  // Source columns from CRM Ready
  SOURCE_COLUMNS: {
    ORGANIZATION: 'organization',
    FIRST_NAME: 'first_name',
    LAST_NAME: 'last_name',
    EMAIL: 'email',
    CONTACT_PHONE1: 'contact_phone1',
    COMPANY_PHONE1: 'company_phone1',
    COMPANY_PHONE2: 'company_phone2',
    CONTACT_MOBILE: 'contact_mobile_phone'
  },
  
  // Output columns (lowercase with underscores)
  OUTPUT_COLUMNS: [
    'organization',
    'first_name',
    'last_name',
    'email',
    'contact_mobile_phone'
  ],
  
  // Phone priority (first to last)
  PHONE_PRIORITY: [
    'contact_mobile_phone',
    'contact_phone1',
    'company_phone1',
    'company_phone2'
  ]
};

/**
 * Main function - Generate CRM Callers Ready sheet
 */
function generateCRMCallersReady() {
  try {
    const startTime = new Date();
    Logger.log('=== Starting CRM Callers Ready Generation ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Step 1: Get source sheet
    const sourceSheet = ss.getSheetByName(CALLERS_CONFIG.SOURCE_SHEET);
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert(
        'Error',
        `Source sheet "${CALLERS_CONFIG.SOURCE_SHEET}" not found!\n\n` +
        'Please make sure the CRM Ready sheet exists.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Step 2: Load and parse data
    Logger.log('Step 1: Loading data from CRM Ready...');
    const sourceData = loadSourceData(sourceSheet);
    Logger.log(`Loaded ${sourceData.length} contacts from CRM Ready`);
    
    // Step 3: Process contacts - extract phone with priority
    Logger.log('Step 2: Processing contacts and selecting phones...');
    const processedData = processContactsWithPhonePriority(sourceData);
    Logger.log(`Processed ${processedData.length} contacts with phones`);
    
    // Step 4: Create or clear output sheet
    Logger.log('Step 3: Creating output sheet...');
    const outputSheet = getOrCreateOutputSheet(ss);
    
    // Step 5: Write data to output
    Logger.log('Step 4: Writing data to CRM Callers Ready...');
    writeCallersData(outputSheet, processedData);
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== CRM Callers Ready Complete in ${duration}s ===`);
    Logger.log(`Total contacts processed: ${sourceData.length}`);
    Logger.log(`Contacts with phones: ${processedData.length}`);
    Logger.log(`Contacts without phones: ${sourceData.length - processedData.length}`);
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      'Success!',
      `✅ CRM Callers Ready generated!\n\n` +
      `📊 Results:\n` +
      `Source contacts: ${sourceData.length}\n` +
      `Contacts with phones: ${processedData.length}\n` +
      `Contacts without phones: ${sourceData.length - processedData.length}\n` +
      `Duration: ${duration.toFixed(2)}s\n\n` +
      `Output: "${CALLERS_CONFIG.OUTPUT_SHEET}" sheet\n\n` +
      `📞 Phone Priority:\n` +
      `1. Contact Mobile Phone\n` +
      `2. Contact Phone 1\n` +
      `3. Company Phone 1\n` +
      `4. Company Phone 2`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to generate CRM Callers Ready: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Load data from source sheet (CRM Ready)
 */
function loadSourceData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error('CRM Ready sheet is empty (no data rows)');
  }
  
  // Get headers
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // Map headers to column indices (case-insensitive)
  const columnMap = {};
  headers.forEach((header, index) => {
    const normalizedHeader = header.toString().trim().toLowerCase();
    columnMap[normalizedHeader] = index;
  });
  
  // Validate required columns exist
  const requiredColumns = [
    'organization',
    'first_name',
    'last_name',
    'email'
  ];
  
  const missingColumns = requiredColumns.filter(col => columnMap[col] === undefined);
  
  if (missingColumns.length > 0) {
    throw new Error(`Missing required columns in CRM Ready: ${missingColumns.join(', ')}`);
  }
  
  // Check if at least one phone column exists
  const phoneColumns = CALLERS_CONFIG.PHONE_PRIORITY.filter(col => columnMap[col] !== undefined);
  
  if (phoneColumns.length === 0) {
    throw new Error('No phone columns found in CRM Ready sheet');
  }
  
  // Get all data
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
  // Convert to objects
  const contacts = data.map(row => {
    const contact = {};
    
    // Get all columns
    Object.keys(columnMap).forEach(colName => {
      const colIndex = columnMap[colName];
      contact[colName] = row[colIndex];
    });
    
    return contact;
  });
  
  return contacts;
}

/**
 * Process contacts and select phone number with priority
 */
function processContactsWithPhonePriority(contacts) {
  const processedContacts = [];
  
  contacts.forEach(contact => {
    // Try to get phone number with priority
    let selectedPhone = '';
    
    for (const phoneColumn of CALLERS_CONFIG.PHONE_PRIORITY) {
      const phone = contact[phoneColumn];
      if (phone && phone.toString().trim()) {
        selectedPhone = phone.toString().trim();
        break; // Found first available phone
      }
    }
    
    // Only include contacts with at least one phone
    if (selectedPhone) {
      processedContacts.push({
        organization: contact.organization || '',
        first_name: contact.first_name || '',
        last_name: contact.last_name || '',
        email: contact.email || '',
        contact_mobile_phone: selectedPhone
      });
    }
  });
  
  return processedContacts;
}

/**
 * Get or create output sheet
 */
function getOrCreateOutputSheet(ss) {
  let sheet = ss.getSheetByName(CALLERS_CONFIG.OUTPUT_SHEET);
  
  if (sheet) {
    // Clear existing sheet
    sheet.clear();
  } else {
    // Create new sheet
    sheet = ss.insertSheet(CALLERS_CONFIG.OUTPUT_SHEET);
  }
  
  return sheet;
}

/**
 * Write data to output sheet
 */
function writeCallersData(sheet, contacts) {
  // Write headers
  const headers = CALLERS_CONFIG.OUTPUT_COLUMNS;
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#34a853')  // Green background
    .setFontColor('#ffffff');
  
  // Write data
  if (contacts.length > 0) {
    const outputData = contacts.map(contact => [
      contact.organization,
      contact.first_name,
      contact.last_name,
      contact.email,
      contact.contact_mobile_phone
    ]);
    
    sheet.getRange(2, 1, outputData.length, headers.length)
      .setValues(outputData);
  }
  
  // Format sheet
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 200);  // organization
  sheet.setColumnWidth(2, 120);  // first_name
  sheet.setColumnWidth(3, 120);  // last_name
  sheet.setColumnWidth(4, 200);  // email
  sheet.setColumnWidth(5, 150);  // contact_mobile_phone
  
  Logger.log(`✓ Data written: ${contacts.length} contacts`);
}

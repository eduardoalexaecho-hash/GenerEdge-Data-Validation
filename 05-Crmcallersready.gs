/**
 * ========================================================================
 * CRM CALLERS READY GENERATOR - V4.2 (FLEXIBLE COLUMNS)
 * ========================================================================
 * 
 * Creates a simplified calling list from CRM Ready with all phone columns
 * 
 * V4.2 CHANGES:
 * - Flexible column matching (case-insensitive, handles spaces)
 * - Skips missing columns instead of throwing error
 * - Works with "Contact Phone 1" or "contact_phone1" format
 * 
 * V4.1 CHANGES:
 * - Removed borders from output sheet (cleaner look)
 * 
 * V4.0 LOGIC:
 * - Simply copies contacts from CRM Ready
 * - Includes ALL phone columns (contact_phone1, company_phone1, company_phone2, contact_mobile_phone)
 * - Excludes: description, website, city, state
 * - No additional filtering (CRM Ready already has valid mobiles only)
 * 
 * Output Columns (8 total, or fewer if some don't exist):
 * 1. organization
 * 2. first_name
 * 3. last_name
 * 4. email
 * 5. contact_phone1 (skipped if not found)
 * 6. company_phone1 (skipped if not found)
 * 7. company_phone2 (skipped if not found)
 * 8. contact_mobile_phone (skipped if not found)
 * 
 * Purpose:
 * - Cleaner view for calling campaigns
 * - All phone numbers in separate columns (try each one)
 * - Remove unnecessary columns (description, website, city, state)
 * - No borders for clean appearance
 * - Flexible with column names
 * 
 * MENU: 🔧 Data Tools → 🎯 CRM Ready → 📞 Generate CRM Callers Ready
 * 
 * @version 4.2
 * ========================================================================
 */

// Configuration
const CALLERS_CONFIG = {
  SOURCE_SHEET: 'CRM Ready',
  OUTPUT_SHEET: 'CRM Callers Ready',
  
  // Output columns to include (in this order)
  OUTPUT_COLUMNS: [
    'organization',
    'first_name',
    'last_name',
    'email',
    'contact_phone1',
    'company_phone1',
    'company_phone2',
    'contact_mobile_phone'
  ]
};

/**
 * Main function - Generates CRM Callers Ready sheet
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
    
    Logger.log(`Source: ${CALLERS_CONFIG.SOURCE_SHEET}`);
    
    // Step 2: Get data
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length <= 1) {
      SpreadsheetApp.getUi().alert(
        'Error',
        'Source sheet is empty or has no data rows.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const totalRows = sourceData.length - 1; // Exclude header
    Logger.log(`Found ${totalRows} contacts in CRM Ready`);
    
    // Step 3: Create output sheet
    const outputSheet = getOrCreateOutputSheet(ss);
    
    // Step 4: Copy selected columns
    const includedColumns = copySelectedColumns(sourceSheet, outputSheet, sourceData);
    
    // Step 5: Format sheet
    formatOutputSheet(outputSheet);
    
    const endTime = new Date();
    const duration = ((endTime - startTime) / 1000).toFixed(1);
    
    Logger.log(`=== CRM Callers Ready Generation Complete ===`);
    Logger.log(`Duration: ${duration} seconds`);
    
    // Build phone columns message
    const phoneColumns = includedColumns.filter(col => 
      col.toLowerCase().includes('phone') || col.toLowerCase().includes('mobile')
    );
    const phoneColumnsMsg = phoneColumns.length > 0 
      ? `Phone columns included:\n${phoneColumns.map(c => '- ' + c).join('\n')}`
      : 'No phone columns found';
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      '✅ Success!',
      `CRM Callers Ready sheet created!\n\n` +
      `Total contacts: ${totalRows}\n` +
      `Columns included: ${includedColumns.length}\n` +
      `Processing time: ${duration} seconds\n\n` +
      phoneColumnsMsg,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Switch to output sheet
    ss.setActiveSheet(outputSheet);
    
  } catch (error) {
    Logger.log('ERROR in generateCRMCallersReady: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to generate CRM Callers Ready:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Get or create output sheet
 */
function getOrCreateOutputSheet(ss) {
  let sheet = ss.getSheetByName(CALLERS_CONFIG.OUTPUT_SHEET);
  
  if (sheet) {
    // Clear existing sheet
    sheet.clear();
    Logger.log('Cleared existing output sheet');
  } else {
    // Create new sheet
    sheet = ss.insertSheet(CALLERS_CONFIG.OUTPUT_SHEET);
    Logger.log('Created new output sheet');
  }
  
  return sheet;
}

/**
 * Copy selected columns from source to output
 * Returns array of included column names
 */
function copySelectedColumns(sourceSheet, outputSheet, sourceData) {
  const sourceHeaders = sourceData[0].map(h => h.toString().trim());
  
  // Create case-insensitive header map
  const headerMap = {};
  sourceHeaders.forEach((header, index) => {
    const normalizedHeader = header.toLowerCase().replace(/\s+/g, '_');
    headerMap[normalizedHeader] = index;
  });
  
  // Find which output columns actually exist in source
  const availableColumns = [];
  const availableHeaders = [];
  
  CALLERS_CONFIG.OUTPUT_COLUMNS.forEach(col => {
    const normalized = col.toLowerCase().replace(/\s+/g, '_');
    
    if (headerMap[normalized] !== undefined) {
      availableColumns.push(col);
      availableHeaders.push(col);
      Logger.log(`Found column: ${col}`);
    } else {
      Logger.log(`Column not found, skipping: ${col}`);
    }
  });
  
  if (availableColumns.length === 0) {
    throw new Error('No matching columns found in CRM Ready sheet!');
  }
  
  Logger.log(`Using ${availableColumns.length} out of ${CALLERS_CONFIG.OUTPUT_COLUMNS.length} columns`);
  
  // Write output headers (only available columns)
  outputSheet.getRange(1, 1, 1, availableHeaders.length)
    .setValues([availableHeaders]);
  Logger.log(`Wrote ${availableHeaders.length} column headers`);
  
  // Copy data rows
  const outputData = [];
  for (let i = 1; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];
    const outputRow = [];
    
    // Extract each available column
    availableColumns.forEach(col => {
      const normalized = col.toLowerCase().replace(/\s+/g, '_');
      const colIndex = headerMap[normalized];
      outputRow.push(sourceRow[colIndex]);
    });
    
    outputData.push(outputRow);
  }
  
  // Write data to sheet
  if (outputData.length > 0) {
    outputSheet.getRange(2, 1, outputData.length, availableHeaders.length)
      .setValues(outputData);
    Logger.log(`Copied ${outputData.length} rows with ${availableHeaders.length} columns`);
  }
  
  return availableColumns;
}

/**
 * Format output sheet
 */
function formatOutputSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0) return;
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Auto-resize columns
  for (let col = 1; col <= lastCol; col++) {
    sheet.autoResizeColumn(col);
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log('Applied formatting to output sheet (no borders)');
}

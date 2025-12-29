/**
 * =============================================================================
 * CRM READY COMPARISON TOOL - ANY EMAIL/PHONE VERSION
 * =============================================================================
 * 
 * Compares "ZB & BP Validation" sheet with "CRM Ready" sheet to show which 
 * contacts were excluded and the specific reasons why they didn't meet the criteria.
 * 
 * UPDATED LOGIC:
 * ==============
 * - Checks ALL email columns (Primary Email, Email 1, Email 2, Personal Email)
 * - Checks ALL phone columns (Contact Phone 1, Company Phone 1, Company Phone 2, Contact Mobile Phone)
 * - Contact is valid if ANY email column is valid AND ANY phone column is valid
 * - Phone doesn't need to be mobile - ANY valid phone type works
 * 
 * IMPORTANT - Column Name Differences:
 * ===================================
 * CRM Ready uses lowercase column names:
 *   - organization, first_name, last_name, email
 *   - contact_phone1, company_phone1, company_phone2, contact_mobile_phone
 * 
 * ZB & BP Validation uses Title Case with spaces:
 *   - Organization, First Name, Last Name, Primary Email, Email 1, Email 2, Personal Email
 *   - Contact Phone 1, Company Phone 1, Company Phone 2, Contact Mobile Phone
 * 
 * The comparison matches:
 *   - CRM Ready "email" ↔ ZB & BP Validation ANY email column (first found)
 *   - Composite key: organization|first_name|last_name|email (case-insensitive)
 * 
 * Note: State is NOT compared (doesn't exist in ZB & BP Validation)
 * 
 * Creates "Excluded Contacts" sheet with:
 * - All original data from ZB & BP Validation
 * - Exclusion reason
 * - ALL email validation statuses
 * - ALL phone validation statuses
 * 
 * @author: Claude
 * @version: 3.0 - ANY EMAIL/PHONE
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const COMPARISON_CONFIG = {
  // Sheet names
  SOURCE_SHEET: 'ZB & BP Validation',  // Source sheet with validation data
  CRM_READY_SHEET: 'CRM Ready',
  EXCLUDED_SHEET: 'Excluded Contacts',
  
  // Column names (for matching and display)
  ORGANIZATION_COL: 'Organization',
  FIRST_NAME_COL: 'First Name',
  LAST_NAME_COL: 'Last Name',
  
  // Email columns - NOTE: Different names in different sheets!
  // CRM Ready: "email" 
  // ZB & BP Validation: "Primary Email"
  PRIMARY_EMAIL_COL_VALIDATION: 'Primary Email',  // In ZB & BP Validation
  EMAIL_COL_CRM_READY: 'email',                   // In CRM Ready (lowercase)
  PERSONAL_EMAIL_COL: 'Personal Email',
  EMAIL_1_COL: 'Email 1',
  EMAIL_2_COL: 'Email 2',
  
  // Phone columns - NOTE: Different names in different sheets!
  // CRM Ready: lowercase with underscores (contact_phone1, company_phone1, etc.)
  // ZB & BP Validation: Title Case with spaces (Contact Phone 1, Company Phone 1, etc.)
  CONTACT_PHONE_1_COL: 'Contact Phone 1',      // In ZB & BP Validation
  COMPANY_PHONE_1_COL: 'Company Phone 1',      // In ZB & BP Validation
  COMPANY_PHONE_2_COL: 'Company Phone 2',      // In ZB & BP Validation
  CONTACT_MOBILE_COL: 'Contact Mobile Phone',  // In ZB & BP Validation
  
  // CRM Ready phone columns (lowercase with underscores)
  CONTACT_PHONE_1_CRM: 'contact_phone1',
  COMPANY_PHONE_1_CRM: 'company_phone1',
  COMPANY_PHONE_2_CRM: 'company_phone2',
  CONTACT_MOBILE_CRM: 'contact_mobile_phone',
  
  // Validation suffixes
  STATUS_SUFFIX: '_Status',
  LINE_TYPE_SUFFIX: '_Line Type',
  
  // Header colors
  EXCLUDED_HEADER_COLOR: '#ea4335',  // Red
  REASON_HEADER_COLOR: '#fbbc04'     // Yellow
};

// =============================================================================
// MAIN FUNCTION
// =============================================================================

/**
 * Creates comparison sheet showing which contacts were excluded from CRM Ready
 */
function compareWithCRMReady() {
  const startTime = new Date();
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log('=== STARTING CRM READY COMPARISON ===');
    
    // Get ZB & BP Validation sheet (source with validation data)
    const sourceSheet = spreadsheet.getSheetByName(COMPARISON_CONFIG.SOURCE_SHEET);
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert(
        'Source Sheet Not Found',
        `Cannot find "${COMPARISON_CONFIG.SOURCE_SHEET}" sheet.\n\n` +
        'This comparison requires the validated data sheet.\n' +
        'Please make sure the sheet name is exactly: "ZB & BP Validation"',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    Logger.log(`Source sheet: ${sourceSheet.getName()}`);
    
    // Check if CRM Ready sheet exists
    const crmReadySheet = spreadsheet.getSheetByName(COMPARISON_CONFIG.CRM_READY_SHEET);
    if (!crmReadySheet) {
      SpreadsheetApp.getUi().alert(
        'CRM Ready Not Found',
        `Cannot find "${COMPARISON_CONFIG.CRM_READY_SHEET}" sheet.\n\n` +
        'Please create CRM Ready sheet first using:\n' +
        '🎯 CRM Ready → ▶️ Create CRM Ready Sheet',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Get included contacts from CRM Ready
    Logger.log('Reading CRM Ready sheet...');
    const includedContacts = getCRMReadyContacts(crmReadySheet);
    Logger.log(`Found ${includedContacts.size} contacts in CRM Ready`);
    
    // Get all contacts from source sheet and compare
    Logger.log('Comparing with source sheet...');
    const excludedData = compareSheets(sourceSheet, includedContacts);
    
    if (excludedData.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Exclusions',
        'All contacts from the source sheet are in CRM Ready!\n\n' +
        'There are no excluded contacts to show.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Create excluded contacts sheet
    Logger.log(`Creating excluded contacts sheet with ${excludedData.length} contacts...`);
    createExcludedSheet(spreadsheet, sourceSheet, excludedData);
    
    const endTime = new Date();
    const duration = Math.round((endTime - startTime) / 1000);
    
    Logger.log('=== COMPARISON COMPLETE ===');
    Logger.log(`Duration: ${duration}s`);
    
    // Show summary
    SpreadsheetApp.getUi().alert(
      'Comparison Complete!',
      `✅ Created "${COMPARISON_CONFIG.EXCLUDED_SHEET}" sheet!\n\n` +
      `📊 Summary:\n` +
      `Source: "${COMPARISON_CONFIG.SOURCE_SHEET}"\n` +
      `Total excluded: ${excludedData.length}\n` +
      `Included in CRM Ready: ${includedContacts.size}\n\n` +
      `The excluded sheet shows:\n` +
      `  • All original contact data from validation sheet\n` +
      `  • Exclusion reason (why not in CRM Ready)\n` +
      `  • Email validation status\n` +
      `  • Phone validation statuses\n\n` +
      `Duration: ${duration}s`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to create comparison: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

// =============================================================================
// COMPARISON FUNCTIONS
// =============================================================================

/**
 * Gets set of contact identifiers from CRM Ready sheet
 */
function getCRMReadyContacts(crmReadySheet) {
  const data = crmReadySheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find key columns in CRM Ready (lowercase column names)
  const orgIndex = headers.indexOf('organization');
  const firstNameIndex = headers.indexOf('first_name');
  const lastNameIndex = headers.indexOf('last_name');
  const emailIndex = headers.indexOf('email');  // CRM Ready uses "email" not "Primary Email"
  
  if (orgIndex === -1 || firstNameIndex === -1 || lastNameIndex === -1 || emailIndex === -1) {
    throw new Error('Could not find required columns in CRM Ready sheet. Expected: organization, first_name, last_name, email');
  }
  
  const includedContacts = new Set();
  
  // Create unique identifier for each contact (use multiple fields for accuracy)
  for (let i = 1; i < data.length; i++) {
    const org = String(data[i][orgIndex] || '').trim().toLowerCase();
    const firstName = String(data[i][firstNameIndex] || '').trim().toLowerCase();
    const lastName = String(data[i][lastNameIndex] || '').trim().toLowerCase();
    const email = String(data[i][emailIndex] || '').trim().toLowerCase();
    
    // Create composite key
    const key = `${org}|${firstName}|${lastName}|${email}`;
    includedContacts.add(key);
  }
  
  return includedContacts;
}

/**
 * Compares source sheet with CRM Ready and finds excluded contacts
 */
function compareSheets(sourceSheet, includedContacts) {
  const data = sourceSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Find key columns in ZB & BP Validation (Title Case)
  const orgIndex = headers.indexOf(COMPARISON_CONFIG.ORGANIZATION_COL);
  const firstNameIndex = headers.indexOf(COMPARISON_CONFIG.FIRST_NAME_COL);
  const lastNameIndex = headers.indexOf(COMPARISON_CONFIG.LAST_NAME_COL);
  
  // Find ALL email columns and their status columns
  // Check: Primary Email, Email 1, Email 2, Personal Email
  const emailColumns = [
    {
      name: COMPARISON_CONFIG.PRIMARY_EMAIL_COL_VALIDATION,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.PRIMARY_EMAIL_COL_VALIDATION),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.PRIMARY_EMAIL_COL_VALIDATION + COMPARISON_CONFIG.STATUS_SUFFIX)
    },
    {
      name: COMPARISON_CONFIG.EMAIL_1_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.EMAIL_1_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.EMAIL_1_COL + COMPARISON_CONFIG.STATUS_SUFFIX)
    },
    {
      name: COMPARISON_CONFIG.EMAIL_2_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.EMAIL_2_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.EMAIL_2_COL + COMPARISON_CONFIG.STATUS_SUFFIX)
    },
    {
      name: COMPARISON_CONFIG.PERSONAL_EMAIL_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.PERSONAL_EMAIL_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.PERSONAL_EMAIL_COL + COMPARISON_CONFIG.STATUS_SUFFIX)
    }
  ].filter(col => col.dataIndex !== -1); // Only keep columns that exist
  
  // Find ALL phone columns and their validation columns
  // Check ANY phone type (not just mobile)
  const phoneColumns = [
    {
      name: COMPARISON_CONFIG.CONTACT_PHONE_1_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.CONTACT_PHONE_1_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.CONTACT_PHONE_1_COL + COMPARISON_CONFIG.STATUS_SUFFIX),
      lineTypeIndex: headers.indexOf(COMPARISON_CONFIG.CONTACT_PHONE_1_COL + COMPARISON_CONFIG.LINE_TYPE_SUFFIX)
    },
    {
      name: COMPARISON_CONFIG.COMPANY_PHONE_1_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.COMPANY_PHONE_1_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.COMPANY_PHONE_1_COL + COMPARISON_CONFIG.STATUS_SUFFIX),
      lineTypeIndex: headers.indexOf(COMPARISON_CONFIG.COMPANY_PHONE_1_COL + COMPARISON_CONFIG.LINE_TYPE_SUFFIX)
    },
    {
      name: COMPARISON_CONFIG.COMPANY_PHONE_2_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.COMPANY_PHONE_2_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.COMPANY_PHONE_2_COL + COMPARISON_CONFIG.STATUS_SUFFIX),
      lineTypeIndex: headers.indexOf(COMPARISON_CONFIG.COMPANY_PHONE_2_COL + COMPARISON_CONFIG.LINE_TYPE_SUFFIX)
    },
    {
      name: COMPARISON_CONFIG.CONTACT_MOBILE_COL,
      dataIndex: headers.indexOf(COMPARISON_CONFIG.CONTACT_MOBILE_COL),
      statusIndex: headers.indexOf(COMPARISON_CONFIG.CONTACT_MOBILE_COL + COMPARISON_CONFIG.STATUS_SUFFIX),
      lineTypeIndex: headers.indexOf(COMPARISON_CONFIG.CONTACT_MOBILE_COL + COMPARISON_CONFIG.LINE_TYPE_SUFFIX)
    }
  ].filter(col => col.dataIndex !== -1); // Only keep columns that exist
  
  
  const excludedData = [];
  
  // Check each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Create composite key (use first email found for matching)
    const org = String(row[orgIndex] || '').trim().toLowerCase();
    const firstName = String(row[firstNameIndex] || '').trim().toLowerCase();
    const lastName = String(row[lastNameIndex] || '').trim().toLowerCase();
    
    // Find first non-empty email for key (priority: Primary, Email 1, Email 2, Personal)
    let emailForKey = '';
    for (const emailCol of emailColumns) {
      const emailValue = String(row[emailCol.dataIndex] || '').trim().toLowerCase();
      if (emailValue) {
        emailForKey = emailValue;
        break;
      }
    }
    
    const key = `${org}|${firstName}|${lastName}|${emailForKey}`;
    
    // If not in CRM Ready, analyze why
    if (!includedContacts.has(key)) {
      // Check ALL email columns for any valid email
      let hasValidEmail = false;
      const emailDetails = [];
      
      for (const emailCol of emailColumns) {
        const emailValue = row[emailCol.dataIndex];
        const status = String(row[emailCol.statusIndex] || '').toLowerCase().trim();
        
        // Check if this email is valid
        if (status === 'valid') {
          hasValidEmail = true;
        }
        
        // Store email details for display
        if (emailValue && emailValue.toString().trim() !== '') {
          emailDetails.push({
            name: emailCol.name,
            email: emailValue,
            status: status || 'not validated'
          });
        }
      }
      
      // Check ALL phone columns for any valid phone (not just mobile)
      let hasValidPhone = false;
      const phoneDetails = [];
      
      for (const phoneCol of phoneColumns) {
        const phoneNumber = row[phoneCol.dataIndex];
        const status = String(row[phoneCol.statusIndex] || '').toLowerCase().trim();
        const lineType = String(row[phoneCol.lineTypeIndex] || '').toLowerCase().trim();
        
        // Check if this is ANY valid phone (not just mobile)
        const isValidPhone = (status === 'valid_confirmed' || status === 'valid confirmed');
        
        if (isValidPhone) {
          hasValidPhone = true;
        }
        
        // Store phone details for display
        if (phoneNumber && phoneNumber.toString().trim() !== '') {
          phoneDetails.push({
            name: phoneCol.name,
            number: phoneNumber,
            status: status || 'not validated',
            lineType: lineType || 'unknown'
          });
        }
      }
      
      // CRITICAL FIX: If contact has BOTH valid email AND valid phone, 
      // they should be in CRM Ready. Don't exclude them - this is likely 
      // a key mismatch issue (different email was used in key creation)
      if (hasValidEmail && hasValidPhone) {
        Logger.log(`WARNING: Row ${i + 1} has valid email AND phone but not in CRM Ready - possible key mismatch. Skipping exclusion.`);
        continue;  // Skip this contact - don't add to excluded list
      }
      
      // Determine exclusion reason (only for contacts missing email OR phone)
      let reason = '';
      if (!hasValidEmail && !hasValidPhone) {
        reason = 'No valid email AND no valid phone';
      } else if (!hasValidEmail) {
        reason = 'No valid email in any column';
      } else if (!hasValidPhone) {
        reason = 'No valid phone in any column';
      }
      
      // Get email status summary for display
      let emailStatusSummary = '';
      if (emailDetails.length > 0) {
        emailStatusSummary = emailDetails.map(e => 
          `${e.name}: ${e.status}`
        ).join('; ');
      } else {
        emailStatusSummary = 'No emails';
      }
      
      excludedData.push({
        rowNumber: i + 1,
        rowData: row,
        reason: reason,
        emailStatus: emailStatusSummary,
        emailDetails: emailDetails,
        hasValidEmail: hasValidEmail,
        hasValidPhone: hasValidPhone,
        phoneDetails: phoneDetails
      });
    }
  }
  
  return excludedData;
}

// =============================================================================
// SHEET CREATION
// =============================================================================

/**
 * Creates the excluded contacts sheet
 */
function createExcludedSheet(spreadsheet, sourceSheet, excludedData) {
  // Delete existing excluded sheet if exists
  const existingSheet = spreadsheet.getSheetByName(COMPARISON_CONFIG.EXCLUDED_SHEET);
  if (existingSheet) {
    spreadsheet.deleteSheet(existingSheet);
    Logger.log('Deleted existing excluded sheet');
  }
  
  // Create new sheet
  const excludedSheet = spreadsheet.insertSheet(COMPARISON_CONFIG.EXCLUDED_SHEET);
  
  // Get source headers
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  // Build headers: Original Row + Exclusion Reason + Email Statuses + original data + Phone Details
  const headers = [
    'Original Row #',
    'Exclusion Reason',
    'Email Validation Statuses',
    ...sourceHeaders,
    'Phone Validation Details'
  ];
  
  // Write headers
  excludedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  const headerRange = excludedSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(COMPARISON_CONFIG.EXCLUDED_HEADER_COLOR);
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  // Highlight exclusion reason column
  excludedSheet.getRange(1, 2).setBackground(COMPARISON_CONFIG.REASON_HEADER_COLOR);
  excludedSheet.getRange(1, 2).setFontColor('#000000');
  
  // Write data
  const outputData = [];
  for (const excluded of excludedData) {
    // Build phone details string
    let phoneDetailsStr = '';
    if (excluded.phoneDetails.length > 0) {
      phoneDetailsStr = excluded.phoneDetails.map(p => 
        `${p.name}: ${p.number} (${p.status}, ${p.lineType})`
      ).join('\n');
    } else {
      phoneDetailsStr = 'No phones';
    }
    
    const row = [
      excluded.rowNumber,
      excluded.reason,
      excluded.emailStatus,
      ...excluded.rowData,
      phoneDetailsStr
    ];
    
    outputData.push(row);
  }
  
  // Write all data at once
  if (outputData.length > 0) {
    excludedSheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
  
  // Format exclusion reason column (highlight)
  const reasonColumnRange = excludedSheet.getRange(2, 2, outputData.length, 1);
  reasonColumnRange.setBackground('#fff3cd');  // Light yellow
  reasonColumnRange.setFontWeight('bold');
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    excludedSheet.autoResizeColumn(i);
  }
  
  // Freeze header row
  excludedSheet.setFrozenRows(1);
  
  // Move sheet to end
  spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());
  
  Logger.log(`Created excluded sheet with ${outputData.length} contacts`);
}

// =============================================================================
// MENU INTEGRATION
// =============================================================================

/**
 * Shows information dialog about the comparison tool
 */
function showComparisonInfo() {
  const html = HtmlService.createHtmlOutput(`
    <h3>CRM Ready Comparison Tool</h3>
    
    <h4>What It Does:</h4>
    <p>Compares "<strong>${COMPARISON_CONFIG.SOURCE_SHEET}</strong>" sheet with "<strong>${COMPARISON_CONFIG.CRM_READY_SHEET}</strong>" sheet to identify which contacts were excluded and why.</p>
    
    <h4>Requirements:</h4>
    <ul>
      <li>"<strong>${COMPARISON_CONFIG.SOURCE_SHEET}</strong>" sheet must exist (with validation data)</li>
      <li>"<strong>${COMPARISON_CONFIG.CRM_READY_SHEET}</strong>" sheet must exist (create it first)</li>
    </ul>
    
    <h4>Why These Sheets?</h4>
    <p><strong>${COMPARISON_CONFIG.SOURCE_SHEET}</strong> has the validated data before cleaning, so it matches properly with CRM Ready which contains cleaned data.</p>
    
    <h4>Output:</h4>
    <p>Creates "<strong>${COMPARISON_CONFIG.EXCLUDED_SHEET}</strong>" sheet with:</p>
    <ul>
      <li><strong>Original Row #:</strong> Row number in ${COMPARISON_CONFIG.SOURCE_SHEET}</li>
      <li><strong>Exclusion Reason:</strong> Why contact was excluded</li>
      <li><strong>Email Statuses:</strong> Validation status of ALL email columns</li>
      <li><strong>All Original Data:</strong> Complete contact information</li>
      <li><strong>Phone Details:</strong> All phone validation results</li>
    </ul>
    
    <h4>Exclusion Reasons:</h4>
    <ul>
      <li>No valid email AND no valid phone</li>
      <li>No valid email in any column (Primary, Email 1, Email 2, Personal)</li>
      <li>No valid phone in any column (any phone type, not just mobile)</li>
    </ul>
    
    <h4>How to Use:</h4>
    <ol>
      <li>Validate data in "${COMPARISON_CONFIG.SOURCE_SHEET}"</li>
      <li>Create CRM Ready sheet</li>
      <li>Click "Compare with CRM Ready"</li>
      <li>Review excluded contacts sheet</li>
      <li>Fix issues and re-validate if needed</li>
    </ol>
    
    <br>
    <p><em>💡 Tip: Use this to understand why specific contacts didn't make it to CRM Ready!</em></p>
  `)
    .setWidth(600)
    .setHeight(650);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Comparison Tool Info');
}

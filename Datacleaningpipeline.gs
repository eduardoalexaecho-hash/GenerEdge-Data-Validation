/**
 * =============================================================================
 * DATA CLEANING PIPELINE V3 - Production-Ready Google Sheets Apps Script
 * =============================================================================
 * 
 * VERSION 3 UPDATES (Critical Fix):
 * - Row-level duplicate detection: Uses COMPOSITE KEY (Name + Company + Website)
 * - Column-level deduplication: ONLY applies to Email and Phone columns
 * - Organization, Website, City, Description: NO duplicate removal (preserved as-is)
 * 
 * DUPLICATE DETECTION LOGIC:
 * A row is considered a duplicate ONLY if ALL THREE match:
 * 1. Contact Full Name
 * 2. Company Name - Cleaned
 * 3. Website
 * 
 * Example - NOT duplicates (same name, different companies):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Whitestone Builders | whitestone.com
 * 
 * Example - ARE duplicates (all three fields match):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Aspire Fine Homes | aspire.com  ← DUPLICATE ROW
 * 
 * IMPORTANT: Source sheet must have "Company Name - Cleaned" column
 *            Output sheet will show this as "Organization"
 * 
 * Previous versions (V1/V2) incorrectly removed duplicates within the same row
 * across ALL columns. V3 fixes this by:
 * 1. Removing duplicate contacts using composite key (Name + Company + Website)
 * 2. Only deduplicating emails/phones within each row
 * 
 * @author: Claude
 * @version: 3.0
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CONFIG = {
  // Sheet names
  OUTPUT_SHEET_NAME: 'Cleaned Data V3',
  
  // Columns to process (by header name)
  COLUMNS: {
    FULL_NAME: 'Contact Full Name',
    ORGANIZATION: 'Company Name - Cleaned',  // Source column name
    WEBSITE: 'Website',
    PRIMARY_EMAIL: 'Primary Email',
    EMAIL_1: 'Email 1',
    EMAIL_2: 'Email 2',
    PERSONAL_EMAIL: 'Personal Email',
    CONTACT_PHONE_1: 'Contact Phone 1',
    COMPANY_PHONE_1: 'Company Phone 1',
    COMPANY_PHONE_2: 'Company Phone 2',
    CONTACT_MOBILE: 'Contact Mobile Phone',
    COMPANY_CITY: 'Company City',
    COMPANY_DESC: 'Company Description'
  },
  
  // Performance settings
  BATCH_SIZE: 1000
};

// =============================================================================
// MAIN EXECUTION
// =============================================================================

/**
 * Main entry point for the data cleaning pipeline V3
 */
function runDataCleaningPipelineV3() {
  try {
    const startTime = new Date();
    Logger.log('=== Starting Data Cleaning Pipeline V3 ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getActiveSheet();
    
    Logger.log('Validating source sheet...');
    
    // Validate source sheet
    if (!validateSourceSheet(sourceSheet)) {
      throw new Error('Source sheet validation failed. Check required columns.');
    }
    
    Logger.log('✓ Validation passed');
    
    // Step 1: Load and parse data
    Logger.log('Step 1: Loading data...');
    const rawData = loadSheetData(sourceSheet);
    const originalRowCount = rawData.data.length;
    
    // Step 2: V3 UPDATED - Remove duplicate ROWS using composite key
    Logger.log('Step 2: Removing duplicate rows (by Name + Company + Website)...');
    const uniqueRowsData = removeDuplicateRows(rawData);
    const removedRows = originalRowCount - uniqueRowsData.data.length;
    Logger.log(`Removed ${removedRows} duplicate rows`);
    
    // Step 3: V3 UPDATED - Clean duplicates ONLY within Phone columns
    Logger.log('Step 3: Cleaning duplicates within Phone columns...');
    const cleanedData = cleanEmailPhoneDuplicates(uniqueRowsData);
    
    // Step 4: REMOVED - Email consolidation now happens ONLY in Step 5 for duplicates
    // No longer moves Primary → Email 1 when Email 1 is empty
    // Only consolidates actual duplicates
    Logger.log('Step 4: Skipped (consolidation only for duplicates in Step 5)');
    const emailFilledData = cleanedData;  // Pass through unchanged
    
    // Step 5: Deduplicate and normalize emails
    Logger.log('Step 5: Deduplicating emails...');
    const emailCleanedData = deduplicateEmails(emailFilledData);
    
    // Step 6: Deduplicate and normalize phones with priority
    Logger.log('Step 6: Deduplicating phones (Mobile priority)...');
    const phoneCleanedData = deduplicatePhonesWithPriority(emailCleanedData);
    
    // Step 7: Parse names
    Logger.log('Step 7: Parsing names...');
    const finalData = parseNames(phoneCleanedData);
    
    // Step 8: Create output sheet
    Logger.log('Step 8: Creating output sheet...');
    const outputSheet = createOutputSheet(ss, sourceSheet);
    
    // Step 9: Write data to output
    Logger.log('Step 9: Writing cleaned data...');
    writeOutputData(outputSheet, finalData, sourceSheet.getName());
    
    // Step 10: Add validation formula - REMOVED (causes timeout)
    // Logger.log('Step 10: Adding validation formula...');
    // addValidationFormula(outputSheet, sourceSheet.getName());
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== Pipeline V3 Complete in ${duration}s ===`);
    Logger.log(`Original rows: ${originalRowCount}`);
    Logger.log(`Removed duplicate rows: ${removedRows}`);
    Logger.log(`Final rows: ${finalData.data.length}`);
    
    // Build column summary
    const validation = CONFIG._VALIDATION_RESULTS || {};
    const foundOptional = validation.foundOptional || [];
    const skippedOptional = validation.skippedOptional || [];
    
    let columnSummary = '';
    if (foundOptional.length > 0) {
      columnSummary += `\n✓ Processed columns:\n`;
      columnSummary += foundOptional.map(col => `  • ${col}`).join('\n');
    }
    if (skippedOptional.length > 0) {
      columnSummary += `\n\n⊗ Skipped columns (not found):\n`;
      columnSummary += skippedOptional.map(col => `  • ${col}`).join('\n');
    }
    
    SpreadsheetApp.getUi().alert(
      'Success! (Version 3)',
      `✅ Data cleaning complete!\n\n` +
      `📊 Results:\n` +
      `Original rows: ${originalRowCount}\n` +
      `Duplicate rows removed: ${removedRows}\n` +
      `Final rows: ${finalData.data.length}\n` +
      `Duration: ${duration.toFixed(2)}s\n` +
      `Output: "${CONFIG.OUTPUT_SHEET_NAME}"\n` +
      columnSummary + `\n\n` +
      `✓ Duplicate Detection:\n` +
      `  • Composite key: Name + Company + Website\n` +
      `  • Row is duplicate only if ALL THREE match\n\n` +
      `✓ Deduplication:\n` +
      `  • Organization/Website preserved\n` +
      `  • Emails/Phones deduplicated within each row`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Explicit return to stop execution immediately after showing alert
    return;
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Pipeline failed: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Explicit return to stop execution immediately after error alert
    return;
  }
}

// =============================================================================
// STEP 1: DATA LOADING
// =============================================================================

/**
 * Validates that the source sheet contains required columns
 * STRICT matching - columns must have exact names
 * - Critical columns: Show warning to rename
 * - Optional columns: Skip silently if missing
 */
function validateSourceSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Define column requirements
  const columnRequirements = {
    // CRITICAL - Must exist with exact name
    critical: {
      'Contact Full Name': 'Contact Full Name',
      'Company Name - Cleaned': 'Company Name - Cleaned'  // Must be exactly this!
    },
    
    // OPTIONAL - Skip if missing
    optional: {
      'Website': 'Website',
      'Primary Email': 'Primary Email',
      'Email 1': 'Email 1',
      'Email 2': 'Email 2',
      'Personal Email': 'Personal Email',
      'Contact Phone 1': 'Contact Phone 1',
      'Company Phone 1': 'Company Phone 1',
      'Company Phone 2': 'Company Phone 2',
      'Contact Mobile Phone': 'Contact Mobile Phone',
      'Company Description': 'Company Description',
      'Company City': 'Company City'
    }
  };
  
  // Check critical columns
  const missingCritical = [];
  const foundColumns = {};
  
  for (const [colName, exactName] of Object.entries(columnRequirements.critical)) {
    if (headers.includes(exactName)) {
      foundColumns[colName] = exactName;
    } else {
      missingCritical.push(exactName);
    }
  }
  
  // If critical columns missing, show warning with rename instructions
  if (missingCritical.length > 0) {
    const warningMessage = 
      '⚠️ CRITICAL COLUMNS MISSING ⚠️\n\n' +
      'The following columns are required with EXACT names:\n' +
      missingCritical.map(col => `  • "${col}"`).join('\n') + '\n\n' +
      'Please rename your columns to match exactly.\n\n' +
      'Available headers in your sheet:\n' +
      headers.join(', ') + '\n\n' +
      'IMPORTANT: "Company Name - Cleaned" must be exactly that name,\n' +
      'including the " - Cleaned" part.';
    
    Logger.log('❌ ' + warningMessage);
    SpreadsheetApp.getUi().alert(
      '⚠️ Critical Columns Missing',
      warningMessage,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
  
  // Check optional columns (just log, don't fail)
  const foundOptional = [];
  const skippedOptional = [];
  
  for (const [colName, exactName] of Object.entries(columnRequirements.optional)) {
    if (headers.includes(exactName)) {
      foundColumns[colName] = exactName;
      foundOptional.push(exactName);
    } else {
      skippedOptional.push(exactName);
    }
  }
  
  // Store results for summary
  CONFIG._VALIDATION_RESULTS = {
    foundOptional: foundOptional,
    skippedOptional: skippedOptional,
    foundColumns: foundColumns
  };
  
  Logger.log(`✅ Validation complete: ${foundOptional.length} optional columns found, ${skippedOptional.length} skipped`);
  return true;
}

/**
 * Loads data from sheet and creates a structured object
 * Only processes columns that were found during validation
 */
function loadSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error('No data rows found in sheet');
  }
  
  // Get all data including headers
  const allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  // Create column index map ONLY for columns that were found
  const columnMap = {};
  
  if (CONFIG._VALIDATION_RESULTS && CONFIG._VALIDATION_RESULTS.foundColumns) {
    // Map only the columns that exist
    for (const [standardName, actualName] of Object.entries(CONFIG._VALIDATION_RESULTS.foundColumns)) {
      const index = headers.indexOf(actualName);
      if (index !== -1) {
        columnMap[actualName] = index;
      }
    }
  } else {
    // Fallback: map all headers
    headers.forEach((header, index) => {
      columnMap[header] = index;
    });
  }
  
  Logger.log(`Column map created with ${Object.keys(columnMap).length} columns`);
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: dataRows
  };
}

// =============================================================================
// STEP 2: ROW-LEVEL DUPLICATE REMOVAL (V3 NEW)
// =============================================================================

/**
 * V3 UPDATED: Removes duplicate ROWS based on composite key:
 * - Contact Full Name + Company Name - Cleaned + Website
 * 
 * A row is considered a duplicate ONLY if ALL THREE match:
 * 1. Same Contact Full Name
 * 2. Same Company Name - Cleaned
 * 3. Same Website
 * 
 * Example - These are NOT duplicates:
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Whitestone Builders | whitestone.com
 * 
 * Example - These ARE duplicates:
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Aspire Fine Homes | aspire.com  ← DUPLICATE
 */
function removeDuplicateRows(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const fullNameIdx = columnMap[CONFIG.COLUMNS.FULL_NAME];
  const companyNameIdx = columnMap[CONFIG.COLUMNS.ORGANIZATION];
  const websiteIdx = columnMap[CONFIG.COLUMNS.WEBSITE];
  const descriptionIdx = columnMap[CONFIG.COLUMNS.DESCRIPTION];
  
  const seenCombinations = new Set();
  const uniqueRows = [];
  
  data.forEach((row, index) => {
    const fullName = row[fullNameIdx];
    const companyName = row[companyNameIdx];
    const website = row[websiteIdx];
    const description = descriptionIdx !== undefined ? row[descriptionIdx] : '';
    
    // Normalize: trim and lowercase for comparison
    const nameKey = fullName ? fullName.toString().trim().toLowerCase() : '';
    const companyKey = companyName ? companyName.toString().trim().toLowerCase() : '';
    const websiteKey = website ? website.toString().trim().toLowerCase() : '';
    const descriptionKey = description ? description.toString().trim().toLowerCase() : '';
    
    // Check if any of the primary 3 fields is empty/blank
    const hasMissingPrimaryField = !nameKey || !companyKey || !websiteKey;
    
    // Create composite key
    let compositeKey;
    
    if (hasMissingPrimaryField && descriptionKey) {
      // If any primary field is blank AND we have description, use 4-field key
      compositeKey = `${nameKey}|${companyKey}|${websiteKey}|${descriptionKey}`;
      // Using enhanced duplicate check (Name + Company + Website + Description)
    } else {
      // All primary fields present OR no description, use standard 3-field key
      compositeKey = `${nameKey}|${companyKey}|${websiteKey}`;
    }
    
    if (seenCombinations.has(compositeKey)) {
      // Duplicate found - skip this entire row (no logging to avoid slowdown)
    } else {
      // First occurrence - keep this row
      seenCombinations.add(compositeKey);
      uniqueRows.push(row);
    }
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: uniqueRows
  };
}

// =============================================================================
// STEP 3: COLUMN-LEVEL DUPLICATE CLEANING (V3 UPDATED)
// =============================================================================

/**
 * V3 UPDATED: Removes duplicate values ONLY within Phone columns
 * Emails are handled separately in Step 5 (deduplicateEmails)
 * Does NOT touch: Organization, Website, Company Description, Company City
 */
function cleanEmailPhoneDuplicates(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  // V3: ONLY clean duplicates in Phone columns
  // Emails are handled in Step 5 with proper business logic
  // PRIORITY ORDER (keep first occurrence):
  // Phones: Contact Mobile Phone > Contact Phone 1 > Company Phone 1 > Company Phone 2
  const columnsToClean = [
    CONFIG.COLUMNS.CONTACT_MOBILE,    // Priority 1 for phones
    CONFIG.COLUMNS.CONTACT_PHONE_1,   // Priority 2 for phones
    CONFIG.COLUMNS.COMPANY_PHONE_1,   // Priority 3 for phones
    CONFIG.COLUMNS.COMPANY_PHONE_2    // Priority 4 for phones
  ];
  
  const cleanedData = data.map(row => {
    const newRow = [...row];
    const seenValues = new Set();
    
    // Get indices for columns to clean (only phones)
    const indicesToClean = columnsToClean
      .map(col => columnMap[col])
      .filter(idx => idx !== undefined);
    
    indicesToClean.forEach(colIdx => {
      const value = row[colIdx];
      
      // Skip empty values
      if (!value || value.toString().trim() === '') {
        return;
      }
      
      // Normalize for comparison: trim and lowercase
      const normalizedValue = value.toString().trim().toLowerCase();
      
      if (seenValues.has(normalizedValue)) {
        // Duplicate found within this row - clear it
        newRow[colIdx] = '';
      } else {
        // First occurrence - keep it and add to set
        seenValues.add(normalizedValue);
        // Keep original value but trimmed
        newRow[colIdx] = value.toString().trim();
      }
    });
    
    return newRow;
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: cleanedData
  };
}

// =============================================================================
// STEP 4: PRIMARY EMAIL FILL RULE
// =============================================================================

/**
 * Move Primary Email to Email 1 column (reverse of old logic)
 * Priority: Email 1 is the main email column
 * - If Email 1 is empty, move Primary Email → Email 1 and clear Primary
 * - If both have values, keep both (deduplication happens in Step 5)
 */
function fillPrimaryEmail(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const primaryIdx = columnMap[CONFIG.COLUMNS.PRIMARY_EMAIL];
  const email1Idx = columnMap[CONFIG.COLUMNS.EMAIL_1];
  
  // Skip if Primary Email column doesn't exist in source
  if (primaryIdx === undefined) {
    Logger.log('⊗ Skipping Primary Email move: Primary Email column not in source');
    return dataObject;
  }
  
  // Skip if Email 1 column doesn't exist
  if (email1Idx === undefined) {
    Logger.log('⊗ Skipping Primary Email move: Email 1 column not in source');
    return dataObject;
  }
  
  const filledData = data.map(row => {
    const newRow = [...row];
    
    const primary = row[primaryIdx] ? row[primaryIdx].toString().trim() : '';
    const email1 = row[email1Idx] ? row[email1Idx].toString().trim() : '';
    
    // If Email 1 is empty AND Primary Email has a value
    // Move Primary Email → Email 1
    if (!email1 && primary) {
      newRow[email1Idx] = primary;  // Move to Email 1
      newRow[primaryIdx] = '';       // Clear Primary
    }
    
    return newRow;
  });
  
  Logger.log('✓ Primary Email move to Email 1 complete');
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: filledData
  };
}

// =============================================================================
// STEP 5: EMAIL DEDUPLICATION
// =============================================================================

/**
 * Deduplicates emails according to business rules and normalizes them
 * Safely handles missing email columns
 */
function deduplicateEmails(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  // Get column indices (may be undefined if column doesn't exist)
  const primaryIdx = columnMap[CONFIG.COLUMNS.PRIMARY_EMAIL];
  const email1Idx = columnMap[CONFIG.COLUMNS.EMAIL_1];
  const email2Idx = columnMap[CONFIG.COLUMNS.EMAIL_2];
  const personalIdx = columnMap[CONFIG.COLUMNS.PERSONAL_EMAIL];
  
  // Check if we have any email columns to process
  if (primaryIdx === undefined && email1Idx === undefined && 
      email2Idx === undefined && personalIdx === undefined) {
    Logger.log('⊗ Skipping email deduplication: No email columns found');
    return dataObject;  // Return unchanged
  }
  
  const cleanedData = data.map(row => {
    const newRow = [...row];
    
    // Normalize and get emails (handle undefined indices)
    let primary = primaryIdx !== undefined ? normalizeEmail(row[primaryIdx]) : '';
    let email1 = email1Idx !== undefined ? normalizeEmail(row[email1Idx]) : '';
    let email2 = email2Idx !== undefined ? normalizeEmail(row[email2Idx]) : '';
    let personal = personalIdx !== undefined ? normalizeEmail(row[personalIdx]) : '';
    
    // Apply deduplication logic
    const result = deduplicateEmailSet(primary, email1, email2);
    
    // Update row with deduplicated values (only if column exists)
    if (primaryIdx !== undefined) newRow[primaryIdx] = result.primary;
    if (email1Idx !== undefined) newRow[email1Idx] = result.email1;
    if (email2Idx !== undefined) newRow[email2Idx] = result.email2;
    if (personalIdx !== undefined) newRow[personalIdx] = personal; // Personal email stays separate
    
    return newRow;
  });
  
  Logger.log('✓ Email deduplication complete');
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: cleanedData
  };
}

/**
 * Normalizes an email: trim and lowercase
 */
function normalizeEmail(email) {
  if (!email || email.toString().trim() === '') {
    return '';
  }
  return email.toString().trim().toLowerCase();
}

/**
 * Implements email deduplication business logic
 * PRIORITY: Email 1 is the main email column
 * All duplicates are consolidated to Email 1
 * - If Primary = Email 1: Clear Primary, keep Email 1
 * - If Email 2 = Email 1: Clear Email 2, keep Email 1
 * - If all three same: Clear Primary and Email 2, keep Email 1
 * (Personal Email stays separate)
 */
function deduplicateEmailSet(primary, email1, email2) {
  // All empty - return empty
  if (!primary && !email1 && !email2) {
    return { primary: '', email1: '', email2: '' };
  }
  
  // All three are the same - consolidate to Email 1
  if (primary && primary === email1 && primary === email2) {
    return { primary: '', email1: email1, email2: '' };
  }
  
  // Primary = Email1 - consolidate to Email 1
  if (primary && primary === email1) {
    return { primary: '', email1: email1, email2: email2 };
  }
  
  // Primary = Email2 - keep both (Primary in Primary, Email 2 in Email 2)
  if (primary && primary === email2) {
    return { primary: primary, email1: email1, email2: email2 };
  }
  
  // Email1 = Email2 - consolidate to Email 1
  if (email1 && email1 === email2) {
    return { primary: primary, email1: email1, email2: '' };
  }
  
  // All different - keep all
  return { primary: primary, email1: email1, email2: email2 };
}

// =============================================================================
// STEP 6: PHONE DEDUPLICATION WITH MOBILE PRIORITY
// =============================================================================

/**
 * Deduplicates phones with Contact Mobile Phone priority
 * - Removes dots from phone numbers
 * - If duplicate exists, always keep Contact Mobile Phone
 */
function deduplicatePhonesWithPriority(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  // Get column indices (may be undefined if column doesn't exist)
  const contactPhone1Idx = columnMap[CONFIG.COLUMNS.CONTACT_PHONE_1];
  const companyPhone1Idx = columnMap[CONFIG.COLUMNS.COMPANY_PHONE_1];
  const companyPhone2Idx = columnMap[CONFIG.COLUMNS.COMPANY_PHONE_2];
  const contactMobileIdx = columnMap[CONFIG.COLUMNS.CONTACT_MOBILE];
  
  // Check if we have any phone columns to process
  if (contactPhone1Idx === undefined && companyPhone1Idx === undefined && 
      companyPhone2Idx === undefined && contactMobileIdx === undefined) {
    Logger.log('⊗ Skipping phone deduplication: No phone columns found');
    return dataObject;  // Return unchanged
  }
  
  const cleanedData = data.map(row => {
    const newRow = [...row];
    
    // Get and normalize all phone numbers (remove dots) - handle undefined indices
    let contactPhone1 = contactPhone1Idx !== undefined ? normalizePhone(row[contactPhone1Idx]) : '';
    let companyPhone1 = companyPhone1Idx !== undefined ? normalizePhone(row[companyPhone1Idx]) : '';
    let companyPhone2 = companyPhone2Idx !== undefined ? normalizePhone(row[companyPhone2Idx]) : '';
    let contactMobile = contactMobileIdx !== undefined ? normalizePhone(row[contactMobileIdx]) : '';
    
    // Create map for comparison (digits only)
    const phoneMap = new Map();
    
    // Priority order: Contact Mobile Phone is highest priority
    let mobileDigits = '';
    if (contactMobile) {
      mobileDigits = extractDigitsOnly(contactMobile);
      phoneMap.set(mobileDigits, 'mobile');
    }
    
    // Now check other phones - if they match mobile, clear them
    if (contactPhone1) {
      const digits = extractDigitsOnly(contactPhone1);
      if (digits === mobileDigits) {
        contactPhone1 = ''; // Clear duplicate
      } else if (!phoneMap.has(digits)) {
        phoneMap.set(digits, 'contact1');
      } else {
        contactPhone1 = ''; // Clear duplicate
      }
    }
    
    if (companyPhone1) {
      const digits = extractDigitsOnly(companyPhone1);
      if (digits === mobileDigits) {
        companyPhone1 = ''; // Clear duplicate
      } else if (!phoneMap.has(digits)) {
        phoneMap.set(digits, 'company1');
      } else {
        companyPhone1 = ''; // Clear duplicate
      }
    }
    
    if (companyPhone2) {
      const digits = extractDigitsOnly(companyPhone2);
      if (digits === mobileDigits) {
        companyPhone2 = ''; // Clear duplicate
      } else if (!phoneMap.has(digits)) {
        phoneMap.set(digits, 'company2');
      } else {
        companyPhone2 = ''; // Clear duplicate
      }
    }
    
    // Update row with cleaned values (only if column exists)
    if (contactPhone1Idx !== undefined) newRow[contactPhone1Idx] = contactPhone1;
    if (companyPhone1Idx !== undefined) newRow[companyPhone1Idx] = companyPhone1;
    if (companyPhone2Idx !== undefined) newRow[companyPhone2Idx] = companyPhone2;
    if (contactMobileIdx !== undefined) newRow[contactMobileIdx] = contactMobile;
    
    return newRow;
  });
  
  Logger.log('✓ Phone deduplication complete');
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: cleanedData
  };
}

/**
 * Normalizes phone by removing dots
 */
function normalizePhone(phone) {
  if (!phone || phone.toString().trim() === '') {
    return '';
  }
  
  const phoneStr = phone.toString().trim();
  
  // Remove dots from phone number
  return phoneStr.replace(/\./g, '');
}

/**
 * Extracts only digits from phone for comparison
 */
function extractDigitsOnly(phone) {
  if (!phone) return '';
  return phone.toString().replace(/\D/g, '');
}

// =============================================================================
// STEP 7: NAME PARSING
// =============================================================================

/**
 * Parses full names into first and last names intelligently
 */
function parseNames(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const fullNameIdx = columnMap[CONFIG.COLUMNS.FULL_NAME];
  
  const parsedData = data.map(row => {
    const fullName = row[fullNameIdx] ? row[fullNameIdx].toString().trim() : '';
    const { firstName, lastName } = parseFullName(fullName);
    
    return {
      originalRow: row,
      firstName: firstName,
      lastName: lastName
    };
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: parsedData
  };
}

/**
 * Intelligently parses a full name into first and last name
 */
function parseFullName(fullName) {
  if (!fullName || fullName.trim() === '') {
    return { firstName: '', lastName: '' };
  }
  
  const parts = fullName.trim().split(/\s+/).filter(p => p.length > 0);
  
  if (parts.length === 0) {
    return { firstName: '', lastName: '' };
  }
  
  if (parts.length === 1) {
    return { firstName: parts[0], lastName: '' };
  }
  
  // For 2+ words: split in the middle
  const midPoint = Math.ceil(parts.length / 2);
  const firstName = parts.slice(0, midPoint).join(' ');
  const lastName = parts.slice(midPoint).join(' ');
  
  return { firstName, lastName };
}

// =============================================================================
// STEP 8: OUTPUT SHEET CREATION
// =============================================================================

/**
 * Creates or clears the output sheet
 */
function createOutputSheet(spreadsheet, sourceSheet) {
  let outputSheet = spreadsheet.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = spreadsheet.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
  }
  
  return outputSheet;
}

/**
 * Writes cleaned data with specified column order
 * Note: Reads "Company Name - Cleaned" from source, outputs as "Organization"
 */
function writeOutputData(outputSheet, dataObject, sourceSheetName) {
  const { headers, columnMap, data } = dataObject;
  
  // Get validation results to know which optional columns were found
  const validation = CONFIG._VALIDATION_RESULTS || {};
  const foundOptional = validation.foundOptional || [];
  
  // Define ALL possible output columns
  const allPossibleColumns = [
    'Contact Full Name',      // Always included (critical)
    'First Name',             // Always included (generated)
    'Last Name',              // Always included (generated)
    'Organization',           // Always included (critical)
    'Primary Email',          // Optional
    'Email 1',                // Optional
    'Email 2',                // Optional
    'Personal Email',         // Optional
    'Contact Phone 1',        // Optional
    'Company Phone 1',        // Optional
    'Company Phone 2',        // Optional
    'Contact Mobile Phone',   // Optional
    'Company Description',    // Optional
    'Website',                // Optional
    'Company City',           // Optional
    'Validation Status'       // Always included
  ];
  
  // Build output headers - only include optional columns that were found
  const outputHeaders = allPossibleColumns.filter(header => {
    // Always include critical columns
    if (header === 'Contact Full Name' || 
        header === 'Organization' || 
        header === 'First Name' || 
        header === 'Last Name' ||
        header === 'Validation Status') {
      return true;
    }
    // Include optional columns only if they were found in source
    return foundOptional.includes(header);
  });
  
  // Write headers
  outputSheet.getRange(1, 1, 1, outputHeaders.length)
    .setValues([outputHeaders])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // Map of output header to source column
  const headerMapping = {
    'Contact Full Name': CONFIG.COLUMNS.FULL_NAME,
    'Organization': CONFIG.COLUMNS.ORGANIZATION,
    'Primary Email': CONFIG.COLUMNS.PRIMARY_EMAIL,
    'Email 1': CONFIG.COLUMNS.EMAIL_1,
    'Email 2': CONFIG.COLUMNS.EMAIL_2,
    'Personal Email': CONFIG.COLUMNS.PERSONAL_EMAIL,
    'Contact Phone 1': CONFIG.COLUMNS.CONTACT_PHONE_1,
    'Company Phone 1': CONFIG.COLUMNS.COMPANY_PHONE_1,
    'Company Phone 2': CONFIG.COLUMNS.COMPANY_PHONE_2,
    'Contact Mobile Phone': CONFIG.COLUMNS.CONTACT_MOBILE,
    'Company Description': CONFIG.COLUMNS.COMPANY_DESC,
    'Website': CONFIG.COLUMNS.WEBSITE,
    'Company City': CONFIG.COLUMNS.COMPANY_CITY
  };
  
  // Prepare output data
  const outputData = data.map(item => {
    const row = item.originalRow;
    const outputRow = [];
    
    // Add columns in exact order
    outputHeaders.forEach(header => {
      if (header === 'First Name') {
        outputRow.push(item.firstName);
      } else if (header === 'Last Name') {
        outputRow.push(item.lastName);
      } else if (header === 'Validation Status') {
        outputRow.push('');
      } else {
        const sourceColumn = headerMapping[header];
        const colIdx = columnMap[sourceColumn];
        outputRow.push(colIdx !== undefined ? row[colIdx] : '');
      }
    });
    
    return outputRow;
  });
  
  // Write data in batches
  if (outputData.length > 0) {
    const batchSize = CONFIG.BATCH_SIZE;
    for (let i = 0; i < outputData.length; i += batchSize) {
      const batch = outputData.slice(i, i + batchSize);
      const startRow = i + 2;
      outputSheet.getRange(startRow, 1, batch.length, outputHeaders.length)
        .setValues(batch);
    }
  }
  
  // Format sheet
  outputSheet.setFrozenRows(1);
  // Auto-resize columns - REMOVED (too slow on large datasets)
  // outputSheet.autoResizeColumns(1, outputHeaders.length);
  
  Logger.log(`✓ Data written successfully: ${outputData.length} rows`);
}

// =============================================================================
// STEP 9: VALIDATION FORMULA
// =============================================================================

/**
 * DISABLED: Adds validation formula to check data integrity
 * This function causes timeout on large datasets and has been disabled
 * @deprecated No longer used in pipeline
 */
function addValidationFormula(outputSheet, sourceSheetName) {
  const lastRow = outputSheet.getLastRow();
  
  if (lastRow < 2) {
    return;
  }
  
  const validationCol = 16;
  
  const formulaTemplate = `=IFERROR(
 IF(
  AND(
   A{ROW} = INDEX('${sourceSheetName}'!A:A, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   D{ROW} = INDEX('${sourceSheetName}'!B:B, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   N{ROW} = INDEX('${sourceSheetName}'!C:C, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   F{ROW} = INDEX('${sourceSheetName}'!E:E, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   TRIM(SUBSTITUTE(M{ROW},CHAR(10)," ")) =
     TRIM(SUBSTITUTE(INDEX('${sourceSheetName}'!M:M, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),CHAR(10)," ")),
   O{ROW} = INDEX('${sourceSheetName}'!L:L, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0)),
   L{ROW} = INDEX('${sourceSheetName}'!K:K, MATCH(E{ROW}, '${sourceSheetName}'!D:D, 0))
  ),
  "MATCH OK",
  "DATA MISMATCH"
 ),
 "NOT FOUND"
)`;
  
  // Apply formula to each row
  for (let row = 2; row <= lastRow; row++) {
    const formula = formulaTemplate.replace(/{ROW}/g, row);
    outputSheet.getRange(row, validationCol).setFormula(formula);
  }
  
  // Format validation column
  const validationRange = outputSheet.getRange(2, validationCol, lastRow - 1, 1);
  
  const matchRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('MATCH OK')
    .setBackground('#d9ead3')
    .setFontColor('#38761d')
    .setRanges([validationRange])
    .build();
  
  const mismatchRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('DATA MISMATCH')
    .setBackground('#f4cccc')
    .setFontColor('#cc0000')
    .setRanges([validationRange])
    .build();
  
  const notFoundRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('NOT FOUND')
    .setBackground('#fff2cc')
    .setFontColor('#bf9000')
    .setRanges([validationRange])
    .build();
  
  const rules = outputSheet.getConditionalFormatRules();
  rules.push(matchRule, mismatchRule, notFoundRule);
  outputSheet.setConditionalFormatRules(rules);
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

/**
 * Creates ONE unified menu in the Google Sheets UI
 * Works on any Google Sheet when this Apps Script project is attached
 * 
 * MENU STRUCTURE:
 * 🔧 Data Tools
 * ├── 📊 Data Cleanup
 * │   └── Run Pipeline V3
 * ├── 🏠 Categorize Companies
 * │   ├── Categorize All Companies
 * │   └── Upload Checked Entries
 * ├── 📧 Email Validation
 * │   ├── Validate Primary Email
 * │   ├── Validate Email 1
 * │   ├── Validate Email 2
 * │   └── Validate Personal Email
 * ├── 📞 Phone Validation
 * │   ├── Validate Contact Phone 1
 * │   ├── Validate Company Phone 1
 * │   ├── Validate Company Phone 2
 * │   └── Validate Contact Mobile Phone
 * └── 🎯 CRM Ready
 *     ├── Create CRM Ready Sheet
 *     ├── Compare with CRM Ready
 *     ├── Comparison Info
 *     └── View Settings
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Create main menu with submenus
  ui.createMenu('🔧 Data Tools')
    
    // Data Cleanup submenu
    .addSubMenu(ui.createMenu('📊 Data Cleanup')
      .addItem('▶️ Run Pipeline V3', 'runDataCleaningPipelineV3')
      .addSeparator()
      .addItem('📋 View Logs', 'showLogs'))
    
    // Categorize Companies submenu
    .addSubMenu(ui.createMenu('🏠 Categorize Companies')
      .addItem('▶️ Categorize All Companies', 'categorizeAllCompanies')
      .addSeparator()
      .addItem('📤 Upload Checked Entries', 'uploadCheckedEntriesCat'))
    
    // Email Validation submenu
    .addSubMenu(ui.createMenu('📧 Email Validation')
      .addItem('📬 Validate Primary Email', 'validatePrimaryEmail')
      .addItem('📧 Validate Email 1', 'validateEmail1')
      .addItem('📧 Validate Email 2', 'validateEmail2')
      .addItem('📨 Validate Personal Email', 'validatePersonalEmail')
      .addSeparator()
      .addItem('⚙️ Configure Settings', 'showEmailConfigDialog')
      .addItem('📋 View Validation Logs', 'showEmailValidationLogs'))
    
    // Phone Validation submenu
    .addSubMenu(ui.createMenu('📞 Phone Validation')
      .addItem('📱 Validate Contact Phone 1', 'validateContactPhone1')
      .addItem('🏢 Validate Company Phone 1', 'validateCompanyPhone1')
      .addItem('🏢 Validate Company Phone 2', 'validateCompanyPhone2')
      .addItem('📱 Validate Contact Mobile Phone', 'validateContactMobilePhone')
      .addSeparator()
      .addItem('⚙️ Configure Settings', 'showConfigDialog')
      .addItem('📋 View Validation Logs', 'showValidationLogs'))
    
    // Filter by Criteria submenu (NEW!)
    .addSubMenu(ui.createMenu('🎯 CRM Ready')
      .addItem('▶️ Create CRM Ready Sheet', 'createCRMReady')
      .addItem('📞 Generate CRM Callers Ready', 'generateCRMCallersReady')
      .addSeparator()
      .addItem('📊 Compare with CRM Ready', 'compareWithCRMReady')
      .addItem('ℹ️ Comparison Info', 'showComparisonInfo')
      .addSeparator()
      .addItem('⚙️ View Settings', 'showCRMConfig'))
    
    .addToUi();
}

/**
 * Shows execution logs
 */
function showLogs() {
  const logs = Logger.getLog();
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Execution Logs',
    logs || 'No logs available',
    ui.ButtonSet.OK
  );
}

// =============================================================================
// END OF SCRIPT V3
// =============================================================================

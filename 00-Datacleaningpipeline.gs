/**
 * =============================================================================
 * DATA CLEANING PIPELINE V3 - Production-Ready Google Sheets Apps Script
 * =============================================================================
 * 
 * VERSION 3.3 CRITICAL FIX:
 * - FIXED: Duplicate detection was treating different people at same company as duplicates
 * - ROOT CAUSE: When Contact Full Name had whitespace, it didn't check First+Last name
 * - SOLUTION: Now checks if Contact Full Name is non-empty after trim, falls back to First+Last
 * 
 * EXAMPLE BUG (BEFORE FIX):
 * - Marty Helgerson | Frase Construction | frase.com
 * - Jeremy Frase | Frase Construction | frase.com
 * ❌ Were marked as duplicates (Match: "Company + Website")
 * ✅ Now correctly kept as separate contacts (Match: "Name + Company + Website")
 * 
 * VERSION 3.2 CRITICAL FIX:
 * - FIXED: Duplicate detection now uses First Name + Last Name when Contact Full Name is empty
 * - BEFORE: Would treat all contacts from same company as duplicates if no Contact Full Name
 * - AFTER: Uses First + Last name to distinguish different people from same company
 * 
 * VERSION 3 UPDATES (Critical Fix):
 * - Row-level duplicate detection: Uses COMPOSITE KEY (Name + Company + Website)
 * - Column-level deduplication: ONLY applies to Email and Phone columns
 * - Organization, Website, City, Description: NO duplicate removal (preserved as-is)
 * 
 * DUPLICATE DETECTION LOGIC (UPDATED):
 * - Priority 1: Contact Full Name + Company + Website (if Contact Full Name non-empty)
 * - Priority 2: First Name + Last Name + Company + Website (if Contact Full Name empty)
 * - Priority 3: First Name + Last Name + Company (no website)
 * - Fallback: Company + Website (ONLY if NO name data at all)
 * 
 * Example - NOT duplicates (different people, same company):
 * - Marty Helgerson | Frase Construction | frase.com
 * - Jeremy Frase | Frase Construction | frase.com  ← DIFFERENT PERSON, NOT DUPLICATE!
 * 
 * Example - ARE duplicates (same person, same company):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Aspire Fine Homes | aspire.com  ← DUPLICATE ROW
 * 
 * REQUIRED COLUMNS:
 * - Company Name - Cleaned (CRITICAL - must exist)
 * 
 * OPTIONAL COLUMNS:
 * - Contact Full Name, First Name, Last Name (at least one recommended)
 * - Website, Email columns, Phone columns, Company Description, Company City
 * 
 * OUTPUT BEHAVIOR:
 * - If Contact Full Name exists → Output includes: Contact Full Name, First Name, Last Name
 * - If Contact Full Name missing → Output includes: First Name, Last Name (if they exist)
 * 
 * @author: Claude
 * @version: 3.3 - CRITICAL FIX: Different people at same company no longer treated as duplicates
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
    FIRST_NAME: 'First Name',  // Added
    LAST_NAME: 'Last Name',    // Added
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
    
    // Step 10: Create duplicates report
    Logger.log('Step 10: Creating duplicates report...');
    const duplicatesSheet = createDuplicatesReport(ss, uniqueRowsData.duplicates);
    
    // Step 11: Add validation formula - REMOVED (causes timeout)
    // Logger.log('Step 11: Adding validation formula...');
    // addValidationFormula(outputSheet, sourceSheet.getName());
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== Pipeline V3 Complete in ${duration}s ===`);
    Logger.log(`Original rows: ${originalRowCount}`);
    Logger.log(`Removed duplicate rows: ${removedRows}`);
    Logger.log(`Duplicates tracked: ${uniqueRowsData.duplicates.length}`);
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
      `Duration: ${duration.toFixed(2)}s\n\n` +
      `📄 Output Sheets:\n` +
      `  • "${CONFIG.OUTPUT_SHEET_NAME}" - Cleaned data\n` +
      `  • "Duplicates Report" - ${uniqueRowsData.duplicates.length} duplicates found\n` +
      columnSummary + `\n\n` +
      `✓ Duplicate Detection:\n` +
      `  • Composite key: Name + Company + Website\n` +
      `  • Row is duplicate only if ALL match\n\n` +
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
      'Company Name - Cleaned': 'Company Name - Cleaned'  // Must be exactly this!
    },
    
    // OPTIONAL - Skip if missing
    optional: {
      'Contact Full Name': 'Contact Full Name',  // Now optional!
      'First Name': 'First Name',  // Added
      'Last Name': 'Last Name',    // Added
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
 * - Contact Full Name + Company Name - Cleaned + Website (if all exist)
 * - Company Name - Cleaned + Website (if Contact Full Name missing)
 * - Falls back to Company + Description if Website missing
 * 
 * A row is considered a duplicate based on available fields:
 * 
 * Example - These are NOT duplicates (different companies):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Whitestone Builders | whitestone.com
 * 
 * Example - These ARE duplicates (all fields match):
 * - David Gordon | Aspire Fine Homes | aspire.com
 * - David Gordon | Aspire Fine Homes | aspire.com  ← DUPLICATE
 * 
 * If no Contact Full Name column:
 * - ABC Company | abc.com
 * - ABC Company | abc.com  ← DUPLICATE (same company + website)
 */
function removeDuplicateRows(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const fullNameIdx = columnMap[CONFIG.COLUMNS.FULL_NAME];
  const firstNameIdx = columnMap[CONFIG.COLUMNS.FIRST_NAME];
  const lastNameIdx = columnMap[CONFIG.COLUMNS.LAST_NAME];
  const companyNameIdx = columnMap[CONFIG.COLUMNS.ORGANIZATION];
  const websiteIdx = columnMap[CONFIG.COLUMNS.WEBSITE];
  const descriptionIdx = columnMap[CONFIG.COLUMNS.COMPANY_DESC];
  
  const seenCombinations = new Map(); // Changed to Map to track first occurrence row
  const uniqueRows = [];
  const duplicates = []; // Track duplicate information
  
  data.forEach((row, index) => {
    // CRITICAL FIX: Build name from Contact Full Name OR First + Last Name
    let nameForKey = '';
    
    // Priority 1: Use Contact Full Name if it exists and has NON-EMPTY value
    if (fullNameIdx !== undefined && row[fullNameIdx]) {
      const fullNameValue = row[fullNameIdx].toString().trim();
      if (fullNameValue) {
        // Contact Full Name has actual content
        nameForKey = fullNameValue;
      }
    }
    
    // Priority 2: Use First Name + Last Name if nameForKey is still empty
    // This runs if: (a) Contact Full Name column doesn't exist, OR
    //               (b) Contact Full Name is empty/whitespace
    if (!nameForKey && (firstNameIdx !== undefined || lastNameIdx !== undefined)) {
      const firstName = firstNameIdx !== undefined ? (row[firstNameIdx] || '').toString().trim() : '';
      const lastName = lastNameIdx !== undefined ? (row[lastNameIdx] || '').toString().trim() : '';
      // Combine First + Last (with space if both exist)
      if (firstName && lastName) {
        nameForKey = `${firstName} ${lastName}`;
      } else if (firstName) {
        nameForKey = firstName;
      } else if (lastName) {
        nameForKey = lastName;
      }
    }
    
    const companyName = row[companyNameIdx];
    const website = websiteIdx !== undefined ? row[websiteIdx] : '';
    const description = descriptionIdx !== undefined ? row[descriptionIdx] : '';
    
    // Normalize: trim and lowercase for comparison
    const nameKey = nameForKey ? nameForKey.toLowerCase() : '';
    const companyKey = companyName ? companyName.toString().trim().toLowerCase() : '';
    const websiteKey = website ? website.toString().trim().toLowerCase() : '';
    const descriptionKey = description ? description.toString().trim().toLowerCase() : '';
    
    // Create composite key based on available fields
    let compositeKey;
    let matchCriteria;
    
    if (nameKey && companyKey && websiteKey) {
      // Standard 3-field key (Name + Company + Website)
      compositeKey = `${nameKey}|${companyKey}|${websiteKey}`;
      matchCriteria = 'Name + Company + Website';
    } else if (nameKey && companyKey) {
      // Name + Company (no website available)
      compositeKey = `${nameKey}|${companyKey}`;
      matchCriteria = 'Name + Company';
    } else if (companyKey && websiteKey) {
      // No name available: Company + Website only
      compositeKey = `${companyKey}|${websiteKey}`;
      matchCriteria = 'Company + Website';
    } else if (companyKey && descriptionKey) {
      // No name or website: Company + Description
      compositeKey = `${companyKey}|${descriptionKey}`;
      matchCriteria = 'Company + Description';
    } else if (companyKey) {
      // Only company name available
      compositeKey = `${companyKey}`;
      matchCriteria = 'Company Name Only';
    } else {
      // No reliable fields - include all available
      compositeKey = `${nameKey}|${companyKey}|${websiteKey}|${descriptionKey}`;
      matchCriteria = 'All Available Fields';
    }
    
    if (seenCombinations.has(compositeKey)) {
      // Duplicate found - track it
      const originalRowNumber = seenCombinations.get(compositeKey);
      duplicates.push({
        duplicateRowNumber: index + 2, // +2 because: +1 for header, +1 for 1-based indexing
        originalRowNumber: originalRowNumber,
        fullName: nameForKey,  // Use the constructed name
        companyName: companyName,
        website: website,
        description: description,
        matchCriteria: matchCriteria,
        rowData: row
      });
    } else {
      // First occurrence - keep this row
      seenCombinations.set(compositeKey, index + 2); // Store row number (+2 for sheet row number)
      uniqueRows.push(row);
    }
  });
  
  return {
    headers: headers,
    columnMap: columnMap,
    data: uniqueRows,
    duplicates: duplicates  // Return duplicate info
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
 * Handles name data intelligently based on what columns exist:
 * 
 * Scenario 1: First Name + Last Name columns exist
 *   → Use them directly (even if Contact Full Name also exists)
 * 
 * Scenario 2: Only Contact Full Name exists (no First/Last columns)
 *   → Parse Contact Full Name into First + Last
 * 
 * Scenario 3: None exist
 *   → Empty names
 * 
 * OUTPUT: Always includes First Name and Last Name in output
 *         (either from source columns or parsed from Contact Full Name)
 */
function parseNames(dataObject) {
  const { headers, columnMap, data } = dataObject;
  
  const fullNameIdx = columnMap[CONFIG.COLUMNS.FULL_NAME];
  const firstNameIdx = columnMap[CONFIG.COLUMNS.FIRST_NAME];
  const lastNameIdx = columnMap[CONFIG.COLUMNS.LAST_NAME];
  
  // Determine strategy
  const hasFirstOrLastColumn = firstNameIdx !== undefined || lastNameIdx !== undefined;
  
  // Debug logging
  Logger.log('parseNames - Strategy:');
  Logger.log(`  Contact Full Name index: ${fullNameIdx}`);
  Logger.log(`  First Name index: ${firstNameIdx}`);
  Logger.log(`  Last Name index: ${lastNameIdx}`);
  
  if (hasFirstOrLastColumn) {
    Logger.log('  → Strategy: Using existing First/Last Name columns');
  } else if (fullNameIdx !== undefined) {
    Logger.log('  → Strategy: Parsing Contact Full Name');
  } else {
    Logger.log('  → Strategy: No name columns found - using empty names');
  }
  
  const parsedData = data.map((row, idx) => {
    let firstName = '';
    let lastName = '';
    
    // PRIORITY 1: Use existing First Name and Last Name columns if they exist
    // (This takes priority even if Contact Full Name also exists)
    if (hasFirstOrLastColumn) {
      firstName = firstNameIdx !== undefined && row[firstNameIdx] 
        ? row[firstNameIdx].toString().trim() 
        : '';
      lastName = lastNameIdx !== undefined && row[lastNameIdx] 
        ? row[lastNameIdx].toString().trim() 
        : '';
      
      // Debug first row
      if (idx === 0) {
        Logger.log(`  Row 1 - Using source columns:`);
        Logger.log(`    First Name: "${firstName}"`);
        Logger.log(`    Last Name: "${lastName}"`);
      }
    } 
    // PRIORITY 2: Parse from Contact Full Name if no First/Last columns exist
    else if (fullNameIdx !== undefined && row[fullNameIdx]) {
      const fullName = row[fullNameIdx].toString().trim();
      const parsed = parseFullName(fullName);
      firstName = parsed.firstName;
      lastName = parsed.lastName;
      
      // Debug first row
      if (idx === 0) {
        Logger.log(`  Row 1 - Parsed from Contact Full Name: "${fullName}"`);
        Logger.log(`    → First: "${firstName}"`);
        Logger.log(`    → Last: "${lastName}"`);
      }
    }
    // PRIORITY 3: No name columns at all - use empty
    else {
      if (idx === 0) {
        Logger.log(`  Row 1 - No name data available (empty)`);
      }
    }
    
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
    'Contact Full Name',      // Optional (if exists in source)
    'First Name',             // Optional (only if Contact Full Name exists)
    'Last Name',              // Optional (only if Contact Full Name exists)
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
    'Company City'            // Optional
  ];
  
  // Check which name columns exist in source
  const hasContactFullName = foundOptional.includes('Contact Full Name');
  const hasFirstName = foundOptional.includes('First Name');
  const hasLastName = foundOptional.includes('Last Name');
  
  // We'll have First/Last Name in output if ANY name data exists
  const hasAnyNameData = hasContactFullName || hasFirstName || hasLastName;
  
  // Build output headers - only include optional columns that were found
  const outputHeaders = allPossibleColumns.filter(header => {
    // Always include Organization
    if (header === 'Organization') {
      return true;
    }
    // ALWAYS include First Name and Last Name in output if we have ANY name data
    // (They'll either come from source columns or be parsed from Contact Full Name)
    if (header === 'First Name' || header === 'Last Name') {
      return hasAnyNameData;
    }
    // Contact Full Name: Only include if it exists in source
    if (header === 'Contact Full Name') {
      return hasContactFullName;
    }
    // Include other optional columns only if they were found in source
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
// STEP 10: DUPLICATES REPORT
// =============================================================================

/**
 * Creates a "Duplicates Report" sheet showing all removed duplicates
 * with their original row numbers and what they duplicated
 */
function createDuplicatesReport(spreadsheet, duplicates) {
  const REPORT_SHEET_NAME = 'Duplicates Report';
  
  // Create or clear report sheet
  let reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = spreadsheet.insertSheet(REPORT_SHEET_NAME);
  }
  
  // Define headers
  const headers = [
    'Duplicate Row #',
    'Original Row #',
    'Contact Full Name',
    'Company Name',
    'Website',
    'Company Description',
    'Match Criteria',
    'Status'
  ];
  
  // Write headers
  reportSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#ea4335')  // Red background for duplicates
    .setFontColor('#ffffff');
  
  // If no duplicates found
  if (!duplicates || duplicates.length === 0) {
    reportSheet.getRange(2, 1, 1, headers.length)
      .setValues([['No duplicates found', '', '', '', '', '', '', '✅ All unique!']])
      .setBackground('#d9ead3')  // Light green
      .setFontStyle('italic');
    
    Logger.log('✓ Duplicates report created: No duplicates found');
    return reportSheet;
  }
  
  // Prepare duplicate data
  const reportData = duplicates.map(dup => [
    dup.duplicateRowNumber,
    dup.originalRowNumber,
    dup.fullName || '',
    dup.companyName || '',
    dup.website || '',
    dup.description || '',
    dup.matchCriteria,
    '🗑️ Removed'
  ]);
  
  // Write data in batches
  if (reportData.length > 0) {
    const batchSize = 1000;
    for (let i = 0; i < reportData.length; i += batchSize) {
      const batch = reportData.slice(i, i + batchSize);
      const startRow = i + 2;
      reportSheet.getRange(startRow, 1, batch.length, headers.length)
        .setValues(batch);
    }
  }
  
  // Format sheet
  reportSheet.setFrozenRows(1);
  
  // Alternate row colors for readability
  if (reportData.length > 0) {
    for (let i = 0; i < reportData.length; i++) {
      const rowNum = i + 2;
      const bgColor = i % 2 === 0 ? '#ffffff' : '#f3f3f3';
      reportSheet.getRange(rowNum, 1, 1, headers.length)
        .setBackground(bgColor);
    }
  }
  
  // Add summary at the bottom
  const summaryRow = reportData.length + 3;
  reportSheet.getRange(summaryRow, 1, 1, 2)
    .setValues([[`Total Duplicates Found:`, duplicates.length]])
    .setFontWeight('bold')
    .setBackground('#fff2cc');  // Light yellow
  
  Logger.log(`✓ Duplicates report created: ${duplicates.length} duplicates tracked`);
  
  return reportSheet;
}

// =============================================================================
// STEP 11: VALIDATION FORMULA
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

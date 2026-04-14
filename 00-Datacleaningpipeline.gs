/**
 * =============================================================================
 * DATA CLEANING PIPELINE V5
 * =============================================================================
 *
 * VERSION 5.0 CHANGES:
 * - Critical column changed from "Company Name - Cleaned" to "Company Name"
 * - New columns added to output: Address 1, Address 2, County, Zip,
 *   Gender, Race/Ethnicity, Industry, Company State
 * - Pre-run rename warning: detects common column name variants and tells
 *   you exactly which columns to rename before running
 * - Step 6.5: Gender normalized to lowercase (Male → male, Female → female)
 * - Step 6.5: Race/Ethnicity normalized to lowercase;
 *   "Asian/Pacific American" → "asian" (special case)
 * - Columns that need renaming:
 *     Contact Name    → Contact Full Name
 *     Contact Email   → Email 1
 *     Contact Phone   → Contact Mobile Phone
 *     City            → Company City
 *     State           → Company State
 * - Duplicate key: Contact Full Name + Company Name + Email 1
 * - Cross-row Email 1 safety net retained
 * - Company name normalization retained (strips punctuation, LLC/Inc/etc.)
 * - Phone normalization retained (digits only, strips leading US country code)
 * - Primary Email removed from all logic
 *
 * REQUIRED COLUMNS:
 *   - Company Name (CRITICAL)
 *
 * OPTIONAL COLUMNS (preserved in output if present):
 *   - Contact Full Name
 *   - Email 1, Email 2, Personal Email
 *   - Contact Mobile Phone, Contact Phone 1, Company Phone 1, Company Phone 2
 *   - Address 1, Address 2, Company City, Company State, County, Zip
 *   - Gender, Race/Ethnicity, Industry
 *   - Website, Company Description
 *
 * COLUMNS THAT NEED RENAMING BEFORE RUNNING:
 *   Contact Name    → Contact Full Name
 *   Contact Email   → Email 1
 *   Contact Phone   → Contact Mobile Phone
 *   City            → Company City
 *   State           → Company State
 *
 * @author: Claude
 * @version: 5.0
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const CONFIG = {
  OUTPUT_SHEET_NAME: 'Cleaned Data V3',

  COLUMNS: {
    FULL_NAME:       'Contact Full Name',
    FIRST_NAME:      'First Name',
    LAST_NAME:       'Last Name',
    ORGANIZATION:    'Company Name',
    EMAIL_1:         'Email 1',
    EMAIL_2:         'Email 2',
    PERSONAL_EMAIL:  'Personal Email',
    CONTACT_MOBILE:  'Contact Mobile Phone',
    CONTACT_PHONE_1: 'Contact Phone 1',
    COMPANY_PHONE_1: 'Company Phone 1',
    COMPANY_PHONE_2: 'Company Phone 2',
    ADDRESS_1:       'Address 1',
    ADDRESS_2:       'Address 2',
    COMPANY_CITY:    'Company City',
    COMPANY_STATE:   'Company State',
    COUNTY:          'County',
    ZIP:             'Zip',
    GENDER:          'Gender',
    RACE_ETHNICITY:  'Race/Ethnicity',
    INDUSTRY:        'Industry',
    WEBSITE:         'Website',
    COMPANY_DESC:    'Company Description'
  },

  // Columns the user might have that need to be renamed before running.
  // Key = what we detect in their sheet, Value = what it should be renamed to.
  RENAME_REQUIRED: {
    'Contact Name':  'Contact Full Name',
    'Contact Email': 'Email 1',
    'Contact Phone': 'Contact Mobile Phone',
    'City':          'Company City',
    'State':         'Company State'
  },

  BATCH_SIZE: 1000
};

// =============================================================================
// MAIN EXECUTION
// =============================================================================

function runDataCleaningPipelineV3() {
  try {
    const startTime = new Date();
    Logger.log('=== Starting Data Cleaning Pipeline V5 ===');

    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getActiveSheet();

    // Step 0: Check for columns that need renaming before anything else
    Logger.log('Step 0: Checking for columns that need renaming...');
    if (!checkForRenameRequired(sourceSheet)) {
      return; // User was shown the warning — stop here
    }

    Logger.log('Validating source sheet...');
    if (!validateSourceSheet(sourceSheet)) {
      throw new Error('Source sheet validation failed. Check required columns.');
    }
    Logger.log('✓ Validation passed');

    // Step 1: Load data
    Logger.log('Step 1: Loading data...');
    const rawData          = loadSheetData(sourceSheet);
    const originalRowCount = rawData.data.length;

    // Step 2: Remove duplicate rows
    Logger.log('Step 2: Removing duplicate rows...');
    const uniqueRowsData = removeDuplicateRows(rawData);
    const removedRows    = originalRowCount - uniqueRowsData.data.length;
    Logger.log(`Removed ${removedRows} duplicate rows`);

    // Step 3: Clean duplicate phone numbers within each row
    Logger.log('Step 3: Cleaning duplicate phones within rows...');
    const cleanedData = cleanEmailPhoneDuplicates(uniqueRowsData);

    // Step 4: Skipped (Primary Email removed)
    Logger.log('Step 4: Skipped (Primary Email removed from pipeline)');
    const emailFilledData = cleanedData;

    // Step 5: Deduplicate emails within each row
    Logger.log('Step 5: Deduplicating emails...');
    const emailCleanedData = deduplicateEmails(emailFilledData);

    // Step 6: Deduplicate and normalize phones with mobile priority
    Logger.log('Step 6: Deduplicating phones (Mobile priority)...');
    const phoneCleanedData = deduplicatePhonesWithPriority(emailCleanedData);

    // Step 6.5: Normalize Gender and Race/Ethnicity values
    Logger.log('Step 6.5: Normalizing Gender and Race/Ethnicity...');
    const normalizedData = normalizeFieldValues(phoneCleanedData);

    // Step 7: Parse names
    Logger.log('Step 7: Parsing names...');
    const finalData = parseNames(normalizedData);

    // Step 8: Create output sheet
    Logger.log('Step 8: Creating output sheet...');
    const outputSheet = createOutputSheet(ss, sourceSheet);

    // Step 9: Write cleaned data
    Logger.log('Step 9: Writing cleaned data...');
    writeOutputData(outputSheet, finalData, sourceSheet.getName());

    // Step 10: Create duplicates report
    Logger.log('Step 10: Creating duplicates report...');
    createDuplicatesReport(ss, uniqueRowsData.duplicates);

    const duration = ((new Date() - startTime) / 1000).toFixed(2);
    Logger.log(`=== Pipeline V5 Complete in ${duration}s ===`);

    const validation      = CONFIG._VALIDATION_RESULTS || {};
    const foundOptional   = validation.foundOptional   || [];
    const skippedOptional = validation.skippedOptional || [];

    let columnSummary = '';
    if (foundOptional.length > 0) {
      columnSummary += `\n✓ Processed columns:\n` + foundOptional.map(c => `  • ${c}`).join('\n');
    }
    if (skippedOptional.length > 0) {
      columnSummary += `\n\n⊗ Skipped (not found):\n` + skippedOptional.map(c => `  • ${c}`).join('\n');
    }

    SpreadsheetApp.getUi().alert(
      '✅ Data Cleaning Complete (V5)',
      `📊 Results:\n` +
      `  Original rows:          ${originalRowCount}\n` +
      `  Duplicate rows removed: ${removedRows}\n` +
      `  Final rows:             ${finalData.data.length}\n` +
      `  Duration:               ${duration}s\n\n` +
      `📄 Output Sheets:\n` +
      `  • "${CONFIG.OUTPUT_SHEET_NAME}" — Cleaned data\n` +
      `  • "Duplicates Report" — ${uniqueRowsData.duplicates.length} duplicates found\n` +
      columnSummary + `\n\n` +
      `✓ Duplicate Detection:\n` +
      `  • Key: Contact Full Name + Company Name + Email 1\n` +
      `  • Cross-row Email 1 safety net enabled\n` +
      `  • Company name punctuation-normalized before compare\n\n` +
      `✓ Phone Normalization:\n` +
      `  • All non-digits stripped for comparison\n` +
      `  • Leading US country code (1) stripped\n` +
      `  • Mobile phone has highest priority`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'Error',
      'Pipeline failed: ' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
}

// =============================================================================
// STEP 0: PRE-RUN RENAME CHECK
// =============================================================================

/**
 * Checks if any columns in the sheet need to be renamed before running.
 * Shows a clear warning with exact rename instructions if any are found.
 * Returns true if it's safe to continue, false if renaming is required.
 *
 * @param {Sheet} sheet
 * @returns {boolean} true = ok to proceed, false = stop and rename first
 */
function checkForRenameRequired(sheet) {
  const headers       = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headersNormal = headers.map(h => h.toString().trim());

  const detected = []; // [{ found, renameTo }]

  for (const [colName, renameTo] of Object.entries(CONFIG.RENAME_REQUIRED)) {
    if (headersNormal.includes(colName)) {
      detected.push({ found: colName, renameTo });
    }
  }

  if (detected.length === 0) {
    Logger.log('✓ No column renaming required');
    return true; // All good
  }

  // Build warning message
  const renameList = detected
    .map(d => `  • "${d.found}"  →  "${d.renameTo}"`)
    .join('\n');

  SpreadsheetApp.getUi().alert(
    '⚠️ Column Rename Required Before Running',
    'The following columns need to be renamed before running the pipeline.\n' +
    'Please rename them in your sheet, then run the pipeline again.\n\n' +
    'Columns to rename:\n' +
    renameList + '\n\n' +
    'Why? The pipeline uses standardized column names to process your data correctly.\n\n' +
    'How to rename:\n' +
    '  1. Click the column header cell\n' +
    '  2. Type the new name exactly as shown above\n' +
    '  3. Press Enter\n' +
    '  4. Run the pipeline again',
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  Logger.log(`Rename required: ${detected.map(d => `"${d.found}" → "${d.renameTo}"`).join(', ')}`);
  return false; // Stop — user must rename first
}

// =============================================================================
// STEP 1: DATA LOADING & VALIDATION
// =============================================================================

function validateSourceSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const columnRequirements = {
    critical: {
      'Company Name': 'Company Name'
    },
    optional: {
      'Contact Full Name':    'Contact Full Name',
      'First Name':           'First Name',
      'Last Name':            'Last Name',
      'Email 1':              'Email 1',
      'Email 2':              'Email 2',
      'Personal Email':       'Personal Email',
      'Contact Mobile Phone': 'Contact Mobile Phone',
      'Contact Phone 1':      'Contact Phone 1',
      'Company Phone 1':      'Company Phone 1',
      'Company Phone 2':      'Company Phone 2',
      'Address 1':            'Address 1',
      'Address 2':            'Address 2',
      'Company City':         'Company City',
      'Company State':        'Company State',
      'County':               'County',
      'Zip':                  'Zip',
      'Gender':               'Gender',
      'Race/Ethnicity':       'Race/Ethnicity',
      'Industry':             'Industry',
      'Website':              'Website',
      'Company Description':  'Company Description'
    }
  };

  const missingCritical = [];
  const foundColumns    = {};

  // Normalize headers once for all comparisons
  const normalizedHeaders = headers.map(h => h.toString().trim().toLowerCase());

  for (const [colName, exactName] of Object.entries(columnRequirements.critical)) {
    const matchIdx = normalizedHeaders.indexOf(exactName.toLowerCase());
    if (matchIdx !== -1) {
      foundColumns[colName] = headers[matchIdx].toString().trim();
    } else {
      missingCritical.push(exactName);
    }
  }

  if (missingCritical.length > 0) {
    SpreadsheetApp.getUi().alert(
      '⚠️ Critical Columns Missing',
      '⚠️ CRITICAL COLUMNS MISSING ⚠️\n\n' +
      'The following columns are required with EXACT names:\n' +
      missingCritical.map(c => `  • "${c}"`).join('\n') + '\n\n' +
      'Available headers in your sheet:\n' +
      headers.map(h => h.toString().trim()).filter(Boolean).join(', '),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }

  const foundOptional   = [];
  const skippedOptional = [];

  // Normalize headers for case-insensitive, trimmed matching
  const headersNormalized = headers.map(h => h.toString().trim().toLowerCase());

  for (const [colName, exactName] of Object.entries(columnRequirements.optional)) {
    const matchIdx = headersNormalized.indexOf(exactName.toLowerCase());
    if (matchIdx !== -1) {
      // Use the actual header value from the sheet (preserves original casing)
      const actualName = headers[matchIdx].toString().trim();
      foundColumns[colName] = actualName;
      foundOptional.push(actualName);
    } else {
      skippedOptional.push(exactName);
    }
  }

  CONFIG._VALIDATION_RESULTS = { foundOptional, skippedOptional, foundColumns };
  Logger.log(`✅ Validation: ${foundOptional.length} optional columns found, ${skippedOptional.length} skipped`);
  return true;
}

function loadSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('No data rows found in sheet');

  const allData   = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers   = allData[0];
  const dataRows  = allData.slice(1);
  const columnMap = {};

  if (CONFIG._VALIDATION_RESULTS && CONFIG._VALIDATION_RESULTS.foundColumns) {
    for (const [, actualName] of Object.entries(CONFIG._VALIDATION_RESULTS.foundColumns)) {
      const index = headers.indexOf(actualName);
      if (index !== -1) columnMap[actualName] = index;
    }
  } else {
    headers.forEach((header, index) => { columnMap[header] = index; });
  }

  Logger.log(`Column map: ${Object.keys(columnMap).length} columns`);
  return { headers, columnMap, data: dataRows };
}

// =============================================================================
// STEP 2: ROW-LEVEL DUPLICATE REMOVAL
// =============================================================================

/**
 * Normalizes a company name for comparison.
 * Strips punctuation and common legal suffixes so
 * "ABC Builders LLC" and "ABC Builders" match as the same company.
 */
function normalizeCompanyName(name) {
  if (!name) return '';
  return name.toString()
    .toLowerCase()
    .trim()
    .replace(/[.,\-\/\\&]/g, ' ')
    .replace(/\b(llc|inc|corp|ltd|co|company|group|associates|services|solutions)\b/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * V5: Duplicate key is Contact Full Name + Company Name + Email 1.
 * Falls back to simpler keys when fields are missing.
 * Followed by a cross-row Email 1 safety net pass.
 */
function removeDuplicateRows(dataObject) {
  const { headers, columnMap, data } = dataObject;

  const fullNameIdx    = columnMap[CONFIG.COLUMNS.FULL_NAME];
  const companyNameIdx = columnMap[CONFIG.COLUMNS.ORGANIZATION];
  const email1Idx      = columnMap[CONFIG.COLUMNS.EMAIL_1];
  const descriptionIdx = columnMap[CONFIG.COLUMNS.COMPANY_DESC];
  const websiteIdx     = columnMap[CONFIG.COLUMNS.WEBSITE];

  const seenCombinations = new Map();
  const uniqueRows       = [];
  const duplicates       = [];

  data.forEach((row, index) => {
    const nameForKey     = fullNameIdx    !== undefined ? (row[fullNameIdx]    || '').toString().trim()           : '';
    const companyRaw     = companyNameIdx !== undefined ? (row[companyNameIdx] || '').toString().trim()           : '';
    const email1Key      = email1Idx      !== undefined ? (row[email1Idx]      || '').toString().trim().toLowerCase() : '';
    const descriptionKey = descriptionIdx !== undefined ? (row[descriptionIdx] || '').toString().trim().toLowerCase() : '';
    const websiteKey     = websiteIdx     !== undefined ? (row[websiteIdx]     || '').toString().trim().toLowerCase() : '';

    const nameKey    = nameForKey.toLowerCase();
    const companyKey = normalizeCompanyName(companyRaw);

    // Build composite key — most specific first
    let compositeKey, matchCriteria;

    if (nameKey && companyKey && email1Key) {
      compositeKey  = `${nameKey}|${companyKey}|${email1Key}`;
      matchCriteria = 'Contact Name + Company + Email 1';
    } else if (nameKey && companyKey) {
      compositeKey  = `${nameKey}|${companyKey}`;
      matchCriteria = 'Contact Name + Company';
    } else if (companyKey && email1Key) {
      compositeKey  = `${companyKey}|${email1Key}`;
      matchCriteria = 'Company + Email 1';
    } else if (companyKey && websiteKey) {
      compositeKey  = `${companyKey}|${websiteKey}`;
      matchCriteria = 'Company + Website';
    } else if (companyKey && descriptionKey) {
      compositeKey  = `${companyKey}|${descriptionKey}`;
      matchCriteria = 'Company + Description';
    } else if (companyKey) {
      compositeKey  = companyKey;
      matchCriteria = 'Company Name Only';
    } else {
      compositeKey  = `${nameKey}|${companyKey}|${email1Key}`;
      matchCriteria = 'All Available Fields';
    }

    if (seenCombinations.has(compositeKey)) {
      duplicates.push({
        duplicateRowNumber: index + 2,
        originalRowNumber:  seenCombinations.get(compositeKey),
        fullName:           nameForKey,
        companyName:        companyRaw,
        email1:             email1Key,
        website:            websiteKey,
        description:        descriptionKey,
        matchCriteria,
        rowData: row
      });
    } else {
      seenCombinations.set(compositeKey, index + 2);
      uniqueRows.push(row);
    }
  });

  // ── CROSS-ROW EMAIL SAFETY NET ──────────────────────────────────────────────
  // After composite key dedup, flag any remaining rows that share Email 1.
  const uniqueAfterEmailCheck = [];
  const seenEmails            = new Map();

  uniqueRows.forEach((row, index) => {
    const email1 = email1Idx !== undefined
      ? (row[email1Idx] || '').toString().trim().toLowerCase()
      : '';

    if (!email1) {
      uniqueAfterEmailCheck.push(row);
      return;
    }

    if (seenEmails.has(email1)) {
      const nameForKey = fullNameIdx !== undefined
        ? (row[fullNameIdx] || '').toString().trim()
        : '';
      const companyRaw = companyNameIdx !== undefined
        ? (row[companyNameIdx] || '').toString().trim()
        : '';
      duplicates.push({
        duplicateRowNumber: index + 2,
        originalRowNumber:  seenEmails.get(email1),
        fullName:           nameForKey,
        companyName:        companyRaw,
        email1,
        website:            '',
        description:        '',
        matchCriteria:      'PROBABLE DUPLICATE — same Email 1',
        rowData:            row
      });
    } else {
      seenEmails.set(email1, index + 2);
      uniqueAfterEmailCheck.push(row);
    }
  });
  // ── END EMAIL SAFETY NET ────────────────────────────────────────────────────

  Logger.log(`Composite key duplicates: ${data.length - uniqueRows.length}`);
  Logger.log(`Email safety net duplicates: ${uniqueRows.length - uniqueAfterEmailCheck.length}`);
  Logger.log(`Total removed: ${data.length - uniqueAfterEmailCheck.length}`);

  return { headers, columnMap, data: uniqueAfterEmailCheck, duplicates };
}

// =============================================================================
// STEP 3: COLUMN-LEVEL PHONE DUPLICATE CLEANING
// =============================================================================

function cleanEmailPhoneDuplicates(dataObject) {
  const { headers, columnMap, data } = dataObject;

  const columnsToClean = [
    CONFIG.COLUMNS.CONTACT_MOBILE,
    CONFIG.COLUMNS.CONTACT_PHONE_1,
    CONFIG.COLUMNS.COMPANY_PHONE_1,
    CONFIG.COLUMNS.COMPANY_PHONE_2
  ];

  const cleanedData = data.map(row => {
    const newRow     = [...row];
    const seenDigits = new Set();

    columnsToClean
      .map(col => columnMap[col])
      .filter(idx => idx !== undefined)
      .forEach(colIdx => {
        const value = row[colIdx];
        if (!value || value.toString().trim() === '') return;
        const digits = normalizePhoneDigits(value.toString());
        if (!digits) return;
        if (seenDigits.has(digits)) {
          newRow[colIdx] = '';
        } else {
          seenDigits.add(digits);
          newRow[colIdx] = value.toString().trim();
        }
      });

    return newRow;
  });

  Logger.log('✓ Phone column deduplication complete');
  return { headers, columnMap, data: cleanedData };
}

// =============================================================================
// STEP 5: EMAIL DEDUPLICATION
// =============================================================================

function deduplicateEmails(dataObject) {
  const { headers, columnMap, data } = dataObject;

  const email1Idx   = columnMap[CONFIG.COLUMNS.EMAIL_1];
  const email2Idx   = columnMap[CONFIG.COLUMNS.EMAIL_2];
  const personalIdx = columnMap[CONFIG.COLUMNS.PERSONAL_EMAIL];

  if (email1Idx === undefined && email2Idx === undefined && personalIdx === undefined) {
    Logger.log('⊗ Skipping email deduplication: no email columns found');
    return dataObject;
  }

  const cleanedData = data.map(row => {
    const newRow   = [...row];
    const email1   = email1Idx   !== undefined ? normalizeEmail(row[email1Idx])   : '';
    const email2   = email2Idx   !== undefined ? normalizeEmail(row[email2Idx])   : '';
    const personal = personalIdx !== undefined ? normalizeEmail(row[personalIdx]) : '';

    let cleanEmail2   = email2;
    let cleanPersonal = personal;
    if (email1 && email1 === email2)    cleanEmail2   = '';
    if (email1 && email1 === personal)  cleanPersonal = '';

    if (email1Idx   !== undefined) newRow[email1Idx]   = email1;
    if (email2Idx   !== undefined) newRow[email2Idx]   = cleanEmail2;
    if (personalIdx !== undefined) newRow[personalIdx] = cleanPersonal;

    return newRow;
  });

  Logger.log('✓ Email deduplication complete');
  return { headers, columnMap, data: cleanedData };
}

function normalizeEmail(email) {
  if (!email || email.toString().trim() === '') return '';
  return email.toString().trim().toLowerCase();
}

// =============================================================================
// STEP 6: PHONE DEDUPLICATION WITH MOBILE PRIORITY
// =============================================================================

/**
 * Strips all non-digit characters and removes the leading US country code "1"
 * so that 15551234567 and 5551234567 compare as equal.
 */
function normalizePhoneDigits(phone) {
  if (!phone) return '';
  let digits = phone.toString().replace(/\D/g, '');
  if (digits.length === 11 && digits.startsWith('1')) digits = digits.substring(1);
  return digits;
}

function normalizePhone(phone) {
  if (!phone || phone.toString().trim() === '') return '';
  return phone.toString().trim().replace(/\./g, '');
}

function deduplicatePhonesWithPriority(dataObject) {
  const { headers, columnMap, data } = dataObject;

  const contactMobileIdx = columnMap[CONFIG.COLUMNS.CONTACT_MOBILE];
  const contactPhone1Idx = columnMap[CONFIG.COLUMNS.CONTACT_PHONE_1];
  const companyPhone1Idx = columnMap[CONFIG.COLUMNS.COMPANY_PHONE_1];
  const companyPhone2Idx = columnMap[CONFIG.COLUMNS.COMPANY_PHONE_2];

  if (contactMobileIdx === undefined && contactPhone1Idx === undefined &&
      companyPhone1Idx === undefined && companyPhone2Idx === undefined) {
    Logger.log('⊗ Skipping phone deduplication: no phone columns found');
    return dataObject;
  }

  const cleanedData = data.map(row => {
    const newRow = [...row];

    let mobile   = contactMobileIdx !== undefined ? normalizePhone(row[contactMobileIdx]) : '';
    let contact1 = contactPhone1Idx !== undefined ? normalizePhone(row[contactPhone1Idx]) : '';
    let company1 = companyPhone1Idx !== undefined ? normalizePhone(row[companyPhone1Idx]) : '';
    let company2 = companyPhone2Idx !== undefined ? normalizePhone(row[companyPhone2Idx]) : '';

    const seenDigits   = new Set();
    const mobileDigits = normalizePhoneDigits(mobile);
    if (mobileDigits) seenDigits.add(mobileDigits);

    const c1d = normalizePhoneDigits(contact1);
    if (c1d)  { if (seenDigits.has(c1d))  contact1 = ''; else seenDigits.add(c1d); }
    const cp1d = normalizePhoneDigits(company1);
    if (cp1d) { if (seenDigits.has(cp1d)) company1 = ''; else seenDigits.add(cp1d); }
    const cp2d = normalizePhoneDigits(company2);
    if (cp2d) { if (seenDigits.has(cp2d)) company2 = ''; else seenDigits.add(cp2d); }

    if (contactMobileIdx !== undefined) newRow[contactMobileIdx] = mobile;
    if (contactPhone1Idx !== undefined) newRow[contactPhone1Idx] = contact1;
    if (companyPhone1Idx !== undefined) newRow[companyPhone1Idx] = company1;
    if (companyPhone2Idx !== undefined) newRow[companyPhone2Idx] = company2;

    return newRow;
  });

  Logger.log('✓ Phone deduplication complete');
  return { headers, columnMap, data: cleanedData };
}


// =============================================================================
// STEP 6.5: NORMALIZE GENDER AND RACE/ETHNICITY
// =============================================================================

/**
 * Normalizes Gender and Race/Ethnicity values:
 *
 * Gender:
 *   - Lowercases the value as-is  (Male → male, Female → female)
 *
 * Race/Ethnicity:
 *   - "Asian/Pacific American" → "asian"  (special case — take first part)
 *   - Everything else → lowercased as-is
 */
function normalizeFieldValues(dataObject) {
  const { headers, columnMap, data } = dataObject;

  const genderIdx    = columnMap[CONFIG.COLUMNS.GENDER];
  const ethnicityIdx = columnMap[CONFIG.COLUMNS.RACE_ETHNICITY];

  if (genderIdx === undefined && ethnicityIdx === undefined) {
    Logger.log('⊗ Skipping field normalization: no Gender or Race/Ethnicity columns found');
    return dataObject;
  }

  const cleanedData = data.map(row => {
    const newRow = [...row];

    // Normalize Gender — lowercase only
    if (genderIdx !== undefined) {
      const val = (row[genderIdx] || '').toString().trim();
      newRow[genderIdx] = val.toLowerCase();
    }

    // Normalize Race/Ethnicity
    if (ethnicityIdx !== undefined) {
      const val = (row[ethnicityIdx] || '').toString().trim();
      if (!val) {
        newRow[ethnicityIdx] = '';
      } else if (val.toLowerCase().includes('asian/pacific')) {
        // Special case: "Asian/Pacific American" → "asian"
        newRow[ethnicityIdx] = 'asian';
      } else {
        // Everything else — just lowercase
        newRow[ethnicityIdx] = val.toLowerCase();
      }
    }

    return newRow;
  });

  Logger.log('✓ Gender and Race/Ethnicity normalization complete');
  return { headers, columnMap, data: cleanedData };
}

// =============================================================================
// STEP 7: NAME PARSING
// =============================================================================

function parseNames(dataObject) {
  const { headers, columnMap, data } = dataObject;

  const fullNameIdx  = columnMap[CONFIG.COLUMNS.FULL_NAME];
  const firstNameIdx = columnMap[CONFIG.COLUMNS.FIRST_NAME];
  const lastNameIdx  = columnMap[CONFIG.COLUMNS.LAST_NAME];

  const hasFirstOrLast = firstNameIdx !== undefined || lastNameIdx !== undefined;

  Logger.log(`parseNames — fullNameIdx: ${fullNameIdx}, firstNameIdx: ${firstNameIdx}, lastNameIdx: ${lastNameIdx}`);

  const parsedData = data.map(row => {
    let firstName = '';
    let lastName  = '';

    if (hasFirstOrLast) {
      // Priority 1: Use First Name / Last Name columns directly
      firstName = firstNameIdx !== undefined && row[firstNameIdx]
        ? row[firstNameIdx].toString().trim() : '';
      lastName  = lastNameIdx  !== undefined && row[lastNameIdx]
        ? row[lastNameIdx].toString().trim()  : '';
    } else if (fullNameIdx !== undefined && row[fullNameIdx]) {
      // Priority 2: Parse Contact Full Name into first + last
      const parsed = parseFullName(row[fullNameIdx].toString().trim());
      firstName = parsed.firstName;
      lastName  = parsed.lastName;
    }

    return { originalRow: row, firstName, lastName };
  });

  Logger.log('✓ Name parsing complete');
  return { headers, columnMap, data: parsedData };
}

function parseFullName(fullName) {
  if (!fullName || fullName.trim() === '') return { firstName: '', lastName: '' };
  const parts = fullName.trim().split(/\s+/).filter(p => p.length > 0);
  if (parts.length === 0) return { firstName: '', lastName: '' };
  if (parts.length === 1) return { firstName: parts[0], lastName: '' };
  // First word = first name, rest = last name
  return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
}

// =============================================================================
// STEP 8: OUTPUT SHEET CREATION
// =============================================================================

function createOutputSheet(spreadsheet, sourceSheet) {
  let outputSheet = spreadsheet.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  if (outputSheet) {
    outputSheet.clear();
  } else {
    outputSheet = spreadsheet.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
  }
  return outputSheet;
}

// =============================================================================
// STEP 9: WRITE OUTPUT DATA
// =============================================================================

function writeOutputData(outputSheet, dataObject, sourceSheetName) {
  const { columnMap, data } = dataObject;

  const validation    = CONFIG._VALIDATION_RESULTS || {};
  const foundOptional = validation.foundOptional || [];

  // All possible output columns in desired order
  // Contact Full Name → First Name + Last Name split in output
  const allPossibleColumns = [
    'Contact Full Name',
    'First Name',
    'Last Name',
    'Organization',
    'Email 1',
    'Email 2',
    'Personal Email',
    'Contact Mobile Phone',
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2',
    'Address 1',
    'Address 2',
    'Company City',
    'Company State',
    'County',
    'Zip',
    'Gender',
    'Race/Ethnicity',
    'Industry',
    'Website',
    'Company Description'
  ];

  const hasAnyNameData = foundOptional.includes('Contact Full Name') ||
                         foundOptional.includes('First Name') ||
                         foundOptional.includes('Last Name');

  // Only include columns that exist in the source sheet
  const hasContactFullName = foundOptional.includes('Contact Full Name');

  const outputHeaders = allPossibleColumns.filter(header => {
    if (header === 'Organization')     return true;      // always required
    if (header === 'Contact Full Name') return hasContactFullName;
    // First Name and Last Name: always include when ANY name data exists
    // (they come from parseNames() even if not separate columns in source)
    if (header === 'First Name' || header === 'Last Name') return hasAnyNameData;
    return foundOptional.includes(header);
  });

  // Header mapping: output column name → CONFIG column key name
  const headerToSourceCol = {
    'Contact Full Name':     CONFIG.COLUMNS.FULL_NAME,
    'Organization':         CONFIG.COLUMNS.ORGANIZATION,
    'Email 1':              CONFIG.COLUMNS.EMAIL_1,
    'Email 2':              CONFIG.COLUMNS.EMAIL_2,
    'Personal Email':       CONFIG.COLUMNS.PERSONAL_EMAIL,
    'Contact Mobile Phone': CONFIG.COLUMNS.CONTACT_MOBILE,
    'Contact Phone 1':      CONFIG.COLUMNS.CONTACT_PHONE_1,
    'Company Phone 1':      CONFIG.COLUMNS.COMPANY_PHONE_1,
    'Company Phone 2':      CONFIG.COLUMNS.COMPANY_PHONE_2,
    'Address 1':            CONFIG.COLUMNS.ADDRESS_1,
    'Address 2':            CONFIG.COLUMNS.ADDRESS_2,
    'Company City':         CONFIG.COLUMNS.COMPANY_CITY,
    'Company State':        CONFIG.COLUMNS.COMPANY_STATE,
    'County':               CONFIG.COLUMNS.COUNTY,
    'Zip':                  CONFIG.COLUMNS.ZIP,
    'Gender':               CONFIG.COLUMNS.GENDER,
    'Race/Ethnicity':       CONFIG.COLUMNS.RACE_ETHNICITY,
    'Industry':             CONFIG.COLUMNS.INDUSTRY,
    'Website':              CONFIG.COLUMNS.WEBSITE,
    'Company Description':  CONFIG.COLUMNS.COMPANY_DESC
  };

  // Write header row
  outputSheet.getRange(1, 1, 1, outputHeaders.length)
    .setValues([outputHeaders])
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  // Build data rows
  const outputData = data.map(item => {
    const row = item.originalRow;
    return outputHeaders.map(header => {
      if (header === 'First Name') return item.firstName;
      if (header === 'Last Name')  return item.lastName;
      const sourceColName = headerToSourceCol[header];
      const colIdx        = sourceColName ? columnMap[sourceColName] : undefined;
      return colIdx !== undefined ? row[colIdx] : '';
    });
  });

  // Write in batches
  if (outputData.length > 0) {
    for (let i = 0; i < outputData.length; i += CONFIG.BATCH_SIZE) {
      const batch = outputData.slice(i, i + CONFIG.BATCH_SIZE);
      outputSheet.getRange(i + 2, 1, batch.length, outputHeaders.length).setValues(batch);
    }
  }

  outputSheet.setFrozenRows(1);
  Logger.log(`✓ Written ${outputData.length} rows with ${outputHeaders.length} columns`);
}

// =============================================================================
// STEP 10: DUPLICATES REPORT
// =============================================================================

function createDuplicatesReport(spreadsheet, duplicates) {
  const REPORT_SHEET_NAME = 'Duplicates Report';

  let reportSheet = spreadsheet.getSheetByName(REPORT_SHEET_NAME);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = spreadsheet.insertSheet(REPORT_SHEET_NAME);
  }

  const headers = [
    'Duplicate Row #',
    'Original Row #',
    'Contact Full Name',
    'Company Name',
    'Email 1',
    'Match Criteria',
    'Status'
  ];

  reportSheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('#ffffff');

  if (!duplicates || duplicates.length === 0) {
    reportSheet.getRange(2, 1, 1, headers.length)
      .setValues([['No duplicates found', '', '', '', '', '', '✅ All unique!']])
      .setBackground('#d9ead3')
      .setFontStyle('italic');
    Logger.log('✓ Duplicates report: no duplicates found');
    return reportSheet;
  }

  const reportData = duplicates.map(dup => [
    dup.duplicateRowNumber,
    dup.originalRowNumber,
    dup.fullName    || '',
    dup.companyName || '',
    dup.email1      || '',
    dup.matchCriteria,
    '🗑️ Removed'
  ]);

  for (let i = 0; i < reportData.length; i += 1000) {
    const batch = reportData.slice(i, i + 1000);
    reportSheet.getRange(i + 2, 1, batch.length, headers.length).setValues(batch);
  }

  reportSheet.setFrozenRows(1);

  // Alternate row shading
  for (let i = 0; i < reportData.length; i++) {
    reportSheet.getRange(i + 2, 1, 1, headers.length)
      .setBackground(i % 2 === 0 ? '#ffffff' : '#f3f3f3');
  }

  // Summary row
  const summaryRow = reportData.length + 3;
  reportSheet.getRange(summaryRow, 1, 1, 2)
    .setValues([[`Total Duplicates Found:`, duplicates.length]])
    .setFontWeight('bold')
    .setBackground('#fff2cc');

  Logger.log(`✓ Duplicates report: ${duplicates.length} duplicates`);
  return reportSheet;
}

// =============================================================================
// UTILITY FUNCTIONS
// =============================================================================

function showLogs() {
  const logs = Logger.getLog();
  SpreadsheetApp.getUi().alert(
    'Execution Logs',
    logs || 'No logs available',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =============================================================================
// MENU
// =============================================================================

/**
 * Creates ONE unified menu in the Google Sheets UI
 *
 * MENU STRUCTURE:
 * 🔧 Data Tools
 * ├── 📊 View Pipeline Dashboard
 * ├── 🧹 Data Cleanup
 * │   ├── ▶️ Run Pipeline V5
 * │   ├── 🚫 Exclude Companies
 * │   ├── 📊 Audit Report
 * │   └── 📋 View Logs
 * ├── 🏠 Categorize Companies
 * │   ├── Categorize All Companies
 * │   └── Upload Checked Entries
 * ├── 📧 Email Validation
 * │   ├── Validate Email 1
 * │   ├── Validate Email 2
 * │   ├── Validate Personal Email
 * │   ├── ─────────────────────────
 * │   ├── 🛑 Cancel Running Validation
 * │   ├── ─────────────────────────
 * │   ├── ⚙️ Configure Settings
 * │   └── 📋 View Validation Logs
 * ├── 📞 Phone Validation
 * │   ├── ▶️ Validate All (Email + Phones)
 * │   ├── ─────────────────────────
 * │   ├── Validate Contact Mobile Phone
 * │   ├── Validate Contact Phone 1
 * │   ├── Validate Company Phone 1
 * │   ├── Validate Company Phone 2
 * │   ├── ─────────────────────────
 * │   ├── 🛑 Cancel Running Validation
 * │   ├── ─────────────────────────
 * │   ├── ⚙️ Configure Settings
 * │   └── 📋 View Validation Logs
 * └── 🎯 CRM Ready
 *     ├── 📞 Generate CRM Callers Ready
 *     ├── ─────────────────────────
 *     ├── 📊 Compare with CRM Ready
 *     ├── ℹ️ Comparison Info
 *     ├── ─────────────────────────
 *     └── ⚙️ View Settings
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🔧 Data Tools')

    .addItem('📊 View Pipeline Dashboard', 'createPipelineDashboard')
    .addSeparator()

    // Data Cleanup
    .addSubMenu(ui.createMenu('🧹 Data Cleanup')
      .addItem('▶️ Run Pipeline V5',    'runDataCleaningPipelineV3')
      .addItem('🚫 Exclude Companies',  'excludeCompanies')
      .addItem('📊 Audit Report',       'generateAuditReport')
      .addSeparator()
      .addItem('📋 View Logs',          'showLogs'))

    // Categorize Companies
    .addSubMenu(ui.createMenu('🏠 Categorize Companies')
      .addItem('▶️ Categorize All Companies', 'categorizeAllCompanies')
      .addSeparator()
      .addItem('📤 Upload Checked Entries', 'uploadCheckedEntriesCat'))

    // Email Validation
    .addSubMenu(ui.createMenu('📧 Email Validation')
      .addItem('📧 Validate Email 1',         'validateEmail1')
      .addItem('📧 Validate Email 2',          'validateEmail2')
      .addItem('📨 Validate Personal Email',   'validatePersonalEmail')
      .addSeparator()
      .addItem('🛑 Cancel Running Validation', 'cancelEmailValidation')
      .addSeparator()
      .addItem('⚙️ Configure Settings',        'showEmailConfigDialog')
      .addItem('📋 View Validation Logs',       'showEmailValidationLogs'))

    // Phone Validation
    .addSubMenu(ui.createMenu('📞 Phone Validation')
      .addItem('▶️ Validate All (Email + Phones)', 'validateAll')
      .addSeparator()
      .addItem('📱 Validate Contact Mobile Phone',  'validateContactMobilePhone')
      .addItem('📱 Validate Contact Phone 1',       'validateContactPhone1')
      .addItem('🏢 Validate Company Phone 1',       'validateCompanyPhone1')
      .addItem('🏢 Validate Company Phone 2',       'validateCompanyPhone2')
      .addSeparator()
      .addItem('🛑 Cancel Running Validation',      'cancelPhoneValidation')
      .addSeparator()
      .addItem('⚙️ Configure Settings',             'showConfigDialog')
      .addItem('📋 View Validation Logs',            'showValidationLogs'))

    // CRM Ready
    .addSubMenu(ui.createMenu('🎯 CRM Ready')
      .addItem('📞 Generate CRM Callers Ready',  'generateCRMCallersReady')
      .addItem('📋 Generate CRM Ready',          'generateCRMReady')
      .addSeparator()
      .addItem('📊 Compare with CRM Ready',      'compareWithCRMReady')
      .addItem('ℹ️ Comparison Info',             'showComparisonInfo')
      .addSeparator()
      .addItem('⚙️ View Settings',               'showCRMConfig'))

    .addToUi();
}

// =============================================================================
// END OF SCRIPT V5
// =============================================================================

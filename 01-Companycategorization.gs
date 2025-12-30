/**
 * HOME BUILDER CATEGORIZATION SYSTEM - V7.1
 * ==========================================
 * IMPROVEMENTS IN V7.1:
 * - Dynamic output columns: Only includes columns that exist in source sheet
 * - No more empty/blank columns in output (Contact Full Name, etc.)
 * - Cleaner output sheets with only relevant data
 * 
 * IMPROVEMENTS IN V7.0:
 * - Works on the ACTIVE spreadsheet (no hardcoded ID)
 * - Integrated with unified menu system
 * - Simplified to 3 categories: Home Builders & Related, Other Companies, No Description
 * - Added restoration/remodeling keywords (categorized as Home Builders)
 * - No popup messages - runs silently
 * - Excludes pool/fence companies even if they mention building
 * - Better keyword-based analysis
 * 
 * CATEGORIES:
 * 1. Home Builders & Related (includes remodeling, restoration, apartments, condos)
 * 2. Other Companies (pools, fences, roofing, everything else)
 * 3. No Description
 * 
 * @author: Claude
 * @version: 7.1 - Dynamic output columns (only includes columns that exist)
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CATEGORIZATION_CONFIG = {
  // HEADER NAMES TO FIND (case-insensitive)
  HEADER_NAMES: {
    contactFullName: 'Contact Full Name',
    firstName: 'First Name',
    lastName: 'Last Name',
    companyName: 'Organization',
    description: 'Company Description',
    primaryEmail: 'Primary Email',
    email1: 'Email 1',
    email2: 'Email 2',
    personalEmail: 'Personal Email',
    contactPhone1: 'Contact Phone 1',
    companyPhone1: 'Company Phone 1',
    companyPhone2: 'Company Phone 2',
    contactMobilePhone: 'Contact Mobile Phone',
    website: 'Website',
    city: 'Company City'
  },
  
  // Output columns will be built dynamically from source sheet
  // Only columns that exist in source will be included in output
  
  // Starting row
  FIRST_DATA_ROW: 2,
  
  // Fuzzy matching threshold
  FUZZY_THRESHOLD: 0.80,
  
  // Keywords for companies with no description
  NO_DESC_KEYWORDS: [
    ['builder', 3.0],
    ['builders', 3.0],
    ['construction', 2.5],
    ['building', 2.0],
    ['homes', 3.5],
    ['home', 3.5],
    ['residential', 2.5],
    ['developer', 2.0],
    ['development', 2.0],
    ['contractor', 1.8],
    ['contracting', 1.8],
    ['remodeling', 3.0],
    ['restoration', 3.0],
  ],
  
  // EXCLUSION KEYWORDS - Pool/Fence companies (categorize as OTHER)
  POOL_KEYWORDS: [
    'pool builder',
    'pool construction',
    'swimming pool',
    'pool installation',
    'pool company',
    'inground pool',
    'custom pool',
    'pool design',
    'pool service',
    'pool contractor',
  ],
  
  FENCE_KEYWORDS: [
    'fence builder',
    'fence construction',
    'fence installation',
    'fence company',
    'fencing contractor',
    'fence design',
    'fence service',
    'custom fence',
  ],
  
  // OTHER TRADES KEYWORDS - Categorize as OTHER
  OTHER_TRADES_KEYWORDS: [
    'roofing contractor',
    'roofing company',
    'plumbing contractor',
    'hvac contractor',
    'electrical contractor',
    'landscaping company',
    'flooring company',
    'painting contractor',
    'carpet installation',
    'tile contractor',
    'cabinet maker',
    'window installation',
    'garage door',
    'concrete contractor',
    'asphalt paving',
    'excavation',
    'tree service',
  ],
  
  // HOME BUILDER KEYWORDS - Strong indicators
  HOME_BUILDER_KEYWORDS: [
    // Home building - very strong
    'home builder',
    'homebuilder',
    'house builder',
    'custom homes',
    'new homes',
    'production builder',
    'tract builder',
    'spec homes',
    
    // Residential development
    'residential developer',
    'residential development',
    'home development',
    'housing development',
    'residential construction',
    'residential building',
    
    // Single/multi family
    'single family',
    'single-family',
    'multifamily',
    'multi-family',
    'multi family',
    
    // Apartment/Condo
    'apartment building',
    'apartment developer',
    'apartment construction',
    'apartment complex',
    'condominium',
    'condo developer',
    'condo construction',
    'townhome',
    'townhouse',
    
    // RESTORATION & REMODELING
    'restoration',
    'restore',
    'remodeling',
    'remodel',
    'renovation',
    'renovate',
    'home improvement',
    'home renovation',
    'house renovation',
    'kitchen remodel',
    'bathroom remodel',
    'whole house remodel',
    'historic restoration',
    'home restoration',
    'fire restoration',
    'water restoration',
    'storm restoration',
    'disaster restoration',
    'property restoration',
    'residential remodeling',
    'residential renovation',
    'addition',
    'home addition',
    
    // General building terms
    'subdivision',
    'master-planned',
    'residential',
    'general contractor',
  ],
  
  // Generic building terms (lower weight)
  GENERIC_BUILDING_KEYWORDS: [
    'builder',
    'developer',
    'construction',
    'building',
    'contractor',
    'contracting',
  ],
};

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Main categorization function - works on ACTIVE spreadsheet
 */
function categorizeAllCompanies() {
  const startTime = new Date();
  
  // Get ACTIVE spreadsheet (not hardcoded)
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
  
  Logger.log('=== STARTING CATEGORIZATION V7.0 ===');
  Logger.log(`Spreadsheet: ${spreadsheet.getName()}`);
  Logger.log(`Active Sheet: ${sourceSheet.getName()}`);
  
  // Find columns
  const columnMap = findColumnsByHeadersCat(sourceSheet);
  if (!columnMap) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Could not find required headers in the active sheet.\n\n' +
      'Required headers:\n' +
      '- Organization (or Company Name)\n' +
      '- Company Description\n\n' +
      'These are the minimum columns needed for categorization.\n' +
      'Other columns are optional but recommended.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  // Show progress
  SpreadsheetApp.getUi().alert(
    'Categorization Started',
    'Processing companies... This may take a few minutes.\n\n' +
    'Check the script logs for progress.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  
  // Process companies
  const results = processAllCompaniesCat(sourceSheet, columnMap);
  
  // Create category sheets
  const categoryStats = createCategorySheetsV7(spreadsheet, sourceSheet, results, columnMap);
  
  const endTime = new Date();
  const duration = Math.round((endTime - startTime) / 1000);
  
  // Log summary
  Logger.log('=== CATEGORIZATION COMPLETE ===');
  Logger.log(`Total Companies: ${results.length}`);
  Logger.log(`Time: ${duration} seconds`);
  Logger.log('CATEGORY BREAKDOWN:');
  
  Object.entries(categoryStats).forEach(([category, count]) => {
    const percentage = Math.round((count / results.length) * 100);
    Logger.log(`  ${category}: ${count} (${percentage}%)`);
  });
  
  // Show completion message
  SpreadsheetApp.getUi().alert(
    'Categorization Complete!',
    `Processed ${results.length} companies in ${duration} seconds.\n\n` +
    'Category Breakdown:\n' +
    Object.entries(categoryStats).map(([cat, count]) => 
      `• ${cat}: ${count} (${Math.round((count/results.length)*100)}%)`
    ).join('\n') +
    '\n\nNew sheets have been created for each category.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// DATA PROCESSING
// ============================================================================

function processAllCompaniesCat(sourceSheet, columnMap) {
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  
  const allData = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const results = [];
  
  const totalCompanies = lastRow - CATEGORIZATION_CONFIG.FIRST_DATA_ROW + 1;
  Logger.log(`Processing ${totalCompanies} companies...`);
  
  for (let i = CATEGORIZATION_CONFIG.FIRST_DATA_ROW - 1; i < lastRow; i++) {
    const row = allData[i];
    const rowNumber = i + 1;
    
    const name = cleanTextCat(String(row[columnMap.nameCol] || ''));
    const description = cleanTextCat(String(row[columnMap.descCol] || ''));
    
    // Try to get email from any available email column
    let email = '';
    if (columnMap.primaryEmailCol !== -1) {
      email = String(row[columnMap.primaryEmailCol] || '');
    } else if (columnMap.email1Col !== -1) {
      email = String(row[columnMap.email1Col] || '');
    }
    
    // Get website if available
    const website = columnMap.webCol !== -1 ? String(row[columnMap.webCol] || '') : '';
    
    const analysis = analyzeCompanyCat(name, description, website, email);
    
    results.push({
      rowNumber: rowNumber,
      originalRow: row,
      name: name,
      description: description,
      website: website,
      email: email,
      analysis: analysis
    });
    
    if ((i + 1) % 1000 === 0) {
      Logger.log(`  Processed ${i + 1 - CATEGORIZATION_CONFIG.FIRST_DATA_ROW + 1} rows...`);
    }
  }
  
  Logger.log(`Finished processing ${results.length} companies`);
  return results;
}

function cleanTextCat(text) {
  if (!text) return '';
  text = text.replace(/[^\x20-\x7E\xA0-\xFF]/g, '');
  text = text.replace(/[ÿ®Ÿ±Ä¢©™§¶]+/g, '');
  text = text.replace(/\s+/g, ' ').trim();
  
  const foreignChars = (text.match(/[^\x00-\x7F]/g) || []).length;
  const totalChars = text.length;
  
  if (totalChars > 0 && (foreignChars / totalChars) > 0.5) {
    return '[FOREIGN_LANGUAGE]';
  }
  
  return text;
}

// ============================================================================
// ANALYSIS
// ============================================================================

function analyzeCompanyCat(name, description, website, email) {
  const nameLower = name.toLowerCase();
  const descLower = description.toLowerCase();
  const domain = extractDomainCat(website) || extractDomainCat(email);
  
  // Check for no description
  if (descLower.length < 10 || descLower === '[foreign_language]') {
    return analyzeNoDescriptionCat(nameLower, domain);
  }
  
  // Check for domain for sale
  if (isPlaceholderSiteCat(descLower)) {
    return {
      category: 'NO_DESCRIPTION',
      probability: 0,
      confidence: 100,
      matchedKeywords: ['domain for sale'],
      reasoning: 'Domain for sale or placeholder'
    };
  }
  
  const fullText = (nameLower + ' ' + descLower + ' ' + (domain || '')).toLowerCase();
  
  // STEP 1: Check for EXCLUSIONS (Pool/Fence companies) - HIGHEST PRIORITY
  const poolMatches = CATEGORIZATION_CONFIG.POOL_KEYWORDS.filter(keyword => fullText.includes(keyword));
  const fenceMatches = CATEGORIZATION_CONFIG.FENCE_KEYWORDS.filter(keyword => fullText.includes(keyword));
  
  if (poolMatches.length > 0) {
    return {
      category: 'OTHER',
      probability: 95,
      confidence: 90,
      matchedKeywords: poolMatches,
      reasoning: `Pool company: ${poolMatches.join(', ')}`
    };
  }
  
  if (fenceMatches.length > 0) {
    return {
      category: 'OTHER',
      probability: 95,
      confidence: 90,
      matchedKeywords: fenceMatches,
      reasoning: `Fence company: ${fenceMatches.join(', ')}`
    };
  }
  
  // STEP 2: Check for other specialized trades
  const tradeMatches = CATEGORIZATION_CONFIG.OTHER_TRADES_KEYWORDS.filter(keyword => fullText.includes(keyword));
  
  if (tradeMatches.length >= 2) {
    const hasHomeBuildingTerms = CATEGORIZATION_CONFIG.HOME_BUILDER_KEYWORDS.some(keyword => fullText.includes(keyword));
    
    if (!hasHomeBuildingTerms) {
      return {
        category: 'OTHER',
        probability: 85,
        confidence: 80,
        matchedKeywords: tradeMatches,
        reasoning: `Specialized trade: ${tradeMatches.join(', ')}`
      };
    }
  }
  
  // STEP 3: Check for HOME BUILDER keywords
  const homeBuilderMatches = CATEGORIZATION_CONFIG.HOME_BUILDER_KEYWORDS.filter(keyword => 
    fuzzyMatchCat(fullText, keyword, CATEGORIZATION_CONFIG.FUZZY_THRESHOLD)
  );
  
  if (homeBuilderMatches.length > 0) {
    const confidence = Math.min(95, 70 + (homeBuilderMatches.length * 5));
    return {
      category: 'HOME_BUILDERS',
      probability: confidence,
      confidence: confidence,
      matchedKeywords: homeBuilderMatches,
      reasoning: `Home builder keywords: ${homeBuilderMatches.join(', ')}`
    };
  }
  
  // STEP 4: Check for GENERIC building keywords
  const genericMatches = CATEGORIZATION_CONFIG.GENERIC_BUILDING_KEYWORDS.filter(keyword => fullText.includes(keyword));
  
  if (genericMatches.length > 0) {
    return {
      category: 'HOME_BUILDERS',
      probability: 70,
      confidence: 60,
      matchedKeywords: genericMatches,
      reasoning: `Generic building terms: ${genericMatches.join(', ')}`
    };
  }
  
  // STEP 5: Default to OTHER if nothing matched
  return {
    category: 'OTHER',
    probability: 30,
    confidence: 40,
    matchedKeywords: [],
    reasoning: 'No specific building keywords found'
  };
}

function analyzeNoDescriptionCat(nameLower, domain) {
  const text = (nameLower + ' ' + (domain || '')).toLowerCase();
  
  const matchedKeywords = [];
  let score = 0;
  
  CATEGORIZATION_CONFIG.NO_DESC_KEYWORDS.forEach(([keyword, weight]) => {
    if (fuzzyMatchCat(text, keyword, CATEGORIZATION_CONFIG.FUZZY_THRESHOLD)) {
      score += weight;
      matchedKeywords.push(keyword);
    }
  });
  
  if (matchedKeywords.length > 0) {
    return {
      category: 'NO_DESCRIPTION',
      probability: Math.min(100, score * 10),
      confidence: 60,
      matchedKeywords: matchedKeywords,
      reasoning: `No description but found keywords: ${matchedKeywords.join(', ')}`
    };
  }
  
  return {
    category: 'NO_DESCRIPTION',
    probability: 0,
    confidence: 0,
    matchedKeywords: [],
    reasoning: 'No description and no relevant keywords'
  };
}

function isPlaceholderSiteCat(description) {
  const placeholderTerms = [
    'domain for sale', 'premium domain', 'buy this domain',
    'available for sale', 'parked domain', 'coming soon'
  ];
  return placeholderTerms.some(term => description.includes(term));
}

// ============================================================================
// CATEGORY SHEET CREATION
// ============================================================================

function createCategorySheetsV7(spreadsheet, sourceSheet, results, columnMap) {
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  
  // Build output columns dynamically - only include columns that exist in source
  const possibleColumns = [
    'Contact Full Name',
    'First Name',
    'Last Name',
    'Organization',
    'Primary Email',
    'Email 1',
    'Email 2',
    'Personal Email',
    'Contact Phone 1',
    'Company Phone 1',
    'Company Phone 2',
    'Contact Mobile Phone',
    'Company Description',
    'Website',
    'Company City'
  ];
  
  // Only keep columns that exist in source headers
  const outputColumns = possibleColumns.filter(colName => {
    return sourceHeaders.some(header => 
      header.toLowerCase().trim() === colName.toLowerCase().trim()
    );
  });
  
  // Build output headers with categorization columns
  const outputHeaders = outputColumns.concat([
    'Category',
    'Confidence %',
    'Matched Keywords',
    'Reasoning',
    'Original Row',
    '☐ Check'
  ]);
  
  // 3 categories only
  const categoryNames = {
    'HOME_BUILDERS': 'Home Builders & Related',
    'OTHER': 'Other Companies',
    'NO_DESCRIPTION': 'No Description'
  };
  
  // Group by category
  const categorized = {
    'HOME_BUILDERS': [],
    'OTHER': [],
    'NO_DESCRIPTION': []
  };
  
  const categoryStats = {
    'Home Builders & Related': 0,
    'Other Companies': 0,
    'No Description': 0
  };
  
  results.forEach(result => {
    const category = result.analysis.category;
    categorized[category].push(result);
    categoryStats[categoryNames[category]]++;
  });
  
  // Get column indices for output columns (all should exist since we filtered)
  const outputColumnIndices = outputColumns.map(colName => {
    return sourceHeaders.findIndex(h => 
      h.toLowerCase().trim() === colName.toLowerCase().trim()
    );
  });
  
  // Create the 3 sheets
  for (const [categoryKey, companies] of Object.entries(categorized)) {
    if (companies.length === 0) continue;
    
    const categoryName = categoryNames[categoryKey];
    
    const existingSheet = spreadsheet.getSheetByName(categoryName);
    if (existingSheet) {
      spreadsheet.deleteSheet(existingSheet);
    }
    
    const newSheet = spreadsheet.insertSheet(categoryName);
    
    // Headers
    newSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
    newSheet.getRange(1, 1, 1, outputHeaders.length).setFontWeight('bold');
    
    // Color-code tabs
    if (categoryKey === 'HOME_BUILDERS') {
      newSheet.getRange(1, 1, 1, outputHeaders.length).setBackground('#34a853');
    } else if (categoryKey === 'OTHER') {
      newSheet.getRange(1, 1, 1, outputHeaders.length).setBackground('#ea4335');
    } else {
      newSheet.getRange(1, 1, 1, outputHeaders.length).setBackground('#fbbc04');
    }
    newSheet.getRange(1, 1, 1, outputHeaders.length).setFontColor('#ffffff');
    
    // Data
    const dataRows = companies.map(company => {
      const selectedData = outputColumnIndices.map(idx => 
        idx >= 0 ? company.originalRow[idx] : ''
      );
      
      const analysisData = [
        categoryName,
        company.analysis.confidence,
        company.analysis.matchedKeywords.join(', '),
        company.analysis.reasoning,
        company.rowNumber,
        false
      ];
      
      return selectedData.concat(analysisData);
    });
    
    newSheet.getRange(2, 1, dataRows.length, outputHeaders.length).setValues(dataRows);
    
    // Format checkbox column
    const checkboxCol = outputHeaders.length;
    newSheet.getRange(2, checkboxCol, dataRows.length, 1).insertCheckboxes();
    
    // Format
    newSheet.setFrozenRows(1);
    newSheet.autoResizeColumns(1, outputHeaders.length);
  }
  
  return categoryStats;
}

// ============================================================================
// UPLOAD ENTRIES FUNCTION
// ============================================================================

function uploadCheckedEntriesCat() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  
  Logger.log('=== UPLOADING CHECKED ENTRIES ===');
  
  const allSheets = spreadsheet.getSheets();
  const categorySheets = allSheets.filter(sheet => {
    const name = sheet.getName();
    return name === 'Home Builders & Related' || 
           name === 'Other Companies' || 
           name === 'No Description';
  });
  
  Logger.log(`Found ${categorySheets.length} category sheets`);
  
  const checkedRows = [];
  let totalChecked = 0;
  
  categorySheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return;
    
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    let checkboxCol = -1;
    let originalRowCol = -1;
    
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].includes('Check')) checkboxCol = i;
      if (headers[i] === 'Original Row') originalRowCol = i;
    }
    
    if (checkboxCol === -1 || originalRowCol === -1) return;
    
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    data.forEach((row, idx) => {
      if (row[checkboxCol] === true) {
        checkedRows.push({
          originalRowNum: Number(row[originalRowCol]),
          categorySheet: sheetName
        });
        totalChecked++;
      }
    });
  });
  
  if (totalChecked === 0) {
    SpreadsheetApp.getUi().alert(
      'No Checked Entries',
      'No entries are checked. Please check some entries first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  Logger.log(`Total checked: ${totalChecked}`);
  
  // Find the source sheet (first sheet that's not a category sheet)
  let sourceSheet = null;
  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (name !== 'Home Builders & Related' && 
        name !== 'Other Companies' && 
        name !== 'No Description' &&
        name !== 'Final Results') {
      sourceSheet = sheet;
      break;
    }
  }
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Could not find source sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastCol = sourceSheet.getLastColumn();
  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0];
  const sourceData = sourceSheet.getRange(1, 1, sourceLastRow, sourceLastCol).getValues();
  
  let finalSheet = spreadsheet.getSheetByName('Final Results');
  if (!finalSheet) {
    finalSheet = spreadsheet.insertSheet('Final Results');
  } else {
    finalSheet.clear();
  }
  
  const finalHeaders = sourceHeaders.concat(['Source Category']);
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setFontWeight('bold');
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setBackground('#34a853');
  finalSheet.getRange(1, 1, 1, finalHeaders.length).setFontColor('#ffffff');
  
  const finalData = [];
  let successCount = 0;
  
  checkedRows.forEach(item => {
    const rowNum = item.originalRowNum;
    const category = item.categorySheet;
    
    if (!rowNum || isNaN(rowNum) || rowNum < 1 || rowNum > sourceLastRow) {
      Logger.log(`ERROR: Invalid row ${rowNum}`);
      return;
    }
    
    const arrayIndex = rowNum - 1;
    const originalRow = sourceData[arrayIndex];
    
    if (originalRow) {
      finalData.push(originalRow.concat([category]));
      successCount++;
    }
  });
  
  if (finalData.length > 0) {
    finalSheet.getRange(2, 1, finalData.length, finalHeaders.length).setValues(finalData);
  }
  
  finalSheet.setFrozenRows(1);
  finalSheet.autoResizeColumns(1, finalHeaders.length);
  
  Logger.log(`=== UPLOAD COMPLETE: ${successCount} rows ===`);
  
  SpreadsheetApp.getUi().alert(
    'Upload Complete',
    `Successfully uploaded ${successCount} checked entries to "Final Results" sheet.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function findColumnsByHeadersCat(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const columnMap = {
    contactFullNameCol: -1,
    firstNameCol: -1,
    lastNameCol: -1,
    nameCol: -1,
    descCol: -1,
    primaryEmailCol: -1,
    email1Col: -1,
    email2Col: -1,
    personalEmailCol: -1,
    contactPhone1Col: -1,
    companyPhone1Col: -1,
    companyPhone2Col: -1,
    contactMobilePhoneCol: -1,
    webCol: -1,
    cityCol: -1
  };
  
  for (let i = 0; i < headerRow.length; i++) {
    const header = String(headerRow[i]).trim();
    
    if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.contactFullName)) {
      columnMap.contactFullNameCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.firstName)) {
      columnMap.firstNameCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.lastName)) {
      columnMap.lastNameCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.companyName)) {
      columnMap.nameCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.description)) {
      columnMap.descCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.primaryEmail)) {
      columnMap.primaryEmailCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.email1)) {
      columnMap.email1Col = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.email2)) {
      columnMap.email2Col = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.personalEmail)) {
      columnMap.personalEmailCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.contactPhone1)) {
      columnMap.contactPhone1Col = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.companyPhone1)) {
      columnMap.companyPhone1Col = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.companyPhone2)) {
      columnMap.companyPhone2Col = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.contactMobilePhone)) {
      columnMap.contactMobilePhoneCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.website)) {
      columnMap.webCol = i;
    } else if (matchesHeaderCat(header, CATEGORIZATION_CONFIG.HEADER_NAMES.city)) {
      columnMap.cityCol = i;
    }
  }
  
  // Only require the essential columns for categorization
  if (columnMap.nameCol === -1 || columnMap.descCol === -1) {
    return null;
  }
  
  return columnMap;
}

function matchesHeaderCat(actual, expected) {
  const normalize = (str) => str.toLowerCase().replace(/\s+/g, ' ').trim();
  return normalize(actual) === normalize(expected);
}

function extractDomainCat(input) {
  if (!input) return null;
  input = input.toLowerCase().trim();
  input = input.replace(/^(https?:\/\/)?(www\.)?/, '');
  if (input.includes('@')) input = input.split('@')[1];
  input = input.split('/')[0].split('?')[0];
  input = input.replace(/\.(com|net|org|co|io|biz|info|us)$/i, '');
  return input || null;
}

function fuzzyMatchCat(text, keyword, threshold) {
  if (text.includes(keyword)) return true;
  
  const variations = generateVariationsCat(keyword);
  for (const variation of variations) {
    if (text.includes(variation)) return true;
  }
  
  return false;
}

function generateVariationsCat(keyword) {
  const variations = [];
  if (keyword.endsWith('s')) variations.push(keyword.slice(0, -1));
  else variations.push(keyword + 's');
  if (keyword.includes('-')) {
    variations.push(keyword.replace(/-/g, ' '));
    variations.push(keyword.replace(/-/g, ''));
  }
  if (keyword.includes(' ')) {
    variations.push(keyword.replace(/\s+/g, '-'));
    variations.push(keyword.replace(/\s+/g, ''));
  }
  return variations;
}

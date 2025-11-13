/**
 * @OnlyCurrentDoc
 *
 * This script adds a custom menu to highlight rows based on potential.
 */

// Define the colors for highlighting
const EXPANSIVE_COLOR = '#a4c2f4'; // Light Blue
const MIDDLE_MARKET_COLOR = '#b6d7a8'; // Green
const LOWER_POTENTIAL_COLOR = '#f6b26b'; // Orange

// Define the category names
const EXPANSIVE_NAME = "1 - Expansive Market";
const MIDDLE_NAME = "2 - Medium Value Market";
const LOWER_NAME = "3 - Lower Volume Market";
const OTHER_NAME = ""; // Or "Other" if you prefer
const CATEGORY_HEADER = "Market Category";

/**
 * Creates a custom menu in the spreadsheet when the file is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Market Analysis')
    .addItem('Highlight Potentials', 'highlightPotentials')
    .addItem('Sort by Potential', 'sortPotentials')
    .addToUi();
}

/**
 * Ensures the "Market Category" column exists, adding it if not.
 * @param {Sheet} sheet The active Google Sheet.
 * @return {number} The 0-based index of the "Market Category" column, or -1 if user needs to re-run.
 */
function ensureCategoryColumn(sheet) {
  const ui = SpreadsheetApp.getUi();
  const headerRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  let colIndex = headers.indexOf(CATEGORY_HEADER);

  if (colIndex === -1) {
    // Column not found, add it
    try {
      const newColIndex = sheet.getLastColumn() + 1;
      sheet.getRange(2, newColIndex).setValue(CATEGORY_HEADER);
      // Immediately fetch the new headers to confirm
      const newHeaders = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      colIndex = newHeaders.indexOf(CATEGORY_HEADER);
      if (colIndex === -1) {
         ui.alert(`Failed to add "${CATEGORY_HEADER}" column. Please add it manually and re-run.`);
         return -1;
      }
      SpreadsheetApp.flush(); // Ensure the change is committed
      return colIndex; // Return the new, correct index
    } catch (e) {
      ui.alert(`Could not add "${CATEGORY_HEADER}" column. Error: ${e.message}. Please add it manually on row 2 and re-run.`);
      return -1;
    }
  }
  return colIndex;
}

/**
 * Re-fetches column indices after potential column addition.
 * @param {Sheet} sheet The active Google Sheet.
 * @return {object} An object mapping rule names to column indices, or null if headers aren't ready.
 */
function getColumnIndices(sheet) {
  const headerRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
  const headers = headerRange.getValues()[0];
  let colIndices = findHeaderColumns(headers);

  // Check if Market Category is missing
  if (colIndices.marketCategory === -1) {
    const categoryColIndex = ensureCategoryColumn(sheet);
    if (categoryColIndex === -1) {
      return null; // Stop execution, user was alerted or error occurred
    }
    // Re-fetch headers if we just added one
    const newHeaders = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    colIndices = findHeaderColumns(newHeaders);
  }
  return colIndices;
}


/**
 * Main function to find and highlight high and medium potential rows.
 */
function highlightPotentials() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (!sheet) {
    ui.alert("No active sheet found.");
    return;
  }

  // Get data starting from Row 2 (headers)
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert("Not enough data in the sheet to process. This script expects headers on row 2 and data starting on row 3.");
    return;
  }

  let colIndices = getColumnIndices(sheet);
  if (!colIndices) {
    return; // Stop if ensureCategoryColumn failed or told user to re-run
  }
  
  // Check if all *other* required columns were found
  const missingCols = Object.keys(colIndices)
                           .filter(key => key !== 'marketCategory' && colIndices[key] === -1);
                           
  if (missingCols.length > 0) {
    ui.alert(`Error: Could not find the following required columns: ${missingCols.join(", ")}. Please check the header names on Row 2.`);
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues(); // values[0] is headers, values[1] is row 3

  let expansiveMarketRanges = [];
  let middleMarketRanges = [];
  let lowerPotentialRanges = [];
  const categoryValues = []; // Array to hold the category names for batch writing
  const numCols = sheet.getLastColumn();

  // Define the range for data rows (excluding headers)
  const dataRowsRange = sheet.getRange(3, 1, values.length - 1, numCols);
  // Clear existing background colors first
  dataRowsRange.setBackground(null);

  // Start loop from row 3 (data starts at index 1 of 'values' array)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // Parse values
    const sellThroughRate = Number(row[colIndices.sellThrough]);
    const totalSold = Number(row[colIndices.totalSold]);
    const percentOver200k = Number(row[colIndices.percentOver200k]);
    const percentUnder10k = Number(row[colIndices.percentUnder10k]);

    // Skip rows with invalid data
    if (isNaN(sellThroughRate) || isNaN(totalSold) || isNaN(percentOver200k) || isNaN(percentUnder10k)) {
      categoryValues.push([OTHER_NAME]); // Push empty category
      continue; 
    }
    
    const rowNum = i + 2; // +2 because values index 1 is Row 3
    const rowRangeA1 = sheet.getRange(rowNum, 1, 1, numCols).getA1Notation();
    
    let categoryName = OTHER_NAME;

    // Rule 1: Expansive Market
    if (sellThroughRate > 0.9 && totalSold > 100 && percentOver200k >= 0.4) {
      expansiveMarketRanges.push(rowRangeA1);
      categoryName = EXPANSIVE_NAME;
    } 
    // Rule 2: Middle Market
    else if (sellThroughRate > 0.9 && totalSold > 150 && percentUnder10k <= 0.1) {
      middleMarketRanges.push(rowRangeA1);
      categoryName = MIDDLE_NAME;
    }
    // Rule 3: Lower Potential
    else if (sellThroughRate > 0.8 && totalSold > 60 && percentUnder10k <= 0.1) {
      lowerPotentialRanges.push(rowRangeA1);
      categoryName = LOWER_NAME;
    }
    
    categoryValues.push([categoryName]); // Add the category name for this row
  }

  // Apply formatting in batches
  if (expansiveMarketRanges.length > 0) {
    sheet.getRangeList(expansiveMarketRanges).setBackground(EXPANSIVE_COLOR);
  }
  if (middleMarketRanges.length > 0) {
    sheet.getRangeList(middleMarketRanges).setBackground(MIDDLE_MARKET_COLOR);
  }
  if (lowerPotentialRanges.length > 0) {
    sheet.getRangeList(lowerPotentialRanges).setBackground(LOWER_POTENTIAL_COLOR);
  }
  
  // Write all category names to the "Market Category" column
  if (categoryValues.length > 0) {
    sheet.getRange(3, colIndices.marketCategory + 1, categoryValues.length, 1).setValues(categoryValues);
  }

  ui.alert(`Highlighting complete! Found ${expansiveMarketRanges.length} Expansive, ${middleMarketRanges.length} Medium Value, and ${lowerPotentialRanges.length} Lower Volume rows.`);
}

/**
 * New function to sort the sheet by potential category and sell-through rate.
 */
function sortPotentials() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (!sheet) {
    ui.alert("No active sheet found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    ui.alert("Not enough data in the sheet to sort. This script expects headers on row 2 and data starting on row 3.");
    return;
  }

  ui.alert("Sorting rows... This may take a moment. Highlights will be reapplied automatically when sorting is complete.");

  let colIndices = getColumnIndices(sheet);
  if (!colIndices) {
    return; // Stop if ensureCategoryColumn failed
  }
  
  // Check if all *other* required columns were found
  const missingCols = Object.keys(colIndices)
                           .filter(key => key !== 'marketCategory' && colIndices[key] === -1);
                           
  if (missingCols.length > 0) {
    ui.alert(`Error: Could not find the following required columns: ${missingCols.join(", ")}. Please check the header names on Row 2.`);
    return;
  }

  // Get data starting from Row 2 (headers)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  let values = dataRange.getValues();
  
  const headers = values[0]; // Row 2
  // Get only the data rows (from row 3 onwards)
  const dataRows = values.slice(1);
  let sortableData = [];

  // Loop through data rows and assign a category and sell-through rate
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    const sellThroughRate = Number(row[colIndices.sellThrough]);
    const totalSold = Number(row[colIndices.totalSold]);
    const percentOver200k = Number(row[colIndices.percentOver200k]);
    const percentUnder10k = Number(row[colIndices.percentUnder10k]);

    let category = 4; // 4 = Other (default)
    let categoryName = OTHER_NAME;
    let sellThrough = isNaN(sellThroughRate) ? -1 : sellThroughRate;

    if (!isNaN(sellThroughRate) && !isNaN(totalSold) && !isNaN(percentOver200k) && !isNaN(percentUnder10k)) {
      // Rule 1: Expansive Market
      if (sellThroughRate > 0.9 && totalSold > 100 && percentOver200k >= 0.4) {
        category = 1;
        categoryName = EXPANSIVE_NAME;
      } 
      // Rule 2: Middle Market
      else if (sellThroughRate > 0.9 && totalSold > 150 && percentUnder10k <= 0.1) {
        category = 2;
        categoryName = MIDDLE_NAME;
      }
      // Rule 3: Lower Potential
      else if (sellThroughRate > 0.8 && totalSold > 60 && percentUnder10k <= 0.1) {
        category = 3;
        categoryName = LOWER_NAME;
      }
    }
    
    // Write the category name directly into the row data
    row[colIndices.marketCategory] = categoryName;
    
    sortableData.push({
      data: row, // This row data now includes the category name
      category: category,
      sellThrough: sellThrough
    });
  }

  // Sort the data
  sortableData.sort((a, b) => {
    // Primary sort: by category (1, 2, 3, 4)
    if (a.category !== b.category) {
      return a.category - b.category;
    }
    // Secondary sort: by sell-through rate (descending)
    return b.sellThrough - a.sellThrough;
  });

  // Get the sorted data back into a 2D array
  const sortedValues = sortableData.map(item => item.data);
  
  // Define the range to write to (Row 3 to end)
  const numRows = sortedValues.length;
  if (numRows === 0) {
    // This should be covered by the lastRow < 3 check, but good to have
    ui.alert("No data rows found to sort.");
    return;
  }
  const numCols = sortedValues[0].length;
  const dataRegion = sheet.getRange(3, 1, numRows, numCols);

  // Write the sorted data back to the sheet
  dataRegion.setValues(sortedValues);

  // Re-apply the highlights, which will now match the new row order
  // A small delay to ensure sorting is fully processed before highlighting
  Utilities.sleep(500);
  highlightPotentials();
}


/**
 * Helper function to find the column index for each required header.
 * @param {Array<string>} headers An array of header names (from row 2).
 * @return {object} An object mapping rule names to column indices.
 */
function findHeaderColumns(headers) {
  return {
    // This is the line that was fixed
    sellThrough: headers.indexOf("Sold/OnMarket (Sell Through Rate) - Bigger the better"),
    totalSold: headers.indexOf("Total # of Sold Properties in 6 months"),
    percentOver200k: headers.indexOf("% (>200k)"),
    percentUnder10k: headers.indexOf("%<10k"),
    marketCategory: headers.indexOf(CATEGORY_HEADER) // New column to find
  };
}
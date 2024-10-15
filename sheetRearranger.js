/**
 * Adds a custom menu to the Google Sheets UI upon opening the spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Rearrange Data', 'rearrangeData')
    .addToUi();
}

/**
 * Rearranges the data from the active sheet by separating fixed columns
 * and question-related data, organizing it into a new sheet with two header rows.
 * Additionally, applies conditional formatting to the marks columns.
 */
function rearrangeData() {
  // Define the name for the output sheet
  const outputSheetName = "Processed Data";
  
  // Get the active spreadsheet and the currently active sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getActiveSheet();
  
  // Get the name of the source sheet for reference
  const sourceSheetName = sourceSheet.getName();
  
  // Get all data from the source sheet
  const sourceData = sourceSheet.getDataRange().getValues();
  
  // Check if there is enough data (at least header + one data row)
  if (sourceData.length < 2) {
    SpreadsheetApp.getUi().alert("The active sheet does not contain enough data.");
    return;
  }
  
  // Get the header row
  const headerRow = sourceData[0];
  
  /**
   * Finds the index of the first occurrence of "QuestionNum" in the headers.
   * @param {Array} headers - The header row as an array.
   * @returns {number} - The zero-based index of "QuestionNum" or -1 if not found.
   */
  function findFirstQuestionNumColumn(headers) {
    for (let i = 0; i < headers.length; i++) {
      if (headers[i].toString().trim().toLowerCase() === 'questionnum') {
        return i;
      }
    }
    return -1; // Not found
  }
  
  // Determine the number of fixed columns by finding "QuestionNum"
  const fixedColumnsIndex = findFirstQuestionNumColumn(headerRow);
  
  if (fixedColumnsIndex === -1) {
    SpreadsheetApp.getUi().alert('The header row does not contain a "QuestionNum" column.');
    return;
  }
  
  const fixedColumns = fixedColumnsIndex;
  
  // Determine the number of question triplets (QuestionNum, Mark, MaxMark)
  const totalColumns = headerRow.length;
  const remainingColumns = totalColumns - fixedColumns;
  
  if (remainingColumns % 3 !== 0) {
    SpreadsheetApp.getUi().alert("The number of columns after fixed columns is not a multiple of 3.");
    return;
  }
  
  const tripletCount = remainingColumns / 3;
  
  // Initialize arrays to store unique question identifiers and their max marks
  const questionIdentifiers = [];
  const questionMaxMarks = {};
  
  // Iterate through each data row to collect unique question identifiers and their max marks
  for (let i = 1; i < sourceData.length; i++) { // Start from 1 to skip header
    const row = sourceData[i];
    
    // Optional: Skip rows that are summaries or do not contain data
    // For example, skip rows where the first cell contains "Max marks" or is empty
    const firstCell = row[0].toString().trim().toLowerCase();
    if (firstCell.includes("max marks") || firstCell === "") {
      continue;
    }
    
    for (let t = 0; t < tripletCount; t++) {
      const qNumCol = fixedColumns + t * 3;
      const markCol = fixedColumns + t * 3 + 1;
      const maxMarkCol = fixedColumns + t * 3 + 2;
      
      const qIdentifier = row[qNumCol];
      const mark = row[markCol];
      const maxMark = row[maxMarkCol];
      
      // Skip empty question identifiers
      if (qIdentifier && qIdentifier.toString().trim() !== "") {
        // If encountering the question for the first time, add to the list
        if (!questionMaxMarks.hasOwnProperty(qIdentifier)) {
          questionIdentifiers.push(qIdentifier);
          questionMaxMarks[qIdentifier] = maxMark;
        }
      }
    }
  }
  
  // Create or clear the output sheet
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  } else {
    outputSheet.clearContents();
    outputSheet.clearConditionalFormatRules(); // Clear existing conditional formatting
  }
  
  // Prepare headers
  const fixedHeaders = headerRow.slice(0, fixedColumns);
  
  // First header row: Fixed headers + Question Identifiers
  const headerRow1 = [...fixedHeaders, ...questionIdentifiers];
  
  // Second header row: Empty for fixed columns + Max Marks
  const headerRow2 = Array(fixedColumns).fill("");
  questionIdentifiers.forEach(q => {
    headerRow2.push(questionMaxMarks[q]);
  });
  
  // Initialize an array to hold all output rows
  const outputRows = [];
  outputRows.push(headerRow1);
  outputRows.push(headerRow2);
  
  // Iterate through each data row to map marks to questions
  for (let i = 1; i < sourceData.length; i++) { // Start from 1 to skip header
    const row = sourceData[i];
    
    // Optional: Skip rows that are summaries or do not contain data
    const firstCell = row[0].toString().trim().toLowerCase();
    if (firstCell.includes("max marks") || firstCell === "") {
      continue;
    }
    
    // Extract fixed column data
    const fixedData = row.slice(0, fixedColumns);
    
    // Initialize an object to map question identifiers to marks
    const marksMap = {};
    
    for (let t = 0; t < tripletCount; t++) {
      const qNumCol = fixedColumns + t * 3;
      const markCol = fixedColumns + t * 3 + 1;
      const maxMarkCol = fixedColumns + t * 3 + 2;
      
      const qIdentifier = row[qNumCol];
      const mark = row[markCol];
      // const maxMark = row[maxMarkCol]; // Not needed here
      
      // Skip invalid question identifiers
      if (qIdentifier && qIdentifier.toString().trim() !== "") {
        marksMap[qIdentifier] = mark;
      }
    }
    
    // Prepare the marks in the order of questionIdentifiers
    const marks = questionIdentifiers.map(q => marksMap[q] !== undefined ? marksMap[q] : "");
    
    // Combine fixed data and marks
    const outputRow = [...fixedData, ...marks];
    
    // Add the row to the outputRows array
    outputRows.push(outputRow);
  }
  
  // Write all rows at once to the output sheet
  if (outputRows.length > 0) {
    outputSheet.getRange(1, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  } else {
    SpreadsheetApp.getUi().alert("No valid data rows found to process.");
    return;
  }
  
  // Apply Conditional Formatting to Marks Columns
  applyConditionalFormatting(outputSheet, fixedColumns, questionIdentifiers.length);
  
  // Optional: Auto-resize columns for better visibility
  outputSheet.autoResizeColumns(1, outputSheet.getLastColumn());
  
  SpreadsheetApp.getUi().alert(`Data rearrangement and conditional formatting complete!\nProcessed data from "${sourceSheetName}" to "${outputSheetName}".`);
}

/**
 * Converts a column number to its corresponding letter (e.g., 1 -> 'A').
 * @param {number} column - The column number (1-based).
 * @returns {string} - The corresponding column letter.
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Applies conditional formatting to the marks columns in the processed sheet.
 * The color scale ranges from Red (0) to Amber (mid-point) to Green (max mark).
 * Additionally, highlights cells containing 'N' with a red background.
 * 
 * @param {Sheet} sheet - The sheet to apply conditional formatting to.
 * @param {number} fixedColumns - The number of fixed columns before the marks columns.
 * @param {number} questionCount - The number of question columns.
 */
function applyConditionalFormatting(sheet, fixedColumns, questionCount) {
  const rules = [];
  
  // Get the maximum row number with data
  const lastRow = sheet.getLastRow();
  
  // Calculate the number of data rows (excluding headers)
  const numDataRows = lastRow - 2; // Assuming headers are in rows 1 and 2
  
  // Iterate through each question column to apply conditional formatting
  for (let i = 0; i < questionCount; i++) {
    const colIndex = fixedColumns + 1 + i; // 1-based indexing
    
    // Get the max mark from the second header row
    const maxMark = sheet.getRange(2, colIndex).getValue();
    
    // Validate maxMark
    const numericMaxMark = parseFloat(maxMark);
    if (isNaN(numericMaxMark) || numericMaxMark <= 0) {
      // Skip applying conditional formatting if maxMark is invalid
      continue;
    }
    
    // Convert column index to letter for A1 notation
    const columnLetter = columnToLetter(colIndex);
    
    // Define the range for conditional formatting (from row 3 to lastRow)
    const rangeA1 = `${columnLetter}3:${columnLetter}${lastRow}`;
    const range = sheet.getRange(rangeA1);
    
    // Define the color scale: Red (0) -> Amber (mid) -> Green (maxMark)
    const colorScaleRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue("#FF0000", SpreadsheetApp.InterpolationType.NUMBER, "0") // Red
      .setGradientMidpointWithValue("#FFC000", SpreadsheetApp.InterpolationType.NUMBER, (numericMaxMark / 2).toString()) // Amber
      .setGradientMaxpointWithValue("#00FF00", SpreadsheetApp.InterpolationType.NUMBER, numericMaxMark.toString()) // Green
      .setRanges([range])
      .build();
    rules.push(colorScaleRule);
    
    // 'N' Text Rule: Highlight cells containing 'N' with red background
    const nValueRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('N')
      .setBackground('#FF0000') // Red
      .setRanges([range])
      .build();
    rules.push(nValueRule);
  }
  
  // Apply all conditional formatting rules at once
  sheet.setConditionalFormatRules(rules);
}

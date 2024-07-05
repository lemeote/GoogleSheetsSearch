function generateCombinedDataFormula() {

  const SEARCH_SHEET_NAME = 'SearchBy3Parameters';
  const COMBO_SHEET_NAME = 'ComboTable';
  const EVERY_TABLE_RANGE = '$B$5:$E' // Note that each table has to start and end in the same column, start in the same row, but can end on a different row

  // Define the names of sheets to exclude
  const excludeSheetNames = [SEARCH_SHEET_NAME, COMBO_SHEET_NAME];
  
  // Get the active spreadsheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get all the sheets in the spreadsheet
  const sheets = spreadsheet.getSheets();
  
  // Initialize the formula string
  let formula = '=ARRAYFORMULA(QUERY({';

  for(let sheet of sheets){
    let sheetName = sheet.getName();
    if(excludeSheetNames.includes(sheetName))
      continue;

    /* This is a try to get the table from the sheet dynamically, but it get the range starting from A1 even if not data there
    let data_range = sheet.getDataRange().getA1Notation(); // Get the table range from the each sheet
    data_range = data_range.substring(0, data_range.indexOf(":")+2) // Remove last row index from range to keep it dynamic
    let table_range = `'${sheetName}'!${data_range}`
    */

    let data_range = EVERY_TABLE_RANGE;
    
    let table_range = `'${sheetName}'!${data_range}`;
    
    formula += `${table_range}, RIGHT(ROW(${table_range})&"${sheetName}", LEN("${sheetName}"));` // We create array of each table with extra rows and an additional column for sheet name (source)

  }
  
  // Remove the trailing semicolon and add the rest of the formula to get the combined table without the extra rows
  formula = formula.slice(0, -1) + '},"SELECT * WHERE Col1 IS NOT NULL"))';
   
  // Set the formula in the first cell in the Combo Sheet
  spreadsheet.getSheetByName(COMBO_SHEET_NAME).getRange(1,1).setFormula(formula);

}



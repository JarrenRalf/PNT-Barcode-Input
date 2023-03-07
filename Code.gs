/**
* This function handles onEdit events in order to control the functioning of the search page and the count pages.
*
* @param {Event Object} e The event object received by the onEdit function
*/
function onEdit(e)
{
  var range = e.range;         // The range of cells edited
  var value = e.value;         // The new value of the cell edited
  var col = range.columnStart; // The first column of the edited range
  var row = range.rowStart;    // The first row of the edited range
  var spreadsheet = e.source;
  var sheetName = spreadsheet.getActiveSheet().getSheetName();
  
  try
  {
    if (sheetName == "Search")
    {
      if (col === 3) // If the search box is edited
      {
        spreadsheet.getRangeByName("ItemSelectionCheckBoxes").uncheck();
        spreadsheet.getRangeByName("WarningMessage").setValue('');
      }
      else if (col === 4 && row === 2 && isCellDeleted(value))  // If the selection cell is being deleted
        range.setValue("Please Make A Selection"); // Populate the data validation on the search page
      else if (col === 5)                 
      {
        if (isCellDeleted(value))   // If the transfer button is deleted
          range.insertCheckboxes();
        else if (value)             // If the transfer button is being edited and it is checked
          copySelectedValues(spreadsheet);
      }
    }
    else if (col === 3 && sheetName != "Dashboard" && sheetName != "Progress" && sheetName != "Progress Search" && !isCellDeleted(value))
      range.offset(1, -2).activate(); // Move the active selection down one cell and left two cells 
  }
  catch (error)
  {
    Browser.msgBox(error);
  }
}

/**
* This function handles the onOpen events (including refresh) which sets all of the dashboard values
*
* @param {Event Object} e The event object received by the onEdit function
*/
function onOpen(e)
{
  const spreadsheet = e.source;
  const s = spreadsheet.getSheets();
  const ss = s.filter(value => value.getSheetName() != "Search"   && value.getSheetName() != "Database"  && value.getSheetName() != "EXPORT TEMPLATE" 
                            && value.getSheetName() != "Progress" && value.getSheetName() != "Dashboard" && value.getSheetName() != "Progress Search");
  const sheets = ss.filter(value => value.getSheetName().substring(value.getSheetName().length - 6, value.getSheetName().length) != "Export");
  const exportSheets = ss.filter(value => value.getSheetName().substring(value.getSheetName().length - 6, value.getSheetName().length) == "Export");
  const dashboardSheet = spreadsheet.getSheetByName("Dashboard"); // Remove the Dashboard sheet from the array
  const sheetNames = sheets.map(value => value.getSheetName());
  const dashboardRange = dashboardSheet.getRange(4, 2, 13, 19);
  var  dashboardValues = dashboardRange.getValues();
  
  try
  {
    dashboardValues[0][17] = 0; // Total Number of SKUs Pending
    dashboardValues[2][17] = 0; // Total Number of SKUs Exported
    dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets,  0, 5, "Richmond");
    dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets,  5, 2, "Parksville");
    dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets, 10, 1, "Rupert");
    dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets, 15, 2, "Trites");
    
    if (exportSheets.length == 0) // There are no exports
    {
      dashboardValues[0][4] = "- - -"; // Date of Last Import
      dashboardValues[2][4] = "- - -"; // Date of Last Export
    }
    
    dashboardRange.setValues(dashboardValues);
  }
  catch (error)
  {
    Browser.msgBox(error);
  }
}

/**
 * This function moves the current searched values on the progress search page to the adagio export page for zeroing a concatenated list.
 * 
 * @author Jarren Ralf
 */
function addItemsToExport()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Progress Search');
  const adagioExportSheet = spreadsheet.getSheetByName('Adagio Export');
  const location = sheet.getRange(2, 1).getValue();
  const numItems = getLastRowSpecial(sheet.getRange('A5:A').getValues());

  if(numItems > 0)
  {
    const values = sheet.getRange(5, 1, numItems, 2).getValues();
    const col = (location === 'Richmond') ? 1 : ( (location === 'Parksville') ? 4 : ( (location === 'Rupert') ? 7 : 10) );
    var lastRow = getLastRowSpecial(adagioExportSheet.getRange(1, col, adagioExportSheet.getMaxRows()).getValues());
    if (lastRow === 2) lastRow = 3;
    spreadsheet.getSheetByName('Adagio Export').getRange(lastRow + 1, col, numItems, 2).setValues(values);
  }
  else
    Browser.msgBox('No Items Found.')
}

/**
 * This function builds a trigger that will run the remainingSKUs function daily between 10 and 11.
 */
function autoUpdateRemainingSKUs()
{
  ScriptApp.newTrigger('remainingSKUs').timeBased().atHour(10).everyDays(1).create();
}

/**
* This function clears all of the data across all of the sheets selected by the checkboxes on the Dashboard.
*
* @author Jarren Ralf
*/
function clearData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const s = spreadsheet.getSheets();
  const ss = s.filter(value => value.getSheetName() != "Search"   && value.getSheetName() != "Database"  && value.getSheetName() != "EXPORT TEMPLATE" 
                            && value.getSheetName() != "Progress" && value.getSheetName() != "Dashboard" && value.getSheetName() != "Progress Search");
  const sheets = ss.filter(value => value.getSheetName().substring(value.getSheetName().length - 6, value.getSheetName().length) != "Export");
  const exportSheets = ss.filter(value => value.getSheetName().substring(value.getSheetName().length - 6, value.getSheetName().length) == "Export");
  const sheetNames = sheets.map(value => value.getSheetName());
  const dashboardRange = spreadsheet.getSheetByName("Dashboard").getRange(4, 2, 13, 19);
  var dashboardValues = dashboardRange.getValues();
  
  dashboardValues = clearLocationData(spreadsheet, dashboardValues,  0, 5); // rich
  dashboardValues = clearLocationData(spreadsheet, dashboardValues,  5, 2); // park
  dashboardValues = clearLocationData(spreadsheet, dashboardValues, 10, 1); // rupt
  dashboardValues = clearLocationData(spreadsheet, dashboardValues, 15, 2); // trit
  
  dashboardValues[0][17] = 0; // Total Number of SKUs Pending
  dashboardValues[2][17] = 0; // Total Number of SKUs Exported
  dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets,  0, 5, "Richmond");
  dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets,  5, 2, "Parksville");
  dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets, 10, 1, "Rupert");
  dashboardValues = updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets, 15, 2, "Trites");
  
  if (exportSheets.length == 0) // There are no exports
  {
    dashboardValues[0][4] = "- - -"; // Function Run Time
    dashboardValues[2][4] = "- - -"; // Date of Last Export
  }

  dashboardRange.setValues(dashboardValues);
}

/**
 * This function clears the adagio export page for the selected location
 * 
 * @author Jarren Ralf
 */
function clearAdagioExport()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName('Adagio Export');
  const location = spreadsheet.getSheetByName('Progress Search').getRange(2, 1).getValue();
  const col = (location === 'Richmond') ? 1 : ( (location === 'Parksville') ? 4 : ( (location === 'Rupert') ? 7 : 10) );
  sheet.getRange(4, col, sheet.getMaxRows(), 2).clearContent();
}

/**
* This function clears all of the data on the selected sheet chosen by the checkbox for the specified location.
*
* @param {Spreadsheet}     spreadsheet : The active spreadsheet
* @param  {Object[][]} dashboardValues : The set of all values on the dashboard
* @param   {Number}        columnIndex : The column index for the dashboard values that correspond with a particular location
* @param   {Number}     numberOfSheets : The number of sheets that correspond to the particular location
* @return {Object[][]} dashboardValues : The updated dashboardValues
* @author Jarren Ralf
*/
function clearLocationData(spreadsheet, dashboardValues, columnIndex, numberOfSheets)
{
  const SHEET_NAME = columnIndex;
  const CHECK_BOX  = columnIndex + 1;
  const NUM_SKUs   = columnIndex + 2;
  const NUM_PIECES = columnIndex + 3;
  const ROW_START  = 5;

  for (var i = 0; i < numberOfSheets; i++)
  {
    if (dashboardValues[ROW_START + i][CHECK_BOX])
    {
      spreadsheet.getSheetByName(dashboardValues[ROW_START + i][SHEET_NAME]).getRangeList(["A3:A", "C3:C"]).clearContent();
      dashboardValues[ROW_START + i][NUM_SKUs]   = 0;
      dashboardValues[ROW_START + i][NUM_PIECES] = 0;
      dashboardValues[ROW_START + i][CHECK_BOX]  = false;
    }
  }
  
  return dashboardValues;
}

/**
* This function collects the export data for the chosen location.
*
* @param {Spreadsheet}    spreadsheet : The active spreadsheet
* @param {Object[][]} dashboardValues : The set of all values on the dashboard
* @param   {Number}       columnIndex : The column index for the dashboard values that correspond with a particular location
* @param   {Number}    numberOfSheets : The number of sheets that correspond to the particular location
* @return {Object[][], Object[][]} [output, dashboardValues] : The export data and the updated dashboardValues
* @author Jarren Ralf
*/
function collectExportData(spreadsheet, dashboardValues, columnIndex, numberOfSheets)
{
  const SHEET_NAME = columnIndex;
  const CHECK_BOX  = columnIndex + 1;
  const ROW_START  = 8;
  var outputData = [], uniqueSKUsData = [];
  
  for (var i = 0; i < numberOfSheets; i++)
  {
    if (dashboardValues[ROW_START + i][CHECK_BOX])
      [outputData, uniqueSKUsData] = getExportData(spreadsheet.getSheetByName(dashboardValues[ROW_START + i][SHEET_NAME]), outputData, uniqueSKUsData);
    
    dashboardValues[ROW_START + i][CHECK_BOX] = false;
  }
  
  Logger.log(outputData)
  var output = outputData.filter(value => value[0] != "#N/A" && !isBlank(value[0]) && isNumber(value[1]));

  return [output, dashboardValues];
}

/**
* This function takes all of the items on the search page that are selected by checkbox in column A, and moves them to 
* the user specified sheet based on which selection is made at cell D2.
*
* @param {Spreadsheet} spreadsheet : The active spreadsheet
* @author Jarren Ralf
*/
function copySelectedValues(spreadsheet)
{
  const ui = SpreadsheetApp.getUi();
  const userSelection = spreadsheet.getRangeByName("UserSelection").getValue();
  const functionTriggerCheckbox = spreadsheet.getRangeByName("FunctionTriggerCheckbox");
  const warningMessage = spreadsheet.getRangeByName("WarningMessage");

  if (userSelection != "Please Make A Selection")
  {
    const searchSheet = spreadsheet.getActiveSheet();
    const destinationSheet = spreadsheet.getSheetByName(userSelection);
    const numRows = getLastRowSpecial(searchSheet.getRange("C4:C").getValues());
    
    if (numRows > 0)
    {
      const ITEM_SELECTED = 0;
      const UPC_CODE = 4;
      const searchValuesRange = searchSheet.getRange(4, 1, numRows, 5);
      const searchValues = searchValuesRange.getValues();
      const  lastRow = getLastRowSpecial(destinationSheet.getRange("A:A").getValues());
      const startRow = (lastRow < 3) ? 3 : lastRow + 1; 
      var UPC_Code = [];
      
      for (var i = 0; i < numRows; i++)
      {
        if (searchValues[i][ITEM_SELECTED])
          UPC_Code.push([searchValues[i][UPC_CODE]]);
      }
      
      const numItems = UPC_Code.length;
      
      if (numItems != 0)
      {
        functionTriggerCheckbox.uncheck();
        spreadsheet.getRangeByName("ItemSelectionCheckBoxes").uncheck();
        destinationSheet.getRange(startRow, 1, numItems).setValues(UPC_Code); // Move the item values to the Order sheet
        searchSheet.getRange(1, 3, 2, 2).setValues([[null, 'Item/s have been moved to ' + userSelection + '\'s sheet.'], 
                                                    [''  , "Please Make A Selection"]]);
        destinationSheet.getRange(startRow, 3).activate();
      }
      else
      {
        warningMessage.setValue('Please select items by clicking the checkboxes in column A.');
        functionTriggerCheckbox.uncheck();
        ui.alert('Please select items by clicking the checkboxes in column A.');
      }
    }
    else
    {
      warningMessage.setValue('There are no items found. Please try your search again.');
      functionTriggerCheckbox.uncheck();
      ui.alert('There are no items found. Please try your search again.');
    }
  }
  else
  {
    warningMessage.setValue('Please make a selection at the top of column D.');
    functionTriggerCheckbox.uncheck();
    ui.alert('Please make a selection at the top of column D.');
  }
}

/**
* This function checks if the given sheet already exists in the spreadsheet.
*
* @param   {Sheet}  sheet : The given sheet object
* @return {Boolean}       : Whether the sheet already exists or not
* @author Jarren Ralf
*/ 
function doesSheetAlreadyExist(sheet)
{
  return sheet != null
}

/**
* This function creates a new sheet for the export data.
*
* @author Jarren Ralf
*/
function exportData()
{
  const INDEX = 2;
  const spreadsheet = SpreadsheetApp.getActive();
  const templateSheet = spreadsheet.getSheetByName("EXPORT TEMPLATE");
  const LOCATION = ["Richmond", "Parksville", "Rupert", "Trites"];
  const FORMATTED_DATE = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "dd MMM yyyy HH:mm");
  const dashboardRange = spreadsheet.getSheetByName("Dashboard").getDataRange();
  var dashboardValues = dashboardRange.getValues();
  var outputs = [];
  var sheetName, exportSheet;

  if (isBlank(dashboardValues[5][18])) // If the total export values are blank
    dashboardValues[5][18] = 0;
  
  [outputs[0], dashboardValues] = collectExportData(spreadsheet, dashboardValues,  1, 5); // rich
  [outputs[1], dashboardValues] = collectExportData(spreadsheet, dashboardValues,  6, 2); // park
  [outputs[2], dashboardValues] = collectExportData(spreadsheet, dashboardValues, 11, 1); // rupt
  [outputs[3], dashboardValues] = collectExportData(spreadsheet, dashboardValues, 16, 2); // trit

  for (var i = outputs.length - 1; i >= 0; i--)
  {
    if (outputs[i].length == 0) // Skip the output with no items
      continue;
    
    sheetName = FORMATTED_DATE + " " + LOCATION[i] + " Export";
    exportSheet = spreadsheet.getSheetByName(sheetName);
    
    if (doesSheetAlreadyExist(exportSheet)) 
      spreadsheet.deleteSheet(exportSheet);
    
    if (isBlank(dashboardValues[15][4 + 5*i])) // If the total export values are blank
      dashboardValues[15][4 + 5*i] = 0;
    
    exportSheet = spreadsheet.insertSheet(sheetName, INDEX, {template: templateSheet});
    exportSheet.getRange(1, 2).setValue(LOCATION[i] + " Export Data");
    exportSheet.getRange(4, 1, outputs[i].length, outputs[i][0].length).setNumberFormat('@').setValues(outputs[i]);
    dashboardValues[15][4 + 5*i] += outputs[i].length;      // The number of exports for each location
    dashboardValues[5][18] += dashboardValues[15][4 + 5*i]; // The total number of exports
    
    // Protect the active sheet, then remove all other users from the list of editors
    var protection = exportSheet.protect();
    protection.removeEditors(protection.getEditors()); // Remove everyone    
    protection.addEditors(["adriangatewood@gmail.com", "lb_blitz_allstar@hotmail.com"]); // Then add dummy and dummy
  }
  
  dashboardValues[5][5] = FORMATTED_DATE;
  dashboardRange.setValues(dashboardValues);
}

/**
* This function collects all of the export data from the chosen sheet.
*
* @param    {Sheet}       sheet : The chosen sheet to collect export data from
* @param  {Object[][]}   output : The output data
* @param  {Object[]} uniqueSKUs : The array of unique SKUs
* @return {Object[][]}   output : The output data
* @author Jarren Ralf
*/
function getExportData(sheet, output, uniqueSKUs)
{
  const ROW_START = 3;
  const COL_START = 2;
  const  COL_ONE = sheet.getRange("A:A").getValues();
  const NUM_ROWS = getLastRowSpecial(COL_ONE) - ROW_START + 1;
  
  if (NUM_ROWS <= 0) // Nothing to export
    return [output, uniqueSKUs];
  
  const NUM_COLS = 6;
  const VALUES = sheet.getRange(ROW_START, COL_START, NUM_ROWS, NUM_COLS).getValues();
  const SHEET_NAME = sheet.getSheetName();
  const  SKU = 0;
  const  QTY = 1;
  const NAME = 4; 
  const  LOC = 5; // The index position of values multi-array that represents the location that the selected item has been counted in

  var index;
  
  for (var i = 0; i < NUM_ROWS; i++)
  {
    index = uniqueSKUs.indexOf(VALUES[i][SKU]); // Get the index position of the sku for the i-th item in the unique array
      
    if (isItemUnique(index))
    {
      uniqueSKUs.push(VALUES[i][SKU]); // Put the unique sku into the unique array
      VALUES[i][NAME] = SHEET_NAME;    // Add the sheet name i.e. The person who counted the items
      output.push(VALUES[i]);          // Save the entire row of data in the output array
    }
    else if (isNonNegative(VALUES[i][QTY])) // sku already in list
    {
      output[index][QTY] += VALUES[i][QTY];         // Add to the previous quantity
      output[index][LOC] += " + " + VALUES[i][LOC]; // Add the location
      
      if (output[index][NAME] != SHEET_NAME)
        output[index][NAME] += " + " + SHEET_NAME; // Add the names of the people who counted
    }
  }

  return [output, uniqueSKUs];
}

/**
* Gets the last row number based on a selected column range values.
*
* @param  {array}  range  : Takes a 2d array of a single column's values
* @return {number} rowNum : The last row number with a value. 
*/ 
function getLastRowSpecial(range)
{
  var rowNum = 0;
  var blank = false;
  
  for (var row = 0; row < range.length; row++)
  {
    if(isBlank(range[row][0]) && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if(!isBlank(range[row][0]))
      blank = false;
  }
  return rowNum;
}

/**
* This function checks if the given string is blank or not.
*
* @param   {String} str : A string
* @return {Boolean}     : Whether the given string is blank or not.
* @author Jarren Ralf
*/ 
function isBlank(str)
{
  return str === '';
}

/**
* This function checks if the given cell value has just been changed to null.
*
* @param   {Range}  cell : A range object with only 1 row and 1 column
* @return {Boolean}      : Whether the given cell is deleted or not.
*/ 
function isCellDeleted(cell)
{
  return cell == undefined;
}

/**
* This function checks if a value is contained in an array by checking the index position, and if it equals -1, it's not in the array.
*
* @param  {Number}  index : The index position of a certain value in an array
* @return {Boolean}       : Whether the item is in the array or not 
* @author Jarren Ralf
*/ 
function isItemUnique(index)
{
  return index === -1;
}

/**
* This function checks if a value is a number or not.
*
* @param   {Object} num : The given object
* @return {Boolean}     : Whether the input is a number or not
* @author Jarren Ralf
*/ 
function isNumber(num)
{
  return Number.isFinite(num);
}

/**
* This function checks if a value is a positive number or zero.
*
* @param   {Object} num : The given object
* @return {Boolean}     : Whether the input is a non-negative or not
* @author Jarren Ralf
*/ 
function isNonNegative(num)
{
  return num >= 0;
}

/**
* This function checks if a value is poitive or not.
*
* @param  {Number}  num : The given number
* @return {Boolean}     : Whether the number is positive or not
* @author Jarren Ralf
*/ 
function isPositive(num)
{
  return num > 0
}

/**
 * This function takes all of the SKUs that have been exported by each of the stores and compares them to the current Adagio data. 
 * If there is a SKU in the adagio data that has not been exported and it has a non-zero quantity, then it is displayed.
 * 
 * @author Jarren Ralf
 */
function remainingSKUs()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const s = spreadsheet.getSheets();
  const progressSheet = spreadsheet.getSheetByName("Progress");
  const exportSheets = s.filter(val => val.getSheetName().substring(val.getSheetName().length - 6, val.getSheetName().length) == "Export"); // Keep sheets if the last word is "Export"
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const  COLOURS = ['	#d9ead3',    '#c9daf8', '#f4cccc', '#fff2cc'];
  const LOCATION = ['Richmond', 'Parksville',  'Rupert',  'Trites'];
  const NUM_LOCATIONS = LOCATION.length;
  const     totalExports = [...new Array(4)].map(e => []);
  const    uniqueExports = [...new Array(4)].map(e => []);
  const  skusNotExported = [...new Array(4)].map(e => []);
  const numSKUsRemaining = [...new Array(3)].map(e => new Array(10));

  // The number of SKUs exported
  for (var j = 0; j < exportSheets.length; j++)
  {
    words = exportSheets[j].getSheetName().split(" ");
    
    for (var i = 0; i < NUM_LOCATIONS; i++)
    {
      if (words[4] == LOCATION[i])
        totalExports[i] = totalExports[i].concat(exportSheets[j].getRange(4, 1, exportSheets[j].getLastRow() - 3).getValues()); 
    }
  }

  progressSheet.getRange(10, 2, progressSheet.getMaxRows() - 9, 11).clear();

  for (var k = 0; k < NUM_LOCATIONS; k++)
  {
      uniqueExports[k] = uniqByKeepFirst(totalExports[k], val => val[0]);
    skusNotExported[k] = adagioData.filter(e => uniqueExports[k].filter(f => e[7] == f[0]).length == 0).map(g => [g[1], g[k + 2]]).filter(h => h[1] != 0);
    var numSKUs = skusNotExported[k].length - 1;
    var exportRange = progressSheet.getRange(9, 2 + 3*k, numSKUs + 1, 2);
    numSKUsRemaining[0][k*3] = uniqueExports[k].length + ' Counted';
    numSKUsRemaining[2][k*3] = numSKUs + ' Remaining';
    exportRange.setWrap(true);
    exportRange.setFontSize(10);
    exportRange.setBackground(COLOURS[k]);
    exportRange.setVerticalAlignment('middle');
    exportRange.setBorder(null, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK);
    if (numSKUs !== 0) exportRange.setValues(skusNotExported[k]);
  }

  progressSheet.getRange(4, 2, 3, numSKUsRemaining[0].length).setValues(numSKUsRemaining);
}

/**
 * This function updates the data via a csv file. It also imports the adagio descriptions as well.
 * 
 * @author Jarren
 */
function resetData()
{  
  const SKU = 4;
  const PRICE_UNIT = 15;
  const DESCRIPTION = 9;
  const spreadsheet = SpreadsheetApp.getActive();
  const db = spreadsheet.getSheetByName("Database");
  const FORMATTED_DATE = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "dd MMM yyyy HH:mm");
  var adagioData = Utilities.parseCsv(DriveApp.getFilesByName("FullAdagioData.csv").next().getBlob().getDataAsString());
  var    upcData = Utilities.parseCsv(DriveApp.getFilesByName("BarcodeInput.csv").next().getBlob().getDataAsString());
  const NUM_ITEMS = upcData.length

  var concatData = upcData.map(val => [val[0], val[1], val[1].concat(" - " + val[2] + " - " + val[3]), val[3]]);

  for (var i = 1; i < NUM_ITEMS; i++)
  {
    for (var j = 1; j < adagioData.length; j++)
    {
      if (concatData[i][1] == adagioData[j][SKU]) // If the SKUs match
        concatData[i][2] = adagioData[j][SKU] + ' - ' + adagioData[j][DESCRIPTION] + ' - ' + adagioData[j][PRICE_UNIT]; // Replace the description with the Adagio one
    }
  }

  concatData.push(concatData.shift()); // Make the header the last element of the data array for sorting purposes
  db.getRange('A:A').activate();
  db.getRange(2, 1, db.getMaxRows() - 1, 4).clearContent();
  db.getRange(1, 1, concatData.length, concatData[0].length).setNumberFormat('@').setValues(concatData.reverse()); // Items sorted by recently entered
  spreadsheet.getSheetByName("Dashboard").getRange(4, 6).setValue(FORMATTED_DATE);
}

/**
* This is a function I found and modified to keep the first instance of an item in a muli-array based on the uniqueness of one of the values.
*
* @param      {Object[][]}    arr The given array
* @param  {Callback Function} key A function that chooses one of the elements of the object or array
* @return     {Object[][]}    The reduced array containing only unique items based on the key
*/
function uniqByKeepFirst(arr, key)
{
  let seen = new Set();

  return arr.filter(item => {
      let k = key(item);
      return seen.has(k) ? false : seen.add(k);
  });
}

/**
* This function updates all of the values on the dashboard associated with a particular location.
*
* @param  {Object[][]} dashboardValues : The set of all values on the dashboard
* @param   {Sheet[]}            sheets : An array of all the relevant counting sheets in this spreadsheet
* @param  {String[]}        sheetNames : An array of all the sheet names of the above sheets
* @param   {Sheet[]}      exportSheets : An array of all of the existing export sheets
* @param   {Number}      locationIndex : The index for the dashboardValues that corresponds to a particular location
* @param   {Number}     numberOfSheets : The number of sheets for a partiucular location
* @param   {String}           location : A string of the particular location names
* @return {Object[][]} dashboardValues : The updated dashboardValues
* @author Jarren Ralf
*/
function updateDashboardData(dashboardValues, sheets, sheetNames, exportSheets, locationIndex, numberOfSheets, location)
{
  const ROW_START = 5;
  var index, range;
  var words = [];

  dashboardValues[11][locationIndex + 2] = 0; // Total SKUs
  dashboardValues[11][locationIndex + 3] = 0; // Total Pieces
  dashboardValues[12][locationIndex + 3] = 0; // Total Exports
  
  for (var i = 0; i < numberOfSheets; i++)
  {
    if (isBlank(dashboardValues[12][locationIndex + 3])) // Number of SKUs exported
      dashboardValues[12][locationIndex + 3] = 0;
    
    index = sheetNames.indexOf(dashboardValues[5 + i][locationIndex]); // Index of the sheet
    range = sheets[index].getRange("A3:C").getValues();
    dashboardValues[ROW_START + i][locationIndex + 2] = getLastRowSpecial(range); // Number of SKUs
    dashboardValues[ROW_START + i][locationIndex + 3] = range.reduce((total, value) => parseInt(total + value[2]), 0); // Quantity
    dashboardValues[11][locationIndex + 2] += dashboardValues[ROW_START + i][locationIndex + 2]; // Total number of SKUs
    dashboardValues[11][locationIndex + 3] += dashboardValues[ROW_START + i][locationIndex + 3]; // Total quantity
  }

  // The number of SKUs exported
  for (var j = 0; j < exportSheets.length; j++)
  {
    words = exportSheets[j].getSheetName().split(" ");
    
    if (words[4] == location)
      dashboardValues[12][locationIndex + 3] += exportSheets[j].getLastRow() - 3; // Number of exports
  }

  dashboardValues[0][17] += dashboardValues[11][locationIndex + 2]; // Total Number of SKUs Pending
  dashboardValues[2][17] += dashboardValues[12][locationIndex + 3]; // Total Number of SKUs Exported
  
  return dashboardValues;
}
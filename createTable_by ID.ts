//this is useful if you do not need dynamic identification of starting row and ending column 
//this allows you to hardcode and specify the two.
//also specifies the Table name 


function main(workbook: ExcelScript.Workbook) {
  // Get the first worksheet 
  const selectedSheet = workbook.getFirstWorksheet();
  //get active range of WorkSheet
  let range = workbook.getActiveWorksheet().getUsedRange();
  // Get last used row of WorkSheet
  let lastrow = range.getRowCount();

  
  // Find first reference of ID in selectedSheet i.e. header row
  let IDCell = selectedSheet.getRange("A1").find("Employee ID", { completeMatch: true, matchCase: true, searchDirection: ExcelScript.SearchDirection.forward });  //to specify the starting cell by its name.
  // Get the current active cell in the workbook.
  //and format current cell without Sheet1! reference
  let activeCell = IDCell.getAddress().replace("Sheet1!", "");


  //get table range  
  const TableRange = `${activeCell}:F${lastrow}`; //to specify the ending column letter.
  // Create a table using the data range.
  let newTable = workbook.addTable(selectedSheet.getRange(TableRange), true);
  newTable.setName("NewTableInExcel");
  // Get the first (and only) table in the worksheet.
  let table = selectedSheet.getTables()[0];
  // Get the data from the table.
  let tableValues = table.getRange().getValues();
  //Return a response to the Cloud Flow
  return tableValues
}



//10.Dec.2023
//find the usedRange() and create a table. the startcell is manually specified.
//can be improved by : dynamic abilty to check that the first row of the range has non-empty column headers. 
//-JSLim


function main(workbook: ExcelScript.Workbook) {
    // Get the current, active worksheet.
    const currentWorksheet = workbook.getActiveWorksheet();
  
    // Specify the starting row.
    let startRow = 5; // Change this to your desired starting row.
  
    // Get the last used cell in the worksheet.
    let lastCell = currentWorksheet.getUsedRange().getLastCell();

    // Log the last cell to the console.
    console.log(`Last cell address: ${lastCell.getAddress()}`);
  
    // Calculate the address for the new range.
    let startCellAddress = "A" + (startRow); 
    let endCellAddress = lastCell.getAddress();
  
    // Get the range starting from the specified row.
    let usedRange = currentWorksheet.getRange(startCellAddress + ":" + endCellAddress);
  
    // Log the range's address to the console.
    console.log(`Used range address: ${usedRange.getAddress()}`);
  
    // Create a table from the used range with headers.
    const table = currentWorksheet.addTable(usedRange, true);
  
    // Set a name for the table (optional) but first check if table creation was successful before setting the name.
    if (table) {
      table.setName("TermList");
      console.log(`Table created successfully. Name: ${table.getName()}`);
    } else {
      console.log("Error creating table.");
    }
  }
  
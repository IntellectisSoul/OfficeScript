//15-Dec.2023 : this creates a table by finding the usedRange. It is also 'smart enough' to be able to
//disregard rows with 'incomplete headers', so it checks the top rows against the usedRange total number of columns.
//Optionally, one might also use "Option 1 : Find ID of cell Method" to identify the starting cell  of row header.
//Next steps to improve : test bottom row to avoid fake endcell. 

//-JSLim  

function main(workbook: ExcelScript.Workbook, tablename: string)) {  //tablename allows passing argument from PA to Officescript.
  // getUsedRange() : Get the active sheet and range.
  const selectedSheet = workbook.getActiveWorksheet();
  const FullrangeAddress = selectedSheet.getUsedRange();
  console.log(`FullrangeAddress : ${FullrangeAddress.getAddress()}`);

  // getColumnCount() : Get the total number of columns.
  const columnCount = FullrangeAddress.getColumnCount();
  console.log(`ColumnCount: ${columnCount}`);

  //CALL FUNCTION to extract the starting and ending cell. because function returns an object of 3 (value pair)
  // using dot notation is how to extract the objects individually
  const extractedCells = extract_StartEndCell(FullrangeAddress.getAddress());
  let startingCell = extractedCells.startingCell; //declare start
  let endingCell = extractedCells.endingCell; //declear end
  console.log(`startingCell: ${extractedCells.startingCell} | endingCell: ${extractedCells.endingCell}`);

  //CALL FUNCTION to Breakdown and split the  cell address into their separate letter and number
  let { startcell_letter, startcell_number, endcell_letter, endcell_number } = celladdress_Breakdown(startingCell, endingCell);

  //CALL FUNCTION to retrieve the values of the starting row to check for Headers completion  : 
  const { values, elementCount } = getRowIndexValues(selectedSheet, (startcell_number - 1), columnCount);  //this function returns 2 objects.
  console.log(`Values: ${values}`);
  console.log(`Element Count: ${elementCount}`);

/*
//test bottom row to avoid fake endcell. identify correct endcell
while (true) {
  let testrangeAddress_columnCount = testrange(startcell_letter, startcell_number, endcell_letter, selectedSheet );

  // Break the loop if testrangeAddress_columnCount equals columnCount
  if (testrangeAddress_columnCount > (columnCount/2)) {
    console.log(`starting Row ${startcell_letter}${startcell_number}, Header not found...moving to next row`);
    // Increment startcell_number
    startcell_number++;
  } else {
    // Define the final Range to be used to create the Table; set finalstartcell
    let newRange = (startcell_letter + startcell_number) + ":" + (endcell_letter + endcell_number);
    const table = selectedSheet.addTable(newRange, true);
    const totalRows = endcell_number - startcell_number;
    console.log(`new Range start : ${startcell_letter}${startcell_number}, Total rows = ${totalRows}`);
    
    if (table) {
      table.setName("TermList");
      console.log(`Table created successfully. Name: ${table.getName()}`);
    } else {
      console.log("Error creating table.");
    }
    break;  // This will break the loop.
  }
}
*/

  //test top row for header completion and then create table.
  while (true) {
    let testrangeAddress_columnCount = testrange(startcell_letter, startcell_number, endcell_letter, selectedSheet );

    // Break the loop if testrangeAddress_columnCount equals columnCount
    if (testrangeAddress_columnCount !== columnCount) {
      console.log(`starting Row ${startcell_letter}${startcell_number}, Header not found...moving to next row`);
      // Increment startcell_number
      startcell_number++;
    } else {
      // Define the final Range to be used to create the Table; set finalstartcell
      let newRange = (startcell_letter + startcell_number) + ":" + (endcell_letter + endcell_number);
      const table = selectedSheet.addTable(newRange, true);
      const totalRows = endcell_number - startcell_number;
      console.log(`new Range start : ${startcell_letter}${startcell_number}, Total rows = ${totalRows}`);
      
      if (table) {
        table.setName(tablename);
        console.log(`Table created successfully. Name: ${table.getName()}`);
      } else {
        console.log("Error creating table.");
      }
      break;  // This will break the loop.
    }
  }


    // Check if cell is null or undefined.
    const iscellEmpty = isCellEmpty("A2", selectedSheet);
    console.log(`Cell is empty: ${iscellEmpty}`);



}

// FUNCTIONS

//1 : Extract starting and ending cell.
function extract_StartEndCell(rangeAddress: string): { startingCell: string; endingCell: string} {
  const parts = rangeAddress.split("!");
  const cellAddress = parts[1];
  const cells = cellAddress.split(":");

  // Extract starting cell
  const startingCell = cells[0];
  // Extract ending cell
  const endingCell = cells[1];

  // Return an object containing all extracted values
  return { startingCell, endingCell };
}



//Unused : Check if cell is empty.
function isCellEmpty(cellAddress: string, sheet: ExcelScript.Worksheet): boolean {
  const cell = sheet.getRange(cellAddress);

  // Check if the cell is null or has an undefined value or is empty.
  return !cell || cell.getText() === '';

}

//2 : this function breaks down the cell address to its start and end cell.
function celladdress_Breakdown(startingCell: string, endingCell: string) {
  //Breakdown and split the  cell address into their separate letter and number
  const startmatch = startingCell.match(/^([A-Z]+)(\d+)$/);  //start
  const startcell_letter = startmatch[1]; // A
  const startcell_number = parseInt(startmatch[2]); // 3
  console.log(`startcell_letter: ${startcell_letter} | startcell_number: ${startcell_number}`);

  const endmatch = endingCell.match(/^([A-Z]+)(\d+)$/);   //end
  const endcell_letter = endmatch[1]; // A
  const endcell_number = parseInt(endmatch[2]); // 3
  console.log(`endcell_letter: ${endcell_letter} | endcell_number: ${endcell_number}`);

  return { startcell_letter, startcell_number, endcell_letter, endcell_number }
}

//3: this function returns the array of the row and its elementCount.
type RangeValues = (string | number | boolean)[][]; //to explicitly define RangeValues because Excel needs this when type of output is not defined and it is unable to understand ExcelScript.RangeValue[]

function getRowIndexValues(selectedSheet: ExcelScript.Worksheet, startcellnumber: number, columnCount: number): { values: RangeValues, elementCount: number } {

  //function getRowIndexValues(selectedSheet: ExcelScript.Worksheet, startcellnumber: number, columnCount: number ): void {
  const range = selectedSheet.getRangeByIndexes(startcellnumber, 0, 1, columnCount);
  const values = range.getValues();
  const elementCount = values[0].filter(cell => cell !== "").length;
  console.log(`getRowIndexvalues of index-based row ${startcellnumber}: ${values}`);
  console.log(`elementCount : ${elementCount}`);
  return { values, elementCount };
}

//4: function to iterate through each row and test/count the number of elements.
function testrange(startcell_letter: string, startcell_number: number, endcell_letter: string, selectedSheet: ExcelScript.Worksheet): number {

  const testrangeAddress = selectedSheet.getRange(startcell_letter + startcell_number + ":" + endcell_letter + startcell_number);
  const testingAdddress_Range = testrangeAddress.getValues();
  const testrangeAddress_columnCount = testingAdddress_Range[0].filter((cell: string | number | boolean) => cell !== "").length; //ensures empty cells are not counted

  console.log(`testrangeAddress: ${testrangeAddress.getValues()}, testrangeAddress_columnCount is  ${testrangeAddress_columnCount},`);
  return testrangeAddress_columnCount

}

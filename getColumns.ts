function main(workbook: ExcelScript.Workbook) {
    // Get the active cell and worksheet.
    let selectedSheet = workbook.getActiveWorksheet();

    // Example range address
  let rangeAddress = selectedSheet.getUsedRange();
  console.log(`first cellAddress : ${rangeAddress.getAddress()}`);

//get values of entire column 3 (zero index base) 
  let column = rangeAddress.getColumn(3);
  const columnvalues = column.getValues(); // Synchronize and retrieve values as a 2D array
  console.log(`Column values: ${columnvalues}`);
 
  //get total column count
  let columnCount = rangeAddress.getColumnCount();
  console.log(`columnCount : ${columnCount}`);


//test columnAfter
  let columnafter = rangeAddress.getColumnsAfter(-1);
  const values = columnafter.getValues(); // Synchronizes and retrieves values
  console.log(`columnafter values: ${values}`);


    // Call the function to extract the starting cell
    let startingCell = extractStartingCell(rangeAddress.getAddress());



    // TODO: Write code or use the Insert action button below.
}

function extractStartingCell(rangeAddress: string): string {
    // Split the range address using '!' as the delimiter
    var parts = rangeAddress.split('!');

    // Extract the second part, which contains the cell address
    var cellAddress = parts[1];
    // Split the cellAddress address using ':' as the delimiter
    var parts = cellAddress.split(':');
    // Extract the first part, which contains the cell address
    var cellAddress = parts[0];
    // Log the result to the console (you can modify this part based on your needs)
    console.log(`Starting cellAddress : ${cellAddress}`);

    // Return the cell address
    return cellAddress;


}
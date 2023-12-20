//swaps the position of 2 column header & contents. must be in range, not in table format.
//input : actual range & startRow
//hardcoded input : columns to swap index-based



function main(workbook: ExcelScript.Workbook, startRow: number = 4) {
    // Get the worksheet.
    let sheet = workbook.getActiveWorksheet();
  
    // Define the range for the columns to be swapped.
    let range = sheet.getRange(`A${startRow}:J10`); // Adjust this to your actual range.
  
    // Get the values in the range.
    let rangeValues = range.getValues();
  
    // Define the new headers.
    let newHeaders = ["Approver GEID", "Approver Email", "Approver LDS", "Approver FullName", "HRM Email", "Count of EC", "Count of RCM", "Count of Concur", "Total"]; // Adjust this to your actual headers.
  
    // Replace the headers in the first row for the columns that exist in the newHeaders array.
    for (let i = 0; i < newHeaders.length && i < rangeValues[0].length; i++) {
      rangeValues[0][i] = newHeaders[i];
    }
  
    // Swap the column headers.
    let tempHeader = rangeValues[0][1]; // 2nd column
    rangeValues[0][1] = rangeValues[0][3]; // 4th column
    rangeValues[0][3] = tempHeader;
  
    // Swap the column contents.
    for (let i = 1; i < rangeValues.length; i++) {
      let temp = rangeValues[i][1]; // 2nd column
      rangeValues[i][1] = rangeValues[i][3]; // 4th column
      rangeValues[i][3] = temp;
    }
  
    // Set the new values to the range.
    range.setValues(rangeValues);
  }
  
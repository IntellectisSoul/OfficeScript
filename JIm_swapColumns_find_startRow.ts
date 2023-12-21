//21.Dec.2023 : this finds the correct column headers as starting point and swaps them.
//requires input on which columns to swap : Swap the column headers and column contents.

function main(workbook: ExcelScript.Workbook, startColumn: string = 'A', endColumn: string = 'J') {
    // Get the worksheet.
    let sheet = workbook.getActiveWorksheet();

    // Get the used range.
    let usedRange = sheet.getUsedRange();

    // Get the values of the used range.
    let usedRangeValues = usedRange.getValues();

    // Get the total number of columns.
    const columnCount = usedRange.getColumnCount();

    // Find the first row where all cells in the specific range are not empty.
    let startRow = 1;
    for (let i = 0; i < usedRangeValues.length; i++) {
        let testRange = sheet.getRange(`${startColumn}${i + 1}:${endColumn}${i + 1}`);
        let testRangeValues = testRange.getValues();
        let testRangeColumnCount = testRangeValues[0].filter((cell: string | number | boolean) => cell !== "").length;

        console.log(`Row ${i + 1} - usedRangeValues ${usedRangeValues.length}`);
        console.log(`Row ${i + 1} - Values: ${JSON.stringify(testRangeValues)}`);
        console.log(`Row ${i + 1} - Non-empty cells: ${testRangeColumnCount}`);  

        if (testRangeColumnCount === columnCount) {
            startRow = i + 1;

            console.log(`Row ${i + 1} - usedRangeValues ${usedRangeValues.length}`);
            console.log(`Row ${i + 1} - Values: ${JSON.stringify(testRangeValues)}`);
            console.log(`Row ${i + 1} - Non-empty cells: ${testRangeColumnCount}`);        
            break;
        }
    }

    // Log the startRow value.
    console.log(`Start Row: ${startRow}`);

    // Get the last row of the used range.
    let endRow = usedRange.getRowCount();

    // Define the range for the headers.
    let headerRange = sheet.getRange(`${startColumn}${startRow}:${endColumn}${startRow}`);

    // Get the values in the header range.
    let headerRangeValues = headerRange.getValues();

    // Define the new headers.
    let newHeaders = ["Approver GEID", "Approver Email", "Approver LDS", "Approver FullName", "HRM Email", "Count of EC", "Count of RCM", "Count of Concur", "Total"];

    // Replace the headers in the startRow for the columns that exist in the newHeaders array.
    for (let i = 0; i < newHeaders.length && i < headerRangeValues[0].length; i++) {
        headerRangeValues[0][i] = newHeaders[i];
    }

    // Swap the column headers.
    let tempHeader = headerRangeValues[0][1]; // 2nd column
    headerRangeValues[0][1] = headerRangeValues[0][3]; // 4th column
    headerRangeValues[0][3] = tempHeader;

    // Set the new values to the header range.
    headerRange.setValues(headerRangeValues);

    // Define the range for the column contents to be swapped.
    let contentRange = sheet.getRange(`${startColumn}${startRow + 1}:${endColumn}${endRow}`);

    // Get the values in the content range.
    let contentRangeValues = contentRange.getValues();

    // Swap the column contents.
    for (let i = 0; i < contentRangeValues.length; i++) {
        let temp = contentRangeValues[i][1]; // 2nd column
        contentRangeValues[i][1] = contentRangeValues[i][3]; // 4th column
        contentRangeValues[i][3] = temp;
    }

    // Set the new values to the content range.
    contentRange.setValues(contentRangeValues);
}

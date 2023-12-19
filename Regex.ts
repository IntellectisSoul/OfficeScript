function main
    (
        workbook: ExcelScript.Workbook, inputString: string, pattern: string, flags: string
    ): Array<string> {
    let regExp = new RegExp(pattern, flags);
    let matches: Array<string> = inputString.match(regExp);

    if (matches) {
        return matches;
    } else {
        return [];
    }

}
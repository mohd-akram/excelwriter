import fs from "fs/promises";

import { Format, Workbook, Worksheet, cell } from "excelwriter";

/**
 * Write some data to the worksheet.
 *
 * @param {Worksheet} worksheet
 * @param {Format} format
 */
function writeWorksheetData(worksheet, format) {
  worksheet.writeString(
    ...cell("A1"),
    "Some examples of data validation in libxlsxwriter",
    format
  );
  worksheet.writeString(...cell("B1"), "Enter values in this column", format);
  worksheet.writeString(...cell("D1"), "Sample Data", format);

  worksheet.writeString(...cell("D3"), "Integers");
  worksheet.writeNumber(...cell("E3"), 1);
  worksheet.writeNumber(...cell("F3"), 10);

  worksheet.writeString(...cell("D4"), "List Data");
  worksheet.writeString(...cell("E4"), "open");
  worksheet.writeString(...cell("F4"), "high");
  worksheet.writeString(...cell("G4"), "close");

  worksheet.writeString(...cell("D5"), "Formula");
  worksheet.writeFormula(...cell("E5"), "=AND(F5=50,G5=60)");
  worksheet.writeNumber(...cell("F5"), 50);
  worksheet.writeNumber(...cell("G5"), 60);
}

/*
 * Create a worksheet with data validations.
 */

const workbook = new Workbook();
const worksheet = workbook.addWorksheet("Sheet1");

/* Add a format to use to highlight the header cells. */
const format = workbook.addFormat();
format.setBorder(Format.THIN_BORDER);
format.setFgColor(0xc6efce);
format.setBold();
format.setTextWrap();
format.setAlign(Format.VERTICAL_CENTER_ALIGN);
format.setIndent(1);

/* Write some data for the validations. */
writeWorksheetData(worksheet, format);

/* Set up layout of the worksheet. */
worksheet.setColumn(0, 0, 55);
worksheet.setColumn(1, 1, 15);
worksheet.setColumn(3, 3, 15);
worksheet.setRow(0, 36);

/*
 * Example 1. Limiting input to an integer in a fixed range.
 */
worksheet.writeString(...cell("A3"), "Enter an integer between 1 and 10");

worksheet.dataValidationCell(...cell("B3"), {
  validate: Worksheet.INTEGER_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumNumber: 1,
  maximumNumber: 10,
});

/*
 * Example 2. Limiting input to an integer outside a fixed range.
 */
worksheet.writeString(
  ...cell("A5"),
  "Enter an integer not between 1 and 10 (using cell references)"
);

worksheet.dataValidationCell(...cell("B5"), {
  validate: Worksheet.INTEGER_FORMULA_VALIDATION_TYPE,
  criteria: Worksheet.NOT_BETWEEN_VALIDATION_CRITERIA,
  minimumFormula: "=E3",
  maximumFormula: "=F3",
});

/*
 * Example 3. Limiting input to an integer greater than a fixed value.
 */
worksheet.writeString(...cell("A7"), "Enter an integer greater than 0");

worksheet.dataValidationCell(...cell("B7"), {
  validate: Worksheet.INTEGER_VALIDATION_TYPE,
  criteria: Worksheet.GREATER_THAN_VALIDATION_CRITERIA,
  valueNumber: 0,
});

/*
 * Example 4. Limiting input to an integer less than a fixed value.
 */
worksheet.writeString(...cell("A9"), "Enter an integer less than 10");

worksheet.dataValidationCell(...cell("B9"), {
  validate: Worksheet.INTEGER_VALIDATION_TYPE,
  criteria: Worksheet.LESS_THAN_VALIDATION_CRITERIA,
  valueNumber: 10,
});

/*
 * Example 5. Limiting input to a decimal in a fixed range.
 */
worksheet.writeString(...cell("A11"), "Enter a decimal between 0.1 and 0.5");

worksheet.dataValidationCell(...cell("B11"), {
  validate: Worksheet.DECIMAL_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumNumber: 0.1,
  maximumNumber: 0.5,
});

/*
 * Example 6. Limiting input to a value in a dropdown list.
 */
worksheet.writeString(...cell("A13"), "Select a value from a drop down list");

worksheet.dataValidationCell(...cell("B13"), {
  validate: Worksheet.LIST_VALIDATION_TYPE,
  valueList: ["open", "high", "close"],
});

/*
 * Example 7. Limiting input to a value in a dropdown list.
 */
worksheet.writeString(
  ...cell("A15"),
  "Select a value from a drop down list (using a cell range)"
);

worksheet.dataValidationCell(...cell("B15"), {
  validate: Worksheet.LIST_FORMULA_VALIDATION_TYPE,
  valueFormula: "=$E$4:$G$4",
});

/*
 * Example 8. Limiting input to a date in a fixed range.
 */
worksheet.writeString(
  ...cell("A17"),
  "Enter a date between 1/1/2008 and 12/12/2008"
);

worksheet.dataValidationCell(...cell("B17"), {
  validate: Worksheet.DATE_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumDatetime: new Date(2008, 0, 1, 0, 0, 0),
  maximumDatetime: new Date(2008, 11, 12, 0, 0, 0),
});

/*
 * Example 9. Limiting input to a time in a fixed range.
 */
worksheet.writeString(...cell("A19"), "Enter a time between 6:00 and 12:00");

worksheet.dataValidationCell(...cell("B19"), {
  validate: Worksheet.TIME_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumDatetime: new Date(0, 0, 0, 6, 0, 0),
  maximumDatetime: new Date(0, 0, 0, 12, 0, 0),
});

/*
 * Example 10. Limiting input to a string greater than a fixed length.
 */
worksheet.writeString(
  ...cell("A21"),
  "Enter a string longer than 3 characters"
);

worksheet.dataValidationCell(...cell("B21"), {
  validate: Worksheet.LENGTH_VALIDATION_TYPE,
  criteria: Worksheet.GREATER_THAN_VALIDATION_CRITERIA,
  valueNumber: 3,
});

/*
 * Example 11. Limiting input based on a formula.
 */
worksheet.writeString(
  ...cell("A23"),
  'Enter a value if the following is true "=AND(F5=50,G5=60)"'
);

worksheet.dataValidationCell(...cell("B23"), {
  validate: Worksheet.CUSTOM_FORMULA_VALIDATION_TYPE,
  valueFormula: "=AND(F5=50,G5=60)",
});

/*
 * Example 12. Displaying and modifying data validation messages.
 */
worksheet.writeString(
  ...cell("A25"),
  "Displays a message when you select the cell"
);

worksheet.dataValidationCell(...cell("B25"), {
  validate: Worksheet.INTEGER_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumNumber: 1,
  maximumNumber: 100,
  inputTitle: "Enter an integer:",
  inputMessage: "between 1 and 100",
});

/*
 * Example 13. Displaying and modifying data validation messages.
 */
worksheet.writeString(
  ...cell("A27"),
  "Display a custom error message when integer isn't between 1 and 100"
);

worksheet.dataValidationCell(...cell("B27"), {
  validate: Worksheet.INTEGER_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumNumber: 1,
  maximumNumber: 100,
  inputTitle: "Enter an integer:",
  inputMessage: "between 1 and 100",
  errorTitle: "Input value is not valid!",
  errorMessage: "It should be an integer between 1 and 100",
});

/*
 * Example 14. Displaying and modifying data validation messages.
 */
worksheet.writeString(
  ...cell("A29"),
  "Display a custom info message when integer isn't between 1 and 100"
);

worksheet.dataValidationCell(...cell("B29"), {
  validate: Worksheet.INTEGER_VALIDATION_TYPE,
  criteria: Worksheet.BETWEEN_VALIDATION_CRITERIA,
  minimumNumber: 1,
  maximumNumber: 100,
  inputTitle: "Enter an integer:",
  inputMessage: "between 1 and 100",
  errorTitle: "Input value is not valid!",
  errorMessage: "It should be an integer between 1 and 100",
  errorType: Worksheet.INFORMATION_VALIDATION_ERROR_TYPE,
});

const data = workbook.close();
await fs.writeFile("data-validate.xlsx", Buffer.from(data));

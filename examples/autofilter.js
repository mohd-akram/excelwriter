import fs from "fs/promises";

import { Format, Workbook, Worksheet } from "excelwriter";

/**
 *
 * @param {Worksheet} worksheet
 * @param {Format} header
 */
function writeWorksheetHeader(worksheet, header) {
  /* Make the columns wider for clarity. */
  worksheet.setColumn(0, 3, 12);

  /* Write the column headers. */
  worksheet.setRow(0, 20, header);
  worksheet.writeString(0, 0, "Region");
  worksheet.writeString(0, 1, "Item");
  worksheet.writeString(0, 2, "Volume");
  worksheet.writeString(0, 3, "Month");
}

const workbook = new Workbook();
const worksheet1 = workbook.addWorksheet("Sheet1");
const worksheet2 = workbook.addWorksheet("Sheet2");
const worksheet3 = workbook.addWorksheet("Sheet3");
const worksheet4 = workbook.addWorksheet("Sheet4");
const worksheet5 = workbook.addWorksheet("Sheet5");
const worksheet6 = workbook.addWorksheet("Sheet6");
const worksheet7 = workbook.addWorksheet("Sheet7");

const data = [
  { region: "East", item: "Apple", volume: 9000, month: "July" },
  { region: "East", item: "Apple", volume: 5000, month: "July" },
  { region: "South", item: "Orange", volume: 9000, month: "September" },
  { region: "North", item: "Apple", volume: 2000, month: "November" },
  { region: "West", item: "Apple", volume: 9000, month: "November" },
  { region: "South", item: "Pear", volume: 7000, month: "October" },
  { region: "North", item: "Pear", volume: 9000, month: "August" },
  { region: "West", item: "Orange", volume: 1000, month: "December" },
  { region: "West", item: "Grape", volume: 1000, month: "November" },
  { region: "South", item: "Pear", volume: 10000, month: "April" },
  { region: "West", item: "Grape", volume: 6000, month: "January" },
  { region: "South", item: "Orange", volume: 3000, month: "May" },
  { region: "North", item: "Apple", volume: 3000, month: "December" },
  { region: "South", item: "Apple", volume: 7000, month: "February" },
  { region: "West", item: "Grape", volume: 1000, month: "December" },
  { region: "East", item: "Grape", volume: 8000, month: "February" },
  { region: "South", item: "Grape", volume: 10000, month: "June" },
  { region: "West", item: "Pear", volume: 7000, month: "December" },
  { region: "South", item: "Apple", volume: 2000, month: "October" },
  { region: "East", item: "Grape", volume: 7000, month: "December" },
  { region: "North", item: "Grape", volume: 6000, month: "April" },
  { region: "East", item: "Pear", volume: 8000, month: "February" },
  { region: "North", item: "Apple", volume: 7000, month: "August" },
  { region: "North", item: "Orange", volume: 7000, month: "July" },
  { region: "North", item: "Apple", volume: 6000, month: "June" },
  { region: "South", item: "Grape", volume: 8000, month: "September" },
  { region: "West", item: "Apple", volume: 3000, month: "October" },
  { region: "South", item: "Orange", volume: 10000, month: "November" },
  { region: "West", item: "Grape", volume: 4000, month: "July" },
  { region: "North", item: "Orange", volume: 5000, month: "August" },
  { region: "East", item: "Orange", volume: 1000, month: "November" },
  { region: "East", item: "Orange", volume: 4000, month: "October" },
  { region: "North", item: "Grape", volume: 5000, month: "August" },
  { region: "East", item: "Apple", volume: 1000, month: "December" },
  { region: "South", item: "Apple", volume: 10000, month: "March" },
  { region: "East", item: "Grape", volume: 7000, month: "October" },
  { region: "West", item: "Grape", volume: 1000, month: "September" },
  { region: "East", item: "Grape", volume: 10000, month: "October" },
  { region: "South", item: "Orange", volume: 8000, month: "March" },
  { region: "North", item: "Apple", volume: 4000, month: "July" },
  { region: "South", item: "Orange", volume: 5000, month: "July" },
  { region: "West", item: "Apple", volume: 4000, month: "June" },
  { region: "East", item: "Apple", volume: 5000, month: "April" },
  { region: "North", item: "Pear", volume: 3000, month: "August" },
  { region: "East", item: "Grape", volume: 9000, month: "November" },
  { region: "North", item: "Orange", volume: 8000, month: "October" },
  { region: "East", item: "Apple", volume: 10000, month: "June" },
  { region: "South", item: "Pear", volume: 1000, month: "December" },
  { region: "North", item: "Grape", volume: 10000, month: "July" },
  { region: "East", item: "Grape", volume: 6000, month: "February" },
];

const hidden = { hidden: true };

const header = workbook.addFormat();
header.setBold();

/*
 * Example 1. Autofilter without conditions.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet1, header);

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet1.writeString(i + 1, 0, item.region);
  worksheet1.writeString(i + 1, 1, item.item);
  worksheet1.writeNumber(i + 1, 2, item.volume);
  worksheet1.writeString(i + 1, 3, item.month);
}

/* Add the autofilter. */
worksheet1.autofilter(0, 0, 50, 3);

/*
 * Example 2. Autofilter with a filter condition in the first column.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet2, header);

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet2.writeString(i + 1, 0, item.region);
  worksheet2.writeString(i + 1, 1, item.item);
  worksheet2.writeNumber(i + 1, 2, item.volume);
  worksheet2.writeString(i + 1, 3, item.month);

  /* It isn't sufficient to just apply the filter condition below. We
   * must also hide the rows that don't match the criteria since Excel
   * doesn't do that automatically. */
  if (item.region == "East") {
    /* Row matches the filter, no further action required. */
  } else {
    /* Hide rows that don't match the filter. */
    worksheet2.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
  }

  /* Note, the if() statement above is written to match the logic of the
   * criteria in worksheet_filter_column() below. However you could get
   * the same results with the following simpler, but reversed, code:
   *
   *     if (item.region != "East") {
   *         worksheet2.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
   *     }
   *
   * The same applies to the Examples 3-6 as well.
   */
}

/* Add the autofilter. */
worksheet2.autofilter(0, 0, 50, 3);

/* Add the filter criteria. */
worksheet2.filterColumn(0, {
  criteria: Worksheet.EQUAL_TO_FILTER_CRITERIA,
  valueString: "East",
});

/*
 * Example 3. Autofilter with a dual filter condition in one of the columns.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet3, header);

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet3.writeString(i + 1, 0, item.region);
  worksheet3.writeString(i + 1, 1, item.item);
  worksheet3.writeNumber(i + 1, 2, item.volume);
  worksheet3.writeString(i + 1, 3, item.month);

  if (item.region == "East" || item.region == "South") {
    /* Row matches the filter, no further action required. */
  } else {
    /* We need to hide rows that don't match the filter. */
    worksheet3.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
  }
}

/* Add the autofilter. */
worksheet3.autofilter(0, 0, 50, 3);

/* Add the filter criteria. */
worksheet3.filterColumn(
  0,
  { criteria: Worksheet.EQUAL_TO_FILTER_CRITERIA, valueString: "East" },
  { criteria: Worksheet.EQUAL_TO_FILTER_CRITERIA, valueString: "South" },
  Worksheet.OR_FILTER
);

/*
 * Example 4. Autofilter with filter conditions in two columns.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet4, header);

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet4.writeString(i + 1, 0, item.region);
  worksheet4.writeString(i + 1, 1, item.item);
  worksheet4.writeNumber(i + 1, 2, item.volume);
  worksheet4.writeString(i + 1, 3, item.month);

  if (item.region == "East" && item.volume > 3000 && item.volume < 8000) {
    /* Row matches the filter, no further action required. */
  } else {
    /* We need to hide rows that don't match the filter. */
    worksheet4.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
  }
}

/* Add the autofilter. */
worksheet4.autofilter(0, 0, 50, 3);

/* Add the filter criteria. */
worksheet4.filterColumn(0, {
  criteria: Worksheet.EQUAL_TO_FILTER_CRITERIA,
  valueString: "East",
});
worksheet4.filterColumn(
  2,
  { criteria: Worksheet.GREATER_THAN_FILTER_CRITERIA, value: 3000 },
  { criteria: Worksheet.LESS_THAN_FILTER_CRITERIA, value: 8000 },
  Worksheet.AND_FILTER
);

/*
 * Example 5. Autofilter with a dual filter condition in one of the columns.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet5, header);

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet5.writeString(i + 1, 0, item.region);
  worksheet5.writeString(i + 1, 1, item.item);
  worksheet5.writeNumber(i + 1, 2, item.volume);
  worksheet5.writeString(i + 1, 3, item.month);

  if (
    item.region == "East" ||
    item.region == "North" ||
    item.region == "South"
  ) {
    /* Row matches the filter, no further action required. */
  } else {
    /* We need to hide rows that don't match the filter. */
    worksheet5.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
  }
}

/* Add the autofilter. */
worksheet5.autofilter(0, 0, 50, 3);

/* Add the filter criteria. */
worksheet5.filterList(0, ["East", "North", "South"]);

/*
 * Example 6. Autofilter with filter for blanks.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet6, header);

/* Simulate one blank cell in the data, to test the filter. */
data[5].region = "";

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet6.writeString(i + 1, 0, item.region);
  worksheet6.writeString(i + 1, 1, item.item);
  worksheet6.writeNumber(i + 1, 2, item.volume);
  worksheet6.writeString(i + 1, 3, item.month);

  if (item.region == "") {
    /* Row matches the filter, no further action required. */
  } else {
    /* We need to hide rows that don't match the filter. */
    worksheet6.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
  }
}

/* Add the autofilter. */
worksheet6.autofilter(0, 0, 50, 3);

/* Add the filter criteria. */
worksheet6.filterColumn(0, { criteria: Worksheet.BLANKS_FILTER_CRITERIA });

/*
 * Example 7. Autofilter with filter for non-blanks.
 */

/* Set up the worksheet data. */
writeWorksheetHeader(worksheet7, header);

/* Write the row data. */
for (const [i, item] of data.entries()) {
  worksheet7.writeString(i + 1, 0, item.region);
  worksheet7.writeString(i + 1, 1, item.item);
  worksheet7.writeNumber(i + 1, 2, item.volume);
  worksheet7.writeString(i + 1, 3, item.month);

  if (item.region) {
    /* Row matches the filter, no further action required. */
  } else {
    /* We need to hide rows that don't match the filter. */
    worksheet7.setRow(i + 1, Worksheet.DEFAULT_ROW_HEIGHT, null, hidden);
  }
}

/* Add the autofilter. */
worksheet7.autofilter(0, 0, 50, 3);

/* Add the filter criteria. */
worksheet7.filterColumn(0, { criteria: Worksheet.NON_BLANKS_FILTER_CRITERIA });

await fs.writeFile("autofilter.xlsx", Buffer.from(workbook.close()));

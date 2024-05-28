import fs from "fs/promises";

import { Format, Workbook } from "excelwriter";

/* Create a new workbook and add some worksheets. */
const workbook = new Workbook();

const worksheet1 = workbook.addWorksheet("Panes 1");
const worksheet2 = workbook.addWorksheet("Panes 2");
const worksheet3 = workbook.addWorksheet("Panes 3");
const worksheet4 = workbook.addWorksheet("Panes 4");

/* Set up some formatting and text to highlight the panes. */
const header = workbook.addFormat();
header.setAlign(Format.CENTER_ALIGN);
header.setAlign(Format.VERTICAL_CENTER_ALIGN);
header.setFgColor(0xd7e4bc);
header.setBold();
header.setBorder(Format.THIN_BORDER);

const center = workbook.addFormat();
center.setAlign(Format.CENTER_ALIGN);

/*
 * Example 1. Freeze pane on the top row.
 */
worksheet1.freezePanes(1, 0);

/* Some sheet formatting. */
worksheet1.setColumn(0, 8, 16);
worksheet1.setRow(0, 20);
worksheet1.setSelection(4, 3, 4, 3);

/* Some worksheet text to demonstrate scrolling. */
for (let col = 0; col < 9; col++) {
  worksheet1.writeString(0, col, "Scroll down", header);
}

for (let row = 1; row < 100; row++) {
  for (let col = 0; col < 9; col++) {
    worksheet1.writeNumber(row, col, row + 1, center);
  }
}

/*
 * Example 2. Freeze pane on the left column.
 */
worksheet2.freezePanes(0, 1);

/* Some sheet formatting. */
worksheet2.setColumn(0, 0, 16);
worksheet2.setSelection(4, 3, 4, 3);

/* Some worksheet text to demonstrate scrolling. */
for (let row = 0; row < 50; row++) {
  worksheet2.writeString(row, 0, "Scroll right", header);

  for (let col = 1; col < 26; col++) {
    worksheet2.writeNumber(row, col, col, center);
  }
}

/*
 * Example 3. Freeze pane on the top row and left column.
 */
worksheet3.freezePanes(1, 1);

/* Some sheet formatting. */
worksheet3.setColumn(0, 25, 16);
worksheet3.setRow(0, 20);
worksheet3.writeString(0, 0, "", header);
worksheet3.setSelection(4, 3, 4, 3);

/* Some worksheet text to demonstrate scrolling. */
for (let col = 1; col < 26; col++) {
  worksheet3.writeString(0, col, "Scroll down", header);
}

for (let row = 1; row < 50; row++) {
  worksheet3.writeString(row, 0, "Scroll right", header);

  for (let col = 1; col < 26; col++) {
    worksheet3.writeNumber(row, col, col, center);
  }
}

/*
 * Example 4. Split pane on the top row and left column.
 *
 * The divisions must be specified in terms of row and column dimensions.
 * The default row height is 15 and the default column width is 8.43
 */
worksheet4.splitPanes(15, 8.43);

/* Some sheet formatting. */

/* Some worksheet text to demonstrate scrolling. */
for (let col = 1; col < 26; col++) {
  worksheet4.writeString(0, col, "Scroll", center);
}

for (let row = 1; row < 50; row++) {
  worksheet4.writeString(row, 0, "Scroll", center);

  for (let col = 1; col < 26; col++) {
    worksheet4.writeNumber(row, col, col, center);
  }
}

const data = workbook.close();
await fs.writeFile("panes.xlsx", Buffer.from(data));

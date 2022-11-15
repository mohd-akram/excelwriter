import fs from "fs/promises";

import { Workbook } from "excelwriter";

async function main() {
  /* A datetime to display. */
  const datetime = new Date(2013, 1, 28, 12, 0, 0);

  /* Create a new workbook and add a worksheet. */
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  /* Add a format with date formatting. */
  const format = workbook.addFormat();
  format.setNumFormat("mmm d yyyy hh:mm AM/PM");

  /* Widen the first column to make the text clearer. */
  worksheet.setColumn(0, 0, 20);

  /* Write the datetime without formatting. */
  worksheet.writeDatetime(0, 0, datetime); // 41333.5

  /* Write the datetime with formatting. */
  worksheet.writeDatetime(1, 0, datetime, format); // Feb 28 2013 12:00 PM

  const data = workbook.close();
  await fs.writeFile("dates-and-times02.xlsx", Buffer.from(data));
}

main();

import fs from "fs/promises";

import { Workbook } from "excelwriter";

async function main() {
  /* Create a new workbook and add a worksheet. */
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  /* Add a format. */
  const format = workbook.addFormat();

  /* Set the bold property for the format */
  format.setBold();

  /* Change the column width for clarity. */
  worksheet.setColumn(0, 0, 20);
  
  let row = 0;

  /* Write some simple text. */
  worksheet.writeString(row++, 0, "Hello");

  /* Text with formatting. */
  worksheet.writeString(row++, 0, "World", format);
  
  /* Write a url. */
  worksheet.writeUrl(row++, 0, "https://www.npmjs.com/package/excelwriter");

  /* Write some numbers. */
  worksheet.writeNumber(row++, 0, 123);
  worksheet.writeNumber(row++, 0, 123.456);
  
  /* Write a formula. */
  worksheet.writeFormula(row++, 0, "=5+5");
  
  /* Write booleans. */
  worksheet.writeBoolean(row++, 0, false);
  worksheet.writeBoolean(row++, 0, true);

  /* Insert an image. */
  worksheet.insertImage(1, 2, await fs.readFile("logo.png"));

  const data = workbook.close();
  await fs.writeFile("demo.xlsx", Buffer.from(data));
}

main();

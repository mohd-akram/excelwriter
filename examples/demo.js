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

  /* Write some simple text. */
  worksheet.writeString(0, 0, "Hello");

  /* Text with formatting. */
  worksheet.writeString(1, 0, "World", format);

  /* Write some numbers. */
  worksheet.writeNumber(2, 0, 123);
  worksheet.writeNumber(3, 0, 123.456);

  /* Insert an image. */
  worksheet.insertImage(1, 2, await fs.readFile("logo.png"));

  const data = workbook.close();
  await fs.writeFile("demo.xlsx", Buffer.from(data));
}

main();

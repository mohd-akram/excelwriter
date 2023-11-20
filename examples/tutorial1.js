import fs from "fs/promises";

import { Workbook } from "excelwriter";

const expenses = [
  { item: "Rent", cost: 1000 },
  { item: "Gas", cost: 100 },
  { item: "Food", cost: 300 },
  { item: "Gym", cost: 50 },
];

async function main() {
  /* Create a workbook and add a worksheet. */
  const workbook = new Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  /* Start from the first cell. Rows and columns are zero indexed. */
  let row = 0;
  let col = 0;

  /* Iterate over the data and write it out element by element. */
  for (row = 0; row < 4; row++) {
    worksheet.writeString(row, col, expenses[row].item);
    worksheet.writeNumber(row, col + 1, expenses[row].cost);
  }

  /* Write a total using a formula. */
  worksheet.writeString(row, col, "Total");
  worksheet.writeFormula(row, col + 1, "=SUM(B1:B4)");

  /* Save the workbook. */
  const data = workbook.close();
  await fs.writeFile("tutorial01.xlsx", Buffer.from(data));
}

main();

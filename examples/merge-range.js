import fs from "fs/promises";

import { Color, Format, Workbook } from "excelwriter";

const workbook = new Workbook();
const worksheet = workbook.addWorksheet("Sheet1");
const mergeFormat = workbook.addFormat();

/* Configure a format for the merged range. */
mergeFormat.setAlign(Format.CENTER_ALIGN);
mergeFormat.setAlign(Format.VERTICAL_CENTER_ALIGN);
mergeFormat.setBold();
mergeFormat.setBgColor(Color.YELLOW_COLOR);
mergeFormat.setBorder(Format.THIN_BORDER);

/* Increase the cell size of the merged cells to highlight the formatting. */
worksheet.setColumn(1, 3, 12);
worksheet.setRow(3, 30);
worksheet.setRow(6, 30);
worksheet.setRow(7, 30);

/* Merge 3 cells. */
worksheet.mergeRange(3, 1, 3, 3, "Merged Range", mergeFormat);

/* Merge 3 cells over two rows. */
worksheet.mergeRange(6, 1, 7, 3, "Merged Range", mergeFormat);

const data = workbook.close();
await fs.writeFile("merge-range.xlsx", Buffer.from(data));

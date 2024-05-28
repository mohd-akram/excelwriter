import fs from "fs/promises";

import { Chart, Color, Workbook } from "excelwriter";

/**
 * Write some data to the worksheet.
 *
 * @param {Worksheet} worksheet
 */
function writeWorksheetData(worksheet) {
  const data = [
    /* Three columns of data. */
    [1, 2, 3],
    [2, 4, 6],
    [3, 6, 9],
    [4, 8, 12],
    [5, 10, 15],
  ];

  for (const [rowNo, row] of data.entries())
    for (const [colNo, num] of row.entries())
      worksheet.writeNumber(rowNo, colNo, num);
}

/* Create a worksheet with a chart. */
const workbook = new Workbook();
const worksheet = workbook.addWorksheet("Sheet1");

/* Write some data for the chart. */
writeWorksheetData(worksheet);

/* Create a chart object. */
const chart = workbook.addChart(Chart.COLUMN_CHART);

/* Configure the chart. In simplest case we just add some value data
 * series. The null categories will default to 1 to 5 like in Excel.
 */
chart.addSeries(null, "Sheet1!$A$1:$A$5");
chart.addSeries(null, "Sheet1!$B$1:$B$5");
chart.addSeries(null, "Sheet1!$C$1:$C$5");

const font = { bold: false, color: Color.BLUE_COLOR };

chart.setTitleName("Year End Results");
chart.setTitleNameFont(font);

/* Insert the chart into the worksheet. */
worksheet.insertChart(6, 1, chart);

const data = workbook.close();
await fs.writeFile("chart.xlsx", Buffer.from(data));

import fs from "fs/promises";

import { Color, Format, Workbook } from "excelwriter";

async function main() {
  /* Create a new workbook. */
  const workbook = new Workbook();

  /* Add a worksheet. */
  const worksheet = workbook.addWorksheet("Sheet1");

  /* Get the default url format (used in the overwriting examples below). */
  const urlFormat = workbook.defaultURLFormat;

  /* Create a user defined link format. */
  const redFormat = workbook.addFormat();
  redFormat.setUnderline(Format.SINGLE_UNDERLINE);
  redFormat.setFontColor(Color.RED_COLOR);

  /* Widen the first column to make the text clearer. */
  worksheet.setColumn(0, 0, 30);

  /* Write a hyperlink. A default blue underline will be used if the format is NULL. */
  worksheet.writeURL(0, 0, "http://libxlsxwriter.github.io");

  /* Write a hyperlink but overwrite the displayed string. Note, we need to
   * specify the format for the string to match the default hyperlink. */
  worksheet.writeURL(2, 0, "http://libxlsxwriter.github.io");
  worksheet.writeString(2, 0, "Read the documentation.", urlFormat);

  /* Write a hyperlink with a different format. */
  worksheet.writeURL(4, 0, "http://libxlsxwriter.github.io", redFormat);

  /* Write a mail hyperlink. */
  worksheet.writeURL(6, 0, "mailto:jmcnamara@cpan.org");

  /* Write a mail hyperlink and overwrite the displayed string. We again
   * specify the format for the string to match the default hyperlink. */
  worksheet.writeURL(8, 0, "mailto:jmcnamara@cpan.org");
  worksheet.writeString(8, 0, "Drop me a line.", urlFormat);

  /* Close the workbook and save the file. */
  const data = workbook.close();
  await fs.writeFile("hyperlinks.xlsx", Buffer.from(data));
}

main();

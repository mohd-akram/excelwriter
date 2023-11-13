# excelwriter

libxlsxwriter bindings for Node.js

## Install

```shell
npm install excelwriter
```

## Usage

```javascript
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
```

## Documentation

See the examples directory.

## Building

To build the excelwriter package, first clone this repository then navigate to its root directory.

```
git clone https://github.com/mohd-akram/excelwriter
cd excelwriter
```

Clone the [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter) git submodule dependency.

```
git submodule update --init --recursive
```

Ensure that [node-gyp](https://github.com/nodejs/node-gyp) is globally installed.

```
npm install -g node-gyp
```

Install JavaScript dependencies. This will also initially build the excelwriter package.

```
npm install
```

To rebuild the package after making changes, run:

```
node-gyp rebuild
```

Run an example script to test, like so:

```
(cd examples && node demo.js)
```

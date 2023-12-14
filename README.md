English| [简体中文](./README-zh.md)

## Introduction
Luckyexcel is an excel import and export library adapted to [Luckysheet](https://github.com/mengshukeji/Luckysheet). It only supports .xlsx format files (not .xls).

## Demo
[Demo](https://mengshukeji.github.io/LuckyexcelDemo/)

## Features
Support excel file import to Luckysheet adaptation list

- Cell style
- Cell border
- Cell format, such as number format, date, percentage, etc.
- Formula

### Plan
The goal is to support all features supported by Luckysheet

- Conditional Formatting
- Pivot table
- Chart
- Sort
- Filter
- Annotation
- Excel export

## Usage

### CDN
```html
<script src="https://cdn.jsdelivr.net/npm/luckyexcel/dist/luckyexcel.umd.js"></script>
<script>
    // Make sure to get the xlsx file first, and then use the global method window.LuckyExcel to convert
    LuckyExcel.transformExcelToLucky(
        file, 
        function(exportJson, luckysheetfile){
            // After obtaining the converted table data, use luckysheet to initialize or update the existing luckysheet workbook
            // Note: Luckysheet needs to introduce a dependency package and initialize the table container before it can be used
            luckysheet.create({
                container: 'luckysheet', // luckysheet is the container id
                data:exportJson.sheets,
                title:exportJson.info.name,
                userInfo:exportJson.info.name.creator
            });
        },
        function(err){
            logger.error('Import failed. Is your fail a valid xlsx?');
        });
</script>
```
> Case [Demo index.html](./src/index.html) shows the detailed usage

### ES and Node.js

#### Installation
```shell
npm install luckyexcel
```

#### ES import
```js
import LuckyExcel from 'luckyexcel'

// After getting the xlsx file
LuckyExcel.transformExcelToLucky(data, 
    function(exportJson, luckysheetfile){
        // Get the worksheet data after conversion
    },
    function(error){
        // handle error if any thrown
    }
)
```
> Case [luckysheet-vue](https://github.com/mengshukeji/luckysheet-vue)

#### Node.js import
```js
var fs = require("fs");
var LuckyExcel = require('luckyexcel');

// Read an xlsx file
fs.readFile("House cleaning checklist.xlsx", function(err, data) {
    if (err) throw err;

    LuckyExcel.transformExcelToLucky(data, function(exportJson, luckysheetfile){
        // Get the worksheet data after conversion
    });
    
});
```
> Case [Luckyexcel-node](https://github.com/mengshukeji/Luckyexcel-node)

## Development

### Requirements
[Node.js](https://nodejs.org/en/) Version >= 6 

### Installation
```
npm install -g gulp-cli
npm install
```
### Development
```
npm run dev
```
### Package
```
npm run build
```

A third-party plug-in is used in the project: [JSZip](https://github.com/Stuk/jszip), thanks!

## Communication

- Any questions or suggestions are welcome to submit [Issues](https://github.com/mengshukeji/Luckyexcel/issues/)

- [Gitter](https://gitter.im/mengshukeji/Luckysheet)

[Chinese community](./README-zh.md)

## Authors and acknowledgment
- [@wbfsa](https://github.com/wbfsa)
- [@wpxp123456](https://github.com/wpxp123456)
- [@Dushusir](https://github.com/Dushusir)
- [@xxxDeveloper](https://github.com/xxxDeveloper)

## License
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, mengshukeji

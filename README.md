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

## Requirements
[Node.js](https://nodejs.org/en/) Version >= 6 

## Installation
```
npm install -g gulp-cli
npm install
gulp
```

A third-party plug-in is used in the project: [JSZip](https://github.com/Stuk/jszip), thanks!

## Usage (under improvement)

#### Step 1
After `gulp build`, copy bundle.js in the `dist` folder to the project directory, and bundle.js is the core code of the project

#### Step 2

Import bundle.js, specify a file upload component on the interface, write a monitoring method similar to the following, call `LuckyExcel.transformExcelToLucky`, and then get the converted JSON data in the callback. This JSON data is in a format that Luckysheet can recognize. Use Luckysheet to initialize.
```js
function demoHandler(){
    let upload = document.getElementById("Luckyexcel-demo-file");
    if(upload){
        
        window.onload = () => {
            
            upload.addEventListener("change", function(evt){
                var files:FileList = (evt.target as any).files;
                LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any){

                    window.luckysheet.destroy();
                    
                    window.luckysheet.create({
                        container:'luckysheet', //luckysheet is the container id
                        data:exportJson.sheets,
                        title:exportJson.info.name,
                        userInfo:exportJson.info.name.creator
                    });
                });
            });
        }
    }
}
```

## Communication

- Any questions or suggestions are welcome to submit [Issues](https://github.com/mengshukeji/Luckyexcel/issues/)

- [Gitter](https://gitter.im/mengshukeji/Luckysheet)

[Chinese community](./README-zh.md)

## Authors and acknowledgment
- [@wbfsa](https://github.com/wbfsa)
- [@wpxp123456](https://github.com/wpxp123456)
- [@Dushusir](https://github.com/Dushusir)

## License
[MIT](http://opensource.org/licenses/MIT)

Copyright (c) 2020-present, mengshukeji

import { LuckyFile } from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import {HandleZip} from './HandleZip';

import {IuploadfileList} from "./ICommon";
import { fstat } from "fs";

// //demo
// function demoHandler(){
//     let upload = document.getElementById("Luckyexcel-demo-file");
//     let selectADemo = document.getElementById("Luckyexcel-select-demo");
//     let downlodDemo = document.getElementById("Luckyexcel-downlod-file");
//     let mask = document.getElementById("lucky-mask-demo");
//     if(upload){
        
//         window.onload = () => {
            
//             upload.addEventListener("change", function(evt){
//                 var files:FileList = (evt.target as any).files;
//                 if(files==null || files.length==0){
//                     alert("No files wait for import");
//                     return;
//                 }

//                 let name = files[0].name;
//                 let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length-1];
//                 if(suffix!="xlsx"){
//                     alert("Currently only supports the import of xlsx files");
//                     return;
//                 }
//                 LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any, luckysheetfile:string){
                    
//                     if(exportJson.sheets==null || exportJson.sheets.length==0){
//                         alert("Failed to read the content of the excel file, currently does not support xls files!");
//                         return;
//                     }
//                     console.log(exportJson, luckysheetfile);
//                     window.luckysheet.destroy();
                    
//                     window.luckysheet.create({
//                         container: 'luckysheet', //luckysheet is the container id
//                         showinfobar:false,
//                         data:exportJson.sheets,
//                         title:exportJson.info.name,
//                         userInfo:exportJson.info.name.creator
//                     });
//                 });
//             });

//             selectADemo.addEventListener("change", function(evt){
//                 var obj:any = selectADemo;
//                 var index = obj.selectedIndex;
//                 var value = obj.options[index].value;
//                 var name = obj.options[index].innerHTML;
//                 if(value==""){
//                     return;
//                 }
//                 mask.style.display = "flex";
//                 LuckyExcel.transformExcelToLuckyByUrl(value, name, function(exportJson:any, luckysheetfile:string){
                    
//                     if(exportJson.sheets==null || exportJson.sheets.length==0){
//                         alert("Failed to read the content of the excel file, currently does not support xls files!");
//                         return;
//                     }
//                     console.log(exportJson, luckysheetfile);
//                     mask.style.display = "none";
//                     window.luckysheet.destroy();
                    
//                     window.luckysheet.create({
//                         container: 'luckysheet', //luckysheet is the container id
//                         showinfobar:false,
//                         data:exportJson.sheets,
//                         title:exportJson.info.name,
//                         userInfo:exportJson.info.name.creator
//                     });
//                 });
//             });

//             downlodDemo.addEventListener("click", function(evt){
//                 var obj:any = selectADemo;
//                 var index = obj.selectedIndex;
//                 var value = obj.options[index].value;

//                 if(value.length==0){
//                     alert("Please select a demo file");
//                     return;
//                 }

//                 var elemIF:any = document.getElementById("Lucky-download-frame");
//                 if(elemIF==null){
//                     elemIF = document.createElement("iframe");
//                     elemIF.style.display = "none";
//                     elemIF.id = "Lucky-download-frame";
//                     document.body.appendChild(elemIF);
//                 }
//                 elemIF.src = value;

//                 // elemIF.parentNode.removeChild(elemIF);
//             });
//         }
//     }
// }
// demoHandler();

// api
export class LuckyExcel{
    static transformExcelToLucky(excelFile: File,
        callback?: (files: IuploadfileList, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip:HandleZip = new HandleZip(excelFile);
        
        handleZip.unzipFile(function (files: IuploadfileList) {
            let luckyFile = new LuckyFile(files, excelFile.name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            if (callback != undefined) {
                callback(exportJson, luckysheetfile);
            }
        },
        function(err:Error){
            if (errorHandler) {
                errorHandler(err);
              } else {
                console.error(err);
              }
        });
    }

    static transformExcelToLuckyByUrl(
        url: string,
        name: string,
        callBack?: (files: IuploadfileList, fs?: string) => void,
        errorHandler?: (err: Error) => void) {
        let handleZip:HandleZip = new HandleZip();
        handleZip.unzipFileByUrl(url, function(files:IuploadfileList){
            let luckyFile = new LuckyFile(files, name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            if(callBack != undefined){
                callBack(exportJson, luckysheetfile);
            }
        },
        function(err:Error){
            if (errorHandler) {
                errorHandler(err);
              } else {
                console.error(err);
              }
        });
    }

    static transformLuckyToExcel(
        LuckyFile: any,
        callBack?: (files: string) => void ,
        errorHandler?: (err: Error) => void){ }
}
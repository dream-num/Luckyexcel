import { LuckyFile } from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import {HandleZip} from './HandleZip';

import {IuploadfileList} from "./ICommon";
import { LuckySheet } from "./ToLuckySheet/LuckySheet";

//demo
function demoHandler(){
    let upload = document.getElementById("Luckyexcel-demo-file");
    if(upload){
        
        window.onload = () => {
            
            upload.addEventListener("change", function(evt){
                var files:FileList = (evt.target as any).files;
                if(files==null || files.length==0){
                    alert("No files wait for import");
                    return;
                }

                let name = files[0].name;
                let suffix = name.split(".")[1];
                if(suffix!="xlsx"){
                    alert("Currently only supports the import of xlsx files");
                    return;
                }
                LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any, luckysheetfile:string){
                    
                    if(exportJson.sheets==null || exportJson.sheets.length==0){
                        alert("Failed to read the content of the excel file, currently does not support xls files!");
                        return;
                    }
                    console.log(exportJson, luckysheetfile);
                    window.luckysheet.destroy();
                    
                    window.luckysheet.create({
                        container: 'luckysheet', //luckysheet is the container id
                        showinfobar:false,
                        data:exportJson.sheets,
                        title:exportJson.info.name,
                        userInfo:exportJson.info.name.creator
                    });
                });
            });
        }
    }
}
demoHandler();

// api
export class LuckyExcel{
    static transformExcelToLucky(excelFile:File, callBack?:(files:IuploadfileList, fs?:string)=>void){
        let handleZip:HandleZip = new HandleZip(excelFile);
        handleZip.unzipFile(function(files:IuploadfileList){
            let luckyFile = new LuckyFile(files, excelFile.name);
            let luckysheetfile = luckyFile.Parse();
            let exportJson = JSON.parse(luckysheetfile);
            if(callBack != undefined){
                callBack(exportJson, luckysheetfile);
            }
            
        },
        function(err:Error){
            console.error(err);
        });
    }

    static transformLuckyToExcel(LuckyFile: any, callBack?: (files: string) => void) {
        
    }
}




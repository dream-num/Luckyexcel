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
                LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any, luckysheetfile:string){
                    console.log(exportJson, luckysheetfile);
                    window.luckysheet.destroy();
                    
                    window.luckysheet.create({
                        container: 'luckysheet', //luckysheet is the container id
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




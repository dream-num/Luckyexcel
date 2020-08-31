import JSZip from "jszip";
import {IuploadfileList, IWorkBook} from "./ICommon";


export class HandleZip{
    uploadFile:File; 
    workBook:IWorkBook; 
    constructor(file:File | IWorkBook){
        if(file instanceof File){
            this.uploadFile = file;
        }
        else {
            this.workBook = file;
        }
    }

    unzipFile(successFunc:(file:IuploadfileList)=>void, errorFunc:(err:Error)=>void):void { 
        var new_zip:JSZip = new JSZip();
        new_zip.loadAsync(this.uploadFile)                                   // 1) read the Blob
        .then(function(zip:any) {
            let fileList:IuploadfileList = <IuploadfileList>{}, lastIndex:number = Object.keys(zip.files).length, index:number=0;
            zip.forEach(function (relativePath:any, zipEntry:any) {  // 2) print entries
                zipEntry.async("string").then(function (data:string) {
                    fileList[zipEntry.name] = data;
                    console.log(lastIndex, index);
                    if(lastIndex==index+1){
                        successFunc(fileList);
                    }
                    index++;
                });
            });
            
        }, function (e:Error) {
            errorFunc(e);
        });
    }

    zipFile(workBook:IWorkBook):void { 
        
    }
}
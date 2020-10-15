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
        new_zip.loadAsync(this.uploadFile,{base64: true})                                   // 1) read the Blob
        .then(function(zip:any) {
            let fileList:IuploadfileList = <IuploadfileList>{}, lastIndex:number = Object.keys(zip.files).length, index:number=0;
            zip.forEach(function (relativePath:any, zipEntry:any) {  // 2) print entries
                let fileName = zipEntry.name;
                let fileNameArr = fileName.split(".");
                let suffix = fileNameArr[fileNameArr.length-1].toLowerCase();
                let fileType = "string";
                if(suffix in {"png":1, "jpeg":1, "jpg":1, "gif":1,"bmp":1,"tif":1,"webp":1,}){
                    fileType = "base64";
                }
                zipEntry.async(fileType).then(function (data:string) {
                    if(fileType=="base64"){
                        data = "data:image/"+ suffix +";base64," + data;
                    }
                    fileList[zipEntry.name] = data;
                    // console.log(lastIndex, index);
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
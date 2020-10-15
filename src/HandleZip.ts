import JSZip from "jszip";
import {IuploadfileList, IWorkBook} from "./ICommon";
import {getBinaryContent} from "./common/method"


export class HandleZip{
    uploadFile:File; 
    workBook:IWorkBook; 
    constructor(file:File | IWorkBook | undefined){
        if(file instanceof File){
            this.uploadFile = file;
        }
        else {
            this.workBook = file;
        }
    }

    unzipFile(successFunc:(file:IuploadfileList)=>void, errorFunc:(err:Error)=>void):void { 
        // var new_zip:JSZip = new JSZip();
        JSZip.loadAsync(this.uploadFile)                                   // 1) read the Blob
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

    unzipFileByUrl(url:string,successFunc:(file:IuploadfileList)=>void, errorFunc:(err:Error)=>void):void { 
        var new_zip:JSZip = new JSZip();
        getBinaryContent(url, function(err:any, data:any) {
            if(err) {
                throw err; // or handle err
            }
        
            JSZip.loadAsync(data).then(function(zip:any) {
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
        });
        
    }

    zipFile(workBook:IWorkBook):void { 
        
    }
}
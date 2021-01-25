import JSZip from "jszip";
import {IdownloadfileList} from "./ICommon"
import {HandleZip} from './HandleZip'

export class ZipFile extends HandleZip{
    downloadFile:IdownloadfileList;

    constructor(fileList:IdownloadfileList){
        super();
        this.downloadFile = fileList;
    }

    // zip all sheets XML data to files
    zipFiles(successFunc:(content:Blob)=>void,errorFunc:(err:Error)=>void):void{
        
        // var zip = new JSZip(); // Create a ZIP file
        // zip.file('_rels/.rels', sheets.toRels());// Add WorkBook RELS   
        // var xl = zip.folder('xl');// Add a XL folder for sheets
        // xl.file('workbook.xml', sheets.toWorkBook());// And a WorkBook
        // xl.file('styles.xml', styles.toStyleSheet());// Add styles
        // xl.file('_rels/workbook.xml.rels', sheets.toWorkBookRels());// Add WorkBook RELs
        // zip.file('[Content_Types].xml', sheets.toContentType());// Add content types
        // sheets.fileData(xl);// Zip the rest    

        const files = this.downloadFile;
        for(let filename in files){
            // todo:paths parse
            this.addToZipFile(filename,files[filename])
        }

        this.workBook.generateAsync({ type: "blob",mimeType:"application/vnd.ms-excel" }).then(function (content:Blob) { 
             successFunc(content)
        },function(err:Error){
            errorFunc(err);
        }); 
    }

}
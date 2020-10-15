import { IluckyImage } from "./ILuck";
import {LuckySheetCelldata} from "./LuckyCell";
import { IuploadfileList, IattributeList } from "../ICommon";
import {getXmlAttibute, getColumnWidthPixel, fromulaRef,getRowHeightPixel,getcellrange} from "../common/method";
import {borderTypes} from "../common/constant";
import { ReadXml, IStyleCollections, Element,getColor } from "./ReadXml";
import { LuckyImageBase } from "./LuckyBase";


export class ImageList {
    private images:IattributeList
    constructor(files:IuploadfileList) {
        if(files==null){
            return;
        }
        this.images = {};
        for(let fileKey in files){
            // let reg = new RegExp("xl/media/image1.png", "g");
            if(fileKey.indexOf("xl/media/")>-1){
                let fileNameArr = fileKey.split(".");
                let suffix = fileNameArr[fileNameArr.length-1].toLowerCase();
                if(suffix in {"png":1, "jpeg":1, "jpg":1, "gif":1,"bmp":1,"tif":1,"webp":1,}){
                    this.images[fileKey] = files[fileKey];
                }
            }
        }
    }

    getImageByName(pathName:string):Image{
        if(pathName in this.images){
            let base64 = this.images[pathName];
            return new Image(pathName, base64);
        }
        return null;
    }
}


class Image extends LuckyImageBase {

    fromCol:number
    fromColOff:number
    fromRow:number
    fromRowOff:number

    toCol:number
    toColOff:number
    toRow:number
    toRowOff:number

    constructor(pathName:string, base64:string) {
        super();
        this.src = base64;
    }

    setDefault(){

    }
}
import { IluckyImage } from "./ILuck";
import {LuckySheetCelldata} from "./LuckyCell";
import { IuploadfileList, IattributeList } from "../ICommon";
import {getXmlAttibute, getColumnWidthPixel, fromulaRef,getRowHeightPixel,getcellrange} from "../common/method";
import {borderTypes} from "../common/constant";
import { ReadXml, IStyleCollections, Element,getColor } from "./ReadXml";
import { LuckyImageBase } from "./LuckyBase";
import { UDOC,FromEMF,ToContext2D  } from "../common/emf";


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
                if(suffix in {"png":1, "jpeg":1, "jpg":1, "gif":1,"bmp":1,"tif":1,"webp":1,"emf":1}){
                    if(suffix=="emf"){
                        var pNum  = 0;  // number of the page, that you want to render
                        var scale = 1;  // the scale of the document
                        var wrt = new ToContext2D(pNum, scale);
                        var inp, out, stt;
                        FromEMF.K = [];
                        inp = FromEMF.C;   out = FromEMF.K;   stt=4;
                        for(var p in inp) out[inp[p]] = p.slice(stt);
                        FromEMF.Parse(files[fileKey], wrt);
                        this.images[fileKey] = wrt.canvas.toDataURL("image/png");
                    }
                    else{
                        this.images[fileKey] = files[fileKey];
                    }
                    
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
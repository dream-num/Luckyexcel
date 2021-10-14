import { ILuckyFile,IluckySheetRowAndColumnHidden,IluckySheetRowAndColumnLen} from "./ILuck";
import { LuckySheet} from "./LuckySheet";
import {IuploadfileList, IattributeList} from "../ICommon";
import {workBookFile, coreFile, appFile, stylesFile, sharedStringsFile,numFmtDefault,theme1File,calcChainFile,workbookRels, numFmtDefaultMap} from "../common/constant";
import { ReadXml,IStyleCollections,Element } from "./ReadXml";
import {getXmlAttibute} from "../common/method";
import { LuckyFileBase,LuckyFileInfo,LuckySheetBase,LuckySheetCelldataBase,LuckySheetCelldataValue,LuckySheetCellFormat } from "./LuckyBase";
import {ImageList} from "./LuckyImage";

export class LuckyFile extends LuckyFileBase {

    private files:IuploadfileList
    private sheetNameList:IattributeList
    private readXml:ReadXml
    private fileName:string
    private styles:IStyleCollections
    private sharedStrings:Element[]
    private calcChain:Element[]
    private imageList:ImageList

    constructor(files:IuploadfileList, fileName:string) {
        super();
        this.files = files;
        this.fileName = fileName;
        this.readXml = new ReadXml(files);
        this.getSheetNameList();

        this.sharedStrings = this.readXml.getElementsByTagName("sst/si", sharedStringsFile);
        this.calcChain = this.readXml.getElementsByTagName("calcChain/c", calcChainFile);
        this.styles = {};
        this.styles["cellXfs"] =  this.readXml.getElementsByTagName("cellXfs/xf", stylesFile);
        this.styles["cellStyleXfs"] =  this.readXml.getElementsByTagName("cellStyleXfs/xf", stylesFile);
        this.styles["cellStyles"] =  this.readXml.getElementsByTagName("cellStyles/cellStyle", stylesFile);
        this.styles["fonts"] =  this.readXml.getElementsByTagName("fonts/font", stylesFile);
        this.styles["fills"] =  this.readXml.getElementsByTagName("fills/fill", stylesFile);
        this.styles["borders"] =  this.readXml.getElementsByTagName("borders/border", stylesFile);
        this.styles["clrScheme"] =  this.readXml.getElementsByTagName("a:clrScheme/a:dk1|a:lt1|a:dk2|a:lt2|a:accent1|a:accent2|a:accent3|a:accent4|a:accent5|a:accent6|a:hlink|a:folHlink", theme1File);
        this.styles["indexedColors"] =  this.readXml.getElementsByTagName("colors/indexedColors/rgbColor", stylesFile);
        this.styles["mruColors"] =  this.readXml.getElementsByTagName("colors/mruColors/color", stylesFile);

        this.imageList = new ImageList(files);

        let numfmts =  this.readXml.getElementsByTagName("numFmt/numFmt", stylesFile);
        let numFmtDefaultC = JSON.parse(JSON.stringify(numFmtDefault));
        for(let i=0;i<numfmts.length;i++){
            let attrList = numfmts[i].attributeList;
            let numfmtid = getXmlAttibute(attrList, "numFmtId", "49");
            let formatcode = getXmlAttibute(attrList, "formatCode", "@");
            // console.log(numfmtid, formatcode);
            if(!(numfmtid in numFmtDefault)){
                numFmtDefaultC[numfmtid] = numFmtDefaultMap[formatcode] || formatcode;
            }
        }

        // console.log(JSON.stringify(numFmtDefaultC), numfmts);
        this.styles["numfmts"] =  numFmtDefaultC;
    }

    /**
    * @return All sheet name of workbook
    */
    private getSheetNameList(){
        let workbookRelList = this.readXml.getElementsByTagName("Relationships/Relationship", workbookRels);
        if(workbookRelList==null){
            return;
        }

        let regex = new RegExp("worksheets/[^/]*?.xml");
        let sheetNames:IattributeList = {};
        for(let i=0;i<workbookRelList.length;i++){
            let rel = workbookRelList[i], attrList = rel.attributeList;
            let id = attrList["Id"], target = attrList["Target"];
            if(regex.test(target)){
                if(target.indexOf('/xl') === 0){
                    sheetNames[id] =   target.substr(1);
                }else{
                    sheetNames[id] =   "xl/" + target;
                }
            }

        }

        this.sheetNameList = sheetNames;
    }

    /**
    * @param sheetName WorkSheet'name
    * @return sheet file name and path in zip
    */
   private getSheetFileBysheetId(sheetId:string){
        // for(let i=0;i<this.sheetNameList.length;i++){
        //     let sheetFileName = this.sheetNameList[i];
        //     if(sheetFileName.indexOf("sheet"+sheetId)>-1){
        //         return sheetFileName;
        //     }
        // }
        return this.sheetNameList[sheetId];
    }

    /**
    * @return workBook information
    */
    getWorkBookInfo(){
        let Company = this.readXml.getElementsByTagName("Company", appFile);
        let AppVersion = this.readXml.getElementsByTagName("AppVersion", appFile);
        let creator = this.readXml.getElementsByTagName("dc:creator", coreFile);
        let lastModifiedBy = this.readXml.getElementsByTagName("cp:lastModifiedBy", coreFile);
        let created = this.readXml.getElementsByTagName("dcterms:created", coreFile);
        let modified = this.readXml.getElementsByTagName("dcterms:modified", coreFile);
        this.info = new LuckyFileInfo();
        this.info.name = this.fileName;
        this.info.creator = creator.length>0?creator[0].value:"";
        this.info.lastmodifiedby = lastModifiedBy.length>0?lastModifiedBy[0].value:"";
        this.info.createdTime = created.length>0?created[0].value:"";
        this.info.modifiedTime = modified.length>0?modified[0].value:"";
        this.info.company = Company.length>0?Company[0].value:"";
        this.info.appversion = AppVersion.length>0?AppVersion[0].value:"";
    }

    /**
    * @return All sheet , include whole information
    */
    getSheetsFull(isInitialCell:boolean=true){
        let sheets = this.readXml.getElementsByTagName("sheets/sheet", workBookFile);
        let sheetList:IattributeList = {};
        for(let key in sheets){
            let sheet = sheets[key];
            sheetList[sheet.attributeList.name] = sheet.attributeList["sheetId"];
        }
        this.sheets = [];
        let order = 0;
        for(let key in sheets){
            let sheet = sheets[key];
            let sheetName = sheet.attributeList.name;
            let sheetId = sheet.attributeList["sheetId"];
            let rid = sheet.attributeList["r:id"];
            let sheetFile = this.getSheetFileBysheetId(rid);
            let hide = sheet.attributeList.state === "hidden" ? 1 : 0;

            let drawing = this.readXml.getElementsByTagName("worksheet/drawing", sheetFile), drawingFile, drawingRelsFile;
            if(drawing!=null && drawing.length>0){
                let attrList = drawing[0].attributeList;
                let rid = getXmlAttibute(attrList, "r:id", null);
                if(rid!=null){
                    drawingFile = this.getDrawingFile(rid, sheetFile);
                    drawingRelsFile = this.getDrawingRelsFile(drawingFile);
                }
            }

            if(sheetFile!=null){
                let sheet = new LuckySheet(sheetName, sheetId, order, isInitialCell,
                    {
                        sheetFile:sheetFile,
                        readXml:this.readXml,
                        sheetList:sheetList,
                        styles:this.styles,
                        sharedStrings:this.sharedStrings,
                        calcChain:this.calcChain,
                        imageList:this.imageList,
                        drawingFile:drawingFile,
                        drawingRelsFile: drawingRelsFile,
                        hide: hide,
                    }
                )
                this.columnWidthSet = [];
                this.rowHeightSet = [];

                this.imagePositionCaculation(sheet);

                this.sheets.push(sheet);
                order++;
            }
        }
    }

    private columnWidthSet:number[] = [];
    private rowHeightSet:number[] = [];

    private extendArray(index:number, sets:number[],def:number, hidden:IluckySheetRowAndColumnHidden, lens:IluckySheetRowAndColumnLen){
        if(index<sets.length){
            return;
        }

        let startIndex = sets.length, endIndex = index;
        let allGap = 0;
        if(startIndex>0){
            allGap = sets[startIndex-1];
        }
        // else{
        //     sets.push(0);
        // }
        for(let i=startIndex;i<=endIndex;i++){
            let gap = def, istring  = i.toString();
            if(istring in hidden){
                gap = 0;
            }
            else if(istring in lens){
                gap = lens[istring];
            }

            allGap += Math.round(gap + 1);

            sets.push(allGap);
        }
    }

    private imagePositionCaculation(sheet:LuckySheet){
        let images = sheet.images, defaultColWidth = sheet.defaultColWidth, defaultRowHeight = sheet.defaultRowHeight;
        let colhidden = {};
        if(sheet.config.colhidden){
            colhidden = sheet.config.colhidden;
        }

        let columnlen = {};
        if(sheet.config.columnlen){
            columnlen = sheet.config.columnlen;
        }

        let rowhidden = {};
        if(sheet.config.rowhidden){
            rowhidden = sheet.config.rowhidden;
        }

        let rowlen = {};
        if(sheet.config.rowlen){
            rowlen = sheet.config.rowlen;
        }

        for(let key in images){
            let imageObject:any = images[key];//Image, luckyImage
            let fromCol = imageObject.fromCol;
            let fromColOff = imageObject.fromColOff;
            let fromRow = imageObject.fromRow;
            let fromRowOff = imageObject.fromRowOff;

            let toCol = imageObject.toCol;
            let toColOff = imageObject.toColOff;
            let toRow = imageObject.toRow;
            let toRowOff = imageObject.toRowOff;

            let x_n =0,y_n = 0;
            let cx_n = 0, cy_n = 0;

            if(fromCol>=this.columnWidthSet.length){
                this.extendArray(fromCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
            }
            if(fromCol==0){
                x_n = 0;
            }
            else{
                x_n = this.columnWidthSet[fromCol-1];
            }
            x_n = x_n + fromColOff;

            if(fromRow>=this.rowHeightSet.length){
                this.extendArray(fromRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
            }
            if(fromRow==0){
                y_n = 0;
            }
            else{
                y_n = this.rowHeightSet[fromRow-1];
            }
            y_n = y_n + fromRowOff;


            if(toCol>=this.columnWidthSet.length){
                this.extendArray(toCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
            }
            if(toCol==0){
                cx_n = 0;
            }
            else{
                cx_n = this.columnWidthSet[toCol-1];
            }
            cx_n = cx_n + toColOff- x_n;

            if(toRow>=this.rowHeightSet.length){
                this.extendArray(toRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
            }
            if(toRow==0){
                cy_n = 0;
            }
            else{
                cy_n = this.rowHeightSet[toRow-1];
            }

            cy_n = cy_n + toRowOff - y_n;

            console.log(defaultColWidth, colhidden , columnlen);
            console.log(fromCol, this.columnWidthSet[fromCol] , fromColOff);
            console.log(toCol, this.columnWidthSet[toCol] , toColOff, JSON.stringify(this.columnWidthSet));

            imageObject.originWidth = cx_n;
            imageObject.originHeight = cy_n;

            imageObject.crop.height = cy_n;
            imageObject.crop.width = cx_n;

            imageObject.default.height = cy_n;
            imageObject.default.left = x_n;
            imageObject.default.top = y_n;
            imageObject.default.width = cx_n;
        }

        console.log(this.columnWidthSet, this.rowHeightSet);
    }

    /**
    * @return drawing file string
    */
   private getDrawingFile(rid:string, sheetFile:string):string{
        let sheetRelsPath = "xl/worksheets/_rels/";
        let sheetFileArr = sheetFile.split("/");
        let sheetRelsName = sheetFileArr[sheetFileArr.length-1];

        let sheetRelsFile = sheetRelsPath + sheetRelsName + ".rels";

        let drawing = this.readXml.getElementsByTagName("Relationships/Relationship", sheetRelsFile);
        if(drawing.length>0){
            for(let i=0;i<drawing.length;i++){
                let relationship = drawing[i];
                let attrList = relationship.attributeList;
                let relationshipId = getXmlAttibute(attrList, "Id", null);
                if(relationshipId==rid){
                    let target = getXmlAttibute(attrList, "Target", null);
                    if(target!=null){
                        return target.replace(/\.\.\//g, "");
                    }
                }
            }
        }

        return null;
    }
    private getDrawingRelsFile(drawingFile:string):string{
        let drawingRelsPath = "xl/drawings/_rels/";
        let drawingFileArr = drawingFile.split("/");
        let drawingRelsName = drawingFileArr[drawingFileArr.length-1];

        let drawingRelsFile = drawingRelsPath + drawingRelsName + ".rels";

        return drawingRelsFile;
    }

    /**
    * @return All sheet base information widthout cell and config
    */
    getSheetsWithoutCell(){
        this.getSheetsFull(false);
    }

    /**
    * @return LuckySheet file json
    */
    Parse():string{
        // let xml = this.readXml;
        // for(let key in this.sheetNameList){
        //     let sheetName=this.sheetNameList[key];
        //     let sheetColumns = xml.getElementsByTagName("row/c/f", sheetName);
        //     console.log(sheetColumns);
        // }
        // return "";

        this.getWorkBookInfo();
        this.getSheetsFull();

        // for(let i=0;i<this.sheets.length;i++){
        //     let sheet = this.sheets[i];
        //     let _borderInfo = sheet.config._borderInfo;
        //     if(_borderInfo==null){
        //         continue;
        //     }
        //     let _borderInfoKeys = Object.keys(_borderInfo);
        //     _borderInfoKeys.sort();
        //     for(let a=0;a<_borderInfoKeys.length;a++){
        //         let key = parseInt(_borderInfoKeys[a]);
        //         let b = _borderInfo[key];
        //         if(b.cells.length==0){
        //             continue;
        //         }
        //         if(sheet.config.borderInfo==null){
        //             sheet.config.borderInfo = [];
        //         }
        //         sheet.config.borderInfo.push(b);
        //     }
        // }

        return this.toJsonString(this);
    }

    private toJsonString(file:ILuckyFile):string{
        let LuckyOutPutFile = new LuckyFileBase();
        LuckyOutPutFile.info = file.info;
        LuckyOutPutFile.sheets = [];

        file.sheets.forEach((sheet)=>{
            let sheetout = new LuckySheetBase();
            //let attrName = ["name","color","config","index","status","order","row","column","luckysheet_select_save","scrollLeft","scrollTop","zoomRatio","showGridLines","defaultColWidth","defaultRowHeight","celldata","chart","isPivotTable","pivotTable","luckysheet_conditionformat_save","freezen","calcChain"];

            if(sheet.name!=null){
                sheetout.name = sheet.name;
            }

            if(sheet.color!=null){
                sheetout.color = sheet.color;
            }

            if(sheet.config!=null){
                sheetout.config = sheet.config;
                // if(sheetout.config._borderInfo!=null){
                //     delete sheetout.config._borderInfo;
                // }
            }

            if(sheet.index!=null){
                sheetout.index = sheet.index;
            }

            if(sheet.status!=null){
                sheetout.status = sheet.status;
            }

            if(sheet.order!=null){
                sheetout.order = sheet.order;
            }

            if(sheet.row!=null){
                sheetout.row = sheet.row;
            }

            if(sheet.column!=null){
                sheetout.column = sheet.column;
            }

            if(sheet.luckysheet_select_save!=null){
                sheetout.luckysheet_select_save = sheet.luckysheet_select_save;
            }

            if(sheet.scrollLeft!=null){
                sheetout.scrollLeft = sheet.scrollLeft;
            }

            if(sheet.scrollTop!=null){
                sheetout.scrollTop = sheet.scrollTop;
            }

            if(sheet.zoomRatio!=null){
                sheetout.zoomRatio = sheet.zoomRatio;
            }

            if(sheet.showGridLines!=null){
                sheetout.showGridLines = sheet.showGridLines;
            }

            if(sheet.defaultColWidth!=null){
                sheetout.defaultColWidth = sheet.defaultColWidth;
            }

            if(sheet.defaultRowHeight!=null){
                sheetout.defaultRowHeight = sheet.defaultRowHeight;
            }

            if(sheet.celldata!=null){
                // sheetout.celldata = sheet.celldata;
                sheetout.celldata = [];
                sheet.celldata.forEach((cell)=>{
                    let cellout = new LuckySheetCelldataBase();
                    cellout.r = cell.r;
                    cellout.c = cell.c;
                    cellout.v = cell.v;
                    sheetout.celldata.push(cellout);
                });
            }

            if(sheet.chart!=null){
                sheetout.chart = sheet.chart;
            }

            if(sheet.isPivotTable!=null){
                sheetout.isPivotTable = sheet.isPivotTable;
            }

            if(sheet.pivotTable!=null){
                sheetout.pivotTable = sheet.pivotTable;
            }

            if(sheet.luckysheet_conditionformat_save!=null){
                sheetout.luckysheet_conditionformat_save = sheet.luckysheet_conditionformat_save;
            }

            if(sheet.freezen!=null){
                sheetout.freezen = sheet.freezen;
            }

            if(sheet.calcChain!=null){
                sheetout.calcChain = sheet.calcChain;
            }

            if(sheet.images!=null){
                sheetout.images = sheet.images;
            }
          
            if (sheet.dataVerification != null){
              sheetout.dataVerification = sheet.dataVerification;
            }
          
            if (sheet.hyperlink != null){
                sheetout.hyperlink = sheet.hyperlink;
            }
          
            if (sheet.hide != null) {
              sheetout.hide = sheet.hide;
            }
            
            LuckyOutPutFile.sheets.push(sheetout);
        });

        return JSON.stringify(LuckyOutPutFile);
    }


}

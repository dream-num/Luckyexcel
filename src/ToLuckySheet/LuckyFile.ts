import { ILuckyFile} from "./ILuck";
import { LuckySheet} from "./LuckySheet";
import {IuploadfileList, IattributeList} from "../ICommon";
import {workBookFile, coreFile, appFile, stylesFile, sharedStringsFile,numFmtDefault,theme1File,calcChainFile,workbookRels} from "../common/constant";
import { ReadXml,IStyleCollections,Element } from "./ReadXml";
import {getXmlAttibute} from "../common/method";
import { LuckyFileBase,LuckyFileInfo,LuckySheetBase,LuckySheetCelldataBase,LuckySheetCelldataValue,LuckySheetCellFormat } from "./LuckyBase";


export class LuckyFile extends LuckyFileBase {

    private files:IuploadfileList
    private sheetNameList:IattributeList
    private readXml:ReadXml
    private fileName:string
    private styles:IStyleCollections
    private sharedStrings:Element[]
    private calcChain:Element[]

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

        let numfmts =  this.readXml.getElementsByTagName("numFmt/numFmt", stylesFile);
        let numFmtDefaultC = numFmtDefault;
        for(let i=0;i<numfmts.length;i++){
            let attrList = numfmts[i].attributeList;
            let numfmtid = getXmlAttibute(attrList, "numFmtId", "49");
            let formatcode = getXmlAttibute(attrList, "formatCode", "@");
            // console.log(numfmtid, formatcode);
            if(!(numfmtid in numFmtDefault)){
                numFmtDefaultC[numfmtid] = formatcode;
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
                sheetNames[id] = "xl/" + target;
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
            if(sheetFile!=null){
                this.sheets.push(new LuckySheet(sheetName, sheetId, order, sheetFile,this.readXml, sheetList, this.styles, this.sharedStrings, this.calcChain,isInitialCell));

                order++;
            }
        }
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

    toJsonString(file:ILuckyFile):string{
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
            
            LuckyOutPutFile.sheets.push(sheetout);
        });

        return JSON.stringify(LuckyOutPutFile);
    }


}


import { LuckyFileBase } from "../ToLuckySheet/LuckyBase";
import { ILuckyFile } from "../ToLuckySheet/ILuck";
import { IdownloadfileList } from "../ICommon";

export class ExcelFile extends LuckyFileBase{
    constructor(luckyFile:ILuckyFile){
        super();
        this.info.name = luckyFile.info.name;
        this.sheets = luckyFile.sheets;
    }

    Parse():void{

        // todo: transform json to xml string
        
        // relsFile toRels()
        // workBookFile toWorkBook()
        // stylesFile toStyles()
        // workbookRels toWorkBookRels()
        // contentTypesFile toContentType()
        // worksheetFilePath toWorkSheets()

        
    }
}
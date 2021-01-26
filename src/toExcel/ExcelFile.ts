

import { LuckyFileBase } from "../ToLuckySheet/LuckyBase";
import { ILuckyFile, ILuckyFileInfo, IluckySheet } from "../ToLuckySheet/ILuck";
import { IdownloadfileList } from "../ICommon";

export class ExcelFile implements ILuckyFile{
    
    info:ILuckyFileInfo;
    sheets:IluckySheet[];
    
    constructor(luckyFile:ILuckyFile){
        // super();
        this.info = luckyFile.info;
        this.sheets = luckyFile.sheets;
    }

    Parse():any{

        // todo: transform json to xml string
        
        // relsFile toRels()
        // workBookFile toWorkBook()
        // stylesFile toStyles()
        // workbookRels toWorkBookRels()
        // contentTypesFile toContentType()
        // worksheetFilePath toWorkSheets()

        return this.sheets;

        
    }
}
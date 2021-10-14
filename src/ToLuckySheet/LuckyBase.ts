import { ILuckyFile, ILuckyFileInfo,IluckySheet,IluckySheetCelldata,IluckySheetConfig,IluckySheetCelldataValue,IluckySheetCelldataValueMerge,ILuckySheetCellFormat,IluckySheetConfigMerges,IluckySheetConfigMerge,IMapluckySheetborderInfoCellForImp,IluckySheetborderInfoCellValue,IluckySheetborderInfoCellValueStyle,IluckySheetborderInfoCellForImp,IluckySheetRowAndColumnLen,IluckySheetRowAndColumnHidden,IluckySheetSelection,IluckysheetFrozen,IluckySheetChart,IluckySheetPivotTable,IluckysheetConditionFormat,IluckysheetCalcChain,ILuckyInlineString,IluckyImage,IluckyImageBorder,IluckyImageCrop,IluckyImageDefault,IluckyImages, IluckysheetHyperlink, IluckysheetDataVerification} from "./ILuck";



export class LuckyFileBase implements ILuckyFile {
    info:ILuckyFileInfo
    sheets:IluckySheet[]
}

export class LuckySheetBase implements IluckySheet{
    name:string
    color:string
    config:IluckySheetConfig
    index:string
    status:string
    order:string
    row:number
    column:number
    luckysheet_select_save:IluckySheetSelection[]
    scrollLeft:number
    scrollTop:number
    zoomRatio:number
    showGridLines:string
    defaultColWidth:number
    defaultRowHeight:number

    celldata:IluckySheetCelldata[]
    chart:IluckySheetChart[]

    isPivotTable:boolean
    pivotTable:IluckySheetPivotTable

    luckysheet_conditionformat_save:IluckysheetConditionFormat[]
    freezen:IluckysheetFrozen

    calcChain:IluckysheetCalcChain[]

    images:IluckyImages
    
    dataVerification: IluckysheetDataVerification;
    hyperlink: IluckysheetHyperlink
    hide: number;
    
}

export class LuckyFileInfo implements ILuckyFileInfo{
    name:string
    creator:string
    lastmodifiedby:string
    createdTime:string
    modifiedTime:string
    company:string
    appversion:string
}

export class LuckySheetCelldataBase implements IluckySheetCelldata{
    r:number
    c:number
    v:IluckySheetCelldataValue | string | null
}

export class LuckySheetCelldataValue implements IluckySheetCelldataValue{
    ct: LuckySheetCellFormat | undefined //celltype,Cell value format: text, time, etc. numfmts
    bg: string | undefined//background,#fff000,	fill
    ff: string | undefined//fontfamily, fonts
    fc: string | undefined//fontcolor fonts
    bl: number | undefined//Bold, fonts
    it: number | undefined//italic, fonts
    fs: number | undefined//font size, fonts
    cl: number | undefined//strike, 0 Regular, 1 strikes, fonts
    un: number | undefined//underline, 0 Regular, 1 underlines, fonts
    vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
    ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
    mc: IluckySheetCelldataValueMerge | undefined //Merge Cells, mergeCells
    tr: number | undefined //Text rotation,0: 0、3 Vertical text alignment
    tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment
    v: string | undefined //Original value, v
    m: string | undefined //Display value, v
    f: string | undefined //formula, f
    rt:number | undefined //text rotation angle 0-180 alignment
    qp:number | undefined //quotePrefix, show number as string
}


export class LuckySheetCellFormat implements ILuckySheetCellFormat {
    fa:string
    t:string
    s:LuckyInlineString[] | undefined
}

export class LuckyInlineString implements ILuckyInlineString {
    ff:string | undefined //font family
    fc:string | undefined//font color
    fs:number | undefined//font size
    cl:number | undefined//strike
    un:number | undefined//underline
    bl:number | undefined//blod
    it:number | undefined//italic
    va:number | undefined//1sub and 2super and 0none
    v:string | undefined
}

export class LuckyConfig implements IluckySheetConfig{
    merge:IluckySheetConfigMerges
    borderInfo:IluckySheetborderInfoCellForImp[]
    // _borderInfo: IMapluckySheetborderInfoCellForImp
    rowlen:IluckySheetRowAndColumnLen
    columnlen:IluckySheetRowAndColumnLen
    rowhidden:IluckySheetRowAndColumnHidden
    colhidden:IluckySheetRowAndColumnHidden

    customHeight:IluckySheetRowAndColumnHidden
    customWidth:IluckySheetRowAndColumnHidden
}

export class LuckySheetborderInfoCellForImp implements IluckySheetborderInfoCellForImp{
    rangeType:string
    // cells:string[]
    value:IluckySheetborderInfoCellValue
}

export class LuckySheetborderInfoCellValue implements IluckySheetborderInfoCellValue{
    row_index: number
    col_index: number
    l: IluckySheetborderInfoCellValueStyle
    r: IluckySheetborderInfoCellValueStyle
    t: IluckySheetborderInfoCellValueStyle
    b: IluckySheetborderInfoCellValueStyle
}

export class LuckySheetborderInfoCellValueStyle implements IluckySheetborderInfoCellValueStyle{
    "style": number
    "color": string
}

export class LuckySheetConfigMerge implements IluckySheetConfigMerge{
    r: number
    c: number
    rs: number
    cs: number
}

export class LuckysheetCalcChain implements IluckysheetCalcChain{
    r:number
    c:number
    index:string | undefined
}


export class LuckyImageBase implements IluckyImage{
    border: IluckyImageBorder
    crop: IluckyImageCrop
    default: IluckyImageDefault

    fixedLeft: number
    fixedTop: number
    isFixedPos: Boolean
    originHeight: number
    originWidth: number
    src: string
    type: string
}
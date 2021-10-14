export interface ILuckyFile{
    info:ILuckyFileInfo,//File information, name,password,date,createor etc. 
    sheets:IluckySheet[],//Sheets, include all sheet data
}

export interface ILuckyFileInfo{
    name:string,// File name
    creator:string,//Create user
    lastmodifiedby:string,//Edit user
    createdTime:string,
    modifiedTime:string,
    company:string,
    appversion:string,//Excel version, Is it necessary?
}

export interface IluckySheet{
    name:string,// Sheet name, it will show on sheet bar, must be unique
    color:string,// Sheet color, it will show on sheet bar
    config?:IluckySheetConfig, // Row height, column width, hidden, and so on
    index:string, //A sheet uniquely identifies, 
    status:string, //If 1 , it means current shown sheet, else means hidden
    order:string, //Order of sheet
    row:number,// Sheet the number of rows, contain blank cell
    column:number, // Sheet the number of columns, contain blank cell
    // visibledatarow:number[],
    // visibledatacolumn:number[],
    luckysheet_select_save?:IluckySheetSelection[],//selection defines
    scrollLeft:number,//horizen scroll offset
    scrollTop:number,//verticel scroll offset

    celldata?:IluckySheetCelldata[],// cells
    chart?:IluckySheetChart[],

    isPivotTable:boolean,
    pivotTable?:IluckySheetPivotTable,

    luckysheet_conditionformat_save?:IluckysheetConditionFormat[],
    freezen?:IluckysheetFrozen,

    calcChain?:IluckysheetCalcChain[],
    
    zoomRatio:number, //sheet zoom ratio 10%-400%

    showGridLines:string, // show grid lines

    defaultColWidth:number, //cloumn width pixel
    defaultRowHeight:number, //row height pixel

    images:IluckyImages,//image list
    
    dataVerification: IluckysheetDataVerification;
		hyperlink: IluckysheetHyperlink, // hyperlinks
		hide: number; // sheet hide
}

//luckysheet general selection
export interface IluckySheetSelection{
    row:number[], //selection start row and end row
    column:number[], //selection start column and end column
    sheetIndex:number
}

export interface IluckySheetChart{
    
}

//pivot table interface
export interface IluckySheetPivotTable{
    pivot_select_save:IluckySheetSelection,//Pivot table data source range
    pivotDataSheetIndex:string | undefined, //data source sheet index, index is unique id
    column:IluckySheetPivotTableField[],// column area, include filed 
    row:IluckySheetPivotTableField[], // row area, include filed 
    filter:IluckySheetPivotTableField[], // filter area, include filed 
    filterparm: IluckySheetPivotTablefilterParam,// save param after apply filter
    values:IluckySheetPivotTableField[],
    showType:string,
    pivotDatas:any[][],
    drawPivotTable:boolean,
    pivotTableBoundary:number[],
}

export interface IluckySheetPivotTableField{
    index: number,
    name: string,
    fullname: string,
    sumtype: string,
    nameindex: number
}

export interface IluckySheetPivotTablefilterParam{
    [index:string]:IluckySheetPivotTablefilterParamItem
}

export interface IluckySheetPivotTablefilterParamItem{
    caljs:IluckySheetPivotTablefilterParamItemCaljs,
    rowhidden:IluckySheetPivotTablefilterParamItemRowhidden,
    selected:IluckySheetPivotTablefilterParamItemSelected,
}

export interface IluckySheetPivotTablefilterParamItemCaljs{
    text: string,
    type: string,
    value: string,
    value1: string,
}

export interface IluckySheetPivotTablefilterParamItemRowhidden{
    [index:number]:number
}

export interface IluckySheetPivotTablefilterParamItemSelected{
    [index:number]:number
}


export interface IluckysheetFrozen{
    horizen:number | undefined, //freeze horizen row number
    vertical:number | undefined, //freeze horizen column number
}

export interface IluckysheetConditionFormat{
    type:string, //Option:defualt,databar,colorGradation,icons,
    cellrange:IluckySheetSelection[],//Valid range
    format:string[] | IluckysheetCFDefaultFormat | IluckysheetCFIconsFormat,//style
    conditionName: string | undefined,//Detailed settings,comparison parameters
    conditionRange:IluckySheetSelection[],//Detailed settings,comparison range
    conditionValue:any[],//Detailed settings,comparison value
}

export interface IluckysheetCFDefaultFormat{
    textColor: string | undefined | null,
	cellColor: string | undefined | null
}

export interface IluckysheetCFIconsFormat{
    len: string | number,
    leftMin: string | number,
    top: string | number,
}

export interface IluckysheetCalcChain{
    r:number,
    c:number,
    index:string | undefined,
    // func?:any[],//[true, 152, "=SUBTOTAL(9,OFFSET(F15,ROW(F15:F18)-ROW(F15),1,3))"] 已经计算、终值、公式
}

export interface IluckySheetCelldata{
    r:number,//cell row number
    c:number,//cell column number
    v:IluckySheetCelldataValue | string | null, //cell value
}

export interface IluckySheetCelldataValue{
    ct:ILuckySheetCellFormat | undefined, //celltype,Cell value format: text, time, etc.
    bg: string | undefined,//background,#fff000	
    ff: string | undefined,//fontfamily,
    fc: string | undefined,//fontcolor
    bl: number | undefined,//Bold
    it: number | undefined,//italic
    fs: number | undefined,//font size	
    cl: number | undefined,//Cancelline, 0 Regular, 1 Cancelline
    un: number | undefined//underline, 0 Regular, 1 underlines, fonts
    vt: number | undefined,//Vertical alignment, 0 middle, 1 up, 2 down
    ht: number | undefined,//Horizontal alignment,0 center, 1 left, 2 right
    mc: IluckySheetCelldataValueMerge | undefined, //Merge Cells
    tr: number | undefined, //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90
    tb: number | undefined, //Text wrap,0 truncation, 1 overflow, 2 word wrap
    v: string | undefined, //Original value	
    m: string | undefined, //Display value	
    rt:number | undefined, //text rotation angle 0-180 alignment
    f: string | undefined, //formula
    qp:number | undefined //quotePrefix, show number as string
}


export interface ILuckySheetCellFormat {
    fa:string //Format definition string
    t:string // Cell Type
}

export interface IluckySheetCelldataValueMerge{
    rs?:number, //row of merge cell  length, only main merge cell, every merge cell has only one main mergeCell
    cs?:number, //column of merge cell  length, only main merge cell, every merge cell has only one main mergeCell
    r:number, // main merge cell row Number, other cell link to main cell
    c:number, // main merge cell column Number, other cell link to main cell
}

//Lucky sheet config attribute
export interface IluckySheetConfig{
    merge?:IluckySheetConfigMerges, //merge handdler
    // _borderInfo?: IMapluckySheetborderInfoCellForImp, //range border
    borderInfo:IluckySheetborderInfoCellForImp[],//range border
    rowlen?:IluckySheetRowAndColumnLen, // every row's height
    columnlen?:IluckySheetRowAndColumnLen,// every column's width
    rowhidden?:IluckySheetRowAndColumnHidden,//setting be hidden rows
    colhidden?:IluckySheetRowAndColumnHidden,//setting be hidden columns

    customHeight:IluckySheetRowAndColumnHidden,//user operate row height
    customWidth:IluckySheetRowAndColumnHidden//user operate column width
}

//Merge cells interface
export interface IluckySheetConfigMerges{
    [firstRange:string]:IluckySheetConfigMerge // "r_s":{ r,c,rs,cs } format, define a main merge cell 
}
//Merge cell interface
export interface IluckySheetConfigMerge{
    r: number,
    c: number,
    rs: number,
    cs: number
}

//Border cell interface
export interface IluckySheetborderInfoCell{
    rangeType:string,
    value:IluckySheetborderInfoCellValue,
}
export interface IluckySheetborderInfoCellValue{
    row_index: number,
    col_index: number,
    l: IluckySheetborderInfoCellValueStyle,
    r: IluckySheetborderInfoCellValueStyle,
    t: IluckySheetborderInfoCellValueStyle,
    b: IluckySheetborderInfoCellValueStyle
}
export interface IluckySheetborderInfoCellValueStyle{
    "style": number,
    "color": string
}

//border range interface
export interface IluckySheetborderInfoRange{
    rangeType:string,
    borderType:string,
    style:string,
    color:string,
    range:IluckySheetSelection[],
}

//Border cell interface for import
export interface IluckySheetborderInfoCellForImp{
    rangeType:string,
    cells?:string[],
    value:IluckySheetborderInfoCellValue,
}

//Border cell interface for import
export interface IMapluckySheetborderInfoCellForImp{
    [index:number]:IluckySheetborderInfoCellForImp 
}

//row length and column length interface
export interface IluckySheetRowAndColumnLen{
    [index:string]:number 
}

//row and column hidden interface
export interface IluckySheetRowAndColumnHidden{
    [index:string]:number
}

export interface IFormulaSI{
    [index:string]:IFormulaCell
}

export interface IFormulaCell{
    [index:string]:IFormulaCellValue
}

export interface IFormulaCellValue{
    t:string
    ref:string
    si:string
    fv:string
    cellValue:IluckySheetCelldata
}

export interface ILuckyInlineString {
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



//Image
export interface IluckyImage {
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

export interface IluckyImageBorder {
    color: string
    radius: number
    style: string
    width: number
}

export interface IluckyImageCrop {
    height: number
    offsetLeft: number
    offsetTop: number
    width: number
}

export interface IluckyImageDefault {
    height: number
    left: number
    top: number
    width: number
}

export interface IluckyImages {
    [index:string]:IluckyImage
}


export interface IcellOtherInfo{
    [index:string]:IformulaList
}

export interface IformulaList{
    [index:string]:IformulaListItem
}

export interface IformulaListItem{
    r:number,
    c:number
}


// DataVerification
export interface IluckysheetDataVerification {
  [key: string]: IluckysheetDataVerificationValue;
}

export interface IluckysheetDataVerificationValue {
  type: IluckysheetDataVerificationType;
  type2: string | null;
  value1: string | number | null;
  value2: string | number | null;
  checked: boolean;
  remote: boolean;
  prohibitInput: boolean;
  hintShow: boolean;
  hintText: string;
}

export type IluckysheetDataVerificationType =
  | "dropdown"
  | "checkbox"
  | "number"
  | "number_integer"
  | "number_decimal"
  | "text_content"
  | "text_length"
  | "date"
  | "validity";

export interface IluckysheetHyperlink {
    [key: string]: IluckysheetHyperlinkValue;
}

export interface IluckysheetHyperlinkValue {
    linkAddress: string;
    linkTooltip: string;
    linkType: IluckysheetHyperlinkType;
    display: string;
}

export type IluckysheetHyperlinkType = "internal" | "external";
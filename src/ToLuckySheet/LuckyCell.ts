import { IluckySheetborderInfoCellForImp,IluckySheetCelldataValue,IluckySheetCelldataValueMerge,ILuckySheetCellFormat } from "./ILuck";
import { ReadXml, Element, IStyleCollections,getColor } from "./ReadXml";
import {getXmlAttibute, getColumnWidthPixel, getRowHeightPixel,getcellrange, escapeCharacter} from "../common/method";
import { ST_CellType, indexedColors, OEM_CHARSET,borderTypes } from "../common/constant"
import { IattributeList, stringToNum } from "../ICommon";
import { LuckySheetborderInfoCellValueStyle,LuckySheetborderInfoCellForImp,LuckySheetborderInfoCellValue,LuckySheetCelldataBase,LuckySheetCelldataValue,LuckySheetCellFormat } from "./LuckyBase";

export class LuckySheetCelldata extends LuckySheetCelldataBase{
    _borderObject:IluckySheetborderInfoCellForImp
    _fomulaRef:string
    _formulaSi:string
    _formulaType:string

    private sheetFile:string
    private readXml:ReadXml
    private cell:Element
    private styles:IStyleCollections
    private sharedStrings:Element[]
    private mergeCells:Element[]

    constructor(cell:Element, styles:IStyleCollections, sharedStrings:Element[], mergeCells:Element[], sheetFile:string, ReadXml:ReadXml){
        //Private
        super();
        this.cell = cell;
        this.sheetFile = sheetFile;
        this.styles = styles;
        this.sharedStrings = sharedStrings;
        this.readXml = ReadXml;
        this.mergeCells = mergeCells;

        let attrList = cell.attributeList;
        let r = attrList.r, s = attrList.s, t = attrList.t;
        let range = getcellrange(r);

        this.r = range.row[0];
        this.c = range.column[0];
        this.v = this.generateValue(s, t);

    }

    /**
    * @param s Style index ,start 1
    * @param t Cell type, Optional value is ST_CellType, it's found at constat.ts
    */
    private generateValue(s:string, t:string){
        let v = this.cell.getInnerElements("v");
        let f = this.cell.getInnerElements("f");

        let cellXfs = this.styles["cellXfs"] as Element[];
        let cellStyleXfs = this.styles["cellStyleXfs"] as Element[];
        let cellStyles = this.styles["cellStyles"] as Element[];
        let fonts = this.styles["fonts"] as Element[];
        let fills = this.styles["fills"] as Element[];
        let borders = this.styles["borders"] as Element[];
        let numfmts = this.styles["numfmts"] as IattributeList;
        let clrScheme = this.styles["clrScheme"] as Element[];

        let sharedStrings = this.sharedStrings;
        let cellValue = new LuckySheetCelldataValue();

        if(f!=null){
            let formula = f[0], attrList = formula.attributeList;
            let t = attrList.t, ref = attrList.ref, si = attrList.si;
            let formulaValue =f[0].value;
            if(t=="shared"){
                this._fomulaRef = ref;
                this._formulaType = t;
                this._formulaSi = si;
            }
            // console.log(ref, t, si);
            if(ref!=null || (formulaValue!=null && formulaValue.length>0)){
                formulaValue = escapeCharacter(formulaValue);
                cellValue.f = "=" + formulaValue;
            }

        }


        let quotePrefix;
        if(s!=null){
            let sNum = parseInt(s);
            let cellXf = cellXfs[sNum];
            let xfId = cellXf.attributeList.xfId;

            let numFmtId,fontId,fillId,borderId;
            let horizontal,vertical, wrapText, textRotation, shrinkToFit, indent,applyProtection;

            if(xfId!=null){
                let cellStyleXf = cellStyleXfs[parseInt(xfId)];
                let attrList = cellStyleXf.attributeList;

                let applyNumberFormat = attrList.applyNumberFormat;
                let applyFont = attrList.applyFont;
                let applyFill = attrList.applyFill;
                let applyBorder = attrList.applyBorder;
                let applyAlignment = attrList.applyAlignment;
                // let applyProtection = attrList.applyProtection;

                applyProtection = attrList.applyProtection;
                quotePrefix = attrList.quotePrefix;

                if(applyNumberFormat!="0" && attrList.numFmtId!=null){
                    // if(attrList.numFmtId!="0"){
                        numFmtId = attrList.numFmtId;
                    // }
                }
                if(applyFont!="0" && attrList.fontId!=null){
                    fontId = attrList.fontId;
                }
                if(applyFill!="0" && attrList.fillId!=null){
                    fillId = attrList.fillId;
                }
                if(applyBorder!="0" && attrList.borderId!=null){
                    borderId = attrList.borderId;
                }
                if(applyAlignment!="0"){
                    let alignment = cellStyleXf.getInnerElements("alignment");
                    if(alignment!=null){
                        let attrList = alignment[0].attributeList;
                        if(attrList.horizontal!=null){
                            horizontal = attrList.horizontal;
                        }
                        if(attrList.vertical!=null){
                            vertical = attrList.vertical;
                        }
                        if(attrList.wrapText!=null){
                            wrapText = attrList.wrapText;
                        }
                        if(attrList.textRotation!=null){
                            textRotation = attrList.textRotation;
                        }
                        if(attrList.shrinkToFit!=null){
                            shrinkToFit = attrList.shrinkToFit;
                        }
                        if(attrList.indent!=null){
                            indent = attrList.indent;
                        }
                    }
                }
            }

            let applyNumberFormat = cellXf.attributeList.applyNumberFormat;
            let applyFont = cellXf.attributeList.applyFont;
            let applyFill = cellXf.attributeList.applyFill;
            let applyBorder = cellXf.attributeList.applyBorder;
            let applyAlignment = cellXf.attributeList.applyAlignment;
            
            if(cellXf.attributeList.applyProtection!=null){
                applyProtection = cellXf.attributeList.applyProtection;
            }
            
            if(cellXf.attributeList.quotePrefix!=null){
                quotePrefix = cellXf.attributeList.quotePrefix;
            }

            if(applyNumberFormat!="0" && cellXf.attributeList.numFmtId!=null){
                numFmtId = cellXf.attributeList.numFmtId;
            }
            if(applyFont!="0"){
                fontId = cellXf.attributeList.fontId;
            }
            if(applyFill!="0"){
                fillId = cellXf.attributeList.fillId;
            }
            if(applyBorder!="0"){
                borderId =cellXf.attributeList.borderId;
            }
            if(applyAlignment!="0"){
                let alignment = cellXf.getInnerElements("alignment");
                if(alignment!=null && alignment.length>0){
                    let attrList = alignment[0].attributeList;
                    if(attrList.horizontal!=null){
                        horizontal = attrList.horizontal;
                    }
                    if(attrList.vertical!=null){
                        vertical = attrList.vertical;
                    }
                    if(attrList.wrapText!=null){
                        wrapText = attrList.wrapText;
                    }
                    if(attrList.textRotation!=null){
                        textRotation = attrList.textRotation;
                    }
                    if(attrList.shrinkToFit!=null){
                        shrinkToFit = attrList.shrinkToFit;
                    }
                    if(attrList.indent!=null){
                        indent = attrList.indent;
                    }
                }
            }

            

            if(numFmtId!=undefined){
                let numf = numfmts[parseInt(numFmtId)];
                let cellFormat = new LuckySheetCellFormat();
                cellFormat.fa = escapeCharacter(numf);
                // console.log(numf, numFmtId, this.v);
                cellFormat.t = t;
                cellValue.ct = cellFormat;
            }

            if(fillId!=undefined){
                let fillIdNum = parseInt(fillId);
                let fill  = fills[fillIdNum];
                // console.log(cellValue.v);
                let bg = this.getBackgroundByFill(fill, clrScheme);
                if(bg!=null){
                    cellValue.bg = bg;
                }
            }

            if(fontId!=undefined){
                let fontIdNum = parseInt(fontId);
                let font = fonts[fontIdNum];
                if(font!=null){
                    let sz = font.getInnerElements("sz");//font size
                    let colors = font.getInnerElements("color");//font color
                    let family = font.getInnerElements("name");//font family
                    let familyOverrides = font.getInnerElements("family");//font family will be overrided by name
                    let charset = font.getInnerElements("charset");//font charset
                    let bolds = font.getInnerElements("b");//font bold
                    let italics = font.getInnerElements("i");//font italic
                    let strikes = font.getInnerElements("strike");//font italic
                    let underlines = font.getInnerElements("u");//font italic

                    if(sz!=null && sz.length>0){
                        let fs = sz[0].attributeList.val;
                        if(fs!=null){
                            cellValue.fs = parseInt(fs);
                        }
                       
                    }

                    if(colors!=null && colors.length>0){
                        let color = colors[0];
                        let fc = getColor(color, clrScheme, "t");
                        if(fc!=null){
                            cellValue.fc = fc;
                        }
                    }

                    let ff;
                    if(familyOverrides!=null && familyOverrides.length>0){
                        let val = familyOverrides[0].attributeList.val;
                        if(val!=null){
                            ff = val;
                        }
                    }
                    if(family!=null && family.length>0){
                        let val = family[0].attributeList.val;
                        if(val!=null){
                            ff = val;
                        }
                    }
                    if(ff!=null){
                        cellValue.ff = ff;
                    }

                    if(bolds!=null && bolds.length>0){
                        let bold = bolds[0].attributeList.val;
                        if(bold=="0"){
                            cellValue.bl =  0;
                        }
                        else{
                            cellValue.bl =  1;
                        }
                    }

                    if(italics!=null && italics.length>0){
                        let italic = italics[0].attributeList.val;
                        if(italic=="0"){
                            cellValue.it =  0;
                        }
                        else{
                            cellValue.it =  1;
                        }
                    }

                    if(strikes!=null && strikes.length>0){
                        let strike = strikes[0].attributeList.val;
                        if(strike=="0"){
                            cellValue.cl =  0;
                        }
                        else{
                            cellValue.cl =  1;
                        }
                    }

                    if(underlines!=null && underlines.length>0){
                        let underline = underlines[0].attributeList.val;
                        if(underline=="0"){
                            cellValue.un =  0;
                        }
                        else{
                            cellValue.un =  1;
                        }
                    }
                }
            }

            // vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
            // ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
            // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
            // tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment

            if(horizontal!=undefined){//Horizontal alignment
                if(horizontal=="center"){
                    cellValue.ht = 0;
                }
                else if(horizontal=="centerContinuous"){
                    cellValue.ht = 0;//luckysheet unsupport
                }
                else if(horizontal=="left"){
                    cellValue.ht = 1;
                }
                else if(horizontal=="right"){
                    cellValue.ht = 2;
                }
                else if(horizontal=="distributed"){
                    cellValue.ht = 0;//luckysheet unsupport
                }
                else if(horizontal=="fill"){
                    cellValue.ht = 1;//luckysheet unsupport
                }
                else if(horizontal=="general"){
                    cellValue.ht = 1;//luckysheet unsupport
                }
                else if(horizontal=="justify"){
                    cellValue.ht = 0;//luckysheet unsupport
                }
                else{
                    cellValue.ht = 1;
                }
            }

            if(vertical!=undefined){//Vertical alignment
                if(vertical=="bottom"){
                    cellValue.vt = 2;
                }
                else if(vertical=="center"){
                    cellValue.vt = 0;
                }
                else if(vertical=="distributed"){
                    cellValue.vt = 0;//luckysheet unsupport
                }
                else if(vertical=="justify"){
                    cellValue.vt = 0;//luckysheet unsupport
                }
                else if(vertical=="top"){
                    cellValue.vt = 1;
                }
                else{
                    cellValue.vt = 1;
                }
            }

            if(wrapText!=undefined){
                if(wrapText=="1"){
                    cellValue.tb = 2;
                }
                else{
                    cellValue.tb = 1;
                }
            }
            else{
                cellValue.tb = 1;
            }

            if(textRotation!=undefined){
                // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
                if(textRotation=="0"){
                    cellValue.tr = 0;
                }
                else if(textRotation=="45"){
                    cellValue.tr = 1;
                }
                else if(textRotation=="90"){
                    cellValue.tr = 4;
                }
                else if(textRotation=="135"){
                    cellValue.tr = 2;
                }
                else if(textRotation=="180"){
                    cellValue.tr = 5;
                }
                else{
                    cellValue.tr = 0;
                }
            }

            if(shrinkToFit!=undefined){//luckysheet unsupport
                
            }

            if(indent!=undefined){//luckysheet unsupport
                
            }

            if(borderId!=undefined){
                let borderIdNum = parseInt(borderId);
                let border = borders[borderIdNum];
                // this._borderId = borderIdNum;

                let borderObject = new LuckySheetborderInfoCellForImp();
                borderObject.rangeType = "cell";
                // borderObject.cells = [];
                let borderCellValue = new LuckySheetborderInfoCellValue();

                borderCellValue.row_index = this.r;
                borderCellValue.col_index = this.c;
                
                let lefts = border.getInnerElements("left");
                let rights = border.getInnerElements("right");
                let tops = border.getInnerElements("top");
                let bottoms = border.getInnerElements("bottom");
                let diagonals = border.getInnerElements("diagonal");

                let starts = border.getInnerElements("start");
                let ends = border.getInnerElements("end");

                let left = this.getBorderInfo(lefts);
                let right = this.getBorderInfo(rights);
                let top = this.getBorderInfo(tops);
                let bottom = this.getBorderInfo(bottoms);
                let diagonal = this.getBorderInfo(diagonals);

                let start = this.getBorderInfo(starts);
                let end = this.getBorderInfo(ends);

                let isAdd = false;

                if(start!=null && start.color!=null){
                    borderCellValue.l = start;
                    isAdd = true;
                }

                if(end!=null && end.color!=null){
                    borderCellValue.r = end;
                    isAdd = true;
                }

                if(left!=null && left.color!=null){
                    borderCellValue.l = left;
                    isAdd = true;
                }

                if(right!=null && right.color!=null){
                    borderCellValue.r = right;
                    isAdd = true;
                }

                if(top!=null && top.color!=null){
                    borderCellValue.t = top;
                    isAdd = true;
                }

                if(bottom!=null && bottom.color!=null){
                    borderCellValue.b = bottom;
                    isAdd = true;
                }

                if(isAdd){
                    borderObject.value = borderCellValue;
                    // this.config._borderInfo[borderId] = borderObject;
                    this._borderObject = borderObject;
                }
            }
            
        }
        else{
            cellValue.tb = 1;
        }

        if(v!=null){
            let value =v[0].value;
            if(t==ST_CellType["SharedString"]){
                let siIndex = parseInt(v[0].value);
                let sharedSI = sharedStrings[siIndex];
                // console.log(siIndex, sharedSI, sharedStrings);
                let tFlag = sharedSI.getInnerElements("t");
                if(tFlag!=null){
                    let text = "";
                    tFlag.forEach((t)=>{
                        text += t.value;
                    });
                    cellValue.v = text;
                    quotePrefix = "1";
                }
            }
            else if(t==ST_CellType["InlineString"] && v!=null){
    
            }
            else {
                cellValue.v = value;
            }
        }

        if(quotePrefix!=null){
            cellValue.qp = parseInt(quotePrefix);
        }

        return cellValue;
    
    }


    private getBackgroundByFill(fill:Element, clrScheme:Element[]):string|null{
        let patternFills = fill.getInnerElements("patternFill");
        if(patternFills!=null){
            let patternFill = patternFills[0];
            let fgColors = patternFill.getInnerElements("fgColor");
            let bgColors = patternFill.getInnerElements("bgColor");
            let fg, bg;
            if(fgColors!=null){
                let fgColor = fgColors[0];
                fg = getColor(fgColor, clrScheme);
            }

            if(bgColors!=null){
                let bgColor = bgColors[0];
                bg = getColor(bgColor, clrScheme);
            }
            // console.log(fgColors,bgColors,clrScheme);
            if(fg!=null){
                return fg;
            }
            else if(bg!=null){
                return bg;
            }
        }
        else{
            let gradientfills = fill.getInnerElements("gradientFill");
            if(gradientfills!=null){
                //graient color fill handler

                return null;
            }
        }
    }

    private getBorderInfo(borders:Element[]):LuckySheetborderInfoCellValueStyle{
        if(borders==null){
            return null;
        }

        let border = borders[0], attrList = border.attributeList;
        let clrScheme = this.styles["clrScheme"] as Element[];
        let style:string = attrList.style;
        if(style==null || style=="none"){
            return null;
        }

        let colors = border.getInnerElements("color");
        let colorRet = "#000000";
        if(colors!=null){
            let color = colors[0];
            colorRet = getColor(color, clrScheme, "b");
            if(colorRet==null){
                colorRet = "#000000";
            }
        }

        let ret = new LuckySheetborderInfoCellValueStyle();
        ret.style = borderTypes[style];
        ret.color = colorRet;

        return ret;
    }

}


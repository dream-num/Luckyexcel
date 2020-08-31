import { IluckySheetChart,IluckySheetPivotTable,IluckysheetConditionFormat,IluckysheetCalcChain,IluckySheetCelldata,IluckySheetCelldataValue,IMapluckySheetborderInfoCellForImp,IluckySheetborderInfoCellValue,IluckySheetborderInfoCellValueStyle,IFormulaSI,IluckySheetRowAndColumnLen,IluckySheetRowAndColumnHidden,IluckySheetSelection,IluckysheetFrozen} from "./ILuck";
import {LuckySheetCelldata} from "./LuckyCell";
import { IattributeList } from "../ICommon";
import {getXmlAttibute, getColumnWidthPixel, fromulaRef,getRowHeightPixel,getcellrange,generateRandomSheetIndex} from "../common/method";
import {borderTypes} from "../common/constant";
import { ReadXml, IStyleCollections, Element,getColor } from "./ReadXml";
import { LuckyFileBase,LuckySheetBase,LuckyConfig,LuckySheetborderInfoCellForImp,LuckySheetborderInfoCellValue,LuckysheetCalcChain,LuckySheetConfigMerge } from "./LuckyBase";


export class LuckySheet extends LuckySheetBase {

    private readXml:ReadXml
    private sheetFile:string
    private isInitialCell:boolean
    private styles:IStyleCollections
    private sharedStrings:Element[]
    private mergeCells:Element[]
    private calcChainEles:Element[]

    private formulaRefList:IFormulaSI

    constructor(sheetName:string, sheetId:string, sheetOrder:number,sheetFile:string, ReadXml:ReadXml, sheets:IattributeList, styles:IStyleCollections, sharedStrings:Element[], calcChain:Element[],isInitialCell:boolean=false){
        //Private
        super();
        this.readXml = ReadXml;
        this.sheetFile = sheetFile;
        this.isInitialCell = isInitialCell;
        this.styles = styles;
        this.sharedStrings = sharedStrings;
        this.calcChainEles = calcChain;

        //Output
        this.name = sheetName;
        this.index = sheetId;
        this.order = sheetOrder.toString();
        this.config = new LuckyConfig();
        this.celldata = [];
        this.mergeCells = this.readXml.getElementsByTagName("mergeCells/mergeCell", sheetFile);
        let sheetView = this.readXml.getElementsByTagName("sheetViews/sheetView", sheetFile);
        let showGridLines = "1", tabSelected="0", zoomScale = "100", activeCell = "A1";
        if(sheetView.length>0){
            let attrList = sheetView[0].attributeList;
            showGridLines = getXmlAttibute(attrList, "showGridLines", "1");
            tabSelected = getXmlAttibute(attrList, "tabSelected", "0");
            zoomScale = getXmlAttibute(attrList, "zoomScale", "100");
            // let colorId = getXmlAttibute(attrList, "colorId", "0");
            let selections = sheetView[0].getInnerElements("selection");
            if(selections!=null && selections.length>0){
                activeCell = getXmlAttibute(selections[0].attributeList, "activeCell", "A1");
                let range:IluckySheetSelection = getcellrange(activeCell, sheets, sheetId);
                this.luckysheet_select_save = [];
                this.luckysheet_select_save.push(range);
            }
        }
        this.showGridLines = showGridLines;
        this.status = tabSelected;
        this.zoomRatio = parseInt(zoomScale)/100;

        let sheetFormatPr = this.readXml.getElementsByTagName("sheetFormatPr", sheetFile);
        let defaultColWidth = "8.38", defaultRowHeight="defaultRowHeight";
        if(sheetFormatPr.length>0){
            let attrList = sheetFormatPr[0].attributeList;
            defaultColWidth = getXmlAttibute(attrList, "defaultColWidth", "8.38");
            defaultRowHeight = getXmlAttibute(attrList, "defaultRowHeight", "19");
        }

        this.defaultColWidth = getColumnWidthPixel(parseFloat(defaultColWidth));
        this.defaultRowHeight = getRowHeightPixel(parseFloat(defaultRowHeight));



        this.generateConfigColumnLenAndHidden();
        this.generateConfigRowLenAndHiddenAddCell();

        if(this.formulaRefList!=null){
            for(let key in this.formulaRefList){
                let funclist = this.formulaRefList[key];
                let mainFunc = funclist["mainRef"], mainCellValue = mainFunc.cellValue;
                let formulaTxt = mainFunc.fv;
                let mainR = mainCellValue.r, mainC = mainCellValue.c;
                // let refRange = getcellrange(ref);
                for(let name in funclist){
                    if(name == "mainRef"){
                        continue;
                    }

                    let funcValue = funclist[name], cellValue = funcValue.cellValue;
                    if(cellValue==null){
                        continue;
                    }
                    let r = cellValue.r, c = cellValue.c;

                    let func = formulaTxt;
                    let offsetRow = r - mainR, offsetCol = c - mainC;

                    
                    if(offsetRow > 0){
                        func = "=" + fromulaRef.functionCopy(func, "down", offsetRow);
                    }
                    else if(offsetRow < 0){
                        func = "=" + fromulaRef.functionCopy(func, "up", Math.abs(offsetRow));
                    }

                    if(offsetCol > 0){
                        func = "=" + fromulaRef.functionCopy(func, "right", offsetCol);
                    }
                    else if(offsetCol < 0){
                        func = "=" + fromulaRef.functionCopy(func, "left", Math.abs(offsetCol));
                    }

                    // console.log(offsetRow, offsetCol, func);

                    (cellValue.v as IluckySheetCelldataValue ).f = func;
                    
                }
            }
        }


        if(this.calcChain==null){
            this.calcChain = [];
        }
        for(let c=0;c<this.calcChainEles.length;c++){
            let calcChainEle = this.calcChainEles[c], attrList = calcChainEle.attributeList;
            if(attrList.i!=sheetId){
                continue;
            }

            let r = attrList.r , i = attrList.i, l = attrList.l, s = attrList.s, a = attrList.a, t = attrList.t;

            let range = getcellrange(r);
            let chain = new LuckysheetCalcChain();
            chain.r = range.row[0];
            chain.c = range.column[0];
            chain.index = this.index;
            this.calcChain.push(chain);
        }

        if(this.mergeCells!=null){
            for(let i=0;i<this.mergeCells.length;i++){
                let merge = this.mergeCells[i], attrList = merge.attributeList;
                let ref = attrList.ref;
                if(ref==null){
                    continue;
                }
                let range = getcellrange(ref, sheets, sheetId);
                let mergeValue = new LuckySheetConfigMerge();
                mergeValue.r = range.row[0];
                mergeValue.c = range.column[0];
                mergeValue.rs = range.row[1]-range.row[0]+1;
                mergeValue.cs = range.column[1]-range.column[0]+1;
                if(this.config.merge==null){
                    this.config.merge = {};
                }
                this.config.merge[range.row[0] + "_" + range.column[0]] = mergeValue;
            }
        }
    }

    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    private generateConfigColumnLenAndHidden(){
        let cols = this.readXml.getElementsByTagName("cols/col", this.sheetFile);
        for(let i=0;i<cols.length;i++){
            let col = cols[i], attrList = col.attributeList;
            let min = getXmlAttibute(attrList, "min", null);
            let max = getXmlAttibute(attrList, "max", null);
            let width = getXmlAttibute(attrList, "width", null);
            let hidden = getXmlAttibute(attrList, "hidden", null);
            let customWidth = getXmlAttibute(attrList, "customWidth", null);


            if(min==null || max==null || customWidth==null){
                continue;
            }

            let minNum = parseInt(min)-1, maxNum=parseInt(max)-1, widthNum=parseFloat(width);
            
            for(let m=minNum;m<=maxNum;m++){
                if(width!=null){
                    if(this.config.columnlen==null){
                        this.config.columnlen = {};
                    }
                    this.config.columnlen[m] = getColumnWidthPixel(widthNum);
                }

                if(hidden=="1"){
                    if(this.config.colhidden==null){
                        this.config.colhidden = {};
                    }
                    this.config.colhidden[m] = 0;

                    delete this.config.columnlen[m];
                }
            } 
        }
    }

    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
   private generateConfigRowLenAndHiddenAddCell(){
        let rows = this.readXml.getElementsByTagName("sheetData/row", this.sheetFile);
        for(let i=0;i<rows.length;i++){
            let row = rows[i], attrList = row.attributeList;
            let rowNo = getXmlAttibute(attrList, "r", null);
            let height = getXmlAttibute(attrList, "ht", null);
            let hidden = getXmlAttibute(attrList, "hidden", null);
            let customHeight = getXmlAttibute(attrList, "customHeight", null);

            if(rowNo==null){
                continue;
            }

            let rowNoNum = parseInt(rowNo) - 1;
            if(height!=null){
                let heightNum = parseFloat(height);
                if(this.config.rowlen==null){
                    this.config.rowlen = {};
                }
                this.config.rowlen[rowNoNum] = getRowHeightPixel(heightNum);
            }

            if(hidden=="1"){
                if(this.config.rowhidden==null){
                    this.config.rowhidden = {};
                }
                this.config.rowhidden[rowNoNum] = 0;

                delete this.config.rowlen[rowNoNum];
            }

            if(customHeight!=null){
                if(this.config.customHeight==null){
                    this.config.customHeight = {};
                }
                this.config.customHeight[rowNoNum] = 1;
            }


            if(this.isInitialCell){
                let cells = row.getInnerElements("c");
                for(let key in cells){
                    let cell = cells[key];
                    let cellValue = new LuckySheetCelldata(cell, this.styles, this.sharedStrings, this.mergeCells,this.sheetFile, this.readXml);
                    if(cellValue._borderObject!=null){
                        if(this.config.borderInfo==null){
                            this.config.borderInfo = [];
                        }
                        this.config.borderInfo.push(cellValue._borderObject);
                        delete cellValue._borderObject;
                    }
                    
                    // let borderId = cellValue._borderId;
                    // if(borderId!=null){
                    //     let borders = this.styles["borders"] as Element[];
                    //     if(this.config._borderInfo==null){
                    //         this.config._borderInfo = {};
                    //     }
                    //     if( borderId in this.config._borderInfo){
                    //         this.config._borderInfo[borderId].cells.push(cellValue.r + "_" + cellValue.c);
                    //     }
                    //     else{
                    //         let border = borders[borderId];
                    //         let borderObject = new LuckySheetborderInfoCellForImp();
                    //         borderObject.rangeType = "cellGroup";
                    //         borderObject.cells = [];
                    //         let borderCellValue = new LuckySheetborderInfoCellValue();
                            
                    //         let lefts = border.getInnerElements("left");
                    //         let rights = border.getInnerElements("right");
                    //         let tops = border.getInnerElements("top");
                    //         let bottoms = border.getInnerElements("bottom");
                    //         let diagonals = border.getInnerElements("diagonal");

                    //         let left = this.getBorderInfo(lefts);
                    //         let right = this.getBorderInfo(rights);
                    //         let top = this.getBorderInfo(tops);
                    //         let bottom = this.getBorderInfo(bottoms);
                    //         let diagonal = this.getBorderInfo(diagonals);

                    //         let isAdd = false;
                    //         if(left!=null && left.color!=null){
                    //             borderCellValue.l = left;
                    //             isAdd = true;
                    //         }

                    //         if(right!=null && right.color!=null){
                    //             borderCellValue.r = right;
                    //             isAdd = true;
                    //         }

                    //         if(top!=null && top.color!=null){
                    //             borderCellValue.t = top;
                    //             isAdd = true;
                    //         }

                    //         if(bottom!=null && bottom.color!=null){
                    //             borderCellValue.b = bottom;
                    //             isAdd = true;
                    //         }

                    //         if(isAdd){
                    //             borderObject.value = borderCellValue;
                    //             this.config._borderInfo[borderId] = borderObject;
                    //         }

                    //     }
                    // }
                    if(cellValue._formulaType=="shared"){
                        if(this.formulaRefList==null){
                            this.formulaRefList = {};
                        }

                        if(this.formulaRefList[cellValue._formulaSi]==null){
                            this.formulaRefList[cellValue._formulaSi] = {}
                        }

                        let fv;
                        if(cellValue.v!=null){
                            fv = (cellValue.v as IluckySheetCelldataValue).f;
                        }

                        let refValue = {
                            t:cellValue._formulaType,
                            ref:cellValue._fomulaRef,
                            si:cellValue._formulaSi,
                            fv:fv,
                            cellValue:cellValue
                        }

                        if(cellValue._fomulaRef!=null){
                            this.formulaRefList[cellValue._formulaSi]["mainRef"] = refValue;
                        }
                        else{
                            this.formulaRefList[cellValue._formulaSi][cellValue.r+"_"+cellValue.c] = refValue;
                        }

                        // console.log(refValue, this.formulaRefList);
                    }

                    this.celldata.push(cellValue);
                }
                
            }
        }
    }

    // private getBorderInfo(borders:Element[]):LuckySheetborderInfoCellValueStyle{
    //     if(borders==null){
    //         return null;
    //     }

    //     let border = borders[0], attrList = border.attributeList;
    //     let clrScheme = this.styles["clrScheme"] as Element[];
    //     let style:string = attrList.style;
    //     if(style==null || style=="none"){
    //         return null;
    //     }

    //     let colors = border.getInnerElements("color");
    //     let colorRet = "#000000";
    //     if(colors!=null){
    //         let color = colors[0];
    //         colorRet = getColor(color, clrScheme);
    //     }

    //     let ret = new LuckySheetborderInfoCellValueStyle();
    //     ret.style = borderTypes[style];
    //     ret.color = colorRet;

    //     return ret;
    // }
}

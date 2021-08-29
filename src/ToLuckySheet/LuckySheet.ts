import { IluckyImageBorder,IluckyImageCrop,IluckyImageDefault,IluckyImages,IluckySheetCelldata,IluckySheetCelldataValue,IMapluckySheetborderInfoCellForImp,IluckySheetborderInfoCellValue,IluckySheetborderInfoCellValueStyle,IFormulaSI,IluckySheetRowAndColumnLen,IluckySheetRowAndColumnHidden,IluckySheetSelection,IcellOtherInfo,IformulaList,IformulaListItem, IluckysheetHyperlink, IluckysheetHyperlinkType, LuckysheetFrozen2TypeEnum} from "./ILuck";
import {LuckySheetCelldata} from "./LuckyCell";
import { IattributeList } from "../ICommon";
import {getXmlAttibute, getColumnWidthPixel, fromulaRef,getRowHeightPixel,getcellrange,generateRandomIndex,getPxByEMUs, getMultiSequenceToNum, getTransR1C1ToSequence} from "../common/method";
import {borderTypes, worksheetFilePath} from "../common/constant";
import { ReadXml, IStyleCollections, Element,getColor } from "./ReadXml";
import { LuckyFileBase,LuckySheetBase,LuckyConfig,LuckySheetborderInfoCellForImp,LuckySheetborderInfoCellValue,LuckysheetCalcChain,LuckySheetConfigMerge } from "./LuckyBase";
import {ImageList} from "./LuckyImage";

export class LuckySheet extends LuckySheetBase {

    private readXml:ReadXml
    private sheetFile:string
    private isInitialCell:boolean
    private styles:IStyleCollections
    private sharedStrings:Element[]
    private mergeCells:Element[]
    private calcChainEles:Element[]
    private sheetList:IattributeList

    private imageList:ImageList

    private formulaRefList:IFormulaSI

    constructor(sheetName:string, sheetId:string, sheetOrder:number,isInitialCell:boolean=false, allFileOption:any){
        //Private
        super();
        this.isInitialCell = isInitialCell;

        this.readXml = allFileOption.readXml;
        this.sheetFile = allFileOption.sheetFile;
        this.styles = allFileOption.styles;
        this.sharedStrings = allFileOption.sharedStrings;
        this.calcChainEles = allFileOption.calcChain;
        this.sheetList = allFileOption.sheetList;
        this.imageList = allFileOption.imageList;
        this.hide = allFileOption.hide;

        //Output
        this.name = sheetName;
        this.index = sheetId;
        this.order = sheetOrder.toString();
        this.config = new LuckyConfig();
        this.celldata = [];
        this.mergeCells = this.readXml.getElementsByTagName("mergeCells/mergeCell", this.sheetFile);
        let clrScheme = this.styles["clrScheme"] as Element[];
        let sheetView = this.readXml.getElementsByTagName("sheetViews/sheetView", this.sheetFile);
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
                let range:IluckySheetSelection = getcellrange(activeCell, this.sheetList, sheetId);
                this.luckysheet_select_save = [];
                this.luckysheet_select_save.push(range);
            }
            // pane
            let panes = sheetView[0].getInnerElements("pane");
            if(panes!=null && panes.length>0){
                let pane=panes[0];
                let paneState=getXmlAttibute(pane.attributeList, "state", "split");
                if(paneState==='frozen'){
                    let xSplit=+getXmlAttibute(pane.attributeList, "xSplit", "0");
                    let ySplit=+getXmlAttibute(pane.attributeList, "ySplit", "0");
                    this.frozen={
                        type:LuckysheetFrozen2TypeEnum.cancel
                    };
                    if(xSplit==0&&ySplit==0)
                    this.frozen.type= LuckysheetFrozen2TypeEnum.cancel;
                    else if(xSplit==1&&ySplit==0)
                    this.frozen.type= LuckysheetFrozen2TypeEnum.column;
                    else if(xSplit==0&&ySplit==1)
                    this.frozen.type= LuckysheetFrozen2TypeEnum.row;
                    else if(xSplit==1&&ySplit==1){
                        this.frozen.type= LuckysheetFrozen2TypeEnum.both;
                    }else if(xSplit>1&&ySplit==0){
                        this.frozen.type= LuckysheetFrozen2TypeEnum.rangeColumn;
                        this.frozen.range={
                            row_focus:0,
                            column_focus:xSplit-1
                        };
                    }else if(xSplit==0&&ySplit>1){
                        this.frozen.type= LuckysheetFrozen2TypeEnum.rangeRow;
                        this.frozen.range={
                            row_focus:ySplit-1,
                            column_focus:0
                        };
                    }else if(xSplit>1&&ySplit>1){
                        this.frozen.type= LuckysheetFrozen2TypeEnum.rangeBoth;
                        this.frozen.range={
                            row_focus:ySplit-1,
                            column_focus:xSplit-1
                        };
                    }  
                }
            }
        }
        this.showGridLines = showGridLines;
        this.status = tabSelected;
        this.zoomRatio = parseInt(zoomScale)/100;

        let tabColors = this.readXml.getElementsByTagName("sheetPr/tabColor", this.sheetFile);
        if(tabColors!=null && tabColors.length>0){
            let tabColor = tabColors[0], attrList = tabColor.attributeList;
            // if(attrList.rgb!=null){
                let tc = getColor(tabColor, this.styles, "b");
                this.color = tc;
            // }
        }

        let sheetFormatPr = this.readXml.getElementsByTagName("sheetFormatPr", this.sheetFile);
        let defaultColWidth, defaultRowHeight;
        if(sheetFormatPr.length>0){
            let attrList = sheetFormatPr[0].attributeList;
            defaultColWidth = getXmlAttibute(attrList, "defaultColWidth", "9.21");
            defaultRowHeight = getXmlAttibute(attrList, "defaultRowHeight", "19");
        }

        this.defaultColWidth = getColumnWidthPixel(parseFloat(defaultColWidth));
        this.defaultRowHeight = getRowHeightPixel(parseFloat(defaultRowHeight));


        this.generateConfigColumnLenAndHidden();
        let cellOtherInfo:IcellOtherInfo =  this.generateConfigRowLenAndHiddenAddCell();

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

        let formulaListExist:IformulaList={};
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
            formulaListExist["r"+r+"c"+c] = null;
        }

        //There may be formulas that do not appear in calcChain
        for(let key in cellOtherInfo.formulaList){
            if(!(key in formulaListExist)){
                let formulaListItem = cellOtherInfo.formulaList[key];
                let chain = new LuckysheetCalcChain();
                chain.r = formulaListItem.r;
                chain.c = formulaListItem.c;
                chain.index = this.index;
                this.calcChain.push(chain);
            }
        }

        // hyperlink config
        this.hyperlink = this.generateConfigHyperlinks();
      
        // sheet hide
        this.hide = this.hide;

        if(this.mergeCells!=null){
            for(let i=0;i<this.mergeCells.length;i++){
                let merge = this.mergeCells[i], attrList = merge.attributeList;
                let ref = attrList.ref;
                if(ref==null){
                    continue;
                }
                let range = getcellrange(ref, this.sheetList, sheetId);
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

        let drawingFile = allFileOption.drawingFile, drawingRelsFile = allFileOption.drawingRelsFile;
        if(drawingFile!=null && drawingRelsFile!=null){
            let twoCellAnchors = this.readXml.getElementsByTagName("xdr:twoCellAnchor", drawingFile);

            if(twoCellAnchors!=null && twoCellAnchors.length>0){
                for(let i=0;i<twoCellAnchors.length;i++){
                    let twoCellAnchor = twoCellAnchors[i];
                    let editAs = getXmlAttibute(twoCellAnchor.attributeList, "editAs", "twoCell");

                    let xdrFroms = twoCellAnchor.getInnerElements("xdr:from"), xdrTos = twoCellAnchor.getInnerElements("xdr:to");

                    let xdr_blipfills = twoCellAnchor.getInnerElements("a:blip");
                    if(xdrFroms!=null && xdr_blipfills!=null && xdrFroms.length>0 && xdr_blipfills.length>0){
                        let xdrFrom = xdrFroms[0], xdrTo = xdrTos[0],xdr_blipfill = xdr_blipfills[0];
                        
                        let rembed = getXmlAttibute(xdr_blipfill.attributeList, "r:embed", null);

                        let imageObject = this.getBase64ByRid(rembed, drawingRelsFile);



                        // let aoff = xdr_xfrm.getInnerElements("a:off"), aext = xdr_xfrm.getInnerElements("a:ext");

                        

                        // if(aoff!=null && aext!=null && aoff.length>0 && aext.length>0){
                        //     let aoffAttribute = aoff[0].attributeList, aextAttribute = aext[0].attributeList;
                        //     let x = getXmlAttibute(aoffAttribute, "x", null);
                        //     let y = getXmlAttibute(aoffAttribute, "y", null);

                        //     let cx = getXmlAttibute(aextAttribute, "cx", null);
                        //     let cy = getXmlAttibute(aextAttribute, "cy", null);

                        //     if(x!=null && y!=null && cx!=null && cy!=null && imageObject !=null){
                        // let x_n = getPxByEMUs(parseInt(x), "c"),y_n = getPxByEMUs(parseInt(y));
                        // let cx_n = getPxByEMUs(parseInt(cx), "c"),cy_n = getPxByEMUs(parseInt(cy));

                        let x_n =0,y_n = 0;
                        let cx_n = 0, cy_n = 0;

                        imageObject.fromCol = this.getXdrValue(xdrFrom.getInnerElements("xdr:col"));
                        imageObject.fromColOff = getPxByEMUs(this.getXdrValue(xdrFrom.getInnerElements("xdr:colOff")));
                        imageObject.fromRow= this.getXdrValue(xdrFrom.getInnerElements("xdr:row"));
                        imageObject.fromRowOff = getPxByEMUs(this.getXdrValue(xdrFrom.getInnerElements("xdr:rowOff")));

                        imageObject.toCol = this.getXdrValue(xdrTo.getInnerElements("xdr:col"));
                        imageObject.toColOff = getPxByEMUs(this.getXdrValue(xdrTo.getInnerElements("xdr:colOff")));
                        imageObject.toRow = this.getXdrValue(xdrTo.getInnerElements("xdr:row"));
                        imageObject.toRowOff = getPxByEMUs(this.getXdrValue(xdrTo.getInnerElements("xdr:rowOff")));

                        imageObject.originWidth = cx_n;
                        imageObject.originHeight = cy_n;
                        
                        if(editAs=="absolute"){
                            imageObject.type = "3";
                        }
                        else if(editAs=="oneCell"){
                            imageObject.type = "2";
                        }
                        else{
                            imageObject.type = "1";
                        }

                        imageObject.isFixedPos = false;
                        imageObject.fixedLeft = 0;
                        imageObject.fixedTop = 0;

                        let imageBorder:IluckyImageBorder = {
                            color: "#000",
                            radius: 0,
                            style: "solid",
                            width: 0
                        }
                        imageObject.border = imageBorder;

                        let imageCrop:IluckyImageCrop = {
                            height: cy_n,
                            offsetLeft: 0,
                            offsetTop: 0,
                            width: cx_n
                        }
                        imageObject.crop = imageCrop;

                        let imageDefault:IluckyImageDefault = {
                            height: cy_n,
                            left: x_n,
                            top: y_n,
                            width: cx_n
                        }
                        imageObject.default = imageDefault;

                        if(this.images==null){
                            this.images = {};
                        }
                        this.images[generateRandomIndex("image")] = imageObject;
                        //     }
                        // }
                    }
                }
            }
            
        } 
    }

    private getXdrValue(ele:Element[]):number{
        if(ele==null || ele.length==0){
            return null;
        }

        return parseInt(ele[0].value);
    }

    private getBase64ByRid(rid:string, drawingRelsFile:string){
        let Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);

        if(Relationships!=null && Relationships.length>0){
            for(let i=0;i<Relationships.length;i++){
                let Relationship = Relationships[i];
                let attrList = Relationship.attributeList;
                let Id = getXmlAttibute(attrList, "Id", null);
                let src = getXmlAttibute(attrList, "Target", null);
                if(Id == rid){
                    src = src.replace(/\.\.\//g, "");
                    src = "xl/" + src;
                    let imgage = this.imageList.getImageByName(src);
                    return imgage;
                }
            }
        }

        return null;
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


            if(min==null || max==null){
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

                    if(this.config.columnlen){
                        delete this.config.columnlen[m];
                    }
                    
                }

                if(customWidth!=null){
                    if(this.config.customWidth==null){
                        this.config.customWidth = {};
                    }
                    this.config.customWidth[m] = 1;
                }
            } 
        }
    }

    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    private generateConfigRowLenAndHiddenAddCell():IcellOtherInfo{
        let rows = this.readXml.getElementsByTagName("sheetData/row", this.sheetFile);
        let cellOtherInfo:IcellOtherInfo = {};
        let formulaList:IformulaList = {};
        cellOtherInfo.formulaList = formulaList;
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
                
                if(this.config.rowlen){
                    delete this.config.rowlen[rowNoNum];
                }
                
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

                    //There may be formulas that do not appear in calcChain
                    if(cellValue.v!=null && (cellValue.v as IluckySheetCelldataValue).f!=null){
                        let formulaCell:IformulaListItem = {
                            r:cellValue.r,
                            c:cellValue.c
                        }
                        cellOtherInfo.formulaList["r"+cellValue.r+"c"+cellValue.c] = formulaCell;
                    }

                    this.celldata.push(cellValue);
                }
                
            }
        }

        return cellOtherInfo;
    }
  
    /**
     * luckysheet config of hyperlink
     * 
     * @returns {IluckysheetHyperlink} - hyperlink config
     */
    private generateConfigHyperlinks(): IluckysheetHyperlink {
      let rows = this.readXml.getElementsByTagName(
        "hyperlinks/hyperlink",
        this.sheetFile
      );
      let hyperlink: IluckysheetHyperlink = {};
      for (let i = 0; i < rows.length; i++) {
        let row = rows[i];
        let attrList = row.attributeList;
        let ref = getXmlAttibute(attrList, "ref", null),
            refArr = getMultiSequenceToNum(ref),
            _display = getXmlAttibute(attrList, "display", null),
            _address = getXmlAttibute(attrList, "location", null),
            _tooltip = getXmlAttibute(attrList, "tooltip", null);
        let _type: IluckysheetHyperlinkType = _address ? "internal" : "external";
  
        // external hyperlink
        if (!_address) {
          let rid = attrList["r:id"];
          let sheetFile = this.sheetFile;
          let relationshipList = this.readXml.getElementsByTagName(
            "Relationships/Relationship",
            `xl/worksheets/_rels/${sheetFile.replace(worksheetFilePath, "")}.rels`
          );
  
          const findRid = relationshipList?.find(
            (e) => e.attributeList["Id"] === rid
          );

          if (findRid) {
            _address = findRid.attributeList["Target"];
            _type = findRid.attributeList[
              "TargetMode"
            ]?.toLocaleLowerCase() as IluckysheetHyperlinkType;
          }
        }

        // match R1C1
        const addressReg = new RegExp(/^.*!R([\d$])+C([\d$])*$/g)
        if (addressReg.test(_address)) {
          _address = getTransR1C1ToSequence(_address);
        }
        
        // dynamically add hyperlinks
        for (const ref of refArr) {
          hyperlink[ref] = {
            linkAddress: _address,
            linkTooltip: _tooltip || "",
            linkType: _type,
            display: _display || "",
          };
        }
      }
      
      return hyperlink;
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

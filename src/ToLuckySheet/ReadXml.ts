import {IuploadfileList, IattributeList, stringToNum} from "../ICommon";
import {indexedColors}  from "../common/constant";
import {LightenDarkenColor}  from "../common/method";


class xmloperation {
    /**
    * @param tag Search xml tag name , div,title etc.
    * @param file Xml string
    * @return Xml element string 
    */
    protected getElementsByOneTag(tag:string, file:string):string[]{
        //<a:[^/>: ]+?>.*?</a:[^/>: ]+?>
        let readTagReg;
        if(tag.indexOf("|")>-1){
            let tags = tag.split("|"), tagsRegTxt="";
            for(let i=0;i<tags.length;i++){
                let t = tags[i];
                tagsRegTxt += "|<"+ t +" [^>]+?[^/]>[\\s\\S]*?</"+ t +">|<"+ t +" [^>]+?/>|<"+ t +">[\\s\\S]*?</"+ t +">|<"+ t +"/>";
            }
            tagsRegTxt = tagsRegTxt.substr(1, tagsRegTxt.length);
            readTagReg = new RegExp(tagsRegTxt, "g");
        }
        else{
            readTagReg = new RegExp("<"+ tag +" [^>]+?[^/]>[\\s\\S]*?</"+ tag +">|<"+ tag +" [^>]+?/>|<"+ tag +">[\\s\\S]*?</"+ tag +">|<"+ tag +"/>", "g");
        }
        
        let ret = file.match(readTagReg);
        if(ret==null){
            return [];
        }
        else{
            return ret;
        }
    }
}

export class ReadXml extends xmloperation{
    originFile:IuploadfileList
    constructor(files:IuploadfileList){
        super();
        this.originFile = files;
    }
    /**
    * @param path Search xml tag group , div,title etc.
    * @param fileName One of uploadfileList, uploadfileList is file group, {key:value}
    * @return Xml element calss
    */
    getElementsByTagName(path:string, fileName:string): Element[]{
        
        let file = this.getFileByName(fileName);
        let pathArr = path.split("/"), ret:string[] | string;
        for(let key in pathArr){
            let path = pathArr[key];
            if(ret==undefined){
                ret = this.getElementsByOneTag(path,file);
            }
            else{
                if(ret instanceof Array){
                    let items:string[]=[];
                    for(let key in ret){
                        let item = ret[key];
                        items = items.concat(this.getElementsByOneTag(path,item));
                    }
                    ret = items;
                }
                else{
                    ret = this.getElementsByOneTag(path,ret);
                }
            }
        }

        let elements:Element[] = [];

        for(let i=0;i<ret.length;i++){
            let ele = new Element(ret[i]);
            elements.push(ele);
        }

        return elements;
    }

    /**
    * @param name One of uploadfileList's name, search for file by this parameter
    * @retrun Select a file from uploadfileList
    */
    private getFileByName(name:string):string{
        for(let fileKey in this.originFile){
            if(fileKey.indexOf(name)>-1){
                return this.originFile[fileKey];
            }
        }
        return "";
    }

    
}

export class Element extends xmloperation {
    elementString:string
    attributeList:IattributeList
    value:string
    container:string
    constructor(str:string){
        super();
        this.elementString = str;
        this.setValue();
        const readAttrReg = new RegExp('[a-zA-Z0-9_:]*?=".*?"', "g");
        let attrList = this.container.match(readAttrReg);
        this.attributeList = {};
        if(attrList!=null){
            for(let key in attrList){
                let attrFull = attrList[key];
                // let al= attrFull.split("=");
                if(attrFull.length==0){
                    continue;
                }
                let attrKey = attrFull.substr(0, attrFull.indexOf('='));
                let attrValue = attrFull.substr(attrFull.indexOf('=') + 1);
                if(attrKey==null || attrValue==null ||attrKey.length==0 || attrValue.length==0){
                    continue;
                }
                this.attributeList[attrKey] = attrValue.substr(1, attrValue.length-2);
            }
        }
    }

    /**
    * @param name Get attribute by key in element
    * @return Single attribute
    */
    get(name:string):string|number|boolean{
        return this.attributeList[name];
    }

    /**
    * @param tag Get elements by tag in elementString
    * @return Element group
    */
    getInnerElements(tag:string):Element[]{
        let ret = this.getElementsByOneTag(tag,this.elementString);
        let elements:Element[] = [];

        for(let i=0;i<ret.length;i++){
            let ele = new Element(ret[i]);
            elements.push(ele);
        }

        if(elements.length==0){
            return null;
        }
        return elements;
    }

    /**
    * @desc get xml dom value and container, <container>value</container>
    */
    private setValue(){
        let str = this.elementString;
        if(str.substr(str.length-2, 2)=="/>"){
            this.value = "";
            this.container = str;
        }
        else{
            let firstTag = this.getFirstTag();
            const firstTagReg = new RegExp("(<"+ firstTag +" [^>]+?[^/]>)([\\s\\S]*?)</"+ firstTag +">|(<"+ firstTag +">)([\\s\\S]*?)</"+ firstTag +">", "g");
            let result = firstTagReg.exec(str);
            if (result != null) {
                if(result[1]!=null){
                    this.container = result[1];
                    this.value = result[2];
                }
                else{
                    this.container = result[3];
                    this.value = result[4];
                }
            }
        }
    }

    /**
    * @desc get xml dom first tag, <a><b></b></a>, get a
    */
    private getFirstTag(){
        let str = this.elementString;
        let firstTag = str.substr(0, str.indexOf(' '));
        if(firstTag=="" || firstTag.indexOf(">")>-1){
            firstTag = str.substr(0, str.indexOf('>'));
        }
        firstTag = firstTag.substr(1,firstTag.length);
        return firstTag;
    }
}


export interface IStyleCollections {
    [index:string]:Element[] | IattributeList
}

function combineIndexedColor(indexedColorsInner:Element[], indexedColors:IattributeList):IattributeList{
    let ret:IattributeList = {};
    if(indexedColorsInner==null || indexedColorsInner.length==0){
        return indexedColors;
    }
    for(let key in indexedColors){
        let value = indexedColors[key], kn = parseInt(key);
        let inner = indexedColorsInner[kn];
        if(inner==null){
            ret[key] = value;
        }
        else{
            let rgb = inner.attributeList.rgb;
            ret[key] = rgb;
        }
    }

    return ret;
}

//clrScheme:Element[]
export function getColor(color:Element, styles:IStyleCollections , type:string="g"){
    let attrList = color.attributeList;
    let clrScheme = styles["clrScheme"] as Element[];
    let indexedColorsInner = styles["indexedColors"] as Element[];
    let mruColorsInner = styles["mruColors"];
    let indexedColorsList = combineIndexedColor(indexedColorsInner, indexedColors);
    let indexed = attrList.indexed, rgb = attrList.rgb, theme = attrList.theme, tint = attrList.tint;
    let bg;
    if(indexed!=null){
        let indexedNum = parseInt(indexed);
        bg = indexedColorsList[indexedNum];
        if(bg!=null){
            bg = bg.substring(bg.length-6, bg.length);
            bg = "#"+bg;
        }
    }
    else if(rgb!=null){
        rgb = rgb.substring(rgb.length-6, rgb.length);
        bg = "#"+rgb;
    }
    else if(theme!=null){
        let themeNum = parseInt(theme);
        if(themeNum==0){
            themeNum = 1;
        }
        else if(themeNum==1){
            themeNum = 0;
        }
        else if(themeNum==2){
            themeNum = 3;
        }
        else if(themeNum==3){
            themeNum = 2;
        }
        let clrSchemeElement = clrScheme[themeNum];
        if(clrSchemeElement!=null){
            let clrs = clrSchemeElement.getInnerElements("a:sysClr|a:srgbClr");
            if(clrs!=null){
                let clr = clrs[0];
                let clrAttrList = clr.attributeList;
                // console.log(clr.container, );
                if(clr.container.indexOf("sysClr")>-1){
                    // if(type=="g" && clrAttrList.val=="windowText"){
                    //     bg = null;
                    // }
                    // else if((type=="t" || type=="b") && clrAttrList.val=="window"){
                    //     bg = null;
                    // }                    
                    // else 
                    if(clrAttrList.lastClr!=null){
                        bg = "#" + clrAttrList.lastClr;
                    }
                    else if(clrAttrList.val!=null){
                        bg = "#" + clrAttrList.val;
                    }

                }
                else if(clr.container.indexOf("srgbClr")>-1){
                    // console.log(clrAttrList.val);
                    bg = "#" + clrAttrList.val;
                }
            }
        }
        
    }
    
    if(tint!=null){
        let tintNum = parseFloat(tint);
        if(bg!=null){
            bg = LightenDarkenColor(bg, tintNum);
        }
    }

    return bg;
}


/** 
 * @dom xml attribute object
 * @attr attribute name
 * @d if attribute is null, return default value 
 * @return attribute value
*/
export function getlineStringAttr(frpr:Element, attr:string):string{
    let attrEle = frpr.getInnerElements(attr), value;

    if(attrEle!=null && attrEle.length>0){
        if(attr=="b" || attr=="i" || attr=="strike"){
            value = "1";
        }
        else if(attr=="u"){
            let v = attrEle[0].attributeList.val;
            if(v=="double"){
                value =  "2";
            }
            else if(v=="singleAccounting"){
                value =  "3";
            }
            else if(v=="doubleAccounting"){
                value =  "4";
            }
            else{
                value = "1";
            }
        }
        else if(attr=="vertAlign"){
            let v = attrEle[0].attributeList.val;
            if(v=="subscript"){
                value = "1";
            }
            else if(v=="superscript"){
                value = "2";
            }
        }
        else{
            value = attrEle[0].attributeList.val;
        }
        
    }

    return value;
}
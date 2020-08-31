import { columeHeader_word, columeHeader_word_index } from "./constant";
import { IluckySheetSelection } from "../ToLuckySheet/ILuck";
import { IattributeList, stringToNum} from "../ICommon";


export function getRangetxt(range:IluckySheetSelection, sheettxt:string) {

    let row0 = range["row"][0], row1 = range["row"][1];
    let column0 = range["column"][0], column1 = range["column"][1];

    if (row0 == null && row1 == null) {
        return sheettxt + chatatABC(column0) + ":" + chatatABC(column1);
    }
    else if (column0 == null && column1 == null) {
        return sheettxt + (row0 + 1) + ":" + (row1 + 1);
    }
    else {
        if (column0 == column1 && row0 == row1) {
            return sheettxt + chatatABC(column0) + (row0 + 1);
        }
        else {
            return sheettxt + chatatABC(column0) + (row0 + 1) + ":" + chatatABC(column1) + (row1 + 1);
        }
    }
}


export function getcellrange (txt:string, sheets:IattributeList={}, sheetId:string="1") {
    let val = txt.split("!");

    let sheettxt = "",
        rangetxt = "",
        sheetIndex = -1;

    if (val.length > 1) {
        sheettxt = val[0];
        rangetxt = val[1];
        
        let si = sheets[sheettxt];
        if(si==null){
            sheetIndex = parseInt(sheetId);
        }
        else{
            sheetIndex = parseInt(si);
        }
    } 
    else {
        sheetIndex = parseInt(sheetId);
        rangetxt = val[0];
    }
    
    if (rangetxt.indexOf(":") == -1) {
        let row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
        let col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));

        if (!isNaN(row) && !isNaN(col)) {
            return {
                "row": [row, row],
                "column": [col, col],
                "sheetIndex": sheetIndex
            };
        }
        else {
            return null;
        }
    } 
    else {
        let rangetxtArray:string[] = rangetxt.split(":");
        let row = [],col = [];
        row[0] = parseInt(rangetxtArray[0].replace(/[^0-9]/g, "")) - 1;
        row[1] = parseInt(rangetxtArray[1].replace(/[^0-9]/g, "")) - 1;
        // if (isNaN(row[0])) {
        //     row[0] = 0;
        // }
        // if (isNaN(row[1])) {
        //     row[1] = sheetdata.length - 1;
        // }
        if (row[0] > row[1]) {
            return null;
        }
        col[0] = ABCatNum(rangetxtArray[0].replace(/[^A-Za-z]/g, ""));
        col[1] = ABCatNum(rangetxtArray[1].replace(/[^A-Za-z]/g, ""));
        // if (isNaN(col[0])) {
        //     col[0] = 0;
        // }
        // if (isNaN(col[1])) {
        //     col[1] = sheetdata[0].length - 1;
        // }
        if (col[0] > col[1]) {
            return null;
        }

        return {
            "row": row,
            "column": col,
            "sheetIndex": sheetIndex
        };
    }
}

//列下标  字母转数字
function ABCatNum(abc:string) {
    abc = abc.toUpperCase();

    let abc_len = abc.length;
    if (abc_len == 0) {
        return NaN;
    }

    let abc_array = abc.split("");
    let wordlen = columeHeader_word.length;
    let ret = 0;

    for (let i = abc_len - 1; i >= 0; i--) {
        if (i == abc_len - 1) {
            ret += columeHeader_word_index[abc_array[i]];
        }
        else {
            ret += Math.pow(wordlen, abc_len - i - 1) * (columeHeader_word_index[abc_array[i]] + 1);
        }
    }

    return ret;
}

//列下标  数字转字母
function chatatABC(index:number) {
    let wordlen = columeHeader_word.length;

    if (index < wordlen) {
        return columeHeader_word[index];
    }
    else {
        let last = 0, pre = 0, ret = "";
        let i = 1, n = 0;

        while (index >= (wordlen / (wordlen - 1)) * (Math.pow(wordlen, i++) - 1)) {
            n = i;
        }

        let index_ab = index - (wordlen / (wordlen - 1)) * (Math.pow(wordlen, n - 1) - 1);//970
        last = index_ab + 1;

        for (let x = n; x > 0; x--) {
            let last1 = last, x1 = x;//-702=268, 3

            if (x == 1) {
                last1 = last1 % wordlen;

                if (last1 == 0) {
                    last1 = 26;
                }

                return ret + columeHeader_word[last1 - 1];
            }

            last1 = Math.ceil(last1 / Math.pow(wordlen, x - 1));
            //last1 = last1 % wordlen;
            ret += columeHeader_word[last1 - 1];

            if (x > 1) {
                last = last - (last1 - 1) * wordlen;
            }
        }
    }
}

/** 
 * @dom xml attribute object
 * @attr attribute name
 * @d if attribute is null, return default value 
 * @return attribute value
*/
export function getXmlAttibute(dom:IattributeList, attr:string, d:string){
    let value = dom[attr];
    value = value==null?d:value;
    return value;
}


/** 
 * @columnWidth Excel column width
 * @return pixel column width
*/
export function getColumnWidthPixel(columnWidth:number){
    let pix = Math.round(columnWidth * 8 + 5);
    return pix;
}

/** 
 * @rowHeight Excel row height
 * @return pixel row height
*/
export function getRowHeightPixel(rowHeight:number){
    let pix = Math.round(rowHeight/0.75);
    return pix;
}

export function LightenDarkenColor(sixColor:string, tint:number){
    let hex:string = sixColor.substring(sixColor.length-6,sixColor.length);
    let rgbArray:number[] = hexToRgbArray("#"+hex);
    let hslArray = rgbToHsl(rgbArray[0], rgbArray[1],rgbArray[2]);
    if(tint>0){
        hslArray[2] = hslArray[2] * (1.0-tint) + tint;
    }
    else if(tint<0){
        hslArray[2] = hslArray[2] * (1.0 + tint)
    }
    else{
        return "#"+hex;
    }

    let newRgbArray = hslToRgb(hslArray[0],hslArray[1],hslArray[2]);

    return rgbToHex("RGB(" + newRgbArray.join(",") + ")");
}


function rgbToHex(rgb:string){
    //十六进制颜色值的正则表达式
    var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
    // 如果是rgb颜色表示
    if (/^(rgb|RGB)/.test(rgb)) {
        var aColor = rgb.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
        var strHex = "#";
        for (var i=0; i<aColor.length; i++) {
            var hex = Number(aColor[i]).toString(16);
            if (hex.length < 2) {
                hex = '0' + hex;    
            }
            strHex += hex;
        }
        if (strHex.length !== 7) {
            strHex = rgb;    
        }
        return strHex;
    } else if (reg.test(rgb)) {
        var aNum = rgb.replace(/#/,"").split("");
        if (aNum.length === 6) {
            return rgb;    
        } else if(aNum.length === 3) {
            var numHex = "#";
            for (var i=0; i<aNum.length; i+=1) {
                numHex += (aNum[i] + aNum[i]);
            }
            return numHex;
        }
    }
    return rgb;
}

function hexToRgb(hex:string){
    var sColor = hex.toLowerCase();
    //十六进制颜色值的正则表达式
    var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
    // 如果是16进制颜色
    if (sColor && reg.test(sColor)) {
        if (sColor.length === 4) {
            var sColorNew = "#";
            for (var i=1; i<4; i+=1) {
                sColorNew += sColor.slice(i, i+1).concat(sColor.slice(i, i+1));    
            }
            sColor = sColorNew;
        }
        //处理六位的颜色值
        var sColorChange = [];
        for (var i=1; i<7; i+=2) {
            sColorChange.push(parseInt("0x"+sColor.slice(i, i+2)));    
        }
        return "RGB(" + sColorChange.join(",") + ")";
    }
    return sColor;
}

function hexToRgbArray(hex:string){
    var sColor = hex.toLowerCase();
    //十六进制颜色值的正则表达式
    var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
    // 如果是16进制颜色
    if (sColor && reg.test(sColor)) {
        if (sColor.length === 4) {
            var sColorNew = "#";
            for (var i=1; i<4; i+=1) {
                sColorNew += sColor.slice(i, i+1).concat(sColor.slice(i, i+1));    
            }
            sColor = sColorNew;
        }
        //处理六位的颜色值
        var sColorChange:number[] = [];
        for (var i=1; i<7; i+=2) {
            sColorChange.push(parseInt("0x"+sColor.slice(i, i+2)));    
        }
        return  sColorChange;
    }
    return null;
}

/**
 * HSL颜色值转换为RGB. 
 * 换算公式改编自 http://en.wikipedia.org/wiki/HSL_color_space.
 * h, s, 和 l 设定在 [0, 1] 之间
 * 返回的 r, g, 和 b 在 [0, 255]之间
 *
 * @param   Number  h       色相
 * @param   Number  s       饱和度
 * @param   Number  l       亮度
 * @return  Array           RGB色值数值
 */
function hslToRgb(h:number, s:number, l:number) {
    var r, g, b;

    if(s == 0) {
        r = g = b = l; // achromatic
    } else {
        var hue2rgb = function hue2rgb(p:number, q:number, t:number) {
            if(t < 0) t += 1;
            if(t > 1) t -= 1;
            if(t < 1/6) return p + (q - p) * 6 * t;
            if(t < 1/2) return q;
            if(t < 2/3) return p + (q - p) * (2/3 - t) * 6;
            return p;
        }

        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        r = hue2rgb(p, q, h + 1/3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1/3);
    }

    return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}


/**
 * RGB 颜色值转换为 HSL.
 * 转换公式参考自 http://en.wikipedia.org/wiki/HSL_color_space.
 * r, g, 和 b 需要在 [0, 255] 范围内
 * 返回的 h, s, 和 l 在 [0, 1] 之间
 *
 * @param   Number  r       红色色值
 * @param   Number  g       绿色色值
 * @param   Number  b       蓝色色值
 * @return  Array           HSL各值数组
 */
function rgbToHsl(r:number, g:number, b:number) {
    r /= 255, g /= 255, b /= 255;
    var max = Math.max(r, g, b), min = Math.min(r, g, b);
    var h, s, l = (max + min) / 2;

    if (max == min){ 
        h = s = 0; // achromatic
    } else {
        var d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch(max) {
            case r: h = (g - b) / d + (g < b ? 6 : 0); break;
            case g: h = (b - r) / d + 2; break;
            case b: h = (r - g) / d + 4; break;
        }
        h /= 6;
    }

    return [h, s, l];
}
 
export function generateRandomSheetIndex(prefix:string):string {
    if(prefix == null){
        prefix = "Sheet";
    }

    let userAgent = window.navigator.userAgent.replace(/[^a-zA-Z0-9]/g, "").split("");

    let mid = "";

    for(let i = 0; i < 5; i++){
        mid += userAgent[Math.round(Math.random() * (userAgent.length - 1))];
    }

    let time = new Date().getTime();

    return prefix + "_" + mid + "_" + time;
}


export function escapeCharacter(str:string){
    if(str==null || str.length==0){
        return str;
    }

    return str.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ');
}


export class fromulaRef {

    static operator = '==|!=|<>|<=|>=|=|+|-|>|<|/|*|%|&|^'
    static error = {
        v: "#VALUE!",    //错误的参数或运算符
        n: "#NAME?",     //公式名称错误
        na: "#N/A",      //函数或公式中没有可用数值
        r: "#REF!",      //删除了由其他公式引用的单元格
        d: "#DIV/0!",    //除数是0或空单元格
        nm: "#NUM!",     //当公式或函数中某个数字有问题时
        nl: "#NULL!",    //交叉运算符（空格）使用不正确
        sp: "#SPILL!"    //数组范围有其它值
    }

    static operatorjson:stringToNum = null

    static trim(str:string) {  
        if(str == null){  
            str = "";  
        }  
        return str.replace(/(^\s*)|(\s*$)/g, "");  
    }

    static functionCopy(txt:string, mode:string, step:number) {
        let _this = this;

        if (_this.operatorjson == null) {
            let arr = _this.operator.split("|"),
                op:stringToNum = {};

            for (let i = 0; i < arr.length; i++) {
                op[arr[i].toString()] = 1;
            }

            _this.operatorjson = op;
        }

        if (mode == null) {
            mode = "down";
        }

        if (step == null) {
            step = 1;
        }

        if (txt.substr(0, 1) == "=") {
            txt = txt.substr(1);
        }

        let funcstack = txt.split("");
        let i = 0,
            str = "",
            function_str = "",
            ispassby = true;
        
        let matchConfig = {
            "bracket": 0,
            "comma": 0,
            "squote": 0,
            "dquote": 0
        };

        while (i < funcstack.length) {
            let s = funcstack[i];

            if (s == "(" && matchConfig.dquote == 0) {
                matchConfig.bracket += 1;

                if (str.length > 0) {
                    function_str += str + "(";
                } 
                else {
                    function_str += "(";
                }

                str = "";
            } 
            else if (s == ")" && matchConfig.dquote == 0) {
                matchConfig.bracket -= 1;
                function_str += _this.functionCopy(str, mode, step) + ")";
                str = "";
            }
            else if (s == '"' && matchConfig.squote == 0) {
                if (matchConfig.dquote > 0) {
                    function_str += str + '"';
                    matchConfig.dquote -= 1;
                    str = "";
                } 
                else {
                    matchConfig.dquote += 1;
                    str += '"';
                }
            } 
            else if (s == ',' && matchConfig.dquote == 0) {
                function_str += _this.functionCopy(str, mode, step) + ',';
                str = "";
            } 
            else if (s == '&' && matchConfig.dquote == 0) {
                if (str.length > 0) {
                    function_str += _this.functionCopy(str, mode, step) + "&";
                    str = "";
                } 
                else {
                    function_str += "&";
                }
            } 
            else if (s in _this.operatorjson && matchConfig.dquote == 0) {
                let s_next = "";

                if ((i + 1) < funcstack.length) {
                    s_next = funcstack[i + 1];
                }

                let p = i - 1, 
                    s_pre = null;

                if(p >= 0){
                    do {
                        s_pre = funcstack[p--];
                    }
                    while (p>=0 && s_pre ==" ")
                }

                if ((s + s_next) in _this.operatorjson) {
                    if (str.length > 0) {
                        function_str += _this.functionCopy(str, mode, step) + s + s_next;
                        str = "";
                    } 
                    else {
                        function_str += s + s_next;
                    }

                    i++;
                }
                else if(!(/[^0-9]/.test(s_next)) && s=="-" && (s_pre=="(" || s_pre == null || s_pre == "," || s_pre == " " || s_pre in _this.operatorjson ) ){
                    str += s;
                }
                else {
                    if (str.length > 0) {
                        function_str += _this.functionCopy(str, mode, step) + s;
                        str = "";
                    } 
                    else {
                        function_str += s;
                    }
                }
            } 
            else {
                str += s;
            }

            if (i == funcstack.length - 1) {
                if (_this.iscelldata(_this.trim(str))) {
                    if (mode == "down") {
                        function_str += _this.downparam(_this.trim(str), step);
                    } 
                    else if (mode == "up") {
                        function_str += _this.upparam(_this.trim(str), step);
                    } 
                    else if (mode == "left") {
                        function_str += _this.leftparam(_this.trim(str), step);
                    } 
                    else if (mode == "right") {
                        function_str += _this.rightparam(_this.trim(str), step);
                    }
                } 
                else {
                    function_str += _this.trim(str);
                }
            }
            
            i++;
        }

        return function_str;
    }


    static downparam(txt:string, step:number) {
        return this.updateparam("d", txt, step);
    }

    static upparam(txt:string, step:number) {
        return this.updateparam("u", txt, step);
    }

    static leftparam(txt:string, step:number) {
        return this.updateparam("l", txt, step);
    }

    static rightparam (txt:string, step:number) {
        return this.updateparam("r", txt, step);
    }


    static updateparam (orient:string, txt:string, step:number) {
        let _this = this;
        let val = txt.split("!"),
            rangetxt, prefix = "";
        
        if (val.length > 1) {
            rangetxt = val[1];
            prefix = val[0] + "!";
        } 
        else {
            rangetxt = val[0];
        }

        if (rangetxt.indexOf(":") == -1) {
            let row = parseInt(rangetxt.replace(/[^0-9]/g, ""));
            let col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
            let freezonFuc = _this.isfreezonFuc(rangetxt);
            let $row = freezonFuc[0] ? "$" : "",
                $col = freezonFuc[1] ? "$" : "";
            
            if (orient == "u" && !freezonFuc[0]) {
                row -= step;
            } 
            else if (orient == "r" && !freezonFuc[1]) {
                col += step;
            } 
            else if (orient == "l" && !freezonFuc[1]) {
                col -= step;
            } 
            else if (!freezonFuc[0]) {
                row += step;
            }

            if(row < 0 || col < 0){
                return _this.error.r;
            }
            
            if (!isNaN(row) && !isNaN(col)) {
                return prefix + $col + chatatABC(col) + $row + (row);
            } 
            else if (!isNaN(row)) {
                return prefix + $row + (row);
            } 
            else if (!isNaN(col)) {
                return prefix + $col + chatatABC(col);
            } 
            else {
                return txt;
            }
        } 
        else {
            rangetxt = rangetxt.split(":");
            let row = [],
                col = [];

            row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, ""));
            row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, ""));
            if (row[0] > row[1]) {
                return txt;
            }
            
            col[0] = ABCatNum(rangetxt[0].replace(/[^A-Za-z]/g, ""));
            col[1] = ABCatNum(rangetxt[1].replace(/[^A-Za-z]/g, ""));
            if (col[0] > col[1]) {
                return txt;
            }

            let freezonFuc0 = _this.isfreezonFuc(rangetxt[0]);
            let freezonFuc1 = _this.isfreezonFuc(rangetxt[1]);
            let $row0 = freezonFuc0[0] ? "$" : "",
                $col0 = freezonFuc0[1] ? "$" : "";
            let $row1 = freezonFuc1[0] ? "$" : "",
                $col1 = freezonFuc1[1] ? "$" : "";
            
            if (orient == "u") {
                if (!freezonFuc0[0]) {
                    row[0] -= step;
                }

                if (!freezonFuc1[0]) {
                    row[1] -= step;
                }
            } 
            else if (orient == "r") {
                if (!freezonFuc0[1]) {
                    col[0] += step;
                }

                if (!freezonFuc1[1]) {
                    col[1] += step;
                }
            } 
            else if (orient == "l") {
                if (!freezonFuc0[1]) {
                    col[0] -= step;
                }

                if (!freezonFuc1[1]) {
                    col[1] -= step;
                }
            } 
            else {
                if (!freezonFuc0[0]) {
                    row[0] += step;
                }

                if (!freezonFuc1[0]) {
                    row[1] += step;
                }
            }

            if(row[0] < 0 || col[0] < 0){
                return _this.error.r;
            }

            if (isNaN(col[0]) && isNaN(col[1])) {
                return prefix + $row0 + (row[0]) + ":" + $row1 + (row[1]);
            } 
            else if (isNaN(row[0]) && isNaN(row[1])) {
                return prefix + $col0 + chatatABC(col[0]) + ":" + $col1 + chatatABC(col[1]);
            } 
            else {
                return prefix + $col0 + chatatABC(col[0]) + $row0 + (row[0]) + ":" + $col1 + chatatABC(col[1]) + $row1 + (row[1]);
            }
        }
    }


    static iscelldata(txt:string) { //判断是否为单元格格式
        let val = txt.split("!"),
            rangetxt;

        if (val.length > 1) {
            rangetxt = val[1];
        } 
        else {
            rangetxt = val[0];
        }

        let reg_cell = /^(([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+))$/g; //增加正则判断单元格为字母+数字的格式：如 A1:B3
        let reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+))))$/g; //增加正则判断单元格为字母+数字或字母的格式：如 A1:B3，A:A
        
        if (rangetxt.indexOf(":") == -1) {
            let row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
            let col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
            
            if (!isNaN(row) && !isNaN(col) && rangetxt.toString().match(reg_cell)) {
                return true;
            } 
            else if (!isNaN(row)) {
                return false;
            } 
            else if (!isNaN(col)) {
                return false;
            } 
            else {
                return false;
            }
        } 
        else {
            reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+)))|((([0-9]+)|([$][0-9]+s))))$/g;

            rangetxt = rangetxt.split(":");

            let row = [],col = [];
            row[0] = parseInt(rangetxt[0].replace(/[^0-9]/g, "")) - 1;
            row[1] = parseInt(rangetxt[1].replace(/[^0-9]/g, "")) - 1;
            if (row[0] > row[1]) {
                return false;
            }

            col[0] = ABCatNum(rangetxt[0].replace(/[^A-Za-z]/g, ""));
            col[1] = ABCatNum(rangetxt[1].replace(/[^A-Za-z]/g, ""));
            if (col[0] > col[1]) {
                return false;
            }

            if(rangetxt[0].toString().match(reg_cellRange) && rangetxt[1].toString().match(reg_cellRange)){
                return true;
            }
            else{
                return false;
            }
        }
    }

    static isfreezonFuc(txt:string) {
        let row = txt.replace(/[^0-9]/g, "");
        let col = txt.replace(/[^A-Za-z]/g, "");
        let row$ = txt.substr(txt.indexOf(row) - 1, 1);
        let col$ = txt.substr(txt.indexOf(col) - 1, 1);
        let ret = [false, false];

        if (row$ == "$") {
            ret[0] = true;
        }
        if (col$ == "$") {
            ret[1] = true;
        }

        return ret;
    }

}
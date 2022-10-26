'use strict';

var JSZip = require('jszip');

function _interopDefaultLegacy (e) { return e && typeof e === 'object' && 'default' in e ? e : { 'default': e }; }

var JSZip__default = /*#__PURE__*/_interopDefaultLegacy(JSZip);

/*! *****************************************************************************
Copyright (c) Microsoft Corporation.

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH
REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY
AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT,
INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM
LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR
OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR
PERFORMANCE OF THIS SOFTWARE.
***************************************************************************** */
/* global Reflect, Promise */

var extendStatics = function(d, b) {
    extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
    return extendStatics(d, b);
};

function __extends(d, b) {
    extendStatics(d, b);
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
}

var columeHeader_word = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
var columeHeader_word_index = { 'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25 };
var coreFile = "docProps/core.xml";
var appFile = "docProps/app.xml";
var workBookFile = "xl/workbook.xml";
var calcChainFile = "xl/calcChain.xml";
var stylesFile = "xl/styles.xml";
var sharedStringsFile = "xl/sharedStrings.xml";
var theme1File = "xl/theme/theme1.xml";
var workbookRels = "xl/_rels/workbook.xml.rels";
//Excel Built-In cell type
var ST_CellType = {
    "Boolean": "b",
    "Date": "d",
    "Error": "e",
    "InlineString": "inlineStr",
    "Number": "n",
    "SharedString": "s",
    "String": "str",
};
var numFmtDefault = {
    "0": 'General',
    "1": '0',
    "2": '0.00',
    "3": '#,##0',
    "4": '#,##0.00',
    "9": '0%',
    "10": '0.00%',
    "11": '0.00E+00',
    "12": '# ?/?',
    "13": '# ??/??',
    "14": 'm/d/yy',
    "15": 'd-mmm-yy',
    "16": 'd-mmm',
    "17": 'mmm-yy',
    "18": 'h:mm AM/PM',
    "19": 'h:mm:ss AM/PM',
    "20": 'h:mm',
    "21": 'h:mm:ss',
    "22": 'm/d/yy h:mm',
    "37": '#,##0 ;(#,##0)',
    "38": '#,##0 ;[Red](#,##0)',
    "39": '#,##0.00;(#,##0.00)',
    "40": '#,##0.00;[Red](#,##0.00)',
    "45": 'mm:ss',
    "46": '[h]:mm:ss',
    "47": 'mmss.0',
    "48": '##0.0E+0',
    "49": '@'
};
var indexedColors = {
    "0": '00000000',
    "1": '00FFFFFF',
    "2": '00FF0000',
    "3": '0000FF00',
    "4": '000000FF',
    "5": '00FFFF00',
    "6": '00FF00FF',
    "7": '0000FFFF',
    "8": '00000000',
    "9": '00FFFFFF',
    "10": '00FF0000',
    "11": '0000FF00',
    "12": '000000FF',
    "13": '00FFFF00',
    "14": '00FF00FF',
    "15": '0000FFFF',
    "16": '00800000',
    "17": '00008000',
    "18": '00000080',
    "19": '00808000',
    "20": '00800080',
    "21": '00008080',
    "22": '00C0C0C0',
    "23": '00808080',
    "24": '009999FF',
    "25": '00993366',
    "26": '00FFFFCC',
    "27": '00CCFFFF',
    "28": '00660066',
    "29": '00FF8080',
    "30": '000066CC',
    "31": '00CCCCFF',
    "32": '00000080',
    "33": '00FF00FF',
    "34": '00FFFF00',
    "35": '0000FFFF',
    "36": '00800080',
    "37": '00800000',
    "38": '00008080',
    "39": '000000FF',
    "40": '0000CCFF',
    "41": '00CCFFFF',
    "42": '00CCFFCC',
    "43": '00FFFF99',
    "44": '0099CCFF',
    "45": '00FF99CC',
    "46": '00CC99FF',
    "47": '00FFCC99',
    "48": '003366FF',
    "49": '0033CCCC',
    "50": '0099CC00',
    "51": '00FFCC00',
    "52": '00FF9900',
    "53": '00FF6600',
    "54": '00666699',
    "55": '00969696',
    "56": '00003366',
    "57": '00339966',
    "58": '00003300',
    "59": '00333300',
    "60": '00993300',
    "61": '00993366',
    "62": '00333399',
    "63": '00333333',
    "64": null,
    "65": null,
};
var borderTypes = {
    "none": 0,
    "thin": 1,
    "hair": 2,
    "dotted": 3,
    "dashed": 4,
    "dashDot": 5,
    "dashDotDot": 6,
    "double": 7,
    "medium": 8,
    "mediumDashed": 9,
    "mediumDashDot": 10,
    "mediumDashDotDot": 11,
    "slantDashDot": 12,
    "thick": 13
};
var fontFamilys = {
    "0": "defualt",
    "1": "Roman",
    "2": "Swiss",
    "3": "Modern",
    "4": "Script",
    "5": "Decorative"
};

function getcellrange(txt, sheets, sheetId) {
    if (sheets === void 0) { sheets = {}; }
    if (sheetId === void 0) { sheetId = "1"; }
    var val = txt.split("!");
    var sheettxt = "", rangetxt = "", sheetIndex = -1;
    if (val.length > 1) {
        sheettxt = val[0];
        rangetxt = val[1];
        var si = sheets[sheettxt];
        if (si == null) {
            sheetIndex = parseInt(sheetId);
        }
        else {
            sheetIndex = parseInt(si);
        }
    }
    else {
        sheetIndex = parseInt(sheetId);
        rangetxt = val[0];
    }
    if (rangetxt.indexOf(":") == -1) {
        var row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
        var col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
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
        var rangetxtArray = rangetxt.split(":");
        var row = [], col = [];
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
function ABCatNum(abc) {
    abc = abc.toUpperCase();
    var abc_len = abc.length;
    if (abc_len == 0) {
        return NaN;
    }
    var abc_array = abc.split("");
    var wordlen = columeHeader_word.length;
    var ret = 0;
    for (var i = abc_len - 1; i >= 0; i--) {
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
function chatatABC(index) {
    var wordlen = columeHeader_word.length;
    if (index < wordlen) {
        return columeHeader_word[index];
    }
    else {
        var last = 0, ret = "";
        var i = 1, n = 0;
        while (index >= (wordlen / (wordlen - 1)) * (Math.pow(wordlen, i++) - 1)) {
            n = i;
        }
        var index_ab = index - (wordlen / (wordlen - 1)) * (Math.pow(wordlen, n - 1) - 1); //970
        last = index_ab + 1;
        for (var x = n; x > 0; x--) {
            var last1 = last; //-702=268, 3
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
 * @return ratio, default 0.75 1in = 2.54cm = 25.4mm = 72pt = 6pc,  pt = 1/72 In, px = 1/dpi In
*/
function getptToPxRatioByDPI() {
    return 72 / 96;
}
/**
 * @emus EMUs, Excel drawing unit
 * @return pixel
*/
function getPxByEMUs(emus) {
    if (emus == null) {
        return 0;
    }
    var inch = emus / 914400;
    var pt = inch * 72;
    var px = pt / getptToPxRatioByDPI();
    return px;
}
/**
 * @dom xml attribute object
 * @attr attribute name
 * @d if attribute is null, return default value
 * @return attribute value
*/
function getXmlAttibute(dom, attr, d) {
    var value = dom[attr];
    value = value == null ? d : value;
    return value;
}
/**
 * @columnWidth Excel column width
 * @return pixel column width
*/
function getColumnWidthPixel(columnWidth) {
    var pix = Math.round((columnWidth - 0.83) * 8 + 5);
    return pix;
}
/**
 * @rowHeight Excel row height
 * @return pixel row height
*/
function getRowHeightPixel(rowHeight) {
    var pix = Math.round(rowHeight / getptToPxRatioByDPI());
    return pix;
}
function LightenDarkenColor(sixColor, tint) {
    var hex = sixColor.substring(sixColor.length - 6, sixColor.length);
    var rgbArray = hexToRgbArray("#" + hex);
    var hslArray = rgbToHsl(rgbArray[0], rgbArray[1], rgbArray[2]);
    if (tint > 0) {
        hslArray[2] = hslArray[2] * (1.0 - tint) + tint;
    }
    else if (tint < 0) {
        hslArray[2] = hslArray[2] * (1.0 + tint);
    }
    else {
        return "#" + hex;
    }
    var newRgbArray = hslToRgb(hslArray[0], hslArray[1], hslArray[2]);
    return rgbToHex("RGB(" + newRgbArray.join(",") + ")");
}
function rgbToHex(rgb) {
    //十六进制颜色值的正则表达式
    var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
    // 如果是rgb颜色表示
    if (/^(rgb|RGB)/.test(rgb)) {
        var aColor = rgb.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
        var strHex = "#";
        for (var i = 0; i < aColor.length; i++) {
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
    }
    else if (reg.test(rgb)) {
        var aNum = rgb.replace(/#/, "").split("");
        if (aNum.length === 6) {
            return rgb;
        }
        else if (aNum.length === 3) {
            var numHex = "#";
            for (var i = 0; i < aNum.length; i += 1) {
                numHex += (aNum[i] + aNum[i]);
            }
            return numHex;
        }
    }
    return rgb;
}
function hexToRgbArray(hex) {
    var sColor = hex.toLowerCase();
    //十六进制颜色值的正则表达式
    var reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/;
    // 如果是16进制颜色
    if (sColor && reg.test(sColor)) {
        if (sColor.length === 4) {
            var sColorNew = "#";
            for (var i = 1; i < 4; i += 1) {
                sColorNew += sColor.slice(i, i + 1).concat(sColor.slice(i, i + 1));
            }
            sColor = sColorNew;
        }
        //处理六位的颜色值
        var sColorChange = [];
        for (var i = 1; i < 7; i += 2) {
            sColorChange.push(parseInt("0x" + sColor.slice(i, i + 2)));
        }
        return sColorChange;
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
function hslToRgb(h, s, l) {
    var r, g, b;
    if (s == 0) {
        r = g = b = l; // achromatic
    }
    else {
        var hue2rgb = function hue2rgb(p, q, t) {
            if (t < 0)
                t += 1;
            if (t > 1)
                t -= 1;
            if (t < 1 / 6)
                return p + (q - p) * 6 * t;
            if (t < 1 / 2)
                return q;
            if (t < 2 / 3)
                return p + (q - p) * (2 / 3 - t) * 6;
            return p;
        };
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        r = hue2rgb(p, q, h + 1 / 3);
        g = hue2rgb(p, q, h);
        b = hue2rgb(p, q, h - 1 / 3);
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
function rgbToHsl(r, g, b) {
    r /= 255, g /= 255, b /= 255;
    var max = Math.max(r, g, b), min = Math.min(r, g, b);
    var h, s, l = (max + min) / 2;
    if (max == min) {
        h = s = 0; // achromatic
    }
    else {
        var d = max - min;
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        switch (max) {
            case r:
                h = (g - b) / d + (g < b ? 6 : 0);
                break;
            case g:
                h = (b - r) / d + 2;
                break;
            case b:
                h = (r - g) / d + 4;
                break;
        }
        h /= 6;
    }
    return [h, s, l];
}
function generateRandomIndex(prefix) {
    if (prefix == null) {
        prefix = "Sheet";
    }
    var userAgent = window.navigator.userAgent.replace(/[^a-zA-Z0-9]/g, "").split("");
    var mid = "";
    for (var i = 0; i < 5; i++) {
        mid += userAgent[Math.round(Math.random() * (userAgent.length - 1))];
    }
    var time = new Date().getTime();
    return prefix + "_" + mid + "_" + time;
}
function escapeCharacter(str) {
    if (str == null || str.length == 0) {
        return str;
    }
    return str.replace(/&amp;/g, "&").replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&apos;/g, "'").replace(/&iexcl;/g, "¡").replace(/&cent;/g, "¢").replace(/&pound;/g, "£").replace(/&curren;/g, "¤").replace(/&yen;/g, "¥").replace(/&brvbar;/g, "¦").replace(/&sect;/g, "§").replace(/&uml;/g, "¨").replace(/&copy;/g, "©").replace(/&ordf;/g, "ª").replace(/&laquo;/g, "«").replace(/&not;/g, "¬").replace(/&shy;/g, "­").replace(/&reg;/g, "®").replace(/&macr;/g, "¯").replace(/&deg;/g, "°").replace(/&plusmn;/g, "±").replace(/&sup2;/g, "²").replace(/&sup3;/g, "³").replace(/&acute;/g, "´").replace(/&micro;/g, "µ").replace(/&para;/g, "¶").replace(/&middot;/g, "·").replace(/&cedil;/g, "¸").replace(/&sup1;/g, "¹").replace(/&ordm;/g, "º").replace(/&raquo;/g, "»").replace(/&frac14;/g, "¼").replace(/&frac12;/g, "½").replace(/&frac34;/g, "¾").replace(/&iquest;/g, "¿").replace(/&times;/g, "×").replace(/&divide;/g, "÷").replace(/&Agrave;/g, "À").replace(/&Aacute;/g, "Á").replace(/&Acirc;/g, "Â").replace(/&Atilde;/g, "Ã").replace(/&Auml;/g, "Ä").replace(/&Aring;/g, "Å").replace(/&AElig;/g, "Æ").replace(/&Ccedil;/g, "Ç").replace(/&Egrave;/g, "È").replace(/&Eacute;/g, "É").replace(/&Ecirc;/g, "Ê").replace(/&Euml;/g, "Ë").replace(/&Igrave;/g, "Ì").replace(/&Iacute;/g, "Í").replace(/&Icirc;/g, "Î").replace(/&Iuml;/g, "Ï").replace(/&ETH;/g, "Ð").replace(/&Ntilde;/g, "Ñ").replace(/&Ograve;/g, "Ò").replace(/&Oacute;/g, "Ó").replace(/&Ocirc;/g, "Ô").replace(/&Otilde;/g, "Õ").replace(/&Ouml;/g, "Ö").replace(/&Oslash;/g, "Ø").replace(/&Ugrave;/g, "Ù").replace(/&Uacute;/g, "Ú").replace(/&Ucirc;/g, "Û").replace(/&Uuml;/g, "Ü").replace(/&Yacute;/g, "Ý").replace(/&THORN;/g, "Þ").replace(/&szlig;/g, "ß").replace(/&agrave;/g, "à").replace(/&aacute;/g, "á").replace(/&acirc;/g, "â").replace(/&atilde;/g, "ã").replace(/&auml;/g, "ä").replace(/&aring;/g, "å").replace(/&aelig;/g, "æ").replace(/&ccedil;/g, "ç").replace(/&egrave;/g, "è").replace(/&eacute;/g, "é").replace(/&ecirc;/g, "ê").replace(/&euml;/g, "ë").replace(/&igrave;/g, "ì").replace(/&iacute;/g, "í").replace(/&icirc;/g, "î").replace(/&iuml;/g, "ï").replace(/&eth;/g, "ð").replace(/&ntilde;/g, "ñ").replace(/&ograve;/g, "ò").replace(/&oacute;/g, "ó").replace(/&ocirc;/g, "ô").replace(/&otilde;/g, "õ").replace(/&ouml;/g, "ö").replace(/&oslash;/g, "ø").replace(/&ugrave;/g, "ù").replace(/&uacute;/g, "ú").replace(/&ucirc;/g, "û").replace(/&uuml;/g, "ü").replace(/&yacute;/g, "ý").replace(/&thorn;/g, "þ").replace(/&yuml;/g, "ÿ");
}
var fromulaRef = /** @class */ (function () {
    function fromulaRef() {
    }
    fromulaRef.trim = function (str) {
        if (str == null) {
            str = "";
        }
        return str.replace(/(^\s*)|(\s*$)/g, "");
    };
    fromulaRef.functionCopy = function (txt, mode, step) {
        var _this = this;
        if (_this.operatorjson == null) {
            var arr = _this.operator.split("|"), op = {};
            for (var i_1 = 0; i_1 < arr.length; i_1++) {
                op[arr[i_1].toString()] = 1;
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
        var funcstack = txt.split("");
        var i = 0, str = "", function_str = "";
        var matchConfig = {
            "bracket": 0,
            "comma": 0,
            "squote": 0,
            "dquote": 0
        };
        while (i < funcstack.length) {
            var s = funcstack[i];
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
                var s_next = "";
                if ((i + 1) < funcstack.length) {
                    s_next = funcstack[i + 1];
                }
                var p = i - 1, s_pre = null;
                if (p >= 0) {
                    do {
                        s_pre = funcstack[p--];
                    } while (p >= 0 && s_pre == " ");
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
                else if (!(/[^0-9]/.test(s_next)) && s == "-" && (s_pre == "(" || s_pre == null || s_pre == "," || s_pre == " " || s_pre in _this.operatorjson)) {
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
    };
    fromulaRef.downparam = function (txt, step) {
        return this.updateparam("d", txt, step);
    };
    fromulaRef.upparam = function (txt, step) {
        return this.updateparam("u", txt, step);
    };
    fromulaRef.leftparam = function (txt, step) {
        return this.updateparam("l", txt, step);
    };
    fromulaRef.rightparam = function (txt, step) {
        return this.updateparam("r", txt, step);
    };
    fromulaRef.updateparam = function (orient, txt, step) {
        var _this = this;
        var val = txt.split("!"), rangetxt, prefix = "";
        if (val.length > 1) {
            rangetxt = val[1];
            prefix = val[0] + "!";
        }
        else {
            rangetxt = val[0];
        }
        if (rangetxt.indexOf(":") == -1) {
            var row = parseInt(rangetxt.replace(/[^0-9]/g, ""));
            var col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
            var freezonFuc = _this.isfreezonFuc(rangetxt);
            var $row = freezonFuc[0] ? "$" : "", $col = freezonFuc[1] ? "$" : "";
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
            if (row < 0 || col < 0) {
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
            var row = [], col = [];
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
            var freezonFuc0 = _this.isfreezonFuc(rangetxt[0]);
            var freezonFuc1 = _this.isfreezonFuc(rangetxt[1]);
            var $row0 = freezonFuc0[0] ? "$" : "", $col0 = freezonFuc0[1] ? "$" : "";
            var $row1 = freezonFuc1[0] ? "$" : "", $col1 = freezonFuc1[1] ? "$" : "";
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
            if (row[0] < 0 || col[0] < 0) {
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
    };
    fromulaRef.iscelldata = function (txt) {
        var val = txt.split("!"), rangetxt;
        if (val.length > 1) {
            rangetxt = val[1];
        }
        else {
            rangetxt = val[0];
        }
        var reg_cell = /^(([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+))$/g; //增加正则判断单元格为字母+数字的格式：如 A1:B3
        var reg_cellRange = /^(((([a-zA-Z]+)|([$][a-zA-Z]+))(([0-9]+)|([$][0-9]+)))|((([a-zA-Z]+)|([$][a-zA-Z]+))))$/g; //增加正则判断单元格为字母+数字或字母的格式：如 A1:B3，A:A
        if (rangetxt.indexOf(":") == -1) {
            var row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
            var col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));
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
            var row = [], col = [];
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
            if (rangetxt[0].toString().match(reg_cellRange) && rangetxt[1].toString().match(reg_cellRange)) {
                return true;
            }
            else {
                return false;
            }
        }
    };
    fromulaRef.isfreezonFuc = function (txt) {
        var row = txt.replace(/[^0-9]/g, "");
        var col = txt.replace(/[^A-Za-z]/g, "");
        var row$ = txt.substr(txt.indexOf(row) - 1, 1);
        var col$ = txt.substr(txt.indexOf(col) - 1, 1);
        var ret = [false, false];
        if (row$ == "$") {
            ret[0] = true;
        }
        if (col$ == "$") {
            ret[1] = true;
        }
        return ret;
    };
    fromulaRef.operator = '==|!=|<>|<=|>=|=|+|-|>|<|/|*|%|&|^';
    fromulaRef.error = {
        v: "#VALUE!",
        n: "#NAME?",
        na: "#N/A",
        r: "#REF!",
        d: "#DIV/0!",
        nm: "#NUM!",
        nl: "#NULL!",
        sp: "#SPILL!" //数组范围有其它值
    };
    fromulaRef.operatorjson = null;
    return fromulaRef;
}());
function isChinese(temp) {
    var re = /[^\u4e00-\u9fa5]/;
    var reg = /[\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5]/;
    if (reg.test(temp))
        return true;
    if (re.test(temp))
        return false;
    return true;
}
function isJapanese(temp) {
    var re = /[^\u0800-\u4e00]/;
    if (re.test(temp))
        return false;
    return true;
}
function isKoera(chr) {
    if (((chr > 0x3130 && chr < 0x318F) ||
        (chr >= 0xAC00 && chr <= 0xD7A3))) {
        return true;
    }
    return false;
}
function getBinaryContent(path, options) {
    var promise, resolve, reject;
    var callback;
    if (!options) {
        options = {};
    }
    // taken from jQuery
    var createStandardXHR = function () {
        try {
            return new window.XMLHttpRequest();
        }
        catch (e) { }
    };
    var createActiveXHR = function () {
        try {
            return new window.ActiveXObject("Microsoft.XMLHTTP");
        }
        catch (e) { }
    };
    // Create the request object
    var createXHR = (typeof window !== "undefined" && window.ActiveXObject) ?
        /* Microsoft failed to properly
        * implement the XMLHttpRequest in IE7 (can't request local files),
        * so we use the ActiveXObject when it is available
        * Additionally XMLHttpRequest can be disabled in IE7/IE8 so
        * we need a fallback.
        */
        function () {
            return createStandardXHR() || createActiveXHR();
        } :
        // For all other browsers, use the standard XMLHttpRequest object
        createStandardXHR;
    // backward compatible callback
    if (typeof options === "function") {
        callback = options;
        options = {};
    }
    else if (typeof options.callback === 'function') {
        // callback inside options object
        callback = options.callback;
    }
    resolve = function (data) { callback(null, data); };
    reject = function (err) { callback(err, null); };
    try {
        var xhr = createXHR();
        xhr.open('GET', path, true);
        // recent browsers
        if ("responseType" in xhr) {
            xhr.responseType = "arraybuffer";
        }
        // older browser
        if (xhr.overrideMimeType) {
            xhr.overrideMimeType("text/plain; charset=x-user-defined");
        }
        xhr.onreadystatechange = function (event) {
            // use `xhr` and not `this`... thanks IE
            if (xhr.readyState === 4) {
                if (xhr.status === 200 || xhr.status === 0) {
                    try {
                        resolve(function (xhr) {
                            // for xhr.responseText, the 0xFF mask is applied by JSZip
                            return xhr.response || xhr.responseText;
                        }(xhr));
                    }
                    catch (err) {
                        reject(new Error(err));
                    }
                }
                else {
                    reject(new Error("Ajax error for " + path + " : " + this.status + " " + this.statusText));
                }
            }
        };
        if (options.progress) {
            xhr.onprogress = function (e) {
                options.progress({
                    path: path,
                    originalEvent: e,
                    percent: e.loaded / e.total * 100,
                    loaded: e.loaded,
                    total: e.total
                });
            };
        }
        xhr.send();
    }
    catch (e) {
        reject(new Error(e), null);
    }
    // returns a promise or undefined depending on whether a callback was
    // provided
    return promise;
}

var xmloperation = /** @class */ (function () {
    function xmloperation() {
    }
    /**
    * @param tag Search xml tag name , div,title etc.
    * @param file Xml string
    * @return Xml element string
    */
    xmloperation.prototype.getElementsByOneTag = function (tag, file) {
        //<a:[^/>: ]+?>.*?</a:[^/>: ]+?>
        var readTagReg;
        if (tag.indexOf("|") > -1) {
            var tags = tag.split("|"), tagsRegTxt = "";
            for (var i = 0; i < tags.length; i++) {
                var t = tags[i];
                tagsRegTxt += "|<" + t + " [^>]+?[^/]>[\\s\\S]*?</" + t + ">|<" + t + " [^>]+?/>|<" + t + ">[\\s\\S]*?</" + t + ">|<" + t + "/>";
            }
            tagsRegTxt = tagsRegTxt.substr(1, tagsRegTxt.length);
            readTagReg = new RegExp(tagsRegTxt, "g");
        }
        else {
            readTagReg = new RegExp("<" + tag + " [^>]+?[^/]>[\\s\\S]*?</" + tag + ">|<" + tag + " [^>]+?/>|<" + tag + ">[\\s\\S]*?</" + tag + ">|<" + tag + "/>", "g");
        }
        var ret = file.match(readTagReg);
        if (ret == null) {
            return [];
        }
        else {
            return ret;
        }
    };
    return xmloperation;
}());
var ReadXml = /** @class */ (function (_super) {
    __extends(ReadXml, _super);
    function ReadXml(files) {
        var _this = _super.call(this) || this;
        _this.originFile = files;
        return _this;
    }
    /**
    * @param path Search xml tag group , div,title etc.
    * @param fileName One of uploadfileList, uploadfileList is file group, {key:value}
    * @return Xml element calss
    */
    ReadXml.prototype.getElementsByTagName = function (path, fileName) {
        var file = this.getFileByName(fileName);
        var pathArr = path.split("/"), ret;
        for (var key in pathArr) {
            var path_1 = pathArr[key];
            if (ret == undefined) {
                ret = this.getElementsByOneTag(path_1, file);
            }
            else {
                if (ret instanceof Array) {
                    var items = [];
                    for (var key_1 in ret) {
                        var item = ret[key_1];
                        items = items.concat(this.getElementsByOneTag(path_1, item));
                    }
                    ret = items;
                }
                else {
                    ret = this.getElementsByOneTag(path_1, ret);
                }
            }
        }
        var elements = [];
        for (var i = 0; i < ret.length; i++) {
            var ele = new Element(ret[i]);
            elements.push(ele);
        }
        return elements;
    };
    /**
    * @param name One of uploadfileList's name, search for file by this parameter
    * @retrun Select a file from uploadfileList
    */
    ReadXml.prototype.getFileByName = function (name) {
        for (var fileKey in this.originFile) {
            if (fileKey.indexOf(name) > -1) {
                return this.originFile[fileKey];
            }
        }
        return "";
    };
    return ReadXml;
}(xmloperation));
var Element = /** @class */ (function (_super) {
    __extends(Element, _super);
    function Element(str) {
        var _this = _super.call(this) || this;
        _this.elementString = str;
        _this.setValue();
        var readAttrReg = new RegExp('[a-zA-Z0-9_:]*?=".*?"', "g");
        var attrList = _this.container.match(readAttrReg);
        _this.attributeList = {};
        if (attrList != null) {
            for (var key in attrList) {
                var attrFull = attrList[key];
                // let al= attrFull.split("=");
                if (attrFull.length == 0) {
                    continue;
                }
                var attrKey = attrFull.substr(0, attrFull.indexOf('='));
                var attrValue = attrFull.substr(attrFull.indexOf('=') + 1);
                if (attrKey == null || attrValue == null || attrKey.length == 0 || attrValue.length == 0) {
                    continue;
                }
                _this.attributeList[attrKey] = attrValue.substr(1, attrValue.length - 2);
            }
        }
        return _this;
    }
    /**
    * @param name Get attribute by key in element
    * @return Single attribute
    */
    Element.prototype.get = function (name) {
        return this.attributeList[name];
    };
    /**
    * @param tag Get elements by tag in elementString
    * @return Element group
    */
    Element.prototype.getInnerElements = function (tag) {
        var ret = this.getElementsByOneTag(tag, this.elementString);
        var elements = [];
        for (var i = 0; i < ret.length; i++) {
            var ele = new Element(ret[i]);
            elements.push(ele);
        }
        if (elements.length == 0) {
            return null;
        }
        return elements;
    };
    /**
    * @desc get xml dom value and container, <container>value</container>
    */
    Element.prototype.setValue = function () {
        var str = this.elementString;
        if (str.substr(str.length - 2, 2) == "/>") {
            this.value = "";
            this.container = str;
        }
        else {
            var firstTag = this.getFirstTag();
            var firstTagReg = new RegExp("(<" + firstTag + " [^>]+?[^/]>)([\\s\\S]*?)</" + firstTag + ">|(<" + firstTag + ">)([\\s\\S]*?)</" + firstTag + ">", "g");
            var result = firstTagReg.exec(str);
            if (result != null) {
                if (result[1] != null) {
                    this.container = result[1];
                    this.value = result[2];
                }
                else {
                    this.container = result[3];
                    this.value = result[4];
                }
            }
        }
    };
    /**
    * @desc get xml dom first tag, <a><b></b></a>, get a
    */
    Element.prototype.getFirstTag = function () {
        var str = this.elementString;
        var firstTag = str.substr(0, str.indexOf(' '));
        if (firstTag == "" || firstTag.indexOf(">") > -1) {
            firstTag = str.substr(0, str.indexOf('>'));
        }
        firstTag = firstTag.substr(1, firstTag.length);
        return firstTag;
    };
    return Element;
}(xmloperation));
function combineIndexedColor(indexedColorsInner, indexedColors) {
    var ret = {};
    if (indexedColorsInner == null || indexedColorsInner.length == 0) {
        return indexedColors;
    }
    for (var key in indexedColors) {
        var value = indexedColors[key], kn = parseInt(key);
        var inner = indexedColorsInner[kn];
        if (inner == null) {
            ret[key] = value;
        }
        else {
            var rgb = inner.attributeList.rgb;
            ret[key] = rgb;
        }
    }
    return ret;
}
//clrScheme:Element[]
function getColor(color, styles, type) {
    var attrList = color.attributeList;
    var clrScheme = styles["clrScheme"];
    var indexedColorsInner = styles["indexedColors"];
    var mruColorsInner = styles["mruColors"];
    var indexedColorsList = combineIndexedColor(indexedColorsInner, indexedColors);
    var indexed = attrList.indexed, rgb = attrList.rgb, theme = attrList.theme, tint = attrList.tint;
    var bg;
    if (indexed != null) {
        var indexedNum = parseInt(indexed);
        bg = indexedColorsList[indexedNum];
        if (bg != null) {
            bg = bg.substring(bg.length - 6, bg.length);
            bg = "#" + bg;
        }
    }
    else if (rgb != null) {
        rgb = rgb.substring(rgb.length - 6, rgb.length);
        bg = "#" + rgb;
    }
    else if (theme != null) {
        var themeNum = parseInt(theme);
        if (themeNum == 0) {
            themeNum = 1;
        }
        else if (themeNum == 1) {
            themeNum = 0;
        }
        else if (themeNum == 2) {
            themeNum = 3;
        }
        else if (themeNum == 3) {
            themeNum = 2;
        }
        var clrSchemeElement = clrScheme[themeNum];
        if (clrSchemeElement != null) {
            var clrs = clrSchemeElement.getInnerElements("a:sysClr|a:srgbClr");
            if (clrs != null) {
                var clr = clrs[0];
                var clrAttrList = clr.attributeList;
                // console.log(clr.container, );
                if (clr.container.indexOf("sysClr") > -1) {
                    // if(type=="g" && clrAttrList.val=="windowText"){
                    //     bg = null;
                    // }
                    // else if((type=="t" || type=="b") && clrAttrList.val=="window"){
                    //     bg = null;
                    // }                    
                    // else 
                    if (clrAttrList.lastClr != null) {
                        bg = "#" + clrAttrList.lastClr;
                    }
                    else if (clrAttrList.val != null) {
                        bg = "#" + clrAttrList.val;
                    }
                }
                else if (clr.container.indexOf("srgbClr") > -1) {
                    // console.log(clrAttrList.val);
                    bg = "#" + clrAttrList.val;
                }
            }
        }
    }
    if (tint != null) {
        var tintNum = parseFloat(tint);
        if (bg != null) {
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
function getlineStringAttr(frpr, attr) {
    var attrEle = frpr.getInnerElements(attr), value;
    if (attrEle != null && attrEle.length > 0) {
        if (attr == "b" || attr == "i" || attr == "strike") {
            value = "1";
        }
        else if (attr == "u") {
            var v = attrEle[0].attributeList.val;
            if (v == "double") {
                value = "2";
            }
            else if (v == "singleAccounting") {
                value = "3";
            }
            else if (v == "doubleAccounting") {
                value = "4";
            }
            else {
                value = "1";
            }
        }
        else if (attr == "vertAlign") {
            var v = attrEle[0].attributeList.val;
            if (v == "subscript") {
                value = "1";
            }
            else if (v == "superscript") {
                value = "2";
            }
        }
        else {
            value = attrEle[0].attributeList.val;
        }
    }
    return value;
}

var LuckyFileBase = /** @class */ (function () {
    function LuckyFileBase() {
    }
    return LuckyFileBase;
}());
var LuckySheetBase = /** @class */ (function () {
    function LuckySheetBase() {
    }
    return LuckySheetBase;
}());
var LuckyFileInfo = /** @class */ (function () {
    function LuckyFileInfo() {
    }
    return LuckyFileInfo;
}());
var LuckySheetCelldataBase = /** @class */ (function () {
    function LuckySheetCelldataBase() {
    }
    return LuckySheetCelldataBase;
}());
var LuckySheetCelldataValue = /** @class */ (function () {
    function LuckySheetCelldataValue() {
    }
    return LuckySheetCelldataValue;
}());
var LuckySheetCellFormat = /** @class */ (function () {
    function LuckySheetCellFormat() {
    }
    return LuckySheetCellFormat;
}());
var LuckyInlineString = /** @class */ (function () {
    function LuckyInlineString() {
    }
    return LuckyInlineString;
}());
var LuckyConfig = /** @class */ (function () {
    function LuckyConfig() {
    }
    return LuckyConfig;
}());
var LuckySheetborderInfoCellForImp = /** @class */ (function () {
    function LuckySheetborderInfoCellForImp() {
    }
    return LuckySheetborderInfoCellForImp;
}());
var LuckySheetborderInfoCellValue = /** @class */ (function () {
    function LuckySheetborderInfoCellValue() {
    }
    return LuckySheetborderInfoCellValue;
}());
var LuckySheetborderInfoCellValueStyle = /** @class */ (function () {
    function LuckySheetborderInfoCellValueStyle() {
    }
    return LuckySheetborderInfoCellValueStyle;
}());
var LuckySheetConfigMerge = /** @class */ (function () {
    function LuckySheetConfigMerge() {
    }
    return LuckySheetConfigMerge;
}());
var LuckysheetCalcChain = /** @class */ (function () {
    function LuckysheetCalcChain() {
    }
    return LuckysheetCalcChain;
}());
var LuckyImageBase = /** @class */ (function () {
    function LuckyImageBase() {
    }
    return LuckyImageBase;
}());

var LuckySheetCelldata = /** @class */ (function (_super) {
    __extends(LuckySheetCelldata, _super);
    function LuckySheetCelldata(cell, styles, sharedStrings, mergeCells, sheetFile, ReadXml) {
        var _this = 
        //Private
        _super.call(this) || this;
        _this.cell = cell;
        _this.sheetFile = sheetFile;
        _this.styles = styles;
        _this.sharedStrings = sharedStrings;
        _this.readXml = ReadXml;
        _this.mergeCells = mergeCells;
        var attrList = cell.attributeList;
        var r = attrList.r, s = attrList.s, t = attrList.t;
        var range = getcellrange(r);
        _this.r = range.row[0];
        _this.c = range.column[0];
        _this.v = _this.generateValue(s, t);
        return _this;
    }
    /**
    * @param s Style index ,start 1
    * @param t Cell type, Optional value is ST_CellType, it's found at constat.ts
    */
    LuckySheetCelldata.prototype.generateValue = function (s, t) {
        var _this = this;
        var v = this.cell.getInnerElements("v");
        var f = this.cell.getInnerElements("f");
        if (v == null) {
            v = this.cell.getInnerElements("t");
        }
        var cellXfs = this.styles["cellXfs"];
        var cellStyleXfs = this.styles["cellStyleXfs"];
        var cellStyles = this.styles["cellStyles"];
        var fonts = this.styles["fonts"];
        var fills = this.styles["fills"];
        var borders = this.styles["borders"];
        var numfmts = this.styles["numfmts"];
        var clrScheme = this.styles["clrScheme"];
        var sharedStrings = this.sharedStrings;
        var cellValue = new LuckySheetCelldataValue();
        if (f != null) {
            var formula = f[0], attrList = formula.attributeList;
            var t_1 = attrList.t, ref = attrList.ref, si = attrList.si;
            var formulaValue = f[0].value;
            if (t_1 == "shared") {
                this._fomulaRef = ref;
                this._formulaType = t_1;
                this._formulaSi = si;
            }
            // console.log(ref, t, si);
            if (ref != null || (formulaValue != null && formulaValue.length > 0)) {
                formulaValue = escapeCharacter(formulaValue);
                cellValue.f = "=" + formulaValue;
            }
        }
        var familyFont = null;
        var quotePrefix;
        if (s != null) {
            var sNum = parseInt(s);
            var cellXf = cellXfs[sNum];
            var xfId = cellXf.attributeList.xfId;
            var numFmtId = void 0, fontId = void 0, fillId = void 0, borderId = void 0;
            var horizontal = void 0, vertical = void 0, wrapText = void 0, textRotation = void 0, shrinkToFit = void 0, indent = void 0, applyProtection = void 0;
            if (xfId != null) {
                var cellStyleXf = cellStyleXfs[parseInt(xfId)];
                var attrList = cellStyleXf.attributeList;
                var applyNumberFormat_1 = attrList.applyNumberFormat;
                var applyFont_1 = attrList.applyFont;
                var applyFill_1 = attrList.applyFill;
                var applyBorder_1 = attrList.applyBorder;
                var applyAlignment_1 = attrList.applyAlignment;
                // let applyProtection = attrList.applyProtection;
                applyProtection = attrList.applyProtection;
                quotePrefix = attrList.quotePrefix;
                if (applyNumberFormat_1 != "0" && attrList.numFmtId != null) {
                    // if(attrList.numFmtId!="0"){
                    numFmtId = attrList.numFmtId;
                    // }
                }
                if (applyFont_1 != "0" && attrList.fontId != null) {
                    fontId = attrList.fontId;
                }
                if (applyFill_1 != "0" && attrList.fillId != null) {
                    fillId = attrList.fillId;
                }
                if (applyBorder_1 != "0" && attrList.borderId != null) {
                    borderId = attrList.borderId;
                }
                if (applyAlignment_1 != null && applyAlignment_1 != "0") {
                    var alignment = cellStyleXf.getInnerElements("alignment");
                    if (alignment != null) {
                        var attrList_1 = alignment[0].attributeList;
                        if (attrList_1.horizontal != null) {
                            horizontal = attrList_1.horizontal;
                        }
                        if (attrList_1.vertical != null) {
                            vertical = attrList_1.vertical;
                        }
                        if (attrList_1.wrapText != null) {
                            wrapText = attrList_1.wrapText;
                        }
                        if (attrList_1.textRotation != null) {
                            textRotation = attrList_1.textRotation;
                        }
                        if (attrList_1.shrinkToFit != null) {
                            shrinkToFit = attrList_1.shrinkToFit;
                        }
                        if (attrList_1.indent != null) {
                            indent = attrList_1.indent;
                        }
                    }
                }
            }
            var applyNumberFormat = cellXf.attributeList.applyNumberFormat;
            var applyFont = cellXf.attributeList.applyFont;
            var applyFill = cellXf.attributeList.applyFill;
            var applyBorder = cellXf.attributeList.applyBorder;
            var applyAlignment = cellXf.attributeList.applyAlignment;
            if (cellXf.attributeList.applyProtection != null) {
                applyProtection = cellXf.attributeList.applyProtection;
            }
            if (cellXf.attributeList.quotePrefix != null) {
                quotePrefix = cellXf.attributeList.quotePrefix;
            }
            if (applyNumberFormat != "0" && cellXf.attributeList.numFmtId != null) {
                numFmtId = cellXf.attributeList.numFmtId;
            }
            if (applyFont != "0") {
                fontId = cellXf.attributeList.fontId;
            }
            if (applyFill != "0") {
                fillId = cellXf.attributeList.fillId;
            }
            if (applyBorder != "0") {
                borderId = cellXf.attributeList.borderId;
            }
            if (applyAlignment != "0") {
                var alignment = cellXf.getInnerElements("alignment");
                if (alignment != null && alignment.length > 0) {
                    var attrList = alignment[0].attributeList;
                    if (attrList.horizontal != null) {
                        horizontal = attrList.horizontal;
                    }
                    if (attrList.vertical != null) {
                        vertical = attrList.vertical;
                    }
                    if (attrList.wrapText != null) {
                        wrapText = attrList.wrapText;
                    }
                    if (attrList.textRotation != null) {
                        textRotation = attrList.textRotation;
                    }
                    if (attrList.shrinkToFit != null) {
                        shrinkToFit = attrList.shrinkToFit;
                    }
                    if (attrList.indent != null) {
                        indent = attrList.indent;
                    }
                }
            }
            if (numFmtId != undefined) {
                var numf = numfmts[parseInt(numFmtId)];
                var cellFormat = new LuckySheetCellFormat();
                cellFormat.fa = escapeCharacter(numf);
                // console.log(numf, numFmtId, this.v);
                cellFormat.t = t;
                cellValue.ct = cellFormat;
            }
            if (fillId != undefined) {
                var fillIdNum = parseInt(fillId);
                var fill = fills[fillIdNum];
                // console.log(cellValue.v);
                var bg = this.getBackgroundByFill(fill, clrScheme);
                if (bg != null) {
                    cellValue.bg = bg;
                }
            }
            if (fontId != undefined) {
                var fontIdNum = parseInt(fontId);
                var font = fonts[fontIdNum];
                if (font != null) {
                    var sz = font.getInnerElements("sz"); //font size
                    var colors = font.getInnerElements("color"); //font color
                    var family = font.getInnerElements("name"); //font family
                    var familyOverrides = font.getInnerElements("family"); //font family will be overrided by name
                    var charset = font.getInnerElements("charset"); //font charset
                    var bolds = font.getInnerElements("b"); //font bold
                    var italics = font.getInnerElements("i"); //font italic
                    var strikes = font.getInnerElements("strike"); //font italic
                    var underlines = font.getInnerElements("u"); //font italic
                    if (sz != null && sz.length > 0) {
                        var fs = sz[0].attributeList.val;
                        if (fs != null) {
                            cellValue.fs = parseInt(fs);
                        }
                    }
                    if (colors != null && colors.length > 0) {
                        var color = colors[0];
                        var fc = getColor(color, this.styles);
                        if (fc != null) {
                            cellValue.fc = fc;
                        }
                    }
                    if (familyOverrides != null && familyOverrides.length > 0) {
                        var val = familyOverrides[0].attributeList.val;
                        if (val != null) {
                            familyFont = fontFamilys[val];
                        }
                    }
                    if (family != null && family.length > 0) {
                        var val = family[0].attributeList.val;
                        if (val != null) {
                            cellValue.ff = val;
                        }
                    }
                    if (bolds != null && bolds.length > 0) {
                        var bold = bolds[0].attributeList.val;
                        if (bold == "0") {
                            cellValue.bl = 0;
                        }
                        else {
                            cellValue.bl = 1;
                        }
                    }
                    if (italics != null && italics.length > 0) {
                        var italic = italics[0].attributeList.val;
                        if (italic == "0") {
                            cellValue.it = 0;
                        }
                        else {
                            cellValue.it = 1;
                        }
                    }
                    if (strikes != null && strikes.length > 0) {
                        var strike = strikes[0].attributeList.val;
                        if (strike == "0") {
                            cellValue.cl = 0;
                        }
                        else {
                            cellValue.cl = 1;
                        }
                    }
                    if (underlines != null && underlines.length > 0) {
                        var underline = underlines[0].attributeList.val;
                        if (underline == "single") {
                            cellValue.un = 1;
                        }
                        else if (underline == "double") {
                            cellValue.un = 2;
                        }
                        else if (underline == "singleAccounting") {
                            cellValue.un = 3;
                        }
                        else if (underline == "doubleAccounting") {
                            cellValue.un = 4;
                        }
                        else {
                            cellValue.un = 0;
                        }
                    }
                }
            }
            // vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
            // ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
            // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
            // tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment
            if (horizontal != undefined) { //Horizontal alignment
                if (horizontal == "center") {
                    cellValue.ht = 0;
                }
                else if (horizontal == "centerContinuous") {
                    cellValue.ht = 0; //luckysheet unsupport
                }
                else if (horizontal == "left") {
                    cellValue.ht = 1;
                }
                else if (horizontal == "right") {
                    cellValue.ht = 2;
                }
                else if (horizontal == "distributed") {
                    cellValue.ht = 0; //luckysheet unsupport
                }
                else if (horizontal == "fill") {
                    cellValue.ht = 1; //luckysheet unsupport
                }
                else if (horizontal == "general") {
                    cellValue.ht = 1; //luckysheet unsupport
                }
                else if (horizontal == "justify") {
                    cellValue.ht = 0; //luckysheet unsupport
                }
                else {
                    cellValue.ht = 1;
                }
            }
            if (vertical != undefined) { //Vertical alignment
                if (vertical == "bottom") {
                    cellValue.vt = 2;
                }
                else if (vertical == "center") {
                    cellValue.vt = 0;
                }
                else if (vertical == "distributed") {
                    cellValue.vt = 0; //luckysheet unsupport
                }
                else if (vertical == "justify") {
                    cellValue.vt = 0; //luckysheet unsupport
                }
                else if (vertical == "top") {
                    cellValue.vt = 1;
                }
                else {
                    cellValue.vt = 1;
                }
            }
            if (wrapText != undefined) {
                if (wrapText == "1") {
                    cellValue.tb = 2;
                }
                else {
                    cellValue.tb = 1;
                }
            }
            else {
                cellValue.tb = 1;
            }
            if (textRotation != undefined) {
                // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
                if (textRotation == "255") {
                    cellValue.tr = 3;
                }
                // else if(textRotation=="45"){
                //     cellValue.tr = 1;
                // }
                // else if(textRotation=="90"){
                //     cellValue.tr = 4;
                // }
                // else if(textRotation=="135"){
                //     cellValue.tr = 2;
                // }
                // else if(textRotation=="180"){
                //     cellValue.tr = 5;
                // }
                else {
                    cellValue.tr = 0;
                    cellValue.rt = parseInt(textRotation);
                }
            }
            if (borderId != undefined) {
                var borderIdNum = parseInt(borderId);
                var border = borders[borderIdNum];
                // this._borderId = borderIdNum;
                var borderObject = new LuckySheetborderInfoCellForImp();
                borderObject.rangeType = "cell";
                // borderObject.cells = [];
                var borderCellValue = new LuckySheetborderInfoCellValue();
                borderCellValue.row_index = this.r;
                borderCellValue.col_index = this.c;
                var lefts = border.getInnerElements("left");
                var rights = border.getInnerElements("right");
                var tops = border.getInnerElements("top");
                var bottoms = border.getInnerElements("bottom");
                var diagonals = border.getInnerElements("diagonal");
                var starts = border.getInnerElements("start");
                var ends = border.getInnerElements("end");
                var left = this.getBorderInfo(lefts);
                var right = this.getBorderInfo(rights);
                var top_1 = this.getBorderInfo(tops);
                var bottom = this.getBorderInfo(bottoms);
                var diagonal = this.getBorderInfo(diagonals);
                var start = this.getBorderInfo(starts);
                var end = this.getBorderInfo(ends);
                var isAdd = false;
                if (start != null && start.color != null) {
                    borderCellValue.l = start;
                    isAdd = true;
                }
                if (end != null && end.color != null) {
                    borderCellValue.r = end;
                    isAdd = true;
                }
                if (left != null && left.color != null) {
                    borderCellValue.l = left;
                    isAdd = true;
                }
                if (right != null && right.color != null) {
                    borderCellValue.r = right;
                    isAdd = true;
                }
                if (top_1 != null && top_1.color != null) {
                    borderCellValue.t = top_1;
                    isAdd = true;
                }
                if (bottom != null && bottom.color != null) {
                    borderCellValue.b = bottom;
                    isAdd = true;
                }
                if (isAdd) {
                    borderObject.value = borderCellValue;
                    // this.config._borderInfo[borderId] = borderObject;
                    this._borderObject = borderObject;
                }
            }
        }
        else {
            cellValue.tb = 1;
        }
        if (v != null) {
            var value = v[0].value;
            if (/&#\d+;/.test(value)) {
                value = this.htmlDecode(value);
            }
            if (t == ST_CellType["SharedString"]) {
                var siIndex = parseInt(v[0].value);
                var sharedSI = sharedStrings[siIndex];
                var rFlag = sharedSI.getInnerElements("r");
                if (rFlag == null) {
                    var tFlag = sharedSI.getInnerElements("t");
                    if (tFlag != null) {
                        var text_1 = "";
                        tFlag.forEach(function (t) {
                            text_1 += t.value;
                        });
                        text_1 = escapeCharacter(text_1);
                        //isContainMultiType(text) &&
                        if (familyFont == "Roman" && text_1.length > 0) {
                            var textArray = text_1.split("");
                            var preWordType = null, wordText = "", preWholef = null;
                            var wholef = "Times New Roman";
                            if (cellValue.ff != null) {
                                wholef = cellValue.ff;
                            }
                            var cellFormat = cellValue.ct;
                            if (cellFormat == null) {
                                cellFormat = new LuckySheetCellFormat();
                            }
                            if (cellFormat.s == null) {
                                cellFormat.s = [];
                            }
                            for (var i = 0; i < textArray.length; i++) {
                                var w = textArray[i];
                                var type = null, ff = wholef;
                                if (isChinese(w)) {
                                    type = "c";
                                    ff = "宋体";
                                }
                                else if (isJapanese(w)) {
                                    type = "j";
                                    ff = "Yu Gothic";
                                }
                                else if (isKoera(w)) {
                                    type = "k";
                                    ff = "Malgun Gothic";
                                }
                                else {
                                    type = "e";
                                }
                                if ((type != preWordType && preWordType != null) || i == textArray.length - 1) {
                                    var InlineString = new LuckyInlineString();
                                    InlineString.ff = preWholef;
                                    if (cellValue.fc != null) {
                                        InlineString.fc = cellValue.fc;
                                    }
                                    if (cellValue.fs != null) {
                                        InlineString.fs = cellValue.fs;
                                    }
                                    if (cellValue.cl != null) {
                                        InlineString.cl = cellValue.cl;
                                    }
                                    if (cellValue.un != null) {
                                        InlineString.un = cellValue.un;
                                    }
                                    if (cellValue.bl != null) {
                                        InlineString.bl = cellValue.bl;
                                    }
                                    if (cellValue.it != null) {
                                        InlineString.it = cellValue.it;
                                    }
                                    if (i == textArray.length - 1) {
                                        if (type == preWordType) {
                                            InlineString.ff = ff;
                                            InlineString.v = wordText + w;
                                        }
                                        else {
                                            InlineString.ff = preWholef;
                                            InlineString.v = wordText;
                                            cellFormat.s.push(InlineString);
                                            var InlineStringLast = new LuckyInlineString();
                                            InlineStringLast.ff = ff;
                                            InlineStringLast.v = w;
                                            if (cellValue.fc != null) {
                                                InlineStringLast.fc = cellValue.fc;
                                            }
                                            if (cellValue.fs != null) {
                                                InlineStringLast.fs = cellValue.fs;
                                            }
                                            if (cellValue.cl != null) {
                                                InlineStringLast.cl = cellValue.cl;
                                            }
                                            if (cellValue.un != null) {
                                                InlineStringLast.un = cellValue.un;
                                            }
                                            if (cellValue.bl != null) {
                                                InlineStringLast.bl = cellValue.bl;
                                            }
                                            if (cellValue.it != null) {
                                                InlineStringLast.it = cellValue.it;
                                            }
                                            cellFormat.s.push(InlineStringLast);
                                            break;
                                        }
                                    }
                                    else {
                                        InlineString.v = wordText;
                                    }
                                    cellFormat.s.push(InlineString);
                                    wordText = w;
                                }
                                else {
                                    wordText += w;
                                }
                                preWordType = type;
                                preWholef = ff;
                            }
                            cellFormat.t = "inlineStr";
                            // cellFormat.s = [InlineString];
                            cellValue.ct = cellFormat;
                            // console.log(cellValue);
                        }
                        else {
                            text_1 = this.replaceSpecialWrap(text_1);
                            if (text_1.indexOf("\r\n") > -1 || text_1.indexOf("\n") > -1) {
                                var InlineString = new LuckyInlineString();
                                InlineString.v = text_1;
                                var cellFormat = cellValue.ct;
                                if (cellFormat == null) {
                                    cellFormat = new LuckySheetCellFormat();
                                }
                                if (cellValue.ff != null) {
                                    InlineString.ff = cellValue.ff;
                                }
                                if (cellValue.fc != null) {
                                    InlineString.fc = cellValue.fc;
                                }
                                if (cellValue.fs != null) {
                                    InlineString.fs = cellValue.fs;
                                }
                                if (cellValue.cl != null) {
                                    InlineString.cl = cellValue.cl;
                                }
                                if (cellValue.un != null) {
                                    InlineString.un = cellValue.un;
                                }
                                if (cellValue.bl != null) {
                                    InlineString.bl = cellValue.bl;
                                }
                                if (cellValue.it != null) {
                                    InlineString.it = cellValue.it;
                                }
                                cellFormat.t = "inlineStr";
                                cellFormat.s = [InlineString];
                                cellValue.ct = cellFormat;
                            }
                            else {
                                cellValue.v = text_1;
                                quotePrefix = "1";
                            }
                        }
                    }
                }
                else {
                    var styles_1 = [];
                    rFlag.forEach(function (r) {
                        var tFlag = r.getInnerElements("t");
                        var rPr = r.getInnerElements("rPr");
                        var InlineString = new LuckyInlineString();
                        if (tFlag != null && tFlag.length > 0) {
                            var text = tFlag[0].value;
                            text = _this.replaceSpecialWrap(text);
                            text = escapeCharacter(text);
                            InlineString.v = text;
                        }
                        if (rPr != null && rPr.length > 0) {
                            var frpr = rPr[0];
                            var sz = getlineStringAttr(frpr, "sz"), rFont = getlineStringAttr(frpr, "rFont"), family = getlineStringAttr(frpr, "family"), charset = getlineStringAttr(frpr, "charset"), scheme = getlineStringAttr(frpr, "scheme"), b = getlineStringAttr(frpr, "b"), i = getlineStringAttr(frpr, "i"), u = getlineStringAttr(frpr, "u"), strike = getlineStringAttr(frpr, "strike"), vertAlign = getlineStringAttr(frpr, "vertAlign"), color = void 0;
                            var cEle = frpr.getInnerElements("color");
                            if (cEle != null && cEle.length > 0) {
                                color = getColor(cEle[0], _this.styles);
                            }
                            var ff = void 0;
                            // if(family!=null){
                            //     ff = fontFamilys[family];
                            // }
                            if (rFont != null) {
                                ff = rFont;
                            }
                            if (ff != null) {
                                InlineString.ff = ff;
                            }
                            else if (cellValue.ff != null) {
                                InlineString.ff = cellValue.ff;
                            }
                            if (color != null) {
                                InlineString.fc = color;
                            }
                            else if (cellValue.fc != null) {
                                InlineString.fc = cellValue.fc;
                            }
                            if (sz != null) {
                                InlineString.fs = parseInt(sz);
                            }
                            else if (cellValue.fs != null) {
                                InlineString.fs = cellValue.fs;
                            }
                            if (strike != null) {
                                InlineString.cl = parseInt(strike);
                            }
                            else if (cellValue.cl != null) {
                                InlineString.cl = cellValue.cl;
                            }
                            if (u != null) {
                                InlineString.un = parseInt(u);
                            }
                            else if (cellValue.un != null) {
                                InlineString.un = cellValue.un;
                            }
                            if (b != null) {
                                InlineString.bl = parseInt(b);
                            }
                            else if (cellValue.bl != null) {
                                InlineString.bl = cellValue.bl;
                            }
                            if (i != null) {
                                InlineString.it = parseInt(i);
                            }
                            else if (cellValue.it != null) {
                                InlineString.it = cellValue.it;
                            }
                            if (vertAlign != null) {
                                InlineString.va = parseInt(vertAlign);
                            }
                            // ff:string | undefined //font family
                            // fc:string | undefined//font color
                            // fs:number | undefined//font size
                            // cl:number | undefined//strike
                            // un:number | undefined//underline
                            // bl:number | undefined//blod
                            // it:number | undefined//italic
                            // v:string | undefined
                        }
                        else {
                            if (InlineString.ff == null && cellValue.ff != null) {
                                InlineString.ff = cellValue.ff;
                            }
                            if (InlineString.fc == null && cellValue.fc != null) {
                                InlineString.fc = cellValue.fc;
                            }
                            if (InlineString.fs == null && cellValue.fs != null) {
                                InlineString.fs = cellValue.fs;
                            }
                            if (InlineString.cl == null && cellValue.cl != null) {
                                InlineString.cl = cellValue.cl;
                            }
                            if (InlineString.un == null && cellValue.un != null) {
                                InlineString.un = cellValue.un;
                            }
                            if (InlineString.bl == null && cellValue.bl != null) {
                                InlineString.bl = cellValue.bl;
                            }
                            if (InlineString.it == null && cellValue.it != null) {
                                InlineString.it = cellValue.it;
                            }
                        }
                        styles_1.push(InlineString);
                    });
                    var cellFormat = cellValue.ct;
                    if (cellFormat == null) {
                        cellFormat = new LuckySheetCellFormat();
                    }
                    cellFormat.t = "inlineStr";
                    cellFormat.s = styles_1;
                    cellValue.ct = cellFormat;
                }
            }
            // else if(t==ST_CellType["InlineString"] && v!=null){
            // }
            else {
                value = escapeCharacter(value);
                cellValue.v = value;
            }
        }
        if (quotePrefix != null) {
            cellValue.qp = parseInt(quotePrefix);
        }
        return cellValue;
    };
    LuckySheetCelldata.prototype.replaceSpecialWrap = function (text) {
        text = text.replace(/_x000D_/g, "").replace(/&#13;&#10;/g, "\r\n").replace(/&#13;/g, "\r").replace(/&#10;/g, "\n");
        return text;
    };
    LuckySheetCelldata.prototype.getBackgroundByFill = function (fill, clrScheme) {
        var patternFills = fill.getInnerElements("patternFill");
        if (patternFills != null) {
            var patternFill = patternFills[0];
            var fgColors = patternFill.getInnerElements("fgColor");
            var bgColors = patternFill.getInnerElements("bgColor");
            var fg = void 0, bg = void 0;
            if (fgColors != null) {
                var fgColor = fgColors[0];
                fg = getColor(fgColor, this.styles);
            }
            if (bgColors != null) {
                var bgColor = bgColors[0];
                bg = getColor(bgColor, this.styles);
            }
            // console.log(fgColors,bgColors,clrScheme);
            if (fg != null) {
                return fg;
            }
            else if (bg != null) {
                return bg;
            }
        }
        else {
            var gradientfills = fill.getInnerElements("gradientFill");
            if (gradientfills != null) {
                //graient color fill handler
                return null;
            }
        }
    };
    LuckySheetCelldata.prototype.getBorderInfo = function (borders) {
        if (borders == null) {
            return null;
        }
        var border = borders[0], attrList = border.attributeList;
        var clrScheme = this.styles["clrScheme"];
        var style = attrList.style;
        if (style == null || style == "none") {
            return null;
        }
        var colors = border.getInnerElements("color");
        var colorRet = "#000000";
        if (colors != null) {
            var color = colors[0];
            colorRet = getColor(color, this.styles);
            if (colorRet == null) {
                colorRet = "#000000";
            }
        }
        var ret = new LuckySheetborderInfoCellValueStyle();
        ret.style = borderTypes[style];
        ret.color = colorRet;
        return ret;
    };
    LuckySheetCelldata.prototype.htmlDecode = function (str) {
        return str.replace(/&#(x)?([^&]{1,5});?/g, function ($, $1, $2) {
            return String.fromCharCode(parseInt($2, $1 ? 16 : 10));
        });
    };
    return LuckySheetCelldata;
}(LuckySheetCelldataBase));

var LuckySheet = /** @class */ (function (_super) {
    __extends(LuckySheet, _super);
    function LuckySheet(sheetName, sheetId, sheetOrder, isInitialCell, allFileOption) {
        if (isInitialCell === void 0) { isInitialCell = false; }
        var _this = 
        //Private
        _super.call(this) || this;
        _this.isInitialCell = isInitialCell;
        _this.readXml = allFileOption.readXml;
        _this.sheetFile = allFileOption.sheetFile;
        _this.styles = allFileOption.styles;
        _this.sharedStrings = allFileOption.sharedStrings;
        _this.calcChainEles = allFileOption.calcChain;
        _this.sheetList = allFileOption.sheetList;
        _this.imageList = allFileOption.imageList;
        //Output
        _this.name = sheetName;
        _this.id = sheetId;
        _this.order = sheetOrder.toString();
        _this.config = new LuckyConfig();
        _this.celldata = [];
        _this.mergeCells = _this.readXml.getElementsByTagName("mergeCells/mergeCell", _this.sheetFile);
        var clrScheme = _this.styles["clrScheme"];
        var sheetView = _this.readXml.getElementsByTagName("sheetViews/sheetView", _this.sheetFile);
        var showGridLines = "1", tabSelected = "0", zoomScale = "100", activeCell = "A1";
        if (sheetView.length > 0) {
            var attrList = sheetView[0].attributeList;
            showGridLines = getXmlAttibute(attrList, "showGridLines", "1");
            tabSelected = getXmlAttibute(attrList, "tabSelected", "0");
            zoomScale = getXmlAttibute(attrList, "zoomScale", "100");
            // let colorId = getXmlAttibute(attrList, "colorId", "0");
            var selections = sheetView[0].getInnerElements("selection");
            if (selections != null && selections.length > 0) {
                activeCell = getXmlAttibute(selections[0].attributeList, "activeCell", "A1");
                var range = getcellrange(activeCell, _this.sheetList, sheetId);
                _this.luckysheet_select_save = [];
                _this.luckysheet_select_save.push(range);
            }
        }
        _this.showGridLines = showGridLines;
        _this.status = tabSelected;
        _this.zoomRatio = parseInt(zoomScale) / 100;
        var tabColors = _this.readXml.getElementsByTagName("sheetPr/tabColor", _this.sheetFile);
        if (tabColors != null && tabColors.length > 0) {
            var tabColor = tabColors[0], attrList = tabColor.attributeList;
            // if(attrList.rgb!=null){
            var tc = getColor(tabColor, _this.styles);
            _this.color = tc;
            // }
        }
        var sheetFormatPr = _this.readXml.getElementsByTagName("sheetFormatPr", _this.sheetFile);
        var defaultColWidth, defaultRowHeight;
        if (sheetFormatPr.length > 0) {
            var attrList = sheetFormatPr[0].attributeList;
            defaultColWidth = getXmlAttibute(attrList, "defaultColWidth", "9.21");
            defaultRowHeight = getXmlAttibute(attrList, "defaultRowHeight", "19");
        }
        _this.defaultColWidth = getColumnWidthPixel(parseFloat(defaultColWidth));
        _this.defaultRowHeight = getRowHeightPixel(parseFloat(defaultRowHeight));
        _this.generateConfigColumnLenAndHidden();
        var cellOtherInfo = _this.generateConfigRowLenAndHiddenAddCell();
        if (_this.formulaRefList != null) {
            for (var key in _this.formulaRefList) {
                var funclist = _this.formulaRefList[key];
                var mainFunc = funclist["mainRef"], mainCellValue = mainFunc.cellValue;
                var formulaTxt = mainFunc.fv;
                var mainR = mainCellValue.r, mainC = mainCellValue.c;
                // let refRange = getcellrange(ref);
                for (var name_1 in funclist) {
                    if (name_1 == "mainRef") {
                        continue;
                    }
                    var funcValue = funclist[name_1], cellValue = funcValue.cellValue;
                    if (cellValue == null) {
                        continue;
                    }
                    var r = cellValue.r, c = cellValue.c;
                    var func = formulaTxt;
                    var offsetRow = r - mainR, offsetCol = c - mainC;
                    if (offsetRow > 0) {
                        func = "=" + fromulaRef.functionCopy(func, "down", offsetRow);
                    }
                    else if (offsetRow < 0) {
                        func = "=" + fromulaRef.functionCopy(func, "up", Math.abs(offsetRow));
                    }
                    if (offsetCol > 0) {
                        func = "=" + fromulaRef.functionCopy(func, "right", offsetCol);
                    }
                    else if (offsetCol < 0) {
                        func = "=" + fromulaRef.functionCopy(func, "left", Math.abs(offsetCol));
                    }
                    // console.log(offsetRow, offsetCol, func);
                    cellValue.v.f = func;
                }
            }
        }
        if (_this.calcChain == null) {
            _this.calcChain = [];
        }
        var formulaListExist = {};
        for (var c = 0; c < _this.calcChainEles.length; c++) {
            var calcChainEle = _this.calcChainEles[c], attrList = calcChainEle.attributeList;
            if (attrList.i != sheetId) {
                continue;
            }
            var r = attrList.r, i = attrList.i, l = attrList.l, s = attrList.s, a = attrList.a, t = attrList.t;
            var range = getcellrange(r);
            var chain = new LuckysheetCalcChain();
            chain.r = range.row[0];
            chain.c = range.column[0];
            chain.id = _this.id;
            _this.calcChain.push(chain);
            formulaListExist["r" + r + "c" + c] = null;
        }
        //There may be formulas that do not appear in calcChain
        for (var key in cellOtherInfo.formulaList) {
            if (!(key in formulaListExist)) {
                var formulaListItem = cellOtherInfo.formulaList[key];
                var chain = new LuckysheetCalcChain();
                chain.r = formulaListItem.r;
                chain.c = formulaListItem.c;
                chain.id = _this.id;
                _this.calcChain.push(chain);
            }
        }
        if (_this.mergeCells != null) {
            for (var i = 0; i < _this.mergeCells.length; i++) {
                var merge = _this.mergeCells[i], attrList = merge.attributeList;
                var ref = attrList.ref;
                if (ref == null) {
                    continue;
                }
                var range = getcellrange(ref, _this.sheetList, sheetId);
                var mergeValue = new LuckySheetConfigMerge();
                mergeValue.r = range.row[0];
                mergeValue.c = range.column[0];
                mergeValue.rs = range.row[1] - range.row[0] + 1;
                mergeValue.cs = range.column[1] - range.column[0] + 1;
                if (_this.config.merge == null) {
                    _this.config.merge = {};
                }
                _this.config.merge[range.row[0] + "_" + range.column[0]] = mergeValue;
            }
        }
        var drawingFile = allFileOption.drawingFile, drawingRelsFile = allFileOption.drawingRelsFile;
        if (drawingFile != null && drawingRelsFile != null) {
            var twoCellAnchors = _this.readXml.getElementsByTagName("xdr:twoCellAnchor", drawingFile);
            if (twoCellAnchors != null && twoCellAnchors.length > 0) {
                for (var i = 0; i < twoCellAnchors.length; i++) {
                    var twoCellAnchor = twoCellAnchors[i];
                    var editAs = getXmlAttibute(twoCellAnchor.attributeList, "editAs", "twoCell");
                    var xdrFroms = twoCellAnchor.getInnerElements("xdr:from"), xdrTos = twoCellAnchor.getInnerElements("xdr:to");
                    var xdr_blipfills = twoCellAnchor.getInnerElements("a:blip");
                    if (xdrFroms != null && xdr_blipfills != null && xdrFroms.length > 0 && xdr_blipfills.length > 0) {
                        var xdrFrom = xdrFroms[0], xdrTo = xdrTos[0], xdr_blipfill = xdr_blipfills[0];
                        var rembed = getXmlAttibute(xdr_blipfill.attributeList, "r:embed", null);
                        var imageObject = _this.getBase64ByRid(rembed, drawingRelsFile);
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
                        var x_n = 0, y_n = 0;
                        var cx_n = 0, cy_n = 0;
                        imageObject.fromCol = _this.getXdrValue(xdrFrom.getInnerElements("xdr:col"));
                        imageObject.fromColOff = getPxByEMUs(_this.getXdrValue(xdrFrom.getInnerElements("xdr:colOff")));
                        imageObject.fromRow = _this.getXdrValue(xdrFrom.getInnerElements("xdr:row"));
                        imageObject.fromRowOff = getPxByEMUs(_this.getXdrValue(xdrFrom.getInnerElements("xdr:rowOff")));
                        imageObject.toCol = _this.getXdrValue(xdrTo.getInnerElements("xdr:col"));
                        imageObject.toColOff = getPxByEMUs(_this.getXdrValue(xdrTo.getInnerElements("xdr:colOff")));
                        imageObject.toRow = _this.getXdrValue(xdrTo.getInnerElements("xdr:row"));
                        imageObject.toRowOff = getPxByEMUs(_this.getXdrValue(xdrTo.getInnerElements("xdr:rowOff")));
                        imageObject.originWidth = cx_n;
                        imageObject.originHeight = cy_n;
                        if (editAs == "absolute") {
                            imageObject.type = "3";
                        }
                        else if (editAs == "oneCell") {
                            imageObject.type = "2";
                        }
                        else {
                            imageObject.type = "1";
                        }
                        imageObject.isFixedPos = false;
                        imageObject.fixedLeft = 0;
                        imageObject.fixedTop = 0;
                        var imageBorder = {
                            color: "#000",
                            radius: 0,
                            style: "solid",
                            width: 0
                        };
                        imageObject.border = imageBorder;
                        var imageCrop = {
                            height: cy_n,
                            offsetLeft: 0,
                            offsetTop: 0,
                            width: cx_n
                        };
                        imageObject.crop = imageCrop;
                        var imageDefault = {
                            height: cy_n,
                            left: x_n,
                            top: y_n,
                            width: cx_n
                        };
                        imageObject.default = imageDefault;
                        if (_this.images == null) {
                            _this.images = {};
                        }
                        _this.images[generateRandomIndex("image")] = imageObject;
                        //     }
                        // }
                    }
                }
            }
        }
        return _this;
    }
    LuckySheet.prototype.getXdrValue = function (ele) {
        if (ele == null || ele.length == 0) {
            return null;
        }
        return parseInt(ele[0].value);
    };
    LuckySheet.prototype.getBase64ByRid = function (rid, drawingRelsFile) {
        var Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);
        if (Relationships != null && Relationships.length > 0) {
            for (var i = 0; i < Relationships.length; i++) {
                var Relationship = Relationships[i];
                var attrList = Relationship.attributeList;
                var Id = getXmlAttibute(attrList, "Id", null);
                var src = getXmlAttibute(attrList, "Target", null);
                if (Id == rid) {
                    src = src.replace(/\.\.\//g, "");
                    src = "xl/" + src;
                    var imgage = this.imageList.getImageByName(src);
                    return imgage;
                }
            }
        }
        return null;
    };
    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    LuckySheet.prototype.generateConfigColumnLenAndHidden = function () {
        var cols = this.readXml.getElementsByTagName("cols/col", this.sheetFile);
        for (var i = 0; i < cols.length; i++) {
            var col = cols[i], attrList = col.attributeList;
            var min = getXmlAttibute(attrList, "min", null);
            var max = getXmlAttibute(attrList, "max", null);
            var width = getXmlAttibute(attrList, "width", null);
            var hidden = getXmlAttibute(attrList, "hidden", null);
            var customWidth = getXmlAttibute(attrList, "customWidth", null);
            if (min == null || max == null) {
                continue;
            }
            var minNum = parseInt(min) - 1, maxNum = parseInt(max) - 1, widthNum = parseFloat(width);
            for (var m = minNum; m <= maxNum; m++) {
                if (width != null) {
                    if (this.config.columnlen == null) {
                        this.config.columnlen = {};
                    }
                    this.config.columnlen[m] = getColumnWidthPixel(widthNum);
                }
                if (hidden == "1") {
                    if (this.config.colhidden == null) {
                        this.config.colhidden = {};
                    }
                    this.config.colhidden[m] = 0;
                    if (this.config.columnlen) {
                        delete this.config.columnlen[m];
                    }
                }
                if (customWidth != null) {
                    if (this.config.customWidth == null) {
                        this.config.customWidth = {};
                    }
                    this.config.customWidth[m] = 1;
                }
            }
        }
    };
    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    LuckySheet.prototype.generateConfigRowLenAndHiddenAddCell = function () {
        var rows = this.readXml.getElementsByTagName("sheetData/row", this.sheetFile);
        var cellOtherInfo = {};
        var formulaList = {};
        cellOtherInfo.formulaList = formulaList;
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i], attrList = row.attributeList;
            var rowNo = getXmlAttibute(attrList, "r", null);
            var height = getXmlAttibute(attrList, "ht", null);
            var hidden = getXmlAttibute(attrList, "hidden", null);
            var customHeight = getXmlAttibute(attrList, "customHeight", null);
            if (rowNo == null) {
                continue;
            }
            var rowNoNum = parseInt(rowNo) - 1;
            if (height != null) {
                var heightNum = parseFloat(height);
                if (this.config.rowlen == null) {
                    this.config.rowlen = {};
                }
                this.config.rowlen[rowNoNum] = getRowHeightPixel(heightNum);
            }
            if (hidden == "1") {
                if (this.config.rowhidden == null) {
                    this.config.rowhidden = {};
                }
                this.config.rowhidden[rowNoNum] = 0;
                if (this.config.rowlen) {
                    delete this.config.rowlen[rowNoNum];
                }
            }
            if (customHeight != null) {
                if (this.config.customHeight == null) {
                    this.config.customHeight = {};
                }
                this.config.customHeight[rowNoNum] = 1;
            }
            if (this.isInitialCell) {
                var cells = row.getInnerElements("c");
                for (var key in cells) {
                    var cell = cells[key];
                    var cellValue = new LuckySheetCelldata(cell, this.styles, this.sharedStrings, this.mergeCells, this.sheetFile, this.readXml);
                    if (cellValue._borderObject != null) {
                        if (this.config.borderInfo == null) {
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
                    if (cellValue._formulaType == "shared") {
                        if (this.formulaRefList == null) {
                            this.formulaRefList = {};
                        }
                        if (this.formulaRefList[cellValue._formulaSi] == null) {
                            this.formulaRefList[cellValue._formulaSi] = {};
                        }
                        var fv = void 0;
                        if (cellValue.v != null) {
                            fv = cellValue.v.f;
                        }
                        var refValue = {
                            t: cellValue._formulaType,
                            ref: cellValue._fomulaRef,
                            si: cellValue._formulaSi,
                            fv: fv,
                            cellValue: cellValue
                        };
                        if (cellValue._fomulaRef != null) {
                            this.formulaRefList[cellValue._formulaSi]["mainRef"] = refValue;
                        }
                        else {
                            this.formulaRefList[cellValue._formulaSi][cellValue.r + "_" + cellValue.c] = refValue;
                        }
                        // console.log(refValue, this.formulaRefList);
                    }
                    //There may be formulas that do not appear in calcChain
                    if (cellValue.v != null && cellValue.v.f != null) {
                        var formulaCell = {
                            r: cellValue.r,
                            c: cellValue.c
                        };
                        cellOtherInfo.formulaList["r" + cellValue.r + "c" + cellValue.c] = formulaCell;
                    }
                    this.celldata.push(cellValue);
                }
            }
        }
        return cellOtherInfo;
    };
    return LuckySheet;
}(LuckySheetBase));

var UDOC = {};
UDOC.G = {
    concat: function (p, r) {
        for (var i = 0; i < r.cmds.length; i++)
            p.cmds.push(r.cmds[i]);
        for (var i = 0; i < r.crds.length; i++)
            p.crds.push(r.crds[i]);
    },
    getBB: function (ps) {
        var x0 = 1e99, y0 = 1e99, x1 = -x0, y1 = -y0;
        for (var i = 0; i < ps.length; i += 2) {
            var x = ps[i], y = ps[i + 1];
            if (x < x0)
                x0 = x;
            else if (x > x1)
                x1 = x;
            if (y < y0)
                y0 = y;
            else if (y > y1)
                y1 = y;
        }
        return [x0, y0, x1, y1];
    },
    rectToPath: function (r) { return { cmds: ["M", "L", "L", "L", "Z"], crds: [r[0], r[1], r[2], r[1], r[2], r[3], r[0], r[3]] }; },
    // a inside b
    insideBox: function (a, b) { return b[0] <= a[0] && b[1] <= a[1] && a[2] <= b[2] && a[3] <= b[3]; },
    isBox: function (p, bb) {
        var sameCrd8 = function (pcrd, crds) {
            for (var o = 0; o < 8; o += 2) {
                var eq = true;
                for (var j = 0; j < 8; j++)
                    if (Math.abs(crds[j] - pcrd[(j + o) & 7]) >= 2) {
                        eq = false;
                        break;
                    }
                if (eq)
                    return true;
            }
            return false;
        };
        if (p.cmds.length > 10)
            return false;
        var cmds = p.cmds.join(""), crds = p.crds;
        var sameRect = false;
        if ((cmds == "MLLLZ" && crds.length == 8)
            || (cmds == "MLLLLZ" && crds.length == 10)) {
            if (crds.length == 10)
                crds = crds.slice(0, 8);
            var x0 = bb[0], y0 = bb[1], x1 = bb[2], y1 = bb[3];
            if (!sameRect)
                sameRect = sameCrd8(crds, [x0, y0, x1, y0, x1, y1, x0, y1]);
            if (!sameRect)
                sameRect = sameCrd8(crds, [x0, y1, x1, y1, x1, y0, x0, y0]);
        }
        return sameRect;
    },
    boxArea: function (a) { var w = a[2] - a[0], h = a[3] - a[1]; return w * h; },
    newPath: function (gst) { gst.pth = { cmds: [], crds: [] }; },
    moveTo: function (gst, x, y) {
        var p = UDOC.M.multPoint(gst.ctm, [x, y]); //if(gst.cpos[0]==p[0] && gst.cpos[1]==p[1]) return;
        gst.pth.cmds.push("M");
        gst.pth.crds.push(p[0], p[1]);
        gst.cpos = p;
    },
    lineTo: function (gst, x, y) {
        var p = UDOC.M.multPoint(gst.ctm, [x, y]);
        if (gst.cpos[0] == p[0] && gst.cpos[1] == p[1])
            return;
        gst.pth.cmds.push("L");
        gst.pth.crds.push(p[0], p[1]);
        gst.cpos = p;
    },
    curveTo: function (gst, x1, y1, x2, y2, x3, y3) {
        var p;
        p = UDOC.M.multPoint(gst.ctm, [x1, y1]);
        x1 = p[0];
        y1 = p[1];
        p = UDOC.M.multPoint(gst.ctm, [x2, y2]);
        x2 = p[0];
        y2 = p[1];
        p = UDOC.M.multPoint(gst.ctm, [x3, y3]);
        x3 = p[0];
        y3 = p[1];
        gst.cpos = p;
        gst.pth.cmds.push("C");
        gst.pth.crds.push(x1, y1, x2, y2, x3, y3);
    },
    closePath: function (gst) { gst.pth.cmds.push("Z"); },
    arc: function (gst, x, y, r, a0, a1, neg) {
        // circle from a0 counter-clock-wise to a1
        if (neg)
            while (a1 > a0)
                a1 -= 2 * Math.PI;
        else
            while (a1 < a0)
                a1 += 2 * Math.PI;
        var th = (a1 - a0) / 4;
        var x0 = Math.cos(th / 2), y0 = -Math.sin(th / 2);
        var x1 = (4 - x0) / 3, y1 = y0 == 0 ? y0 : (1 - x0) * (3 - x0) / (3 * y0);
        var x2 = x1, y2 = -y1;
        var x3 = x0, y3 = -y0;
        var p1 = [x1, y1], p2 = [x2, y2], p3 = [x3, y3];
        var pth = { cmds: [(gst.pth.cmds.length == 0) ? "M" : "L", "C", "C", "C", "C"], crds: [x0, y0, x1, y1, x2, y2, x3, y3] };
        var rot = [1, 0, 0, 1, 0, 0];
        UDOC.M.rotate(rot, -th);
        for (var i = 0; i < 3; i++) {
            p1 = UDOC.M.multPoint(rot, p1);
            p2 = UDOC.M.multPoint(rot, p2);
            p3 = UDOC.M.multPoint(rot, p3);
            pth.crds.push(p1[0], p1[1], p2[0], p2[1], p3[0], p3[1]);
        }
        var sc = [r, 0, 0, r, x, y];
        UDOC.M.rotate(rot, -a0 + th / 2);
        UDOC.M.concat(rot, sc);
        UDOC.M.multArray(rot, pth.crds);
        UDOC.M.multArray(gst.ctm, pth.crds);
        UDOC.G.concat(gst.pth, pth);
        var y = pth.crds.pop();
        x = pth.crds.pop();
        gst.cpos = [x, y];
    },
    toPoly: function (p) {
        if (p.cmds[0] != "M" || p.cmds[p.cmds.length - 1] != "Z")
            return null;
        for (var i = 1; i < p.cmds.length - 1; i++)
            if (p.cmds[i] != "L")
                return null;
        var out = [], cl = p.crds.length;
        if (p.crds[0] == p.crds[cl - 2] && p.crds[1] == p.crds[cl - 1])
            cl -= 2;
        for (var i = 0; i < cl; i += 2)
            out.push([p.crds[i], p.crds[i + 1]]);
        if (UDOC.G.polyArea(p.crds) < 0)
            out.reverse();
        return out;
    },
    fromPoly: function (p) {
        var o = { cmds: [], crds: [] };
        for (var i = 0; i < p.length; i++) {
            o.crds.push(p[i][0], p[i][1]);
            o.cmds.push(i == 0 ? "M" : "L");
        }
        o.cmds.push("Z");
        return o;
    },
    polyArea: function (p) {
        if (p.length < 6)
            return 0;
        var l = p.length - 2;
        var sum = (p[0] - p[l]) * (p[l + 1] + p[1]);
        for (var i = 0; i < l; i += 2)
            sum += (p[i + 2] - p[i]) * (p[i + 1] + p[i + 3]);
        return -sum * 0.5;
    },
    polyClip: function (p0, p1) {
        var cp1, cp2, s, e;
        var inside = function (p) {
            return (cp2[0] - cp1[0]) * (p[1] - cp1[1]) > (cp2[1] - cp1[1]) * (p[0] - cp1[0]);
        };
        var isc = function () {
            var dc = [cp1[0] - cp2[0], cp1[1] - cp2[1]], dp = [s[0] - e[0], s[1] - e[1]], n1 = cp1[0] * cp2[1] - cp1[1] * cp2[0], n2 = s[0] * e[1] - s[1] * e[0], n3 = 1.0 / (dc[0] * dp[1] - dc[1] * dp[0]);
            return [(n1 * dp[0] - n2 * dc[0]) * n3, (n1 * dp[1] - n2 * dc[1]) * n3];
        };
        var out = p0;
        cp1 = p1[p1.length - 1];
        for (var j in p1) {
            var cp2 = p1[j];
            var inp = out;
            out = [];
            s = inp[inp.length - 1]; //last on the input list
            for (var i in inp) {
                var e = inp[i];
                if (inside(e)) {
                    if (!inside(s)) {
                        out.push(isc());
                    }
                    out.push(e);
                }
                else if (inside(s)) {
                    out.push(isc());
                }
                s = e;
            }
            cp1 = cp2;
        }
        return out;
    }
};
UDOC.M = {
    getScale: function (m) { return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])); },
    translate: function (m, x, y) { UDOC.M.concat(m, [1, 0, 0, 1, x, y]); },
    rotate: function (m, a) { UDOC.M.concat(m, [Math.cos(a), -Math.sin(a), Math.sin(a), Math.cos(a), 0, 0]); },
    scale: function (m, x, y) { UDOC.M.concat(m, [x, 0, 0, y, 0, 0]); },
    concat: function (m, w) {
        var a = m[0], b = m[1], c = m[2], d = m[3], tx = m[4], ty = m[5];
        m[0] = (a * w[0]) + (b * w[2]);
        m[1] = (a * w[1]) + (b * w[3]);
        m[2] = (c * w[0]) + (d * w[2]);
        m[3] = (c * w[1]) + (d * w[3]);
        m[4] = (tx * w[0]) + (ty * w[2]) + w[4];
        m[5] = (tx * w[1]) + (ty * w[3]) + w[5];
    },
    invert: function (m) {
        var a = m[0], b = m[1], c = m[2], d = m[3], tx = m[4], ty = m[5], adbc = a * d - b * c;
        m[0] = d / adbc;
        m[1] = -b / adbc;
        m[2] = -c / adbc;
        m[3] = a / adbc;
        m[4] = (c * ty - d * tx) / adbc;
        m[5] = (b * tx - a * ty) / adbc;
    },
    multPoint: function (m, p) { var x = p[0], y = p[1]; return [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]]; },
    multArray: function (m, a) { for (var i = 0; i < a.length; i += 2) {
        var x = a[i], y = a[i + 1];
        a[i] = x * m[0] + y * m[2] + m[4];
        a[i + 1] = x * m[1] + y * m[3] + m[5];
    } }
};
UDOC.C = {
    srgbGamma: function (x) { return x < 0.0031308 ? 12.92 * x : 1.055 * Math.pow(x, 1.0 / 2.4) - 0.055; },
    cmykToRgb: function (clr) {
        var c = clr[0], m = clr[1], y = clr[2], k = clr[3];
        // return [1-Math.min(1,c+k), 1-Math.min(1, m+k), 1-Math.min(1,y+k)];
        var r = 255
            + c * (-4.387332384609988 * c + 54.48615194189176 * m + 18.82290502165302 * y + 212.25662451639585 * k + -285.2331026137004)
            + m * (1.7149763477362134 * m - 5.6096736904047315 * y + -17.873870861415444 * k - 5.497006427196366)
            + y * (-2.5217340131683033 * y - 21.248923337353073 * k + 17.5119270841813)
            + k * (-21.86122147463605 * k - 189.48180835922747);
        var g = 255
            + c * (8.841041422036149 * c + 60.118027045597366 * m + 6.871425592049007 * y + 31.159100130055922 * k + -79.2970844816548)
            + m * (-15.310361306967817 * m + 17.575251261109482 * y + 131.35250912493976 * k - 190.9453302588951)
            + y * (4.444339102852739 * y + 9.8632861493405 * k - 24.86741582555878)
            + k * (-20.737325471181034 * k - 187.80453709719578);
        var b = 255
            + c * (0.8842522430003296 * c + 8.078677503112928 * m + 30.89978309703729 * y - 0.23883238689178934 * k + -14.183576799673286)
            + m * (10.49593273432072 * m + 63.02378494754052 * y + 50.606957656360734 * k - 112.23884253719248)
            + y * (0.03296041114873217 * y + 115.60384449646641 * k + -193.58209356861505)
            + k * (-22.33816807309886 * k - 180.12613974708367);
        return [Math.max(0, Math.min(1, r / 255)), Math.max(0, Math.min(1, g / 255)), Math.max(0, Math.min(1, b / 255))];
        //var iK = 1-c[3];  
        //return [(1-c[0])*iK, (1-c[1])*iK, (1-c[2])*iK];  
    },
    labToRgb: function (lab) {
        var k = 903.3, e = 0.008856, L = lab[0], a = lab[1], b = lab[2];
        var fy = (L + 16) / 116, fy3 = fy * fy * fy;
        var fz = fy - b / 200, fz3 = fz * fz * fz;
        var fx = a / 500 + fy, fx3 = fx * fx * fx;
        var zr = fz3 > e ? fz3 : (116 * fz - 16) / k;
        var yr = fy3 > e ? fy3 : (116 * fy - 16) / k;
        var xr = fx3 > e ? fx3 : (116 * fx - 16) / k;
        var X = xr * 96.72, Y = yr * 100, Z = zr * 81.427, xyz = [X / 100, Y / 100, Z / 100];
        var x2s = [3.1338561, -1.6168667, -0.4906146, -0.9787684, 1.9161415, 0.0334540, 0.0719453, -0.2289914, 1.4052427];
        var rgb = [x2s[0] * xyz[0] + x2s[1] * xyz[1] + x2s[2] * xyz[2],
            x2s[3] * xyz[0] + x2s[4] * xyz[1] + x2s[5] * xyz[2],
            x2s[6] * xyz[0] + x2s[7] * xyz[1] + x2s[8] * xyz[2]];
        for (var i = 0; i < 3; i++)
            rgb[i] = Math.max(0, Math.min(1, UDOC.C.srgbGamma(rgb[i])));
        return rgb;
    }
};
UDOC.getState = function (crds) {
    return {
        font: UDOC.getFont(),
        dd: { flat: 1 },
        space: "/DeviceGray",
        // fill
        ca: 1,
        colr: [0, 0, 0],
        sspace: "/DeviceGray",
        // stroke
        CA: 1,
        COLR: [0, 0, 0],
        bmode: "/Normal",
        SA: false, OPM: 0, AIS: false, OP: false, op: false, SMask: "/None",
        lwidth: 1,
        lcap: 0,
        ljoin: 0,
        mlimit: 10,
        SM: 0.1,
        doff: 0,
        dash: [],
        ctm: [1, 0, 0, 1, 0, 0],
        cpos: [0, 0],
        pth: { cmds: [], crds: [] },
        cpth: crds ? UDOC.G.rectToPath(crds) : null // clipping path
    };
};
UDOC.getFont = function () {
    return {
        Tc: 0,
        Tw: 0,
        Th: 100,
        Tl: 0,
        Tf: "Helvetica-Bold",
        Tfs: 1,
        Tmode: 0,
        Trise: 0,
        Tk: 0,
        Tal: 0,
        Tun: 0,
        Tm: [1, 0, 0, 1, 0, 0],
        Tlm: [1, 0, 0, 1, 0, 0],
        Trm: [1, 0, 0, 1, 0, 0]
    };
};
var FromEMF = function () {
};
FromEMF.Parse = function (buff, genv) {
    buff = new Uint8Array(buff);
    var off = 0;
    //console.log(buff.slice(0,32));
    var prms = { fill: false, strk: false, bb: [0, 0, 1, 1], wbb: [0, 0, 1, 1], fnt: { nam: "Arial", hgh: 25, und: false, orn: 0 }, tclr: [0, 0, 0], talg: 0 }, gst, tab = [], sts = [];
    var rI = FromEMF.B.readShort, rU = FromEMF.B.readUshort, rI32 = FromEMF.B.readInt, rU32 = FromEMF.B.readUint, rF32 = FromEMF.B.readFloat;
    while (true) {
        var fnc = rU32(buff, off);
        off += 4;
        var fnm = FromEMF.K[fnc];
        var siz = rU32(buff, off);
        off += 4;
        //if(gst && isNaN(gst.ctm[0])) throw "e";
        //console.log(fnc,fnm,siz);
        var loff = off;
        //if(opn++==253) break;
        var obj = null, oid = 0;
        //console.log(fnm, siz);
        if (fnm == "EOF") {
            break;
        }
        else if (fnm == "HEADER") {
            prms.bb = FromEMF._readBox(buff, loff);
            loff += 16; //console.log(fnm, prms.bb);
            genv.StartPage(prms.bb[0], prms.bb[1], prms.bb[2], prms.bb[3]);
            gst = UDOC.getState(prms.bb);
        }
        else if (fnm == "SAVEDC")
            sts.push(JSON.stringify(gst), JSON.stringify(prms));
        else if (fnm == "RESTOREDC") {
            var dif = rI32(buff, loff);
            loff += 4;
            while (dif < -1) {
                sts.pop();
                sts.pop();
            }
            prms = JSON.parse(sts.pop());
            gst = JSON.parse(sts.pop());
        }
        else if (fnm == "SELECTCLIPPATH") {
            gst.cpth = JSON.parse(JSON.stringify(gst.pth));
        }
        else if (["SETMAPMODE", "SETPOLYFILLMODE", "SETBKMODE" /*,"SETVIEWPORTEXTEX"*/, "SETICMMODE", "SETROP2", "EXTSELECTCLIPRGN"].indexOf(fnm) != -1) ;
        //else if(fnm=="INTERSECTCLIPRECT") {  var r=prms.crct=FromEMF._readBox(buff, loff);  /*var y0=r[1],y1=r[3]; if(y0>y1){r[1]=y1; r[3]=y0;}*/ console.log(prms.crct);  }
        else if (fnm == "SETMITERLIMIT")
            gst.mlimit = rU32(buff, loff);
        else if (fnm == "SETTEXTCOLOR")
            prms.tclr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
        else if (fnm == "SETTEXTALIGN")
            prms.talg = rU32(buff, loff);
        else if (fnm == "SETVIEWPORTEXTEX" || fnm == "SETVIEWPORTORGEX") {
            if (prms.vbb == null)
                prms.vbb = [];
            var coff = fnm == "SETVIEWPORTORGEX" ? 0 : 2;
            prms.vbb[coff] = rI32(buff, loff);
            loff += 4;
            prms.vbb[coff + 1] = rI32(buff, loff);
            loff += 4;
            //console.log(prms.vbb);
            if (fnm == "SETVIEWPORTEXTEX")
                FromEMF._updateCtm(prms, gst);
        }
        else if (fnm == "SETWINDOWEXTEX" || fnm == "SETWINDOWORGEX") {
            var coff = fnm == "SETWINDOWORGEX" ? 0 : 2;
            prms.wbb[coff] = rI32(buff, loff);
            loff += 4;
            prms.wbb[coff + 1] = rI32(buff, loff);
            loff += 4;
            if (fnm == "SETWINDOWEXTEX")
                FromEMF._updateCtm(prms, gst);
        }
        //else if(fnm=="SETMETARGN") {}
        else if (fnm == "COMMENT") {
            var ds = rU32(buff, loff);
            loff += 4;
        }
        else if (fnm == "SELECTOBJECT") {
            var ind = rU32(buff, loff);
            loff += 4;
            //console.log(ind.toString(16), tab, tab[ind]);
            if (ind == 0x80000000) {
                prms.fill = true;
                gst.colr = [1, 1, 1];
            } // white brush
            else if (ind == 0x80000005) {
                prms.fill = false;
            } // null brush
            else if (ind == 0x80000007) {
                prms.strk = true;
                prms.lwidth = 1;
                gst.COLR = [0, 0, 0];
            } // black pen
            else if (ind == 0x80000008) {
                prms.strk = false;
            } // null  pen
            else if (ind == 0x8000000d) ; // system font
            else if (ind == 0x8000000e) ; // device default font
            else {
                var co = tab[ind]; //console.log(ind, co);
                if (co.t == "b") {
                    prms.fill = co.stl != 1;
                    if (co.stl == 0) ;
                    else if (co.stl == 1) ;
                    else
                        throw co.stl + " e";
                    gst.colr = co.clr;
                }
                else if (co.t == "p") {
                    prms.strk = co.stl != 5;
                    gst.lwidth = co.wid;
                    gst.COLR = co.clr;
                }
                else if (co.t == "f") {
                    prms.fnt = co;
                    gst.font.Tf = co.nam;
                    gst.font.Tfs = Math.abs(co.hgh);
                    gst.font.Tun = co.und;
                }
                else
                    throw "e";
            }
        }
        else if (fnm == "DELETEOBJECT") {
            var ind = rU32(buff, loff);
            loff += 4;
            if (tab[ind] != null)
                tab[ind] = null;
            else
                throw "e";
        }
        else if (fnm == "CREATEBRUSHINDIRECT") {
            oid = rU32(buff, loff);
            loff += 4;
            obj = { t: "b" };
            obj.stl = rU32(buff, loff);
            loff += 4;
            obj.clr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
            loff += 4;
            obj.htc = rU32(buff, loff);
            loff += 4;
            //console.log(oid, obj);
        }
        else if (fnm == "CREATEPEN" || fnm == "EXTCREATEPEN") {
            oid = rU32(buff, loff);
            loff += 4;
            obj = { t: "p" };
            if (fnm == "EXTCREATEPEN") {
                loff += 16;
                obj.stl = rU32(buff, loff);
                loff += 4;
                obj.wid = rU32(buff, loff);
                loff += 4;
                //obj.stl = rU32(buff, loff);  
                loff += 4;
            }
            else {
                obj.stl = rU32(buff, loff);
                loff += 4;
                obj.wid = rU32(buff, loff);
                loff += 4;
                loff += 4;
            }
            obj.clr = [buff[loff] / 255, buff[loff + 1] / 255, buff[loff + 2] / 255];
            loff += 4;
        }
        else if (fnm == "EXTCREATEFONTINDIRECTW") {
            oid = rU32(buff, loff);
            loff += 4;
            obj = { t: "f", nam: "" };
            obj.hgh = rI32(buff, loff);
            loff += 4;
            loff += 4 * 2;
            obj.orn = rI32(buff, loff) / 10;
            loff += 4;
            var wgh = rU32(buff, loff);
            loff += 4; //console.log(fnm, obj.orn, wgh);
            //console.log(rU32(buff,loff), rU32(buff,loff+4), buff.slice(loff,loff+8));
            obj.und = buff[loff + 1];
            obj.stk = buff[loff + 2];
            loff += 4 * 2;
            while (rU(buff, loff) != 0) {
                obj.nam += String.fromCharCode(rU(buff, loff));
                loff += 2;
            }
            if (wgh > 500)
                obj.nam += "-Bold";
            //console.log(wgh, obj.nam);
        }
        else if (fnm == "EXTTEXTOUTW") {
            //console.log(buff.slice(loff-8, loff-8+siz));
            loff += 16;
            var mod = rU32(buff, loff);
            loff += 4; //console.log(mod);
            var scx = rF32(buff, loff);
            loff += 4;
            var scy = rF32(buff, loff);
            loff += 4;
            var rfx = rI32(buff, loff);
            loff += 4;
            var rfy = rI32(buff, loff);
            loff += 4;
            //console.log(mod, scx, scy,rfx,rfy);
            gst.font.Tm = [1, 0, 0, -1, 0, 0];
            UDOC.M.rotate(gst.font.Tm, prms.fnt.orn * Math.PI / 180);
            UDOC.M.translate(gst.font.Tm, rfx, rfy);
            var alg = prms.talg; //console.log(alg.toString(2));
            if ((alg & 6) == 6)
                gst.font.Tal = 2;
            else if ((alg & 7) == 0)
                gst.font.Tal = 0;
            else
                throw alg + " e";
            if ((alg & 24) == 24) ; // baseline
            else if ((alg & 24) == 0)
                UDOC.M.translate(gst.font.Tm, 0, gst.font.Tfs);
            else
                throw "e";
            var crs = rU32(buff, loff);
            loff += 4;
            var ofs = rU32(buff, loff);
            loff += 4;
            var ops = rU32(buff, loff);
            loff += 4; //if(ops!=0) throw "e";
            //console.log(ofs,ops,crs);
            loff += 16;
            var ofD = rU32(buff, loff);
            loff += 4; //console.log(ops, ofD, loff, ofs+off-8);
            ofs += off - 8; //console.log(crs, ops);
            var str = "";
            for (var i = 0; i < crs; i++) {
                var cc = rU(buff, ofs + i * 2);
                str += String.fromCharCode(cc);
            }
            var oclr = gst.colr;
            gst.colr = prms.tclr;
            //console.log(str, gst.colr, gst.font.Tm);
            //var otfs = gst.font.Tfs;  gst.font.Tfs *= 1/gst.ctm[0];
            genv.PutText(gst, str, str.length * gst.font.Tfs * 0.5);
            gst.colr = oclr;
            //gst.font.Tfs = otfs;
            //console.log(rfx, rfy, scx, ops, rcX, rcY, rcW, rcH, offDx, str);
        }
        else if (fnm == "BEGINPATH") {
            UDOC.G.newPath(gst);
        }
        else if (fnm == "ENDPATH") ;
        else if (fnm == "CLOSEFIGURE")
            UDOC.G.closePath(gst);
        else if (fnm == "MOVETOEX") {
            UDOC.G.moveTo(gst, rI32(buff, loff), rI32(buff, loff + 4));
        }
        else if (fnm == "LINETO") {
            if (gst.pth.cmds.length == 0) {
                var im = gst.ctm.slice(0);
                UDOC.M.invert(im);
                var p = UDOC.M.multPoint(im, gst.cpos);
                UDOC.G.moveTo(gst, p[0], p[1]);
            }
            UDOC.G.lineTo(gst, rI32(buff, loff), rI32(buff, loff + 4));
        }
        else if (fnm == "POLYGON" || fnm == "POLYGON16" || fnm == "POLYLINE" || fnm == "POLYLINE16" || fnm == "POLYLINETO" || fnm == "POLYLINETO16") {
            loff += 16;
            var ndf = fnm.startsWith("POLYGON"), isTo = fnm.indexOf("TO") != -1;
            var cnt = rU32(buff, loff);
            loff += 4;
            if (!isTo)
                UDOC.G.newPath(gst);
            loff = FromEMF._drawPoly(buff, loff, cnt, gst, fnm.endsWith("16") ? 2 : 4, ndf, isTo);
            if (!isTo)
                FromEMF._draw(genv, gst, prms, ndf);
            //console.log(prms, gst.lwidth);
            //console.log(JSON.parse(JSON.stringify(gst.pth)));
        }
        else if (fnm == "POLYPOLYGON16") {
            loff += 16;
            var ndf = fnm.startsWith("POLYPOLYGON"), isTo = fnm.indexOf("TO") != -1;
            var nop = rU32(buff, loff);
            loff += 4;
            loff += 4;
            var pi = loff;
            loff += nop * 4;
            if (!isTo)
                UDOC.G.newPath(gst);
            for (var i = 0; i < nop; i++) {
                var ppp = rU(buff, pi + i * 4);
                loff = FromEMF._drawPoly(buff, loff, ppp, gst, fnm.endsWith("16") ? 2 : 4, ndf, isTo);
            }
            if (!isTo)
                FromEMF._draw(genv, gst, prms, ndf);
        }
        else if (fnm == "POLYBEZIER" || fnm == "POLYBEZIER16" || fnm == "POLYBEZIERTO" || fnm == "POLYBEZIERTO16") {
            loff += 16;
            var is16 = fnm.endsWith("16"), rC = is16 ? rI : rI32, nl = is16 ? 2 : 4;
            var cnt = rU32(buff, loff);
            loff += 4;
            if (fnm.indexOf("TO") == -1) {
                UDOC.G.moveTo(gst, rC(buff, loff), rC(buff, loff + nl));
                loff += 2 * nl;
                cnt--;
            }
            while (cnt > 0) {
                UDOC.G.curveTo(gst, rC(buff, loff), rC(buff, loff + nl), rC(buff, loff + 2 * nl), rC(buff, loff + 3 * nl), rC(buff, loff + 4 * nl), rC(buff, loff + 5 * nl));
                loff += 6 * nl;
                cnt -= 3;
            }
            //console.log(JSON.parse(JSON.stringify(gst.pth)));
        }
        else if (fnm == "RECTANGLE" || fnm == "ELLIPSE") {
            UDOC.G.newPath(gst);
            var bx = FromEMF._readBox(buff, loff);
            if (fnm == "RECTANGLE") {
                UDOC.G.moveTo(gst, bx[0], bx[1]);
                UDOC.G.lineTo(gst, bx[2], bx[1]);
                UDOC.G.lineTo(gst, bx[2], bx[3]);
                UDOC.G.lineTo(gst, bx[0], bx[3]);
            }
            else {
                var x = (bx[0] + bx[2]) / 2, y = (bx[1] + bx[3]) / 2;
                UDOC.G.arc(gst, x, y, (bx[2] - bx[0]) / 2, 0, 2 * Math.PI, false);
            }
            UDOC.G.closePath(gst);
            FromEMF._draw(genv, gst, prms, true);
            //console.log(prms, gst.lwidth);
        }
        else if (fnm == "FILLPATH")
            genv.Fill(gst, false);
        else if (fnm == "STROKEPATH")
            genv.Stroke(gst);
        else if (fnm == "STROKEANDFILLPATH") {
            genv.Fill(gst, false);
            genv.Stroke(gst);
        }
        else if (fnm == "SETWORLDTRANSFORM" || fnm == "MODIFYWORLDTRANSFORM") {
            var mat = [];
            for (var i = 0; i < 6; i++)
                mat.push(rF32(buff, loff + i * 4));
            loff += 24;
            //console.log(fnm, gst.ctm.slice(0), mat);
            if (fnm == "SETWORLDTRANSFORM")
                gst.ctm = mat;
            else {
                var mod = rU32(buff, loff);
                loff += 4;
                if (mod == 2) {
                    var om = gst.ctm;
                    gst.ctm = mat;
                    UDOC.M.concat(gst.ctm, om);
                }
                else
                    throw "e";
            }
        }
        else if (fnm == "SETSTRETCHBLTMODE") {
            var sm = rU32(buff, loff);
            loff += 4;
        }
        else if (fnm == "STRETCHDIBITS") {
            var bx = FromEMF._readBox(buff, loff);
            loff += 16;
            var xD = rI32(buff, loff);
            loff += 4;
            var yD = rI32(buff, loff);
            loff += 4;
            var xS = rI32(buff, loff);
            loff += 4;
            var yS = rI32(buff, loff);
            loff += 4;
            var wS = rI32(buff, loff);
            loff += 4;
            var hS = rI32(buff, loff);
            loff += 4;
            var ofH = rU32(buff, loff) + off - 8;
            loff += 4;
            var szH = rU32(buff, loff);
            loff += 4;
            var ofB = rU32(buff, loff) + off - 8;
            loff += 4;
            var szB = rU32(buff, loff);
            loff += 4;
            var usg = rU32(buff, loff);
            loff += 4;
            if (usg != 0)
                throw "e";
            var bop = rU32(buff, loff);
            loff += 4;
            var wD = rI32(buff, loff);
            loff += 4;
            var hD = rI32(buff, loff);
            loff += 4; //console.log(bop, wD, hD);
            //console.log(ofH, szH, ofB, szB, ofH+40);
            //console.log(bx, xD,yD,wD,hD);
            //console.log(xS,yS,wS,hS);
            //console.log(ofH,szH,ofB,szB,usg,bop);
            var hl = rU32(buff, ofH);
            ofH += 4;
            var w = rU32(buff, ofH);
            ofH += 4;
            var h = rU32(buff, ofH);
            ofH += 4;
            if (w != wS || h != hS)
                throw "e";
            var ps = rU(buff, ofH);
            ofH += 2;
            var bc = rU(buff, ofH);
            ofH += 2;
            if (bc != 8 && bc != 24 && bc != 32)
                throw bc + " e";
            var cpr = rU32(buff, ofH);
            ofH += 4;
            if (cpr != 0)
                throw cpr + " e";
            var sz = rU32(buff, ofH);
            ofH += 4;
            var xpm = rU32(buff, ofH);
            ofH += 4;
            var ypm = rU32(buff, ofH);
            ofH += 4;
            var cu = rU32(buff, ofH);
            ofH += 4;
            var ci = rU32(buff, ofH);
            ofH += 4; //console.log(hl, w, h, ps, bc, cpr, sz, xpm, ypm, cu, ci);
            //console.log(hl,w,h,",",xS,yS,wS,hS,",",xD,yD,wD,hD,",",xpm,ypm);
            var rl = Math.floor(((w * ps * bc + 31) & ~31) / 8);
            var img = new Uint8Array(w * h * 4);
            if (bc == 8) {
                for (var y = 0; y < h; y++)
                    for (var x = 0; x < w; x++) {
                        var qi = (y * w + x) << 2, ind = buff[ofB + (h - 1 - y) * rl + x] << 2;
                        img[qi] = buff[ofH + ind + 2];
                        img[qi + 1] = buff[ofH + ind + 1];
                        img[qi + 2] = buff[ofH + ind + 0];
                        img[qi + 3] = 255;
                    }
            }
            if (bc == 24) {
                for (var y = 0; y < h; y++)
                    for (var x = 0; x < w; x++) {
                        var qi = (y * w + x) << 2, ti = ofB + (h - 1 - y) * rl + x * 3;
                        img[qi] = buff[ti + 2];
                        img[qi + 1] = buff[ti + 1];
                        img[qi + 2] = buff[ti + 0];
                        img[qi + 3] = 255;
                    }
            }
            if (bc == 32) {
                for (var y = 0; y < h; y++)
                    for (var x = 0; x < w; x++) {
                        var qi = (y * w + x) << 2, ti = ofB + (h - 1 - y) * rl + x * 4;
                        img[qi] = buff[ti + 2];
                        img[qi + 1] = buff[ti + 1];
                        img[qi + 2] = buff[ti + 0];
                        img[qi + 3] = buff[ti + 3];
                    }
            }
            var ctm = gst.ctm.slice(0);
            gst.ctm = [1, 0, 0, 1, 0, 0];
            UDOC.M.scale(gst.ctm, wD, -hD);
            UDOC.M.translate(gst.ctm, xD, yD + hD);
            UDOC.M.concat(gst.ctm, ctm);
            genv.PutImage(gst, img, w, h);
            gst.ctm = ctm;
        }
        else {
            console.log(fnm, siz);
        }
        if (obj != null)
            tab[oid] = obj;
        off += siz - 8;
    }
    //genv.Stroke(gst);
    genv.ShowPage();
    genv.Done();
};
FromEMF._readBox = function (buff, off) { var b = []; for (var i = 0; i < 4; i++)
    b[i] = FromEMF.B.readInt(buff, off + i * 4); return b; };
FromEMF._updateCtm = function (prms, gst) {
    var mat = [1, 0, 0, 1, 0, 0];
    var wbb = prms.wbb, bb = prms.bb, vbb = (prms.vbb && prms.vbb.length == 4) ? prms.vbb : prms.bb;
    //var y0 = bb[1], y1 = bb[3];  bb[1]=Math.min(y0,y1);  bb[3]=Math.max(y0,y1);
    UDOC.M.translate(mat, -wbb[0], -wbb[1]);
    UDOC.M.scale(mat, 1 / wbb[2], 1 / wbb[3]);
    UDOC.M.scale(mat, vbb[2], vbb[3]);
    //UDOC.M.scale(mat, vbb[2]/(bb[2]-bb[0]), vbb[3]/(bb[3]-bb[1]));
    //UDOC.M.scale(mat, bb[2]-bb[0],bb[3]-bb[1]);
    gst.ctm = mat;
};
FromEMF._draw = function (genv, gst, prms, needFill) {
    if (prms.fill && needFill)
        genv.Fill(gst, false);
    if (prms.strk && gst.lwidth != 0)
        genv.Stroke(gst);
};
FromEMF._drawPoly = function (buff, off, ppp, gst, nl, clos, justLine) {
    var rS = nl == 2 ? FromEMF.B.readShort : FromEMF.B.readInt;
    for (var j = 0; j < ppp; j++) {
        var px = rS(buff, off);
        off += nl;
        var py = rS(buff, off);
        off += nl;
        if (j == 0 && !justLine)
            UDOC.G.moveTo(gst, px, py);
        else
            UDOC.G.lineTo(gst, px, py);
    }
    if (clos)
        UDOC.G.closePath(gst);
    return off;
};
FromEMF.B = {
    uint8: new Uint8Array(4),
    readShort: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; return FromEMF.B.int16[0]; },
    readUshort: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; return FromEMF.B.uint16[0]; },
    readInt: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; u8[2] = buff[p + 2]; u8[3] = buff[p + 3]; return FromEMF.B.int32[0]; },
    readUint: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; u8[2] = buff[p + 2]; u8[3] = buff[p + 3]; return FromEMF.B.uint32[0]; },
    readFloat: function (buff, p) { var u8 = FromEMF.B.uint8; u8[0] = buff[p]; u8[1] = buff[p + 1]; u8[2] = buff[p + 2]; u8[3] = buff[p + 3]; return FromEMF.B.flot32[0]; },
    readASCII: function (buff, p, l) { var s = ""; for (var i = 0; i < l; i++)
        s += String.fromCharCode(buff[p + i]); return s; }
};
FromEMF.B.int16 = new Int16Array(FromEMF.B.uint8.buffer);
FromEMF.B.uint16 = new Uint16Array(FromEMF.B.uint8.buffer);
FromEMF.B.int32 = new Int32Array(FromEMF.B.uint8.buffer);
FromEMF.B.uint32 = new Uint32Array(FromEMF.B.uint8.buffer);
FromEMF.B.flot32 = new Float32Array(FromEMF.B.uint8.buffer);
FromEMF.C = {
    EMR_HEADER: 0x00000001,
    EMR_POLYBEZIER: 0x00000002,
    EMR_POLYGON: 0x00000003,
    EMR_POLYLINE: 0x00000004,
    EMR_POLYBEZIERTO: 0x00000005,
    EMR_POLYLINETO: 0x00000006,
    EMR_POLYPOLYLINE: 0x00000007,
    EMR_POLYPOLYGON: 0x00000008,
    EMR_SETWINDOWEXTEX: 0x00000009,
    EMR_SETWINDOWORGEX: 0x0000000A,
    EMR_SETVIEWPORTEXTEX: 0x0000000B,
    EMR_SETVIEWPORTORGEX: 0x0000000C,
    EMR_SETBRUSHORGEX: 0x0000000D,
    EMR_EOF: 0x0000000E,
    EMR_SETPIXELV: 0x0000000F,
    EMR_SETMAPPERFLAGS: 0x00000010,
    EMR_SETMAPMODE: 0x00000011,
    EMR_SETBKMODE: 0x00000012,
    EMR_SETPOLYFILLMODE: 0x00000013,
    EMR_SETROP2: 0x00000014,
    EMR_SETSTRETCHBLTMODE: 0x00000015,
    EMR_SETTEXTALIGN: 0x00000016,
    EMR_SETCOLORADJUSTMENT: 0x00000017,
    EMR_SETTEXTCOLOR: 0x00000018,
    EMR_SETBKCOLOR: 0x00000019,
    EMR_OFFSETCLIPRGN: 0x0000001A,
    EMR_MOVETOEX: 0x0000001B,
    EMR_SETMETARGN: 0x0000001C,
    EMR_EXCLUDECLIPRECT: 0x0000001D,
    EMR_INTERSECTCLIPRECT: 0x0000001E,
    EMR_SCALEVIEWPORTEXTEX: 0x0000001F,
    EMR_SCALEWINDOWEXTEX: 0x00000020,
    EMR_SAVEDC: 0x00000021,
    EMR_RESTOREDC: 0x00000022,
    EMR_SETWORLDTRANSFORM: 0x00000023,
    EMR_MODIFYWORLDTRANSFORM: 0x00000024,
    EMR_SELECTOBJECT: 0x00000025,
    EMR_CREATEPEN: 0x00000026,
    EMR_CREATEBRUSHINDIRECT: 0x00000027,
    EMR_DELETEOBJECT: 0x00000028,
    EMR_ANGLEARC: 0x00000029,
    EMR_ELLIPSE: 0x0000002A,
    EMR_RECTANGLE: 0x0000002B,
    EMR_ROUNDRECT: 0x0000002C,
    EMR_ARC: 0x0000002D,
    EMR_CHORD: 0x0000002E,
    EMR_PIE: 0x0000002F,
    EMR_SELECTPALETTE: 0x00000030,
    EMR_CREATEPALETTE: 0x00000031,
    EMR_SETPALETTEENTRIES: 0x00000032,
    EMR_RESIZEPALETTE: 0x00000033,
    EMR_REALIZEPALETTE: 0x00000034,
    EMR_EXTFLOODFILL: 0x00000035,
    EMR_LINETO: 0x00000036,
    EMR_ARCTO: 0x00000037,
    EMR_POLYDRAW: 0x00000038,
    EMR_SETARCDIRECTION: 0x00000039,
    EMR_SETMITERLIMIT: 0x0000003A,
    EMR_BEGINPATH: 0x0000003B,
    EMR_ENDPATH: 0x0000003C,
    EMR_CLOSEFIGURE: 0x0000003D,
    EMR_FILLPATH: 0x0000003E,
    EMR_STROKEANDFILLPATH: 0x0000003F,
    EMR_STROKEPATH: 0x00000040,
    EMR_FLATTENPATH: 0x00000041,
    EMR_WIDENPATH: 0x00000042,
    EMR_SELECTCLIPPATH: 0x00000043,
    EMR_ABORTPATH: 0x00000044,
    EMR_COMMENT: 0x00000046,
    EMR_FILLRGN: 0x00000047,
    EMR_FRAMERGN: 0x00000048,
    EMR_INVERTRGN: 0x00000049,
    EMR_PAINTRGN: 0x0000004A,
    EMR_EXTSELECTCLIPRGN: 0x0000004B,
    EMR_BITBLT: 0x0000004C,
    EMR_STRETCHBLT: 0x0000004D,
    EMR_MASKBLT: 0x0000004E,
    EMR_PLGBLT: 0x0000004F,
    EMR_SETDIBITSTODEVICE: 0x00000050,
    EMR_STRETCHDIBITS: 0x00000051,
    EMR_EXTCREATEFONTINDIRECTW: 0x00000052,
    EMR_EXTTEXTOUTA: 0x00000053,
    EMR_EXTTEXTOUTW: 0x00000054,
    EMR_POLYBEZIER16: 0x00000055,
    EMR_POLYGON16: 0x00000056,
    EMR_POLYLINE16: 0x00000057,
    EMR_POLYBEZIERTO16: 0x00000058,
    EMR_POLYLINETO16: 0x00000059,
    EMR_POLYPOLYLINE16: 0x0000005A,
    EMR_POLYPOLYGON16: 0x0000005B,
    EMR_POLYDRAW16: 0x0000005C,
    EMR_CREATEMONOBRUSH: 0x0000005D,
    EMR_CREATEDIBPATTERNBRUSHPT: 0x0000005E,
    EMR_EXTCREATEPEN: 0x0000005F,
    EMR_POLYTEXTOUTA: 0x00000060,
    EMR_POLYTEXTOUTW: 0x00000061,
    EMR_SETICMMODE: 0x00000062,
    EMR_CREATECOLORSPACE: 0x00000063,
    EMR_SETCOLORSPACE: 0x00000064,
    EMR_DELETECOLORSPACE: 0x00000065,
    EMR_GLSRECORD: 0x00000066,
    EMR_GLSBOUNDEDRECORD: 0x00000067,
    EMR_PIXELFORMAT: 0x00000068,
    EMR_DRAWESCAPE: 0x00000069,
    EMR_EXTESCAPE: 0x0000006A,
    EMR_SMALLTEXTOUT: 0x0000006C,
    EMR_FORCEUFIMAPPING: 0x0000006D,
    EMR_NAMEDESCAPE: 0x0000006E,
    EMR_COLORCORRECTPALETTE: 0x0000006F,
    EMR_SETICMPROFILEA: 0x00000070,
    EMR_SETICMPROFILEW: 0x00000071,
    EMR_ALPHABLEND: 0x00000072,
    EMR_SETLAYOUT: 0x00000073,
    EMR_TRANSPARENTBLT: 0x00000074,
    EMR_GRADIENTFILL: 0x00000076,
    EMR_SETLINKEDUFIS: 0x00000077,
    EMR_SETTEXTJUSTIFICATION: 0x00000078,
    EMR_COLORMATCHTOTARGETW: 0x00000079,
    EMR_CREATECOLORSPACEW: 0x0000007A
};
FromEMF.K = [];
// (function() {
//     var inp, out, stt;
//     inp = FromEMF.C;   out = FromEMF.K;   stt=4;
//     for(var p in inp) out[inp[p]] = p.slice(stt);
// }  )();
var ToContext2D = function (needPage, scale) {
    this.canvas = document.createElement("canvas");
    this.ctx = this.canvas.getContext("2d");
    this.bb = null;
    this.currPage = 0;
    this.needPage = needPage;
    this.scale = scale;
};
ToContext2D.prototype.StartPage = function (x, y, w, h) {
    if (this.currPage != this.needPage)
        return;
    this.bb = [x, y, w, h];
    var scl = this.scale, dpr = window.devicePixelRatio;
    var cnv = this.canvas, ctx = this.ctx;
    cnv.width = Math.round(w * scl);
    cnv.height = Math.round(h * scl);
    ctx.translate(0, h * scl);
    ctx.scale(scl, -scl);
    cnv.setAttribute("style", "border:1px solid; width:" + (cnv.width / dpr) + "px; height:" + (cnv.height / dpr) + "px");
};
ToContext2D.prototype.Fill = function (gst, evenOdd) {
    if (this.currPage != this.needPage)
        return;
    var ctx = this.ctx;
    ctx.beginPath();
    this._setStyle(gst, ctx);
    this._draw(gst.pth, ctx);
    ctx.fill();
};
ToContext2D.prototype.Stroke = function (gst) {
    if (this.currPage != this.needPage)
        return;
    var ctx = this.ctx;
    ctx.beginPath();
    this._setStyle(gst, ctx);
    this._draw(gst.pth, ctx);
    ctx.stroke();
};
ToContext2D.prototype.PutText = function (gst, str, stw) {
    if (this.currPage != this.needPage)
        return;
    var scl = this._scale(gst.ctm);
    var ctx = this.ctx;
    this._setStyle(gst, ctx);
    ctx.save();
    var m = [1, 0, 0, -1, 0, 0];
    this._concat(m, gst.font.Tm);
    this._concat(m, gst.ctm);
    //console.log(str, m, gst);  throw "e";
    ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
    ctx.fillText(str, 0, 0);
    ctx.restore();
};
ToContext2D.prototype.PutImage = function (gst, buff, w, h, msk) {
    if (this.currPage != this.needPage)
        return;
    var ctx = this.ctx;
    if (buff.length == w * h * 4) {
        buff = buff.slice(0);
        if (msk && msk.length == w * h * 4)
            for (var i = 0; i < buff.length; i += 4)
                buff[i + 3] = msk[i + 1];
        var cnv = document.createElement("canvas"), cctx = cnv.getContext("2d");
        cnv.width = w;
        cnv.height = h;
        var imgd = cctx.createImageData(w, h);
        for (var i = 0; i < buff.length; i++)
            imgd.data[i] = buff[i];
        cctx.putImageData(imgd, 0, 0);
        ctx.save();
        var m = [1, 0, 0, 1, 0, 0];
        this._concat(m, [1 / w, 0, 0, -1 / h, 0, 1]);
        this._concat(m, gst.ctm);
        ctx.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
        ctx.drawImage(cnv, 0, 0);
        ctx.restore();
    }
};
ToContext2D.prototype.ShowPage = function () { this.currPage++; };
ToContext2D.prototype.Done = function () { };
function _flt(n) { return "" + parseFloat(n.toFixed(2)); }
ToContext2D.prototype._setStyle = function (gst, ctx) {
    var scl = this._scale(gst.ctm);
    ctx.fillStyle = this._getFill(gst.colr, gst.ca, ctx);
    ctx.strokeStyle = this._getFill(gst.COLR, gst.CA, ctx);
    ctx.lineCap = ["butt", "round", "square"][gst.lcap];
    ctx.lineJoin = ["miter", "round", "bevel"][gst.ljoin];
    ctx.lineWidth = gst.lwidth * scl;
    var dsh = gst.dash.slice(0);
    for (var i = 0; i < dsh.length; i++)
        dsh[i] = _flt(dsh[i] * scl);
    ctx.setLineDash(dsh);
    ctx.miterLimit = gst.mlimit * scl;
    var fn = gst.font.Tf, ln = fn.toLowerCase();
    var p0 = ln.indexOf("bold") != -1 ? "bold " : "";
    var p1 = (ln.indexOf("italic") != -1 || ln.indexOf("oblique") != -1) ? "italic " : "";
    ctx.font = p0 + p1 + gst.font.Tfs + "px \"" + fn + "\"";
};
ToContext2D.prototype._getFill = function (colr, ca, ctx) {
    if (colr.typ == null)
        return this._colr(colr, ca);
    else {
        var grd = colr, crd = grd.crds, mat = grd.mat, scl = this._scale(mat), gf;
        if (grd.typ == "lin") {
            var p0 = this._multPoint(mat, crd.slice(0, 2)), p1 = this._multPoint(mat, crd.slice(2));
            gf = ctx.createLinearGradient(p0[0], p0[1], p1[0], p1[1]);
        }
        else if (grd.typ == "rad") {
            var p0 = this._multPoint(mat, crd.slice(0, 2)), p1 = this._multPoint(mat, crd.slice(3));
            gf = ctx.createRadialGradient(p0[0], p0[1], crd[2] * scl, p1[0], p1[1], crd[5] * scl);
        }
        for (var i = 0; i < grd.grad.length; i++)
            gf.addColorStop(grd.grad[i][0], this._colr(grd.grad[i][1], ca));
        return gf;
    }
};
ToContext2D.prototype._colr = function (c, a) { return "rgba(" + Math.round(c[0] * 255) + "," + Math.round(c[1] * 255) + "," + Math.round(c[2] * 255) + "," + a + ")"; };
ToContext2D.prototype._scale = function (m) { return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])); };
ToContext2D.prototype._concat = function (m, w) {
    var a = m[0], b = m[1], c = m[2], d = m[3], tx = m[4], ty = m[5];
    m[0] = (a * w[0]) + (b * w[2]);
    m[1] = (a * w[1]) + (b * w[3]);
    m[2] = (c * w[0]) + (d * w[2]);
    m[3] = (c * w[1]) + (d * w[3]);
    m[4] = (tx * w[0]) + (ty * w[2]) + w[4];
    m[5] = (tx * w[1]) + (ty * w[3]) + w[5];
};
ToContext2D.prototype._multPoint = function (m, p) { var x = p[0], y = p[1]; return [x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5]]; },
    ToContext2D.prototype._draw = function (path, ctx) {
        var c = 0, crds = path.crds;
        for (var j = 0; j < path.cmds.length; j++) {
            var cmd = path.cmds[j];
            if (cmd == "M") {
                ctx.moveTo(crds[c], crds[c + 1]);
                c += 2;
            }
            else if (cmd == "L") {
                ctx.lineTo(crds[c], crds[c + 1]);
                c += 2;
            }
            else if (cmd == "C") {
                ctx.bezierCurveTo(crds[c], crds[c + 1], crds[c + 2], crds[c + 3], crds[c + 4], crds[c + 5]);
                c += 6;
            }
            else if (cmd == "Q") {
                ctx.quadraticCurveTo(crds[c], crds[c + 1], crds[c + 2], crds[c + 3]);
                c += 4;
            }
            else if (cmd == "Z") {
                ctx.closePath();
            }
        }
    };

var ImageList = /** @class */ (function () {
    function ImageList(files) {
        if (files == null) {
            return;
        }
        this.images = {};
        for (var fileKey in files) {
            // let reg = new RegExp("xl/media/image1.png", "g");
            if (fileKey.indexOf("xl/media/") > -1) {
                var fileNameArr = fileKey.split(".");
                var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, "emf": 1 }) {
                    if (suffix == "emf") {
                        var pNum = 0; // number of the page, that you want to render
                        var scale = 1; // the scale of the document
                        var wrt = new ToContext2D(pNum, scale);
                        var inp, out, stt;
                        FromEMF.K = [];
                        inp = FromEMF.C;
                        out = FromEMF.K;
                        stt = 4;
                        for (var p in inp)
                            out[inp[p]] = p.slice(stt);
                        FromEMF.Parse(files[fileKey], wrt);
                        this.images[fileKey] = wrt.canvas.toDataURL("image/png");
                    }
                    else {
                        this.images[fileKey] = files[fileKey];
                    }
                }
            }
        }
    }
    ImageList.prototype.getImageByName = function (pathName) {
        if (pathName in this.images) {
            var base64 = this.images[pathName];
            return new Image(pathName, base64);
        }
        return null;
    };
    return ImageList;
}());
var Image = /** @class */ (function (_super) {
    __extends(Image, _super);
    function Image(pathName, base64) {
        var _this = _super.call(this) || this;
        _this.src = base64;
        return _this;
    }
    Image.prototype.setDefault = function () {
    };
    return Image;
}(LuckyImageBase));

var LuckyFile = /** @class */ (function (_super) {
    __extends(LuckyFile, _super);
    function LuckyFile(files, fileName) {
        var _this = _super.call(this) || this;
        _this.columnWidthSet = [];
        _this.rowHeightSet = [];
        _this.files = files;
        _this.fileName = fileName;
        _this.readXml = new ReadXml(files);
        _this.getSheetNameList();
        _this.sharedStrings = _this.readXml.getElementsByTagName("sst/si", sharedStringsFile);
        _this.calcChain = _this.readXml.getElementsByTagName("calcChain/c", calcChainFile);
        _this.styles = {};
        _this.styles["cellXfs"] = _this.readXml.getElementsByTagName("cellXfs/xf", stylesFile);
        _this.styles["cellStyleXfs"] = _this.readXml.getElementsByTagName("cellStyleXfs/xf", stylesFile);
        _this.styles["cellStyles"] = _this.readXml.getElementsByTagName("cellStyles/cellStyle", stylesFile);
        _this.styles["fonts"] = _this.readXml.getElementsByTagName("fonts/font", stylesFile);
        _this.styles["fills"] = _this.readXml.getElementsByTagName("fills/fill", stylesFile);
        _this.styles["borders"] = _this.readXml.getElementsByTagName("borders/border", stylesFile);
        _this.styles["clrScheme"] = _this.readXml.getElementsByTagName("a:clrScheme/a:dk1|a:lt1|a:dk2|a:lt2|a:accent1|a:accent2|a:accent3|a:accent4|a:accent5|a:accent6|a:hlink|a:folHlink", theme1File);
        _this.styles["indexedColors"] = _this.readXml.getElementsByTagName("colors/indexedColors/rgbColor", stylesFile);
        _this.styles["mruColors"] = _this.readXml.getElementsByTagName("colors/mruColors/color", stylesFile);
        _this.imageList = new ImageList(files);
        var numfmts = _this.readXml.getElementsByTagName("numFmt/numFmt", stylesFile);
        var numFmtDefaultC = JSON.parse(JSON.stringify(numFmtDefault));
        for (var i = 0; i < numfmts.length; i++) {
            var attrList = numfmts[i].attributeList;
            var numfmtid = getXmlAttibute(attrList, "numFmtId", "49");
            var formatcode = getXmlAttibute(attrList, "formatCode", "@");
            // console.log(numfmtid, formatcode);
            if (!(numfmtid in numFmtDefault)) {
                numFmtDefaultC[numfmtid] = formatcode;
            }
        }
        // console.log(JSON.stringify(numFmtDefaultC), numfmts);
        _this.styles["numfmts"] = numFmtDefaultC;
        return _this;
    }
    /**
    * @return All sheet name of workbook
    */
    LuckyFile.prototype.getSheetNameList = function () {
        var workbookRelList = this.readXml.getElementsByTagName("Relationships/Relationship", workbookRels);
        if (workbookRelList == null) {
            return;
        }
        var regex = new RegExp("worksheets/[^/]*?.xml");
        var sheetNames = {};
        for (var i = 0; i < workbookRelList.length; i++) {
            var rel = workbookRelList[i], attrList = rel.attributeList;
            var id = attrList["Id"], target = attrList["Target"];
            if (regex.test(target)) {
                if (target.startsWith('/xl')) {
                    sheetNames[id] = target.substr(1);
                }
                else {
                    sheetNames[id] = "xl/" + target;
                }
            }
        }
        this.sheetNameList = sheetNames;
    };
    /**
    * @param sheetName WorkSheet'name
    * @return sheet file name and path in zip
    */
    LuckyFile.prototype.getSheetFileBysheetId = function (sheetId) {
        // for(let i=0;i<this.sheetNameList.length;i++){
        //     let sheetFileName = this.sheetNameList[i];
        //     if(sheetFileName.indexOf("sheet"+sheetId)>-1){
        //         return sheetFileName;
        //     }
        // }
        return this.sheetNameList[sheetId];
    };
    /**
    * @return workBook information
    */
    LuckyFile.prototype.getWorkBookInfo = function () {
        var Company = this.readXml.getElementsByTagName("Company", appFile);
        var AppVersion = this.readXml.getElementsByTagName("AppVersion", appFile);
        var creator = this.readXml.getElementsByTagName("dc:creator", coreFile);
        var lastModifiedBy = this.readXml.getElementsByTagName("cp:lastModifiedBy", coreFile);
        var created = this.readXml.getElementsByTagName("dcterms:created", coreFile);
        var modified = this.readXml.getElementsByTagName("dcterms:modified", coreFile);
        this.info = new LuckyFileInfo();
        this.info.name = this.fileName;
        this.info.creator = creator.length > 0 ? creator[0].value : "";
        this.info.lastmodifiedby = lastModifiedBy.length > 0 ? lastModifiedBy[0].value : "";
        this.info.createdTime = created.length > 0 ? created[0].value : "";
        this.info.modifiedTime = modified.length > 0 ? modified[0].value : "";
        this.info.company = Company.length > 0 ? Company[0].value : "";
        this.info.appversion = AppVersion.length > 0 ? AppVersion[0].value : "";
    };
    /**
    * @return All sheet , include whole information
    */
    LuckyFile.prototype.getSheetsFull = function (isInitialCell) {
        if (isInitialCell === void 0) { isInitialCell = true; }
        var sheets = this.readXml.getElementsByTagName("sheets/sheet", workBookFile);
        var sheetList = {};
        for (var key in sheets) {
            var sheet = sheets[key];
            sheetList[sheet.attributeList.name] = sheet.attributeList["sheetId"];
        }
        this.sheets = [];
        var order = 0;
        for (var key in sheets) {
            var sheet = sheets[key];
            var sheetName = sheet.attributeList.name;
            var sheetId = sheet.attributeList["sheetId"];
            var rid = sheet.attributeList["r:id"];
            var sheetFile = this.getSheetFileBysheetId(rid);
            var drawing = this.readXml.getElementsByTagName("worksheet/drawing", sheetFile), drawingFile = void 0, drawingRelsFile = void 0;
            if (drawing != null && drawing.length > 0) {
                var attrList = drawing[0].attributeList;
                var rid_1 = getXmlAttibute(attrList, "r:id", null);
                if (rid_1 != null) {
                    drawingFile = this.getDrawingFile(rid_1, sheetFile);
                    drawingRelsFile = this.getDrawingRelsFile(drawingFile);
                }
            }
            if (sheetFile != null) {
                var sheet_1 = new LuckySheet(sheetName, sheetId, order, isInitialCell, {
                    sheetFile: sheetFile,
                    readXml: this.readXml,
                    sheetList: sheetList,
                    styles: this.styles,
                    sharedStrings: this.sharedStrings,
                    calcChain: this.calcChain,
                    imageList: this.imageList,
                    drawingFile: drawingFile,
                    drawingRelsFile: drawingRelsFile,
                });
                this.columnWidthSet = [];
                this.rowHeightSet = [];
                this.imagePositionCaculation(sheet_1);
                this.sheets.push(sheet_1);
                order++;
            }
        }
    };
    LuckyFile.prototype.extendArray = function (index, sets, def, hidden, lens) {
        if (index < sets.length) {
            return;
        }
        var startIndex = sets.length, endIndex = index;
        var allGap = 0;
        if (startIndex > 0) {
            allGap = sets[startIndex - 1];
        }
        // else{
        //     sets.push(0);
        // }
        for (var i = startIndex; i <= endIndex; i++) {
            var gap = def, istring = i.toString();
            if (istring in hidden) {
                gap = 0;
            }
            else if (istring in lens) {
                gap = lens[istring];
            }
            allGap += Math.round(gap + 1);
            sets.push(allGap);
        }
    };
    LuckyFile.prototype.imagePositionCaculation = function (sheet) {
        var images = sheet.images, defaultColWidth = sheet.defaultColWidth, defaultRowHeight = sheet.defaultRowHeight;
        var colhidden = {};
        if (sheet.config.colhidden) {
            colhidden = sheet.config.colhidden;
        }
        var columnlen = {};
        if (sheet.config.columnlen) {
            columnlen = sheet.config.columnlen;
        }
        var rowhidden = {};
        if (sheet.config.rowhidden) {
            rowhidden = sheet.config.rowhidden;
        }
        var rowlen = {};
        if (sheet.config.rowlen) {
            rowlen = sheet.config.rowlen;
        }
        for (var key in images) {
            var imageObject = images[key]; //Image, luckyImage
            var fromCol = imageObject.fromCol;
            var fromColOff = imageObject.fromColOff;
            var fromRow = imageObject.fromRow;
            var fromRowOff = imageObject.fromRowOff;
            var toCol = imageObject.toCol;
            var toColOff = imageObject.toColOff;
            var toRow = imageObject.toRow;
            var toRowOff = imageObject.toRowOff;
            var x_n = 0, y_n = 0;
            var cx_n = 0, cy_n = 0;
            if (fromCol >= this.columnWidthSet.length) {
                this.extendArray(fromCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
            }
            if (fromCol == 0) {
                x_n = 0;
            }
            else {
                x_n = this.columnWidthSet[fromCol - 1];
            }
            x_n = x_n + fromColOff;
            if (fromRow >= this.rowHeightSet.length) {
                this.extendArray(fromRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
            }
            if (fromRow == 0) {
                y_n = 0;
            }
            else {
                y_n = this.rowHeightSet[fromRow - 1];
            }
            y_n = y_n + fromRowOff;
            if (toCol >= this.columnWidthSet.length) {
                this.extendArray(toCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
            }
            if (toCol == 0) {
                cx_n = 0;
            }
            else {
                cx_n = this.columnWidthSet[toCol - 1];
            }
            cx_n = cx_n + toColOff - x_n;
            if (toRow >= this.rowHeightSet.length) {
                this.extendArray(toRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
            }
            if (toRow == 0) {
                cy_n = 0;
            }
            else {
                cy_n = this.rowHeightSet[toRow - 1];
            }
            cy_n = cy_n + toRowOff - y_n;
            console.log(defaultColWidth, colhidden, columnlen);
            console.log(fromCol, this.columnWidthSet[fromCol], fromColOff);
            console.log(toCol, this.columnWidthSet[toCol], toColOff, JSON.stringify(this.columnWidthSet));
            imageObject.originWidth = cx_n;
            imageObject.originHeight = cy_n;
            imageObject.crop.height = cy_n;
            imageObject.crop.width = cx_n;
            imageObject.default.height = cy_n;
            imageObject.default.left = x_n;
            imageObject.default.top = y_n;
            imageObject.default.width = cx_n;
        }
        console.log(this.columnWidthSet, this.rowHeightSet);
    };
    /**
    * @return drawing file string
    */
    LuckyFile.prototype.getDrawingFile = function (rid, sheetFile) {
        var sheetRelsPath = "xl/worksheets/_rels/";
        var sheetFileArr = sheetFile.split("/");
        var sheetRelsName = sheetFileArr[sheetFileArr.length - 1];
        var sheetRelsFile = sheetRelsPath + sheetRelsName + ".rels";
        var drawing = this.readXml.getElementsByTagName("Relationships/Relationship", sheetRelsFile);
        if (drawing.length > 0) {
            for (var i = 0; i < drawing.length; i++) {
                var relationship = drawing[i];
                var attrList = relationship.attributeList;
                var relationshipId = getXmlAttibute(attrList, "Id", null);
                if (relationshipId == rid) {
                    var target = getXmlAttibute(attrList, "Target", null);
                    if (target != null) {
                        return target.replace(/\.\.\//g, "");
                    }
                }
            }
        }
        return null;
    };
    LuckyFile.prototype.getDrawingRelsFile = function (drawingFile) {
        var drawingRelsPath = "xl/drawings/_rels/";
        var drawingFileArr = drawingFile.split("/");
        var drawingRelsName = drawingFileArr[drawingFileArr.length - 1];
        var drawingRelsFile = drawingRelsPath + drawingRelsName + ".rels";
        return drawingRelsFile;
    };
    /**
    * @return All sheet base information widthout cell and config
    */
    LuckyFile.prototype.getSheetsWithoutCell = function () {
        this.getSheetsFull(false);
    };
    /**
    * @return LuckySheet file json
    */
    LuckyFile.prototype.Parse = function () {
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
    };
    LuckyFile.prototype.toJsonString = function (file) {
        var LuckyOutPutFile = new LuckyFileBase();
        LuckyOutPutFile.info = file.info;
        LuckyOutPutFile.sheets = [];
        file.sheets.forEach(function (sheet) {
            var sheetout = new LuckySheetBase();
            //let attrName = ["name","color","config","index","status","order","row","column","luckysheet_select_save","scrollLeft","scrollTop","zoomRatio","showGridLines","defaultColWidth","defaultRowHeight","celldata","chart","isPivotTable","pivotTable","luckysheet_conditionformat_save","freezen","calcChain"];
            if (sheet.name != null) {
                sheetout.name = sheet.name;
            }
            if (sheet.color != null) {
                sheetout.color = sheet.color;
            }
            if (sheet.config != null) {
                sheetout.config = sheet.config;
                // if(sheetout.config._borderInfo!=null){
                //     delete sheetout.config._borderInfo;
                // }
            }
            if (sheet.id != null) {
                sheetout.id = sheet.id;
            }
            if (sheet.status != null) {
                sheetout.status = sheet.status;
            }
            if (sheet.order != null) {
                sheetout.order = sheet.order;
            }
            if (sheet.row != null) {
                sheetout.row = sheet.row;
            }
            if (sheet.column != null) {
                sheetout.column = sheet.column;
            }
            if (sheet.luckysheet_select_save != null) {
                sheetout.luckysheet_select_save = sheet.luckysheet_select_save;
            }
            if (sheet.scrollLeft != null) {
                sheetout.scrollLeft = sheet.scrollLeft;
            }
            if (sheet.scrollTop != null) {
                sheetout.scrollTop = sheet.scrollTop;
            }
            if (sheet.zoomRatio != null) {
                sheetout.zoomRatio = sheet.zoomRatio;
            }
            if (sheet.showGridLines != null) {
                sheetout.showGridLines = sheet.showGridLines;
            }
            if (sheet.defaultColWidth != null) {
                sheetout.defaultColWidth = sheet.defaultColWidth;
            }
            if (sheet.defaultRowHeight != null) {
                sheetout.defaultRowHeight = sheet.defaultRowHeight;
            }
            if (sheet.celldata != null) {
                // sheetout.celldata = sheet.celldata;
                sheetout.celldata = [];
                sheet.celldata.forEach(function (cell) {
                    var cellout = new LuckySheetCelldataBase();
                    cellout.r = cell.r;
                    cellout.c = cell.c;
                    cellout.v = cell.v;
                    sheetout.celldata.push(cellout);
                });
            }
            if (sheet.chart != null) {
                sheetout.chart = sheet.chart;
            }
            if (sheet.isPivotTable != null) {
                sheetout.isPivotTable = sheet.isPivotTable;
            }
            if (sheet.pivotTable != null) {
                sheetout.pivotTable = sheet.pivotTable;
            }
            if (sheet.luckysheet_conditionformat_save != null) {
                sheetout.luckysheet_conditionformat_save = sheet.luckysheet_conditionformat_save;
            }
            if (sheet.freezen != null) {
                sheetout.freezen = sheet.freezen;
            }
            if (sheet.calcChain != null) {
                sheetout.calcChain = sheet.calcChain;
            }
            if (sheet.images != null) {
                sheetout.images = sheet.images;
            }
            LuckyOutPutFile.sheets.push(sheetout);
        });
        return JSON.stringify(LuckyOutPutFile);
    };
    return LuckyFile;
}(LuckyFileBase));

var HandleZip = /** @class */ (function () {
    function HandleZip(file) {
        // Support nodejs fs to read files
        // if(file instanceof File){
        this.uploadFile = file;
        // }
    }
    HandleZip.prototype.unzipFile = function (successFunc, errorFunc) {
        // var new_zip:JSZip = new JSZip();
        JSZip__default['default'].loadAsync(this.uploadFile) // 1) read the Blob
            .then(function (zip) {
            var fileList = {}, lastIndex = Object.keys(zip.files).length, index = 0;
            zip.forEach(function (relativePath, zipEntry) {
                var fileName = zipEntry.name;
                var fileNameArr = fileName.split(".");
                var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                var fileType = "string";
                if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, }) {
                    fileType = "base64";
                }
                else if (suffix == "emf") {
                    fileType = "arraybuffer";
                }
                zipEntry.async(fileType).then(function (data) {
                    if (fileType == "base64") {
                        data = "data:image/" + suffix + ";base64," + data;
                    }
                    fileList[zipEntry.name] = data;
                    // console.log(lastIndex, index);
                    if (lastIndex == index + 1) {
                        successFunc(fileList);
                    }
                    index++;
                });
            });
        }, function (e) {
            errorFunc(e);
        });
    };
    HandleZip.prototype.unzipFileByUrl = function (url, successFunc, errorFunc) {
        var new_zip = new JSZip__default['default']();
        getBinaryContent(url, function (err, data) {
            if (err) {
                throw err; // or handle err
            }
            JSZip__default['default'].loadAsync(data).then(function (zip) {
                var fileList = {}, lastIndex = Object.keys(zip.files).length, index = 0;
                zip.forEach(function (relativePath, zipEntry) {
                    var fileName = zipEntry.name;
                    var fileNameArr = fileName.split(".");
                    var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                    var fileType = "string";
                    if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, }) {
                        fileType = "base64";
                    }
                    else if (suffix == "emf") {
                        fileType = "arraybuffer";
                    }
                    zipEntry.async(fileType).then(function (data) {
                        if (fileType == "base64") {
                            data = "data:image/" + suffix + ";base64," + data;
                        }
                        fileList[zipEntry.name] = data;
                        // console.log(lastIndex, index);
                        if (lastIndex == index + 1) {
                            successFunc(fileList);
                        }
                        index++;
                    });
                });
            }, function (e) {
                errorFunc(e);
            });
        });
    };
    HandleZip.prototype.newZipFile = function () {
        var zip = new JSZip__default['default']();
        this.workBook = zip;
    };
    //title:"nested/hello.txt", content:"Hello Worldasdfasfasdfasfasfasfasfasdfas"
    HandleZip.prototype.addToZipFile = function (title, content) {
        if (this.workBook == null) {
            var zip = new JSZip__default['default']();
            this.workBook = zip;
        }
        this.workBook.file(title, content);
    };
    return HandleZip;
}());

// //demo
// function demoHandler(){
//     let upload = document.getElementById("Luckyexcel-demo-file");
//     let selectADemo = document.getElementById("Luckyexcel-select-demo");
//     let downlodDemo = document.getElementById("Luckyexcel-downlod-file");
//     let mask = document.getElementById("lucky-mask-demo");
//     if(upload){
//         window.onload = () => {
//             upload.addEventListener("change", function(evt){
//                 var files:FileList = (evt.target as any).files;
//                 if(files==null || files.length==0){
//                     alert("No files wait for import");
//                     return;
//                 }
//                 let name = files[0].name;
//                 let suffixArr = name.split("."), suffix = suffixArr[suffixArr.length-1];
//                 if(suffix!="xlsx"){
//                     alert("Currently only supports the import of xlsx files");
//                     return;
//                 }
//                 LuckyExcel.transformExcelToLucky(files[0], function(exportJson:any, luckysheetfile:string){
//                     if(exportJson.sheets==null || exportJson.sheets.length==0){
//                         alert("Failed to read the content of the excel file, currently does not support xls files!");
//                         return;
//                     }
//                     console.log(exportJson, luckysheetfile);
//                     window.luckysheet.destroy();
//                     window.luckysheet.create({
//                         container: 'luckysheet', //luckysheet is the container id
//                         showinfobar:false,
//                         data:exportJson.sheets,
//                         title:exportJson.info.name,
//                         userInfo:exportJson.info.name.creator
//                     });
//                 });
//             });
//             selectADemo.addEventListener("change", function(evt){
//                 var obj:any = selectADemo;
//                 var index = obj.selectedIndex;
//                 var value = obj.options[index].value;
//                 var name = obj.options[index].innerHTML;
//                 if(value==""){
//                     return;
//                 }
//                 mask.style.display = "flex";
//                 LuckyExcel.transformExcelToLuckyByUrl(value, name, function(exportJson:any, luckysheetfile:string){
//                     if(exportJson.sheets==null || exportJson.sheets.length==0){
//                         alert("Failed to read the content of the excel file, currently does not support xls files!");
//                         return;
//                     }
//                     console.log(exportJson, luckysheetfile);
//                     mask.style.display = "none";
//                     window.luckysheet.destroy();
//                     window.luckysheet.create({
//                         container: 'luckysheet', //luckysheet is the container id
//                         showinfobar:false,
//                         data:exportJson.sheets,
//                         title:exportJson.info.name,
//                         userInfo:exportJson.info.name.creator
//                     });
//                 });
//             });
//             downlodDemo.addEventListener("click", function(evt){
//                 var obj:any = selectADemo;
//                 var index = obj.selectedIndex;
//                 var value = obj.options[index].value;
//                 if(value.length==0){
//                     alert("Please select a demo file");
//                     return;
//                 }
//                 var elemIF:any = document.getElementById("Lucky-download-frame");
//                 if(elemIF==null){
//                     elemIF = document.createElement("iframe");
//                     elemIF.style.display = "none";
//                     elemIF.id = "Lucky-download-frame";
//                     document.body.appendChild(elemIF);
//                 }
//                 elemIF.src = value;
//                 // elemIF.parentNode.removeChild(elemIF);
//             });
//         }
//     }
// }
// demoHandler();
// api
var LuckyExcel = /** @class */ (function () {
    function LuckyExcel() {
    }
    LuckyExcel.transformExcelToLucky = function (excelFile, callBack) {
        var handleZip = new HandleZip(excelFile);
        handleZip.unzipFile(function (files) {
            var luckyFile = new LuckyFile(files, excelFile.name);
            var luckysheetfile = luckyFile.Parse();
            var exportJson = JSON.parse(luckysheetfile);
            if (callBack != undefined) {
                callBack(exportJson, luckysheetfile);
            }
        }, function (err) {
            console.error(err);
        });
    };
    LuckyExcel.transformExcelToLuckyByUrl = function (url, name, callBack) {
        var handleZip = new HandleZip();
        handleZip.unzipFileByUrl(url, function (files) {
            var luckyFile = new LuckyFile(files, name);
            var luckysheetfile = luckyFile.Parse();
            var exportJson = JSON.parse(luckysheetfile);
            if (callBack != undefined) {
                callBack(exportJson, luckysheetfile);
            }
        }, function (err) {
            console.error(err);
        });
    };
    LuckyExcel.transformLuckyToExcel = function (LuckyFile, callBack) {
    };
    return LuckyExcel;
}());

module.exports = LuckyExcel;

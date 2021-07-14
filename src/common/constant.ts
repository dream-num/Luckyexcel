import {stringToNum, IattributeList, numTostring, IDataVerificationMap, IDataVerificationType2Map} from "../ICommon";

export const columeHeader_word:string[] = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

export const columeHeader_word_index:stringToNum = { 'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8, 'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16, 'R': 17, 'S': 18, 'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24, 'Z': 25 };

export const coreFile = "docProps/core.xml";
export const appFile = "docProps/app.xml";
export const contentTypesFile = "[Content_Types].xml";
export const workBookFile = "xl/workbook.xml";
export const calcChainFile = "xl/calcChain.xml";
export const stylesFile = "xl/styles.xml";
export const sharedStringsFile = "xl/sharedStrings.xml";
export const worksheetFilePath = "xl/worksheets/";
export const theme1File = "xl/theme/theme1.xml";
export const workbookRels= "xl/_rels/workbook.xml.rels";




//Excel Built-In cell type
export const ST_CellType:IattributeList ={
    "Boolean":"b",
    "Date":"d",
    "Error":"e",
    "InlineString":"inlineStr",
    "Number":"n",
    "SharedString":"s",
    "String":"str",
}

//Excel Built-In cell style
export const BuiltInCellStyles:IattributeList = {
    "0":"Normal",

}


export let numFmtDefault:IattributeList = {
    "0":'General',
    "1": '0',
    "2": '0.00',
    "3": '#,##0',
    "4": '#,##0.00',
    "9":'0%',
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
}

export const indexedColors:IattributeList = {
    "0":'00000000',
    "1":'00FFFFFF',
    "2":'00FF0000',
    "3":'0000FF00',
    "4":'000000FF',
    "5":'00FFFF00',
    "6":'00FF00FF',
    "7":'0000FFFF',
    "8":'00000000',
    "9":'00FFFFFF',
    "10":'00FF0000',
    "11":'0000FF00',
    "12":'000000FF',
    "13":'00FFFF00',
    "14":'00FF00FF',
    "15":'0000FFFF',
    "16":'00800000',
    "17":'00008000',
    "18":'00000080',
    "19":'00808000',
    "20":'00800080',
    "21":'00008080',
    "22":'00C0C0C0',
    "23":'00808080',
    "24":'009999FF',
    "25":'00993366',
    "26":'00FFFFCC',
    "27":'00CCFFFF',
    "28":'00660066',
    "29":'00FF8080',
    "30":'000066CC',
    "31":'00CCCCFF',
    "32":'00000080',
    "33":'00FF00FF',
    "34":'00FFFF00',
    "35":'0000FFFF',
    "36":'00800080',
    "37":'00800000',
    "38":'00008080',
    "39":'000000FF',
    "40":'0000CCFF',
    "41":'00CCFFFF',
    "42":'00CCFFCC',
    "43":'00FFFF99',
    "44":'0099CCFF',
    "45":'00FF99CC',
    "46":'00CC99FF',
    "47":'00FFCC99',
    "48":'003366FF',
    "49":'0033CCCC',
    "50":'0099CC00',
    "51":'00FFCC00',
    "52":'00FF9900',
    "53":'00FF6600',
    "54":'00666699',
    "55":'00969696',
    "56":'00003366',
    "57":'00339966',
    "58":'00003300',
    "59":'00333300',
    "60":'00993300',
    "61":'00993366',
    "62":'00333399',
    "63":'00333333',
    "64":null,//system Foreground n/a
    "65":null,//system Background n/a
}

export const OEM_CHARSET:IattributeList = {
    "0": "ANSI_CHARSET",
    "1": "DEFAULT_CHARSET",
    "2": "SYMBOL_CHARSET",
    "77": "MAC_CHARSET",
    "128": "SHIFTJIS_CHARSET",
    "129": "HANGUL_CHARSET",
    "130": "JOHAB_CHARSET",
    "134": "GB2312_CHARSET",
    "136": "CHINESEBIG5_CHARSET",
    "161": "GREEK_CHARSET",
    "162": "TURKISH_CHARSET",
    "163": "VIETNAMESE_CHARSET",
    "177": "HEBREW_CHARSET",
    "178": "ARABIC_CHARSET",
    "186": "BALTIC_CHARSET",
    "204": "RUSSIAN_CHARSET",
    "222": "THAI_CHARSET",
    "238": "EASTEUROPE_CHARSET",
    "255": "OEM_CHARSET",
}


export const borderTypes:stringToNum = {
    "none":0,
    "thin":1, 
    "hair":2, 
    "dotted":3, 
    "dashed":4, 
    "dashDot":5, 
    "dashDotDot":6, 
    "double":7, 
    "medium":8, 
    "mediumDashed":9, 
    "mediumDashDot":10, 
    "mediumDashDotDot":11, 
    "slantDashDot":12, 
    "thick":13
}


export let numFmtDefaultMap: IattributeList = {
    "yyyy/m/d;@": "yyyy/MM/dd",
    "yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;;@":
      "yyyy&quot;年&quot;MM&quot;月&quot;dd&quot;日&quot;",
    "[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@": "yyyy/MM/dd hh:mm AM/PM",
};

export const fontFamilys:IattributeList = {
    "0":"defualt",
    "1":"Roman",
    "2":"Swiss",
    "3":"Modern",
    "4":"Script",
    "5":"Decorative"
}

export const DATA_VERIFICATION_MAP: IDataVerificationMap = {
    list: "dropdown",
    whole: "number_integer",
    decimal: "number_decimal",
    custom: "text_content",
    textLength: "text_length",
    date: "date",
    "unknown1": "number", // no match yet
    "unknown2": "checkbox", // no match yet
    "unknown3": "validity", // no match yet
};

export const COMMON_TYPE2: string[] = [
    "number",
    "number_integer",
    "number_decimal",
    "text_length",
];

export const DATA_VERIFICATION_TYPE2_MAP: IDataVerificationType2Map = {
    common: {
      between: "bw",
      notBetween: "nb",
      equal: "eq",
      notEqualTo: "ne",
      moreThanThe: "gt",
      lessThan: "lt",
      greaterOrEqualTo: "gte",
      lessThanOrEqualTo: "lte",
    },
    text_content: {
      include: "include",
      exclude: "exclude",
      equal: "equal",
    },
    date: {
      between: "bw",
      notBetween: "nb",
      equal: "eq",
      notEqualTo: "ne",
      earlierThan: "bf",
      noEarlierThan: "nbf",
      laterThan: "af",
      noLaterThan: "naf",
    },
    validity: {
      card: "card",
      phone: "phone",
    },
};

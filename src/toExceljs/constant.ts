import Excel from 'exceljs'
import { stringToNum, IattributeList, stringToBoolean } from "../ICommon";

export const CellTypeMap: IattributeList = {
  "Boolean": "b",
  "Date": "d",
  "Error": "e",
  "InlineString": "inlineStr",
  "Number": "n",
  "SharedString": "s",
  "String": "str",
}

export const FontFamilyMap: Record<string, string> = {
  '0': '微软雅黑',
  '1': '宋体（Song）',
  '2': '黑体（ST Heiti）',
  '3': '楷体（ST Kaiti）',
  '4': '仿宋（ST FangSong）',
  '5': '新宋体（ST Song）',
  '6': '华文新魏',
  '7': '华文行楷',
  '8': '华文隶书',
  '9': 'Arial',
  '10': 'Times New Roman ',
  '11': 'Tahoma ',
  '12': 'Verdana',
}

export const ErrorValueMap: Record<string, Excel.CellErrorValue['error']> = {
  '#N/A': Excel.ErrorValue.NotApplicable,
  '#REF!': Excel.ErrorValue.Ref,
  '#NAME?': Excel.ErrorValue.Name,
  '#DIV/0!': Excel.ErrorValue.DivZero,
  '#NULL!': Excel.ErrorValue.Null,
  '#VALUE!': Excel.ErrorValue.Value,
  '#NUM!': Excel.ErrorValue.Num,
}

export const excelBorderPositions: IattributeList = {
  t: "top",
  b: "bottom",
  l: "left",
  r: "right",
}

export const verticalMap: Record<string, Excel.Alignment['vertical']> = {
  '0': 'middle',
  '1': 'top',
  '2': 'bottom',
}

export const horizontalMap: Record<string, Excel.Alignment['horizontal']> = {
  '0': 'center',
  '1': 'left',
  '2': 'right',
}

export const wrapTextMap: Record<string, Excel.Alignment['wrapText']> = {
  '0': false,
  '1': false,
  '2': true,
}

export const textRotationMap: Record<string, Excel.Alignment['textRotation']> = {
  '0': 0,
  '1': 45,
  '2': -45,
  '3': 'vertical',
  '4': 90,
  '5': -90,
}

export const excelBorderStyles: IattributeList = {
  '1': 'thin',
  '2': 'hair',
  '3': 'dotted',
  '4': 'dashDot', // dashed exceljs不支持
  '5': 'dashDot',
  '6': 'dashDotDot',
  '7': 'double',
  '8': 'medium',
  '9': 'mediumDashed',
  '10': 'mediumDashDot',
  '11': 'mediumDashDotDot',
  '12': 'slantDashDot',
  '13': 'thick',
}
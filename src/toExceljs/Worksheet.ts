import { WorkCell } from './Workcell'
import {BorderInfo} from './BorderInfo'
import Excel from 'exceljs'
import {chatatABC, encodeRange} from '../common/method'
import { IluckySheetCelldata } from '../ToLuckySheet/ILuck'
interface IMergeCell {
  e: { c: number; r: number },
  s: { c: number; r: number }
}
export class WorkSheetBase {
  worksheet: Excel.Worksheet;
  data: Record<string, any>;
  border: Record<string, any>;
  hyperlink: Record<string, any>;
  ref: string;
  cols: any[];
  rows: any[];
  name: string;
  merges: IMergeCell[];
}
export class WorkSheet extends WorkSheetBase {
  constructor(workbook: Excel.Workbook, sheetData: any) {
    super()
    const { config, column: maxCol, row: maxRow, hyperlink, celldata, name, data, hide, defaultColWidth, defaultRowHeight, color } = sheetData
    const { columnlen, rowlen, merge, borderInfo } = config || {}
    this.data = {}
    this.border = {}
    this.hyperlink = hyperlink
    this.name = name

    this.worksheet = workbook.addWorksheet(name, {
      state: hide === 1 ? 'hidden' : 'visible',
      properties: {
        tabColor: color === '' ? undefined : { argb: color },
      }
    })

    // 解析数据
    celldata?.forEach((data: IluckySheetCelldata) => {
      const cell = new WorkCell(this, data)
      this.data[cell.key] = cell.cell
    })

    // 规定行列宽高
    if (columnlen) {
      Object.keys(columnlen).forEach((_col) => {
        if (parseInt(_col) > maxCol) {
          return;
        }
        const colKey = chatatABC(parseInt(_col))
        const col = this.worksheet.getColumn(colKey);
        col.width =  Math.round(parseFloat(columnlen?.[_col] ?? 70) * 0.17)
      })
    }
    if (rowlen) {
      Object.keys(rowlen).forEach((_row) => {
        if (parseInt(_row) > maxRow) {
          return;
        }
        const row = this.worksheet.getRow(parseInt(_row) + 1);
        row.height = Math.round(parseFloat(rowlen?.[_row] ?? 20) * 0.5)
      })
    }

    // 合并单元格
    if (merge) {
      Object.keys(merge).forEach((key) => {
        // rs cs 为所占行数
        const {r, c, rs, cs} = merge[key]
        // s为扩展的offset
        const range = encodeRange({r, c}, {r: r + rs - 1, c: c + cs - 1})
        this.worksheet.mergeCells(range);
      })
    }

    // 解析borderInfo
    const border = new BorderInfo(borderInfo)
    for (const [key, borderValue] of Object.entries(border.data)) {
      this.border[key] = borderValue
      this.worksheet.getCell(key).border = borderValue
    }

  }

  toData() {
    const result = Object.assign({}, this.data)
    return result
  }
}
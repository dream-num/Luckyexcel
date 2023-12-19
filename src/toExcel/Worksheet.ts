import XLSX from 'xlsx-js-style';
import {WorkCell} from './Workcell'
import {BorderInfo} from './BorderInfo'
import {IluckySheetCelldata} from '../ToLuckySheet/ILuck'
interface IMergeCell {
  e: { c: number; r: number },
  s: { c: number; r: number }
}
export class WorkSheetBase {
  data: Record<string, any>;
  border: Record<string, any>;
  ref: string;
  cols: any[];
  rows: any[];
  name: string;
  merges: IMergeCell[];
}
export class WorkSheet extends WorkSheetBase {
  constructor(sheetData: any) {
    super()
    const { config, column, row, celldata, name, data } = sheetData
    const { columnlen, rowlen, merge, borderInfo } = config || {}
    this.data = {}
    this.border = {}
    this.name = name

    // 解析borderInfo
    const border = new BorderInfo(borderInfo)
    for (const [key, borderValue] of Object.entries(border.data)) {
      this.data[key] = {}
      this.border[key] = borderValue
    }

    // 解析单元格内容
    // const rowLen = data.length
    // const columnLen = data?.[0]?.length || 0
    // for(let row = 0; row <= rowLen; row++){
    //   for(let col = 0; col <= columnLen; col++){
    //     const cellValue = data[row][col] || {r: row, c: col}
    //     const cell = new WorkCell(cellValue, border)
    //     this.data[cell.key] = cell.toData()
    //   }
    // }
    celldata?.forEach((data: IluckySheetCelldata) => {
      const cell = new WorkCell(data)
      if (this.border[cell.key]) {
        this.data[cell.key] = {
          ...(cell.toData() || {}),
          s: {
            ...(cell.toData()?.s || {}),
            border: this.border[cell.key] || {}
          }
        }
      } else {
        this.data[cell.key] = cell.toData()
      }
    })

    // 规定数据范围
    const endCellRange = XLSX.utils.encode_cell({ c: column - 1, r: row - 1 })
    this.ref = `A1:${endCellRange}`

    // 规定行列宽高
    this.cols = [];
    if (columnlen) {
      Object.keys(columnlen).forEach((index) => {
        this.cols[Number(index)] = { wpx: Math.round(parseFloat(columnlen[index]) * 0.75) }
      })
    }
    this.rows = [];
    if (rowlen) {
      Object.keys(rowlen).forEach((index) => {
        this.rows[Number(index)] = { hpx: Math.round(parseFloat(rowlen[index]) * 0.75) }
      })
    }

    // 合并单元格范围
    this.merges = []
    if (merge) {
      Object.keys(merge).forEach((key) => {
        // rs cs 为所占行数
        const {r, c, rs, cs} = merge[key]
        // s为合并单元格开始行列  e为结束
        const mergeValue = {
          s: { r, c },
          e: { r: r + ( rs - 1), c: c + ( cs - 1) }
        }
        this.merges.push(mergeValue)
      })
    }
  }

  toData() {
    const result = Object.assign({}, this.data)
    result['!ref'] = this.ref
    result['!cols'] = this.cols
    result['!rows'] = this.rows
    result['!merges'] = this.merges
    return result
  }
}
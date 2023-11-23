import XLSX from 'xlsx-js-style';
import { excelBorderPositions } from '../common/constant'
import { rgbToHex } from '../common/method'
import { excelBorderStyles } from './constant'

export class BorderInfoBase {
  data: Record<string, any>
}

export class BorderInfo extends BorderInfoBase {
  constructor(data: any) {
    super()
    this.data = {}
    data?.forEach((item: any) => {
      const { rangeType } = item;
      // 旧版excel导入时生成的borderinfo
      if (rangeType === 'cell') {
        const { value } = item;
        const cellKey = XLSX.utils.encode_cell({ c: value.col_index, r: value.row_index });
        this.data[cellKey] = {};
        ['t', 'b', 'l', 'r'].forEach((p) => {
          const borderStyle = value[p]
          if (borderStyle && excelBorderPositions[p]) {
            this.data[cellKey][excelBorderPositions[p]] = {
              style: excelBorderStyles[borderStyle.style] || 'thin',
              color: { argb: rgbToHex(borderStyle.color)?.replace('#', '') || '000000' }
            }
          }
        })
      }
      // 新版luckysheet操作生成的borderinfo
      if (rangeType === 'range') {
        const { color, style, range = [], borderType } = item
        const borderStyle = {
          style: excelBorderStyles[style] || 'thin',
          color: { argb: rgbToHex(color)?.replace('#', '') || '000000' }
        }

        for (let rangeItem of range) {
          const [rangeStartRow, rangeEndRow] = rangeItem.row
          const [rangeStartCol, rangeEndCol] = rangeItem.column

          for (let row = rangeStartRow; row <= rangeEndRow; row++) {
            for (let col = rangeStartCol; col <= rangeEndCol; col++) {
              const cellKey = XLSX.utils.encode_cell({ c: col, r: row });
              this.data[cellKey] = {}
              // 检查上下左右边的情况，按需添加边框
              if (borderType === 'border-none') {
                this.data[cellKey] = null
              } else if (borderType === 'border-all') {
                this.data[cellKey].top = borderStyle;
                this.data[cellKey].bottom = borderStyle;
                this.data[cellKey].left = borderStyle;
                this.data[cellKey].right = borderStyle;
              } else {
                const isOnTop = isOnTopBorder({ row: row, column: col }, rangeItem.row, rangeItem.column)
                const isOnBottom = isOnBottomBorder({ row: row, column: col }, rangeItem.row, rangeItem.column)
                const isOnRight = isOnRightBorder({ row: row, column: col }, rangeItem.row, rangeItem.column)
                const isOnLeft = isOnLeftBorder({ row: row, column: col }, rangeItem.row, rangeItem.column)

                if (['border-top', 'border-outside'].includes(borderType) && isOnTop) {
                  this.data[cellKey].top = borderStyle;
                }
                if (['border-bottom', 'border-outside'].includes(borderType) && isOnBottom) {
                  this.data[cellKey].bottom = borderStyle;
                }
                if (['border-right', 'border-outside'].includes(borderType) && isOnRight) {
                  this.data[cellKey].right = borderStyle;
                }
                if (['border-left', 'border-outside'].includes(borderType) && isOnLeft) {
                  this.data[cellKey].left = borderStyle;
                }
                if (borderType === 'border-inside') {
                  !isOnTop && (this.data[cellKey].top = borderStyle)
                  !isOnBottom && (this.data[cellKey].bottom = borderStyle)
                  !isOnRight && (this.data[cellKey].right = borderStyle)
                  !isOnLeft && (this.data[cellKey].left = borderStyle)
                }
                if (borderType === 'border-horizontal') {
                  !isOnTop && (this.data[cellKey].top = borderStyle)
                  !isOnBottom && (this.data[cellKey].bottom = borderStyle)
                }
                if (borderType === 'border-vertical') {
                  !isOnRight && (this.data[cellKey].right = borderStyle)
                  !isOnLeft && (this.data[cellKey].left = borderStyle)
                }
              }
            }
          }
        }
      }
    })
  }

  get(key: string) {
    return this.data[key]
  }
}

type Cell = {
  row: number;
  column: number;
}

type Range = [number, number]

function isOnTopBorder(cell: Cell, rowRange: Range, columnRange: Range) {
  const { row: cellRow, column: cellCol } = cell;
  const [rangeStartRow, rangeEndRow] = rowRange
  const [rangeStartCol, rangeEndCol] = columnRange
  return cellRow === rangeStartRow
}

function isOnBottomBorder(cell: Cell, rowRange: Range, columnRange: Range) {
  const { row: cellRow, column: cellCol } = cell;
  const [rangeStartRow, rangeEndRow] = rowRange
  const [rangeStartCol, rangeEndCol] = columnRange
  return cellRow === rangeEndRow
}

function isOnRightBorder(cell: Cell, rowRange: Range, columnRange: Range) {
  const { row: cellRow, column: cellCol } = cell;
  const [rangeStartRow, rangeEndRow] = rowRange
  const [rangeStartCol, rangeEndCol] = columnRange
  return cellCol === rangeEndCol
}

function isOnLeftBorder(cell: Cell, rowRange: Range, columnRange: Range) {
  const { row: cellRow, column: cellCol } = cell;
  const [rangeStartRow, rangeEndRow] = rowRange
  const [rangeStartCol, rangeEndCol] = columnRange
  return cellCol === rangeStartCol
}
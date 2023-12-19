import Excel from 'exceljs'
import { WorkSheet } from './Worksheet'
import { rgbToHex, chatatABC, encodeCell } from '../common/method'
import { IluckySheetCelldata, IluckySheetCelldataValue } from '../ToLuckySheet/ILuck'
import { FontFamilyMap, ErrorValueMap, verticalMap, horizontalMap, wrapTextMap, textRotationMap } from './constant'

export class WorkCellBase {
  key: string
  worksheet: WorkSheet
  cell: Excel.Cell;
}

export class WorkCell extends WorkCellBase {
  constructor(worksheet: WorkSheet, data: IluckySheetCelldata) {
    super()
    const { r, c } = data

    this.key = encodeCell({ c, r })
    this.worksheet = worksheet
    this.cell = worksheet.worksheet.getCell(this.key)

    this._parseValue(data)
    this.cell.font = this._parseFont(data?.v)
    this._parseStyle(data?.v)
  }

  _parseValue(data: IluckySheetCelldata) {
    const { r, c } = data
    if (data?.v === undefined) {
      return;
    }
    if (typeof data?.v === 'string' || data?.v === null) {
      this.cell.value = data.v as string | null;
      return
    }
    const { ct, v, f } = data.v as IluckySheetCelldataValue
    if (ct) {
      // 自动格式、数字、字符串、布尔值 直接赋值
      if (['g', 'n', 's', 'b'].includes(ct?.t)) {
        this.cell.value = v
        this.cell.numFmt = ct.fa
      }
      // 日期格式
      if (ct?.t === 'd') {
        const date0 = new Date(0);
        const utcOffset = date0.getTimezoneOffset();
        const cellValue = new Date(0, 0, parseInt(v) - 1, 0, -utcOffset, 0);
        this.cell.value = cellValue
        this.cell.numFmt = ct.fa
      }

      // 公式
      if (f) {
        this.cell.value = {
          formula: f,
          result: v
        }
      }

      // 富文本
      if (ct?.t === 'inlineStr' && ct?.s?.length > 0) {
        this.cell.value = {
          richText: ct?.s?.map((item: IluckySheetCelldataValue) => {
            return {
              text: item.v,
              font: this._parseFont(item)
            }
          })
        };
      }

      // // 错误值
      if (ct?.t === 'e' && !f) {
        this.cell.value = { error: ErrorValueMap[v] || Excel.ErrorValue.Value };
      }
    }

    // 超链接
    if (this.worksheet.hyperlink?.[`${r}_${c}`]) {
      const hyperlink = this.worksheet.hyperlink?.[`${r}_${c}`]
      let linkAddress = hyperlink.linkAddress
      if (hyperlink.linkType === 'internal') {
        const [sheet, cell] = hyperlink.linkAddress?.split('!')
        linkAddress = `#'${sheet}'!${cell}`
      }
      this.cell.value = {
        text: v,
        hyperlink: linkAddress,
        tooltip: hyperlink.linkTooltip
      };
    }
  }

  // 格式化字体样式
  _parseFont(data: IluckySheetCelldataValue | string) {
    if (typeof data === 'string' || data === undefined) {
      return {};
    }
    const { fc, bl, it, cl, un, fs, ff } = data || {}
    return {
      name: FontFamilyMap[ff] || ff,
      size: fs ?? 10,
      family: 1,
      color: { argb: rgbToHex(fc)?.replace('#', '') || '000000' },
      bold: !!bl,
      italic: !!it,
      underline: !!un,
      strike: !!cl,
    }
  }

  _parseStyle(data: IluckySheetCelldataValue | string) {
    if (typeof data === 'string' || data === undefined) {
      return {};
    }
    const { vt, ht, tb, tr, ti, bg } = data || {}
    // 对齐
    const alignment: Partial<Excel.Alignment> = {}
    alignment.vertical = verticalMap[vt] || 'middle'
    alignment.horizontal = horizontalMap[ht] || 'left'
    tb !== undefined && (alignment.wrapText = wrapTextMap[tb])
    ti !== undefined && (alignment.indent = ti ?? 0)
    tr !== undefined && (alignment.textRotation = textRotationMap[tr] ?? 0)
    this.cell.alignment = alignment

    // 背景色
    if (bg) {
      this.cell.fill = {
        type: 'pattern',
	      pattern: 'solid',
        fgColor:{argb:rgbToHex(bg)?.replace('#', '') || 'ffffff'}
      }
    }

  }

  toData() {
    const data = this
    return data
  }
}
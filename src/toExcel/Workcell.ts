import XLSX from 'xlsx-js-style';
import {verticalMap, horizontalMap, wrapTextMap, textRotationMap, ErrorValueMap} from '../common/constant'
import {rgbToHex} from '../common/method'

export class WorkCellBase {
  key: string
  data: any
}

export class WorkCell extends WorkCellBase {
  constructor(data: any) {
    super()
    const {r, c} = data
    const {ct, v, m, f, vt, ht, tb, tr, bg, fc, bl, it, cl, un, fs, ff} = data.v || {}

    this.key = XLSX.utils.encode_cell({ c, r })
    this.data = {v: '', t: 's', s: {}}

    if (ct) {
      if (v) {
        // 错误时 值有特定mapping
        if (ErrorValueMap[v] !== undefined) {
          this.data.v = ErrorValueMap[v]
        } else {
          this.data.v = v
        }
        
        this.data.t = ct.t
      } else if (Array.isArray(ct?.s)) {
        // inline string格式的简易处理
        this.data.t = 's'
        this.data.v = ct?.s?.reduce((prev: any, cur: { v: any; }) => prev + cur.v, '');
      }

      // 数字 & 错误
      if (ct?.t === 'n' && ct?.t === 'e') {
        this.data.t = ct.t
        this.data.z = ct.fa
        f && (this.data.f = f)
      }

      // 日期
      if (ct?.t === 'd') {
        // 如果类型规定为d 解析日期不对
        this.data.t = 'n'
        this.data.z = 'm/d/yy'
        this.data.s = {numFmt: 'm/d/yy'}
      }

      // 字符串
      if (ct?.t === 'g') {
        this.data.t = 's'
        this.data.w = ct.fa
      }

    }

    // 对齐方式 换行 文字方向
    if ([vt, ht, tb, tr].some((item) => item !== undefined)) {
      this.data.s.alignment = {}
      vt !== undefined && (this.data.s.alignment.vertical = verticalMap[vt])
      ht !== undefined && (this.data.s.alignment.horizontal = horizontalMap[ht])
      tb !== undefined && (this.data.s.alignment.wrapText = wrapTextMap[tb])
      tr !== undefined && (this.data.s.alignment.textRotation = textRotationMap[tr])
    }

    // 背景色
    if (bg) {
      this.data.s.fill = {
        fgColor: {rgb: rgbToHex(bg)?.replace('#', '') || 'ffffff'}
      }
    }

    // 文本样式
    if ([fc, bl, it, cl, un, fs].some((item) => item !== undefined)) {
      this.data.s.font = {}
      fc !== undefined && (this.data.s.font.color = {rgb: rgbToHex(fc)?.replace('#', '') || '000000'})
      bl !== undefined && (this.data.s.font.bold = !!bl)
      bl !== undefined && (this.data.s.font.italic = !!it)
      cl !== undefined && (this.data.s.font.strike = !!cl)
      un !== undefined && (this.data.s.font.underline = !!un)
      fs !== undefined && (this.data.s.font.sz = fs)
      // 字体种类可能缺失造成不匹配 暂时去掉
      // ff !== undefined && (this.data.s.font.name = ff)
    }
  }

  toData() {
    const data = this.data
    // const borderInfo = this.border.get(this.key)
    // if (borderInfo && Object.keys(borderInfo).length !== 0) {
    //   data.s = {border: {...borderInfo}}
    // }
    return data
  }
}
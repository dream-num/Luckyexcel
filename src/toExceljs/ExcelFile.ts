import Excel from 'exceljs'
import { WorkSheet } from './Worksheet';

export class ExcelFileBase {
  fileName: string
  sheets: any[]
  workbook: Excel.Workbook
}

export class ExcelFile extends ExcelFileBase {
  /**
   * @param luckysheetJson luckysheet.toJson() 返回数据
   * @param fileName 
   */
  constructor(luckysheetJson: any, fileName?: string) {
    super();
    const { title, data } = luckysheetJson
    this.fileName = fileName || luckysheetJson.title
    this.workbook = new Excel.Workbook()
    this.sheets = data?.map((item: any) => new WorkSheet(this.workbook, item))
  }

  async export() {
    const buffer = await this.workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = this.fileName.indexOf('.xlsx') !== -1 ? this.fileName : `${this.fileName}.xlsx`;
    link.click();
    URL.revokeObjectURL(url);

    return this.workbook
  }
}
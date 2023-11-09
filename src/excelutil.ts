function createUtil(sheetName: string, excelRuner: ExcelRunner):ExcelUtil{
  return new ExcelUtil(sheetName, excelRuner);
}

export class ExcelUtil {
    excelRuner: ExcelRunner;
    sheetName: string = "";
    constructor(sheetName: string, excelRuner: ExcelRunner) {
      this.sheetName = sheetName;
      this.excelRuner = excelRuner;
      
    }
    async replaceSheet() {
      let sheet = await this.getSheet();
      if (!!sheet) {
        await this.excelRuner(async (context: any) => {
          sheet = context.workbook.worksheets.getItem(this.sheetName);
          sheet.delete();
          await context.sync();
        });
      }
      //await this.ensureSheet();
    }
    async getSheet(): Promise<any> {
      let sheet: any;
      await this.excelRuner(async (context: any) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
  
        const idx = sheets.items.findIndex((sh: any) => sh.name == this.sheetName);
        if (idx == -1) {
          return;
        }
        sheet = context.workbook.worksheets.getItem(this.sheetName);
      });
      return sheet;
    }
    async ensureSheet(): Promise<any> {
      const sheet = await this.getSheet();
      if (!sheet) {
        await this.excelRuner(async (context:any) => {
          context.workbook.worksheets.add(this.sheetName);
          await context.sync();
          return context.workbook.worksheets.getItem(this.sheetName);
        });
        return sheet;
      }
    }
    async readRows(range: string): Promise<any> {
      const vals: any[][] = await this.readValues(range);
      const rowVal = vals.map((row) => row[0]);
      return rowVal;
    }
    async readCols(range: string): Promise<any> {
      const vals: any[][] = await this.readValues(range);
      if (!vals) return;
      if (vals.length <= 0) return;
      return vals[0];
    }
    async readCell(range: string): Promise<any> {
      const vals: any[][] = await this.readValues(range);
      if (!vals) return;
      if (vals.length <= 0) return;
      return vals[0][0];
    }
  
    async readValues(range: string): Promise<any> {
      this.ensureSheet();
      let val: any;
      await this.excelRuner(async (context:any) => {
        const sheet = context.workbook.worksheets.getItem(this.sheetName);
        const filterRange = sheet.getRange(range);
        filterRange.load(["values"]);
        await context.sync();
        val = filterRange.values;
      });
      return val;
    }
  
    async writeValues(values: any[][], baseCell: string) {
      if (!values) return;
      if (values.length <= 0) return;
      const sheet = await this.ensureSheet();
      await this.excelRuner(async (context:any) => {
        const sheet = context.workbook.worksheets.getItem(this.sheetName);
        const range = ExcelUtil.calcRange(values, baseCell);
        const filterRange = sheet.getRange(range);
        filterRange.values = values;
        await context.sync();
      });
    }
    async writeCell(value: any, baseCell: string) {
      if (!value) return;
      await this.writeValues([[value]], baseCell);
    }
    async writeRows(values: any[], baseCell: string) {
      const newValues: any[][] = new Array();
      for (const val of values) {
        newValues.push([val]);
      }
      await this.writeValues(newValues, baseCell);
    }
    async writeCols(values: any[], baseCell: string) {
      const newValues: any[][] = new Array();
      newValues.push(values);
      await this.writeValues(newValues, baseCell);
    }
    startWriteSession(values: any[][], baseCell: string): WriteSessionChain {
      this.writeSessionTemp = new Array();
      return this.addWriteChain(values, baseCell);
    }
    addWriteChain(values: any[][], baseCell: string): WriteSessionChain {
      this.writeSessionTemp.push({ values: values, baseCell: baseCell });
      return this;
    }
    static calcRange(values: any[][], baseCell: string): string {
      const rangeWidth = values[0].length;
      const baseCol = (baseCell.match(/[a-zA-Z]/g)||[])[0] || "";
      const baseRowNum = parseInt(baseCell.replace(baseCol, ""));
      const baseColNum = ExcelUtil.toColumnNumber(baseCol);
      const endColNum = baseColNum + rangeWidth - 1;
      const endRowNum = baseRowNum + values.length - 1;
      const endColName = ExcelUtil.toColumnName(endColNum);
      const range = `${baseCell}:${endColName}${endRowNum}`;
      return range;
    }
    writeSessionTemp: WriteSession[] = Array();
    async sessionWrite(): Promise<void> {
      const sheet = await this.ensureSheet();
      await this.excelRuner(async (context: any) => {
        const sheet = context.workbook.worksheets.getItem(this.sheetName);
        const ranges: any[] = new Array();
        for (const session of this.writeSessionTemp) {
          const range = ExcelUtil.calcRange(session.values, session.baseCell);
          const sheetRange = sheet.getRange(range);
          sheetRange.values = session.values;
          ranges.push(sheetRange);
        }
        await context.sync();
        this.writeSessionTemp = new Array();
      });
    }
    static toColumnName(index: number): string {
      const ordA = "a".charCodeAt(0);
      const ordZ = "z".charCodeAt(0);
      const len = ordZ - ordA + 1;
      let s = "";
      let n: number = index;
      while (n >= 0) {
        s = String.fromCharCode((n % len) + ordA) + s;
        n = Math.floor(n / len) - 1;
      }
      return s.toLocaleUpperCase();
    }
    static toColumnNumber(val: string): number {
      val = val.toLocaleUpperCase();
      let base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
        i,
        j,
        result = 0;
      for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {
        result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
      }
      return result - 1;
    }
  }
  export type SettingsChain = {
    addSettings: { (key: string, val: any): SettingsChain };
    write: { ():void };
  };
  export type WriteSessionChain = {
    addWriteChain: { (values: any[][], baseCell: string): WriteSessionChain };
    sessionWrite: { (): Promise<void> };
  };
 export type WriteSession = {
    values: any[][];
    baseCell: string;
  };
 export type ExcelRunner = {(context: any):Promise<void>};
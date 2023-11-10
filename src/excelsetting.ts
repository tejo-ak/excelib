//import { ExcelUtil, ExcelRunner } from "./excelutil";
export class ExcelSetting {
    sheetName: string;
    baseCell: string;
    excelUtil: ExcelUtil;
    paramKeys: any;
    constructor(sheetName: string, paramKeys: any,runner:ExcelRunner, baseCell: string = "A2") {
      this.sheetName = sheetName;
      this.baseCell = baseCell;
      this.paramKeys = paramKeys;
      this.excelUtil = new ExcelUtil(sheetName, runner);
    }
    settingsTemp: Map<string, any> = new Map();
    writeSettings(key: string, val: any): SettingsChain {
      this.settingsTemp = new Map();
      return this.addSettings(key, val);
    }
    addSettings(key: string, val: any): SettingsChain {
      if (!key) return this;
      if (!this.paramKeys[key]) return this;
      this.settingsTemp.set(this.paramKeys[key], val);
      return this;
    }
    async write() {
      //this.settingsTemp.entries()
    }
    async readAsMap(range: string): Promise<Map<string, any>> {
      const vals: any[][] = await this.excelUtil.readValues(range);
      const map: Map<string, any> = new Map();
      if (!vals) return map;
      if (vals.length <= 0) return map;
      for (const v of vals) {
        if (v.length < 2) continue;
        map.set(v[0], v[1]);
      }
      return map;
    }
  }

  export type SettingsChain = {
    addSettings: { (key: string, val: any): SettingsChain };
    write: { ():void };
  };
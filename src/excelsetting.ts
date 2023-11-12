//import { ExcelUtil, ExcelRunner } from "./excelutil";
class ExcelSetting {
  sheetName: string;
  baseCell: string;
  excelUtil: ExcelUtil;
  paramKeys: any;
  loadedSettings: Map<string, any> = new Map();
  settingFieldIndex: Map<string, number> = new Map();
  paramRange: string;
  constructor(sheetName: string, paramKeys: any, runner:ExcelRunner, baseCell: string = "A2") {
    this.sheetName = sheetName;
    this.baseCell = baseCell;
    this.paramKeys = paramKeys;
    const fieldKeys = Object.keys(this.paramKeys);
    let i = 0;
    for (const k of fieldKeys) {
      this.settingFieldIndex.set(this.paramKeys[k], i);
      i++;
    }
    this.excelUtil = new ExcelUtil(sheetName, runner);
    this.paramRange = ExcelUtil.calcRangeDimension(Object.keys(paramKeys).length, 2, baseCell);
  }
  async getJson():Promise<any> {
    await this.sync();
    const keyArray: string[] = Object.keys(this.paramKeys);
    const fieldArray: string[] = Array.from(this.loadedSettings.keys());
    const postData: any = {};
    for (const f of fieldArray) {
      const val = this.loadedSettings.get(f);
      if (!val) continue;
      const idx:number|undefined = this.settingFieldIndex.get(f);
      if (!idx && idx!=0) continue;
      postData[keyArray[idx]] = val;
    }
    return postData
  }
  initializeSettings() {
    const settingsFields: string[] = new Array();
    for (const k of Object.keys(this.paramKeys)) {
      settingsFields.push(this.paramKeys[k]);
    }
    this.excelUtil.writeRows(settingsFields, this.baseCell);
  }

  settingsTemp: Map<string, any> = new Map();
  writeSettings(fieldKey: string, val: any): SettingsChain {
    this.settingsTemp = new Map();
    return this.addSettings(fieldKey, val);
  }
  async commitSettings(fieldKey: string, val: any): Promise<void> {
    this.settingsTemp = new Map();
    const ses = this.addSettings(fieldKey, val);
    await ses.write();
  }
  addSettings(fieldKey: string, val: any): SettingsChain {
    if (!fieldKey) return this;
    //if (!this.paramKeys[key]) return this;
    this.settingsTemp.set(fieldKey, val);
    return this;
  }
  async write() {
    const ses = this.excelUtil.startWriteSession();
    const sts: string[] = Array.from(this.settingsTemp.keys());
    for (const st of sts) {
      const idx: number|undefined = this.settingFieldIndex.get(st);
      if(!idx) continue;
      ses.addWriteChain([[this.settingsTemp.get(st)]], ExcelUtil.calcAddress(idx + 1, 2, this.baseCell));
    }
    await ses.sessionWrite();
    await this.sync();
    this.settingsTemp = new Map();
  }
  async sync(): Promise<Map<string, any>> {
    const vals: any[][] = await this.excelUtil.readValues(this.paramRange);
    const map: Map<string, any> = new Map();
    if (!vals) return map;
    if (vals.length <= 0) return map;
    for (const v of vals) {
      if (v.length < 2) continue;
      if (!v[0]) continue;
      map.set(v[0], v[1]);
    }
    this.loadedSettings = map;
    return map;
  }
}

type SettingsChain = {
  addSettings: { (key: string, val: any): SettingsChain };
  write: { ():void };
};
"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
//import { ExcelUtil, ExcelRunner } from "./excelutil";
class ExcelSetting {
    constructor(sheetName, paramKeys, runner, baseCell = "A2") {
        this.loadedSettings = new Map();
        this.settingFieldIndex = new Map();
        this.settingsTemp = new Map();
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
    getJson() {
        return __awaiter(this, void 0, void 0, function* () {
            yield this.sync();
            const keyArray = Object.keys(this.paramKeys);
            const fieldArray = Array.from(this.loadedSettings.keys());
            const postData = {};
            for (const f of fieldArray) {
                const val = this.loadedSettings.get(f);
                if (!val)
                    continue;
                const idx = this.settingFieldIndex.get(f);
                if (!idx && idx != 0)
                    continue;
                postData[keyArray[idx]] = val;
            }
            return postData;
        });
    }
    initializeSettings() {
        const settingsFields = new Array();
        for (const k of Object.keys(this.paramKeys)) {
            settingsFields.push(this.paramKeys[k]);
        }
        this.excelUtil.writeRows(settingsFields, this.baseCell);
    }
    writeSettings(fieldKey, val) {
        this.settingsTemp = new Map();
        return this.addSettings(fieldKey, val);
    }
    commitSettings(fieldKey, val) {
        return __awaiter(this, void 0, void 0, function* () {
            this.settingsTemp = new Map();
            const ses = this.addSettings(fieldKey, val);
            yield ses.write();
        });
    }
    addSettings(fieldKey, val) {
        if (!fieldKey)
            return this;
        //if (!this.paramKeys[key]) return this;
        this.settingsTemp.set(fieldKey, val);
        return this;
    }
    write() {
        return __awaiter(this, void 0, void 0, function* () {
            const ses = this.excelUtil.startWriteSession();
            const sts = Array.from(this.settingsTemp.keys());
            for (const st of sts) {
                const idx = this.settingFieldIndex.get(st);
                if (!idx)
                    continue;
                ses.addWriteChain([[this.settingsTemp.get(st)]], ExcelUtil.calcAddress(idx + 1, 2, this.baseCell));
            }
            yield ses.sessionWrite();
            yield this.sync();
            this.settingsTemp = new Map();
        });
    }
    sync() {
        return __awaiter(this, void 0, void 0, function* () {
            const vals = yield this.excelUtil.readValues(this.paramRange);
            const map = new Map();
            if (!vals)
                return map;
            if (vals.length <= 0)
                return map;
            for (const v of vals) {
                if (v.length < 2)
                    continue;
                if (!v[0])
                    continue;
                map.set(v[0], v[1]);
            }
            this.loadedSettings = map;
            return map;
        });
    }
}

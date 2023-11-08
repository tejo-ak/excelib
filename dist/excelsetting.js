var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
define(["require", "exports", "./excelutil"], function (require, exports, excelutil_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ExcelSetting = void 0;
    class ExcelSetting {
        constructor(sheetName, paramKeys, runner, baseCell = "A2") {
            this.settingsTemp = new Map();
            this.sheetName = sheetName;
            this.baseCell = baseCell;
            this.paramKeys = paramKeys;
            this.excelUtil = new excelutil_1.ExcelUtil(sheetName, runner);
        }
        writeSettings(key, val) {
            this.settingsTemp = new Map();
            return this.addSettings(key, val);
        }
        addSettings(key, val) {
            if (!key)
                return this;
            if (!this.paramKeys[key])
                return this;
            this.settingsTemp.set(this.paramKeys[key], val);
            return this;
        }
        write() {
            return __awaiter(this, void 0, void 0, function* () {
                //this.settingsTemp.entries()
            });
        }
        readAsMap(range) {
            return __awaiter(this, void 0, void 0, function* () {
                const vals = yield this.excelUtil.readValues(range);
                const map = new Map();
                if (!vals)
                    return map;
                if (vals.length <= 0)
                    return map;
                for (const v of vals) {
                    if (v.length < 2)
                        continue;
                    map.set(v[0], v[1]);
                }
                return map;
            });
        }
    }
    exports.ExcelSetting = ExcelSetting;
});

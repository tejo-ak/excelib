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
var __classPrivateFieldSet = (this && this.__classPrivateFieldSet) || function (receiver, state, value, kind, f) {
    if (kind === "m") throw new TypeError("Private method is not writable");
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a setter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot write private member to an object whose class did not declare it");
    return (kind === "a" ? f.call(receiver, value) : f ? f.value = value : state.set(receiver, value)), value;
};
var __classPrivateFieldGet = (this && this.__classPrivateFieldGet) || function (receiver, state, kind, f) {
    if (kind === "a" && !f) throw new TypeError("Private accessor was defined without a getter");
    if (typeof state === "function" ? receiver !== state || !f : !state.has(receiver)) throw new TypeError("Cannot read private member from an object whose class did not declare it");
    return kind === "m" ? f : kind === "a" ? f.call(receiver) : f ? f.value : state.get(receiver);
};
var _ExcelUtil_writeSessionTemp;
function createUtil(sheetName, excelRuner) {
    const xl = new ExcelUtil(sheetName, excelRuner);
    return xl;
}
class ExcelUtil {
    constructor(sheetName, excelRuner) {
        this.sheetName = "";
        _ExcelUtil_writeSessionTemp.set(this, Array());
        this.sheetName = sheetName;
        this.excelRuner = excelRuner;
    }
    replaceSheet() {
        return __awaiter(this, void 0, void 0, function* () {
            let sheet = yield this.getSheet();
            if (!!sheet) {
                yield this.excelRuner((context) => __awaiter(this, void 0, void 0, function* () {
                    sheet = context.workbook.worksheets.getItem(this.sheetName);
                    sheet.delete();
                    yield context.sync();
                }));
            }
            //await this.ensureSheet();
        });
    }
    getSheet() {
        return __awaiter(this, void 0, void 0, function* () {
            let sheet;
            yield this.excelRuner((context) => __awaiter(this, void 0, void 0, function* () {
                const sheets = context.workbook.worksheets;
                sheets.load("items/name");
                yield context.sync();
                const idx = sheets.items.findIndex((sh) => sh.name == this.sheetName);
                if (idx == -1) {
                    return;
                }
                sheet = context.workbook.worksheets.getItem(this.sheetName);
            }));
            return sheet;
        });
    }
    ensureSheet() {
        return __awaiter(this, void 0, void 0, function* () {
            const sheet = yield this.getSheet();
            if (!sheet) {
                yield this.excelRuner((context) => __awaiter(this, void 0, void 0, function* () {
                    context.workbook.worksheets.add(this.sheetName);
                    yield context.sync();
                    return context.workbook.worksheets.getItem(this.sheetName);
                }));
                return sheet;
            }
        });
    }
    readRows(range) {
        return __awaiter(this, void 0, void 0, function* () {
            const vals = yield this.readValues(range);
            const rowVal = vals.map((row) => row[0]);
            return rowVal;
        });
    }
    readCols(range) {
        return __awaiter(this, void 0, void 0, function* () {
            const vals = yield this.readValues(range);
            if (!vals)
                return;
            if (vals.length <= 0)
                return;
            return vals[0];
        });
    }
    readCell(range) {
        return __awaiter(this, void 0, void 0, function* () {
            const vals = yield this.readValues(range);
            if (!vals)
                return;
            if (vals.length <= 0)
                return;
            return vals[0][0];
        });
    }
    readValues(range) {
        return __awaiter(this, void 0, void 0, function* () {
            this.ensureSheet();
            let val;
            yield this.excelRuner((context) => __awaiter(this, void 0, void 0, function* () {
                const sheet = context.workbook.worksheets.getItem(this.sheetName);
                const filterRange = sheet.getRange(range);
                filterRange.load(["values"]);
                yield context.sync();
                val = filterRange.values;
            }));
            return val;
        });
    }
    writeValues(values, baseCell) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!values)
                return;
            if (values.length <= 0)
                return;
            const sheet = yield this.ensureSheet();
            yield this.excelRuner((context) => __awaiter(this, void 0, void 0, function* () {
                const sheet = context.workbook.worksheets.getItem(this.sheetName);
                const range = ExcelUtil.calcRange(values, baseCell);
                const filterRange = sheet.getRange(range);
                filterRange.values = values;
                yield context.sync();
            }));
        });
    }
    writeCell(value, baseCell) {
        return __awaiter(this, void 0, void 0, function* () {
            if (!value)
                return;
            yield this.writeValues([[value]], baseCell);
        });
    }
    writeRows(values, baseCell) {
        return __awaiter(this, void 0, void 0, function* () {
            const newValues = new Array();
            for (const val of values) {
                newValues.push([val]);
            }
            yield this.writeValues(newValues, baseCell);
        });
    }
    writeCols(values, baseCell) {
        return __awaiter(this, void 0, void 0, function* () {
            const newValues = new Array();
            newValues.push(values);
            yield this.writeValues(newValues, baseCell);
        });
    }
    startWriteSession(values, baseCell) {
        __classPrivateFieldSet(this, _ExcelUtil_writeSessionTemp, new Array(), "f");
        return this.addWriteChain(values, baseCell);
    }
    addWriteChain(values, baseCell) {
        __classPrivateFieldGet(this, _ExcelUtil_writeSessionTemp, "f").push({ values: values, baseCell: baseCell });
        return this;
    }
    addRowWriteChain(values, baseCell) {
        const valuesArrays = new Array();
        for (const val of values) {
            valuesArrays.push([values]);
        }
        this.addWriteChain(valuesArrays, baseCell);
        return this;
    }
    addColWriteChain(values, baseCell) {
        const valuesArrays = new Array();
        valuesArrays.push(values);
        this.addWriteChain(valuesArrays, baseCell);
        return this;
    }
    static calcRange(values, baseCell) {
        const rangeWidth = values[0].length;
        const baseCol = (baseCell.match(/[a-zA-Z]/g) || [])[0] || "";
        const baseRowNum = parseInt(baseCell.replace(baseCol, ""));
        const baseColNum = ExcelUtil.toColumnNumber(baseCol);
        const endColNum = baseColNum + rangeWidth - 1;
        const endRowNum = baseRowNum + values.length - 1;
        const endColName = ExcelUtil.toColumnName(endColNum);
        const range = `${baseCell}:${endColName}${endRowNum}`;
        return range;
    }
    sessionWrite() {
        return __awaiter(this, void 0, void 0, function* () {
            const sheet = yield this.ensureSheet();
            yield this.excelRuner((context) => __awaiter(this, void 0, void 0, function* () {
                const sheet = context.workbook.worksheets.getItem(this.sheetName);
                const ranges = new Array();
                for (const session of __classPrivateFieldGet(this, _ExcelUtil_writeSessionTemp, "f")) {
                    const range = ExcelUtil.calcRange(session.values, session.baseCell);
                    const sheetRange = sheet.getRange(range);
                    sheetRange.values = session.values;
                    ranges.push(sheetRange);
                }
                yield context.sync();
                __classPrivateFieldSet(this, _ExcelUtil_writeSessionTemp, new Array(), "f");
            }));
        });
    }
    static toColumnName(index) {
        const ordA = "a".charCodeAt(0);
        const ordZ = "z".charCodeAt(0);
        const len = ordZ - ordA + 1;
        let s = "";
        let n = index;
        while (n >= 0) {
            s = String.fromCharCode((n % len) + ordA) + s;
            n = Math.floor(n / len) - 1;
        }
        return s.toLocaleUpperCase();
    }
    static toColumnNumber(val) {
        val = val.toLocaleUpperCase();
        let base = "ABCDEFGHIJKLMNOPQRSTUVWXYZ", i, j, result = 0;
        for (i = 0, j = val.length - 1; i < val.length; i += 1, j -= 1) {
            result += Math.pow(base.length, j) * (base.indexOf(val[i]) + 1);
        }
        return result - 1;
    }
}
_ExcelUtil_writeSessionTemp = new WeakMap();

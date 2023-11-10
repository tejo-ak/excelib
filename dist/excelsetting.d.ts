export declare class ExcelSetting {
    sheetName: string;
    baseCell: string;
    excelUtil: ExcelUtil;
    paramKeys: any;
    constructor(sheetName: string, paramKeys: any, runner: ExcelRunner, baseCell?: string);
    settingsTemp: Map<string, any>;
    writeSettings(key: string, val: any): SettingsChain;
    addSettings(key: string, val: any): SettingsChain;
    write(): Promise<void>;
    readAsMap(range: string): Promise<Map<string, any>>;
}
export type SettingsChain = {
    addSettings: {
        (key: string, val: any): SettingsChain;
    };
    write: {
        (): void;
    };
};

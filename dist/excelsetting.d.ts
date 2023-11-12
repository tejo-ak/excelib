declare class ExcelSetting {
    sheetName: string;
    baseCell: string;
    excelUtil: ExcelUtil;
    paramKeys: any;
    loadedSettings: Map<string, any>;
    settingFieldIndex: Map<string, number>;
    paramRange: string;
    constructor(sheetName: string, paramKeys: any, runner: ExcelRunner, baseCell?: string);
    getJson(): Promise<any>;
    initializeSettings(): void;
    settingsTemp: Map<string, any>;
    writeSettings(fieldKey: string, val: any): SettingsChain;
    commitSettings(fieldKey: string, val: any): Promise<void>;
    addSettings(fieldKey: string, val: any): SettingsChain;
    write(): Promise<void>;
    sync(): Promise<Map<string, any>>;
}
type SettingsChain = {
    addSettings: {
        (key: string, val: any): SettingsChain;
    };
    write: {
        (): void;
    };
};

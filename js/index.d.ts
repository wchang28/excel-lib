export interface Options {
    report: string;
    visible?: boolean;
    addins?: string[];
    open_after_save?: boolean;
}
export declare class ExcelReport {
    private __options;
    constructor(options: Options);
    private static getExcelFilePath();
    private static openExcelFile(excel_file);
    generate(): void;
    saveWorkbook(wrkbk: any, filepath: string): void;
    populate(excelApp: any): void;
}

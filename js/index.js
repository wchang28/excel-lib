"use strict";
exports.__esModule = true;
var ExcelReport = /** @class */ (function () {
    function ExcelReport(options) {
        if (typeof options == "undefined")
            throw "options is not optional";
        if (typeof options.visible == "undefined")
            options.visible = false;
        if (typeof options.addins == "undefined")
            options.addins = [];
        if (typeof options.report != "string")
            throw "options.report is not optional";
        if (typeof options.open_after_save == "undefined")
            options.open_after_save = false;
        this.__options = options;
    }
    // returns Excel.exe's file path
    ExcelReport.getExcelFilePath = function () {
        var shell = new ActiveXObject("WScript.Shell");
        var excel_path = shell.RegRead("HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe\\Path");
        if (excel_path.substr(excel_path.length - 1) == "\\")
            excel_path = excel_path.substr(0, excel_path.length - 1);
        excel_path += "\\EXCEL.EXE";
        return excel_path;
    };
    ExcelReport.openExcelFile = function (excel_file) {
        if (typeof excel_file != "string")
            throw "excel_file is not optional";
        try {
            var excel_filepath = ExcelReport.getExcelFilePath();
            var cmd = '"' + excel_filepath + '"';
            cmd += " ";
            cmd += '"' + excel_file + '"';
            var shell = new ActiveXObject("WScript.Shell");
            shell.Run(cmd, 5, false);
        }
        catch (e) {
            throw "unable to open the Excel file: " + excel_file;
        }
    };
    ExcelReport.prototype.generate = function () {
        var excelApp = new ActiveXObject("Excel.Application");
        var options = this.__options;
        try {
            excelApp.Visible = options.visible; // show shide Excel during report population
            excelApp.DisplayAlerts = false;
            excelApp.AlertBeforeOverwriting = false;
            // loading addins required by the report generation
            for (var i in options.addins) {
                var addin = options.addins[i];
                excelApp.AddIns.Item(addin).Installed = true;
                excelApp.Workbooks.Open(excelApp.AddIns.Item(addin).FullName);
            }
            var wrkbk = this.populate(excelApp); // call the derived class populate method to populate the report
            WScript.Echo("saving the report to " + options.report + "...");
            this.saveWorkbook(wrkbk, options.report); // save the report
            WScript.Echo("report successfully saved");
            excelApp.Quit(); // quit Excel
            if (options.open_after_save)
                ExcelReport.openExcelFile(options.report);
        }
        catch (e) {
            excelApp.Quit(); // quit Excel
            throw e;
        }
    };
    // save workbook to a file
    ExcelReport.prototype.saveWorkbook = function (wrkbk, filepath) {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        try {
            fso.DeleteFile(filepath);
        }
        catch (e) { } // try to delete it first
        try {
            wrkbk.SaveAs(filepath);
        }
        catch (e) {
            throw "unable to save the file " + filepath + ". the file is probabily locked or already openned in Excel";
        }
    };
    ExcelReport.prototype.populate = function (excelApp) {
        return excelApp.Workbooks.Add(); // create a blank work book
    };
    return ExcelReport;
}());
exports.ExcelReport = ExcelReport;
//# sourceMappingURL=index.js.map
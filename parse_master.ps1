
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

public class ExcelMasterParser {
    public static void Parse() {
        Application excel = null;
        Workbooks wbs = null;
        Workbook wb = null;

        try {
            excel = new Application();
            wbs = excel.Workbooks;
            string dir = @"C:\Users\307984\.openclaw\workspace\temp_attachments";
            string fileName = "";
            foreach(string f in Directory.GetFiles(dir, "*.xlsx")) {
                if (f.Contains("2") && f.Contains("AI")) { fileName = f; break; }
            }

            if (string.IsNullOrEmpty(fileName)) { Console.WriteLine("FILE_NOT_FOUND"); return; }

            wb = wbs.Open(fileName);
            foreach (Worksheet sheet in wb.Sheets) {
                Console.WriteLine("=== SHEET_START: " + sheet.Name + " ===");
                Range range = sheet.UsedRange;
                object[,] values = (object[,])range.Value2;
                if (values != null) {
                    int rows = values.GetLength(0);
                    int cols = values.GetLength(1);
                    for (int r = 1; r <= rows; r++) {
                        List<string> rowData = new List<string>();
                        bool hasValue = false;
                        for (int c = 1; c <= cols; c++) {
                            object raw = values[r, c];
                            string val = (raw != null) ? raw.ToString().Trim() : "";
                            rowData.Add(val);
                            if (!string.IsNullOrEmpty(val)) hasValue = true;
                        }
                        if (hasValue) Console.WriteLine(string.Join(" | ", rowData));
                    }
                }
                Console.WriteLine("=== SHEET_END ===");
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(sheet);
            }
        } catch (Exception e) {
            Console.WriteLine("ERROR: " + e.Message);
        } finally {
            if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); }
            if (wbs != null) Marshal.ReleaseComObject(wbs);
            if (excel != null) { excel.Quit(); Marshal.ReleaseComObject(excel); }
        }
    }
}
"@ -ReferencedAssemblies "Microsoft.Office.Interop.Excel", "System.Runtime.InteropServices"
[ExcelMasterParser]::Parse()

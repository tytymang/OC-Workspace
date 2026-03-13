
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

public class ExcelReaderFinal {
    public static void Read() {
        Application excel = null;
        Workbooks wbs = null;
        Workbook wb = null;

        try {
            excel = new Application();
            wbs = excel.Workbooks;
            
            string dir = @"C:\Users\307984\.openclaw\workspace\temp_attachments";
            string fileName = "";
            foreach(string f in Directory.GetFiles(dir, "*.xlsx")) {
                if (f.Contains("2") && f.Contains("AI")) {
                    fileName = f;
                    break;
                }
            }

            if (string.IsNullOrEmpty(fileName)) {
                Console.WriteLine("FILE_NOT_FOUND");
                return;
            }

            wb = wbs.Open(fileName);
            foreach (Worksheet sheet in wb.Sheets) {
                Console.WriteLine("### SHEET: " + sheet.Name);
                Range range = sheet.UsedRange;
                object[,] values = (object[,])range.Value2;
                if (values == null) continue;

                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                for (int r = 1; r <= rowCount; r++) {
                    bool hasData = false;
                    string line = "";
                    for (int c = 1; c <= colCount; c++) {
                        object rawVal = values[r, c];
                        string val = (rawVal != null) ? rawVal.ToString() : "";
                        if (!string.IsNullOrEmpty(val)) {
                            line += val + " | ";
                            if (val.Length > 0) hasData = true; 
                        }
                    }
                    if (hasData) Console.WriteLine(line);
                }
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
[ExcelReaderFinal]::Read()

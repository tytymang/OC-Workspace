
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

public class ExcelReader {
    public static void Read() {
        Application excel = null;
        Workbooks wbs = null;
        Workbook wb = null;
        Worksheet sheet = null;
        Range range = null;

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
            sheet = (Worksheet)wb.Sheets[1];
            range = sheet.UsedRange;

            int rowCount = Math.Min(range.Rows.Count, 100); // 안전을 위해 100행까지만
            int colCount = Math.Min(range.Columns.Count, 20); // 안전을 위해 20열까지만

            for (int r = 1; r <= rowCount; r++) {
                string line = "";
                for (int c = 1; c <= colCount; c++) {
                    Range cell = (Range)range.Cells[r, c];
                    line += cell.Text + " | ";
                    Marshal.ReleaseComObject(cell);
                }
                Console.WriteLine(line);
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
[ExcelReader]::Read()

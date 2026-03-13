
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

public class ExcelReaderV3 {
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
            string[] files = Directory.GetFiles(dir, "*.xlsx");
            foreach(string f in files) {
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

            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            if (rows > 100) rows = 100;
            if (cols > 20) cols = 20;

            for (int r = 1; r <= rows; r++) {
                string line = "";
                for (int c = 1; c <= cols; c++) {
                    Range cell = (Range)range.Cells[r, c];
                    line += cell.Text + " | ";
                    Marshal.ReleaseComObject(cell);
                }
                Console.WriteLine(line);
            }
        } catch (Exception e) {
            Console.WriteLine("ERROR: " + e.Message);
        } finally {
            if (wb != null) { try { wb.Close(false); Marshal.ReleaseComObject(wb); } catch {} }
            if (wbs != null) Marshal.ReleaseComObject(wbs);
            if (excel != null) { try { excel.Quit(); Marshal.ReleaseComObject(excel); } catch {} }
        }
    }
}
"@ -ReferencedAssemblies "Microsoft.Office.Interop.Excel", "System.Runtime.InteropServices"
[ExcelReaderV3]::Read()

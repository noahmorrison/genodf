using System;
using System.IO;

using Genodf;

namespace GenodfTest
{
    public class GenodfTest
    {
        public static void Main()
        {
            var cwd = Directory.GetCurrentDirectory();
            if (!cwd.EndsWith("build"))
                cwd = Path.Combine(cwd, "build");
            var filePath = Path.Combine(cwd, "genodf.ods");

            var sheet = new Spreadsheet();
            sheet.SetCell("A1", "2.5");
            sheet.SetCell("B1", "3.5");
            sheet.SetCell("C4", "=SUM(A1:B1)");

            sheet.GetCell("C4").Bg = "#ff0000";
            sheet.GetCell("C4").Fg = "#0000ff";
            sheet.GetCell("C4").TextAlign = "center";
            sheet.GetCell("C4").Bold = true;

            sheet.SetCell("B6", "I'm really big!");
            sheet.GetCell("B6").SpannedColumns = 2;
            sheet.GetCell("B6").SpannedRows = 2;

            sheet.SetCell("B9", "all");
            sheet.GetCell("B9").Border = true;
            sheet.SetCell("D9", "top");
            sheet.GetCell("D9").BorderTop = true;
            sheet.SetCell("F9", "bottom");
            sheet.GetCell("F9").BorderBottom = true;
            sheet.SetCell("H9", "left");
            sheet.GetCell("H9").BorderLeft = true;
            sheet.SetCell("J9", "right");
            sheet.GetCell("J9").BorderRight = true;

            var notSet = sheet.GetCell("E1");
            notSet.value = "hey!";

            var neverSet = sheet.GetCell("F1");
            neverSet.Bg = "#aa55aa";

            sheet.GetColumn(5).Width = 0.5;

            sheet.Write(filePath);
            Console.WriteLine("Done with test");
        }
    }
}

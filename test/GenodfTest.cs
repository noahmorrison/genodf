using System;
using System.IO;
using System.Reflection;

using Genodf;

namespace GenodfTest
{
    public class GenodfTest
    {
        public static void Main()
        {
            var path = Assembly.GetExecutingAssembly().Location;
            var dir = Path.GetDirectoryName(path);
            var filePath = Path.Combine(dir, "genodf.ods");

            var ods = new Spreadsheet();
            var sheet = ods.NewSheet("Genodf Test");

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

            sheet.SetCell("AU171", "wow, I'm far out");

            sheet.GetCells("a2:h2").ForEach(c => c.Bg = "#000000");

            sheet.SetCell("F20", "104.23491");
            sheet.GetCell("F20").Format = new NumberFormat("000.00");

            sheet.SetCell("F21", "0.64");
            sheet.GetCell("F21").Format = new NumberFormat("0%");

            var multisheet = ods.NewSheet("MultiSheet");

            multisheet.SetCell("A1", "Change me");
            var format = new NumberFormat("0");
            format.AddCondition("-0", "value()<0").Fg = "#ff0000";
            multisheet.GetCell("A1").Format = format;

            ods.Write(filePath);
            Console.WriteLine("Done with test");
        }
    }
}

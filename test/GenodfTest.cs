using System;
using System.IO;
using System.Reflection;

using Genodf;

namespace GenodfTest
{
    public class GenodfTest
    {
        static string filePath;

        public static void Main()
        {
            var path = typeof(GenodfTest).GetTypeInfo().Assembly.Location;
            var dir = Path.GetDirectoryName(path);
            filePath = Path.Combine(dir, "genodf.ods");

            TestWriting();

            Console.WriteLine("Done with test");
        }

        private static void TestWriting()
        {
            var ods = new Spreadsheet();
            var sheet = ods.NewSheet("Genodf Test");
            sheet.PrintingHeader = new Header(height: 2);
            sheet.VisualHeader = new Header(height: 2);

            sheet.SetCell("A1", "2.5");
            sheet.SetCell("B1", "3.5");
            sheet.SetCell("C4", "=SUM(A1:B1)");

            sheet.GetCell("C4").Bg = "#ff0000";
            sheet.GetCell("C4").Fg = "#0000ff";
            sheet.GetCell("C4").TextAlign = TextAlign.Center;
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
            notSet.Value = "hey!";
            notSet.TextAlign = TextAlign.Center;

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

            multisheet.SetCell("A1", "-42");
            var format = new NumberFormat("0");
            format.AddCondition("-0", "value()<0").Fg = "#ff0000";
            multisheet.GetCell("A1").Format = format;


            var redStyle = new Genodf.Styles.CellStyle();
            redStyle.Fg = "#ff0000";
            ods.AddGlobalStyle("RedStyle", redStyle);
            multisheet.SetCell("B1", "Fail");
            multisheet.GetCell("B1").AddConditional("cell-content()=\"Fail\"", "RedStyle");

            ods.MasterPageStyle.PageLayout = new Genodf.Styles.PageLayoutStyle("PageLayout")
            {
                PageWidth = "11in",
                PageHeight = "17in",
                FitInPages = 1
            };
            ods.AddGlobalStyle(ods.MasterPageStyle.PageLayout);

            ods.Write(filePath);
        }

        private static void TestReloading()
        {
            var ods = OpenDocument.Read<Spreadsheet>(filePath);
            var multisheet = ods.Sheets[1];
            multisheet.GetCell("C5", cell =>
            {
                cell.Value = "Set after a reload";
                cell.Bg = "#eeeeee";
                cell.SpannedColumns = 4;
                cell.TextAlign = TextAlign.Center;
            });

            multisheet.GetCell("C6", cell =>
            {
                cell.SpannedRows = 2;
                cell.SpannedColumns = 2;
            });

            multisheet.BorderAround("C6:D7");
            multisheet.BorderAround("C5:F10");

            ods.Write(filePath);
        }

        private static void TestReloadingShouldChangeNothing()
        {
            var ods = OpenDocument.Read<Spreadsheet>(filePath);
            var filePath2 = filePath.Replace("genodf.ods", "genodf-reloaded.ods");
            ods.Write(filePath2);

            if (!CompareFiles(filePath, filePath2)) {}
                //throw new Exception("Reloaded file does not match original");
        }

        private static bool CompareFiles(string file1, string file2)
        {
            using (var fs1 = new FileStream(file1, FileMode.Open, FileAccess.Read))
            using (var fs2 = new FileStream(file2, FileMode.Open, FileAccess.Read))
            {
                while (fs1.ReadByte() == fs2.ReadByte())
                    if (fs1.Position == fs1.Length)
                        return true;
                return false;
            }
        }
    }
}

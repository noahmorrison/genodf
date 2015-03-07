using System;
using System.IO;

namespace Genodf
{
    public class Genodf
    {
        public static void Main()
        {
            var cwd = Directory.GetCurrentDirectory();
            var build = Path.Combine(cwd, "build");
            var filePath = Path.Combine(build, "genodf.ods");

            var sheet = new Spreadsheet();
            sheet.SetCell("A1", "2.5");
            sheet.SetCell("B1", "3.5");
            sheet.SetCell("C4", "=SUM(A1:B1)");

            sheet.GetCell("C4").Bg = "#ff0000";
            sheet.GetCell("C4").Fg = "#0000ff";
            sheet.GetCell("C4").Aligned = "center";

            sheet.Write(filePath);
            Console.WriteLine("Done with test");
        }
    }
}

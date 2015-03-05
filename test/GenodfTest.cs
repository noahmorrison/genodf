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
            sheet.Write(filePath);
            Console.WriteLine("Done with test");
        }
    }
}

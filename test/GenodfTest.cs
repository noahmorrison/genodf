using System;
using System.IO;

namespace Genodf
{
    public class Genodf
    {
        public static void Main()
        {
            var cwd = Directory.GetCurrentDirectory();
            var root = Directory.GetParent(cwd).FullName;
            var build = Path.Combine(root, "build");
            var filePath = Path.Combine(build, "genodf.ods");

            var sheet = new Spreadsheet();
            sheet.Write(filePath);
            Console.WriteLine("Done with test");
        }
    }
}

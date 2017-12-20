using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace _6502Annotate
{
    class Program
    {
        static void Main(string[] args)
        {
            var fi = new FileInfo(args[0]);
            using (var ef = new ExcelPackage(fi))
            {
                Annotator.Annotate(ef.Workbook.Worksheets.First());
                Console.Write("Saving... ");
                ef.Save();
                Console.WriteLine("done");
            }
        }
    }
}

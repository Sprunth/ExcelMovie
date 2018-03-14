using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelMovie
{
    class Program
    {
        static void Main(string[] args)
        {
            var outputFile = "test.xlsx";

            var vfm = new VideoFrameManager("films/painting");

            // delete if exists, always start fresh
            System.IO.File.Delete(outputFile);
            using (var package = new ExcelPackage(new FileInfo(outputFile)))
            {
                var style = package.Workbook.Styles.CreateNamedStyle("BaseStyle");
                style.Style.Fill.PatternType = ExcelFillStyle.Solid;

                foreach (var frameName in vfm.FrameNames.GetRange(11,4))
                {
                    Console.WriteLine($"Processing {frameName}");
                    var bitmap = vfm.GetFrame(frameName);
                    var worksheet = package.Workbook.Worksheets.Add(frameName);
                    BitmapToWorksheet.DrawBitmapOnWorksheet(bitmap, worksheet);
                    bitmap.Dispose();
                }
                Console.WriteLine($"Saving to {outputFile}");
                package.Save();
            }
        }
    }
}

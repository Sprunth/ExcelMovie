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
            var vfm = new VideoFrameManager("films/painting");

            // delete if exists, always start fresh
            System.IO.File.Delete("test.xlsx");
            using (var package = new ExcelPackage(new FileInfo("test.xlsx")))
            {
                //var worksheet = package.Workbook.Worksheets.Add("TestPage1");
                //worksheet.Cells[2, 1].Value = "B1, Bold!";
                //worksheet.Cells[2, 1].Style.Font.Bold = true;
                //worksheet.Cells[2, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //worksheet.Cells[2, 3].Style.Fill.BackgroundColor.SetColor(Color.Orange);
                //package.Save();

                var bitmap = vfm.GetFrame(10);
                var worksheet2 = package.Workbook.Worksheets.Add("10");

                BitmapToWorksheet.DrawBitmapOnWorksheet(bitmap, worksheet2);
                package.Save();
            }
        }
    }
}

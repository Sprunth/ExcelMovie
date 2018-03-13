using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelMovie
{
    static class BitmapToWorksheet
    {
        public static void DrawBitmapOnWorksheet(Bitmap bitmap, ExcelWorksheet worksheet)
        {
            var width = bitmap.Size.Width;
            var height = bitmap.Size.Height;

            var rgbArray = new int[height * width];
            bitmap.getRGB(0, 0, width, height, rgbArray, 0, width);
            var colorArray = ConvertRgbArrayToColors(rgbArray, width, height);

            for (var x = 0; x < width; x++)
            {
                for (var y = 0; y < height; y++)
                {
                    worksheet.Cells[y+1, x+1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[y+1, x+1].Style.Fill.BackgroundColor.SetColor(colorArray[x,y]);
                }
                Console.WriteLine($"{x}/{width} Done");
            }
        }

        private static Color[,] ConvertRgbArrayToColors(int[] rgbArray, int width, int height)
        {
            var ret = new Color[width, height];
            for (var x=0; x<width; x++)
            {
                for (var y = 0; y < height; y++)
                {
                    ret[x, y] = Color.FromArgb(rgbArray[x + y*width]);
                }
            }
            return ret;
        } 
    }
}

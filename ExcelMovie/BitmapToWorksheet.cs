using System;
using System.Collections.Generic;
using System.Diagnostics;
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

            var scaleDown = 2;
            var targetWidth = width / scaleDown;
            var targetHeight = height / scaleDown;

            var rgbArray = new int[height * width];
            bitmap.getRGB(0, 0, width, height, rgbArray, 0, width);
            var colorArray = ConvertRgbArrayToColors(rgbArray, width, height, scaleDown);

            for (var x = 0; x < targetWidth; x++)
            {
                worksheet.Column(x+1).Width = 3;
                for (var y = 0; y < targetHeight; y++)
                {
                    var cell = worksheet.Cells[y + 1, x + 1];
                    cell.StyleName = "BaseStyle";
                    //worksheet.Cells[y+1, x+1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(colorArray[x,y]);
                    //cell.Value = "\u2588";
                }
                if (x%40 == 0)
                    Console.WriteLine($"{x}/{targetWidth} Done");
            }
        }

        private static Color[,] ConvertRgbArrayToColors(int[] rgbArray, int width, int height, int scaleDown)
        {
            var targetWidth = width / scaleDown;
            var targetHeight = height / scaleDown;

            var ret = new Color[targetWidth, targetHeight];
            for (var x = 0; x < targetWidth; x++)
            {
                for (var y = 0; y < targetHeight; y++)
                {
                    var a = 0;
                    var r = 0;
                    var g = 0;
                    var b = 0;
                    for (var i = 0; i < scaleDown; i++)
                    {
                        for (var j = 0; j < scaleDown; j++)
                        {
                            var col = Color.FromArgb(rgbArray[(x*scaleDown + i) + (y*scaleDown + j) * width]);
                            a += col.A;
                            r += col.R;
                            g += col.G;
                            b += col.B;
                        }
                    }

                    a /= scaleDown * scaleDown;
                    r /= scaleDown * scaleDown;
                    g /= scaleDown * scaleDown;
                    b /= scaleDown * scaleDown;
                    Debug.Assert(a <= 255 && a >= 0);
                    Debug.Assert(r <= 255 && r >= 0);
                    Debug.Assert(g <= 255 && g >= 0);
                    Debug.Assert(b <= 255 && b >= 0);
                    ret[x, y] = Color.FromArgb(a, r, g, b);
                }
            }
            return ret;
        } 
    }
}

﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using BitmapData = System.Drawing.Imaging.BitmapData;

namespace ExcelMovie
{
    public static class BitmapExtensions
    {
        // https://stackoverflow.com/a/4748383
        public static void getRGB(this Bitmap image, int startX, int startY, int w, int h, int[] rgbArray, int offset, int scansize)
        {
            const int PixelWidth = 3;
            const PixelFormat Format = PixelFormat.Format24bppRgb;

            // En garde!
            if (image == null) throw new ArgumentNullException("image");
            if (rgbArray == null) throw new ArgumentNullException("rgbArray");
            if (startX < 0 || startX + w > image.Width) throw new ArgumentOutOfRangeException("startX");
            if (startY < 0 || startY + h > image.Height) throw new ArgumentOutOfRangeException("startY");
            if (w < 0 || w > scansize || w > image.Width) throw new ArgumentOutOfRangeException("w");
            if (h < 0 || (rgbArray.Length < offset + h * scansize) || h > image.Height) throw new ArgumentOutOfRangeException("h");

            BitmapData data = image.LockBits(new Rectangle(startX, startY, w, h), ImageLockMode.ReadOnly, Format);
            try
            {
                byte[] pixelData = new Byte[data.Stride];
                for (int scanline = 0; scanline < data.Height; scanline++)
                {
                    Marshal.Copy(data.Scan0 + (scanline * data.Stride), pixelData, 0, data.Stride);
                    for (int pixeloffset = 0; pixeloffset < data.Width; pixeloffset++)
                    {
                        // PixelFormat.Format32bppRgb means the data is stored
                        // in memory as BGR. We want RGB, so we must do some 
                        // bit-shuffling.
                        rgbArray[offset + (scanline * scansize) + pixeloffset] =
                            (pixelData[pixeloffset * PixelWidth + 2] << 16) +   // R 
                            (pixelData[pixeloffset * PixelWidth + 1] << 8) +    // G
                            pixelData[pixeloffset * PixelWidth];                // B
                    }
                }
            }
            finally
            {
                image.UnlockBits(data);
            }
        }
    }
}

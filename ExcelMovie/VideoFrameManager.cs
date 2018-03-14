using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMovie
{
    class VideoFrameManager
    {
        public readonly List<string> FrameNames;
        public VideoFrameManager(string imageFolder)
        {
            FrameNames = Directory.GetFiles(imageFolder, "*.jpg").ToList();
        }

        public Bitmap GetFrame(string frameName)
        {
            var bitmap = Image.FromFile(frameName) as Bitmap;
            return bitmap;
        }
    }
}

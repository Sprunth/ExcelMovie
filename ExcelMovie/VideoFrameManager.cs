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
        private List<string> frames;
        public VideoFrameManager(string imageFolder)
        {
            frames = Directory.GetFiles(imageFolder, "*.jpg").ToList();
        }

        public Bitmap GetFrame(int index)
        {
            var bitmap = Image.FromFile(frames[index]) as Bitmap;
            return bitmap;
        }
    }
}

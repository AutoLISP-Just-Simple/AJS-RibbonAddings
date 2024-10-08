﻿using System;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;

namespace AJS_RibbonAddings
{
    internal static class Extensions
    {
        public static string Left(this string input, int count)
        {
            try
            {
                return input.Substring(0, Math.Min(input.Length, count));
            }
            catch { }

            return "";
        }

        public static string Mid(this string input, int start, int count)
        {
            return input.Substring(Math.Min(start, input.Length), Math.Min(count, Math.Max(input.Length - start, 0)));
        }

        public static BitmapImage ToImageSource(this Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }
    }
}
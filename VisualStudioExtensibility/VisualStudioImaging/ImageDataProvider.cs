using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using ImageAttributes = Microsoft.VisualStudio.Imaging.Interop.ImageAttributes;

namespace VisualStudioImaging
{
    public interface IImageDataProvider
    {
        Bitmap GetImage(Bitmap image, uint themedColor);
        Bitmap GetImageMoniker(IVsUIObject vsUiObject);
        byte[] GetIconBytes(byte[] imageBytes, int imageMaxDimension);
        byte[] GetImageBytes(Image image);
        IVsUIObject GetVsUiObject(ImageMoniker imageMoniker, ImageAttributes imageAttributes);
        uint? GetThemedColor(uint colorType);
    }

    public class ImageDataProvider : IImageDataProvider
    {
        private readonly IVsImageService2 _vsImageService;
        private readonly IVsUIShell6 _vsUiShell;

        public ImageDataProvider(IVsImageService2 vsImageService, IVsUIShell6 vsUiShell)
        {
            _vsImageService = vsImageService;
            _vsUiShell = vsUiShell;
        }

        public Bitmap GetImage(Bitmap image, uint themedColor)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var rectangle = new Rectangle(default(Point), image.Size);
                var bitmapData = image.LockBits(rectangle, ImageLockMode.ReadWrite, image.PixelFormat);
                var bitmapDataPointer = bitmapData.Scan0;
                var length = Math.Abs(bitmapData.Stride) * image.Height;
                var bytes = new byte[length];

                Marshal.Copy(source: bitmapDataPointer, destination: bytes, startIndex: 0, length: length);

                _vsUiShell.ThemeDIBits((uint)bytes.Length, bytes, (uint)image.Width, (uint)image.Height, fIsTopDownBitmap: true, crBackground: themedColor);

                Marshal.Copy(source: bytes, startIndex: 0, destination: bitmapDataPointer, length: length);

                image.UnlockBits(bitmapData);

                return image;
            }
            catch
            {
                return null;
            }
        }

        public Bitmap GetImageMoniker(IVsUIObject vsUiObject)
        {
            var objectData = Utilities.GetObjectData(vsUiObject);
            var image = (Bitmap)objectData;

            image.MakeTransparent(Color.Transparent);

            return image;
        }

        public byte[] GetIconBytes(byte[] imageBytes, int imageMaxDimension)
        {
            var iconBytes = new byte[byte.MaxValue];

            iconBytes[02] = 01;
            iconBytes[04] = 01;
            iconBytes[06] = (byte)imageMaxDimension;
            iconBytes[07] = (byte)imageMaxDimension;
            iconBytes[10] = 01;
            iconBytes[12] = 24;
            iconBytes[14] = (byte)(imageBytes.Length & iconBytes.Length);
            iconBytes[15] = (byte)(imageBytes.Length / (iconBytes.Length + 1));
            iconBytes[18] = (byte)iconBytes.Length;

            return iconBytes;
        }

        public byte[] GetImageBytes(Image image)
        {
            byte[] imageBytes;

            using (var memoryStream = new MemoryStream())
            {
                image.Save(memoryStream, ImageFormat.Png);
                memoryStream.Position = 0;
                imageBytes = memoryStream.ToArray();
            }

            return imageBytes;
        }

        public IVsUIObject GetVsUiObject(ImageMoniker imageMoniker, ImageAttributes imageAttributes)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var vsUiObject = _vsImageService.GetImage(imageMoniker, imageAttributes);

                return vsUiObject;
            }
            catch
            {
                return null;
            }
        }

        public uint? GetThemedColor(uint colorType)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var themeResourceKey = EnvironmentColors.EnvironmentBackgroundBrushKey;
                var colorCategory = themeResourceKey.Category;
                var colorName = themeResourceKey.Name;
                var themedColor = _vsUiShell.GetThemedColor(colorCategory, colorName, colorType);

                return themedColor;
            }
            catch
            {
                return null;
            }
        }
    }
}
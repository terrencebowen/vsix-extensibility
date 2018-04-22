using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Runtime.InteropServices;

namespace VisualStudioImaging
{
    public interface IMonikerAttributes
    {
        _UIImageType UiImageType { get; }
        ColorTheme ColorTheme { get; }
        ImageMoniker ImageMoniker { get; }
        int Height { get; }
        int Width { get; }
        ImageAttributes GetImageAttributes();
        int GetMaxDimension();
        uint GetColorType();
    }

    public class MonikerAttributes : IMonikerAttributes
    {
        public ImageMoniker ImageMoniker { get; }
        public int Height { get; }
        public int Width { get; }
        public _UIImageType UiImageType { get; }
        public ColorTheme ColorTheme { get; }

        public MonikerAttributes(ImageMoniker imageMoniker, ColorTheme colorTheme, _UIImageType uiImageType, int width, int height)
        {
            ColorTheme = colorTheme;
            Height = height;
            ImageMoniker = imageMoniker;
            UiImageType = uiImageType;
            Width = width;
        }

        public ImageAttributes GetImageAttributes()
        {
            const uint flags = (uint)_ImageAttributesFlags.IAF_RequiredFlags;
            const uint format = (uint)_UIDataFormat.DF_WinForms;
            const uint uiImageType = (uint)_UIImageType.IT_Bitmap;

            var logicalDpiX = DpiHelper.LogicalDpiX;
            var logicalDpiY = DpiHelper.LogicalDpiY;
            var dpi = (int)((logicalDpiX + logicalDpiY) / 2);
            var structSize = Marshal.SizeOf(typeof(ImageAttributes));

            var imageAttributes = new ImageAttributes
            {
                Dpi = dpi,
                Flags = flags,
                Format = format,
                ImageType = uiImageType,
                LogicalHeight = Height,
                LogicalWidth = Width,
                StructSize = structSize
            };

            return imageAttributes;
        }

        public int GetMaxDimension()
        {
            var maxDimension = Width >= Height ? Width : Height;

            if (UiImageType == _UIImageType.IT_Icon)
            {
                maxDimension = maxDimension < byte.MaxValue + 1 ? maxDimension : 0;
            }

            return maxDimension;
        }

        public uint GetColorType()
        {
            __THEMEDCOLORTYPE themedColorType;

            switch (ColorTheme)
            {
                case ColorTheme.Dark:
                    themedColorType = __THEMEDCOLORTYPE.TCT_Foreground;
                    break;

                case ColorTheme.Light:
                    themedColorType = __THEMEDCOLORTYPE.TCT_Background;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(ColorTheme));
            }

            var colorType = (uint)themedColorType;

            return colorType;
        }
    }
}
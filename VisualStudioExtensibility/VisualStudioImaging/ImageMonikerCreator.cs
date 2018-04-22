using Microsoft.VisualStudio.Shell.Interop;
using stdole;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace VisualStudioImaging
{
    public interface IImageMonikerCreator
    {
        Icon CreateIcon(IMonikerAttributes monikerAttributes, Image image);
        Image CreateImage(IMonikerAttributes monikerAttributes, IVsUIObject vsUiObject);
        StdPicture CreateStandardPicture(IMonikerAttributes monikerAttributes, Image image);
    }

    public class ImageMonikerCreator : AxHost, IImageMonikerCreator
    {
        private readonly IImageDataProvider _imageDataProvider;

        public ImageMonikerCreator(IImageDataProvider imageDataProvider) : this()
        {
            _imageDataProvider = imageDataProvider;
        }

        private ImageMonikerCreator() : base(Guid.Empty.ToString())
        {
        }

        public Icon CreateIcon(IMonikerAttributes monikerAttributes, Image image)
        {
            Icon icon;

            var maxDimension = monikerAttributes.GetMaxDimension();
            var imageBytes = _imageDataProvider.GetImageBytes(image);

            image.Dispose();

            using (var memoryStream = new MemoryStream())
            {
                var iconBytes = _imageDataProvider.GetIconBytes(imageBytes, maxDimension);

                memoryStream.Write(iconBytes, offset: 0, count: iconBytes.Length);
                memoryStream.Write(imageBytes, offset: 0, count: imageBytes.Length);
                memoryStream.Position = 0;

                icon = new Icon(memoryStream);
            }

            return icon;
        }

        public Image CreateImage(IMonikerAttributes monikerAttributes, IVsUIObject vsUiObject)
        {
            try
            {
                if (vsUiObject == null)
                {
                    return null;
                }

                var colorType = monikerAttributes.GetColorType();
                var themedColor = _imageDataProvider.GetThemedColor(colorType);

                if (!themedColor.HasValue)
                {
                    return null;
                }

                var imageMoniker = _imageDataProvider.GetImageMoniker(vsUiObject);
                var image = _imageDataProvider.GetImage(imageMoniker, themedColor.Value);

                return image;
            }
            catch
            {
                return null;
            }
        }

        public StdPicture CreateStandardPicture(IMonikerAttributes monikerAttributes, Image image)
        {
            var pictureDisp = (IPictureDisp)GetIPictureDispFromPicture(image);
            var standardPicture = (dynamic)pictureDisp;

            return standardPicture;
        }
    }
}
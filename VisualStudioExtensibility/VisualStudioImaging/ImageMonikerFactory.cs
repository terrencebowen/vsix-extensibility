using stdole;
using System.Drawing;

namespace VisualStudioImaging
{
    public interface IImageMonikerFactory
    {
        Icon GetIcon(IMonikerAttributes monikerAttributes);
        Image GetImage(IMonikerAttributes monikerAttributes);
        StdPicture GetStandardPicture(IMonikerAttributes monikerAttributes);
    }

    public class ImageMonikerFactory : IImageMonikerFactory
    {
        private readonly IImageMonikerCreator _imageMonikerCreator;
        private readonly IImageDataProvider _imageDataProvider;

        public ImageMonikerFactory(IImageMonikerCreator imageMonikerCreator, IImageDataProvider imageDataProvider)
        {
            _imageMonikerCreator = imageMonikerCreator;
            _imageDataProvider = imageDataProvider;
        }

        public Icon GetIcon(IMonikerAttributes monikerAttributes)
        {
            var imageMoniker = monikerAttributes.ImageMoniker;
            var imageAttributes = monikerAttributes.GetImageAttributes();
            var vsUiObject = _imageDataProvider.GetVsUiObject(imageMoniker, imageAttributes);
            var image = _imageMonikerCreator.CreateImage(monikerAttributes, vsUiObject);
            var icon = _imageMonikerCreator.CreateIcon(monikerAttributes, image);

            return icon;
        }

        public Image GetImage(IMonikerAttributes monikerAttributes)
        {
            var imageMoniker = monikerAttributes.ImageMoniker;
            var imageAttributes = monikerAttributes.GetImageAttributes();
            var vsUiObject = _imageDataProvider.GetVsUiObject(imageMoniker, imageAttributes);
            var image = _imageMonikerCreator.CreateImage(monikerAttributes, vsUiObject);

            return image;
        }

        public StdPicture GetStandardPicture(IMonikerAttributes monikerAttributes)
        {
            var imageMoniker = monikerAttributes.ImageMoniker;
            var imageAttributes = monikerAttributes.GetImageAttributes();
            var vsUiObject = _imageDataProvider.GetVsUiObject(imageMoniker, imageAttributes);
            var image = _imageMonikerCreator.CreateImage(monikerAttributes, vsUiObject);
            var standardPicture = _imageMonikerCreator.CreateStandardPicture(monikerAttributes, image);

            return standardPicture;
        }
    }
}
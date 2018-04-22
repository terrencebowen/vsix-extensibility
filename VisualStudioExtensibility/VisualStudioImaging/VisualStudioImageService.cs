using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using stdole;
using System;
using System.Drawing;

namespace VisualStudioImaging
{
    public interface IVisualStudioImageService
    {
        Icon GetIcon(ImageMoniker imageMoniker, ColorTheme colorTheme, int squareSize);
        Icon GetIcon(ImageMoniker imageMoniker, ColorTheme colorTheme, int width, int height);
        Image GetImage(ImageMoniker imageMoniker, ColorTheme colorTheme, int squareSize);
        Image GetImage(ImageMoniker imageMoniker, ColorTheme colorTheme, int width, int height);
        StdPicture GetStandardPicture(ImageMoniker imageMoniker, ColorTheme colorTheme, int squareSize);
        StdPicture GetStandardPicture(ImageMoniker imageMoniker, ColorTheme colorTheme, int width, int height);
    }

    public class VisualStudioImageService : Package, IVisualStudioImageService
    {
        private readonly IImageMonikerFactory _imageMonikerFactory;

        public VisualStudioImageService(IVsPackage vsPackage)
        {
            var isServiceUnavailable = vsPackage == null;

            if (isServiceUnavailable)
            {
                throw new ArgumentNullException($"{nameof(VisualStudioImageService)} must be instantiated in {nameof(Initialize)} method of {nameof(Package)}.");
            }

            ThreadHelper.ThrowIfNotOnUIThread();

            var vsUiShell = GetGlobalService(typeof(SVsUIShell)) as IVsUIShell6;
            var vsImageService = GetGlobalService(typeof(SVsImageService)) as IVsImageService2;
            var imageDataProvider = new ImageDataProvider(vsImageService, vsUiShell);
            var imageMonikerCreator = new ImageMonikerCreator(imageDataProvider);

            _imageMonikerFactory = new ImageMonikerFactory(imageMonikerCreator, imageDataProvider);
        }

        public Icon GetIcon(ImageMoniker imageMoniker, ColorTheme colorTheme, int squareSize)
        {
            const _UIImageType uiImageType = _UIImageType.IT_Icon;

            var monikerAttributes = new MonikerAttributes(imageMoniker, colorTheme, uiImageType, squareSize, squareSize);
            var icon = _imageMonikerFactory.GetIcon(monikerAttributes);

            return icon;
        }

        public Icon GetIcon(ImageMoniker imageMoniker, ColorTheme colorTheme, int width, int height)
        {
            var monikerAttributes = new MonikerAttributes(imageMoniker, colorTheme, _UIImageType.IT_Icon, width, height);
            var icon = _imageMonikerFactory.GetIcon(monikerAttributes);

            return icon;
        }

        public Image GetImage(ImageMoniker imageMoniker, ColorTheme colorTheme, int squareSize)
        {
            var monikerAttributes = new MonikerAttributes(imageMoniker, colorTheme, _UIImageType.IT_Bitmap, width: squareSize, height: squareSize);
            var image = _imageMonikerFactory.GetImage(monikerAttributes);

            return image;
        }

        public Image GetImage(ImageMoniker imageMoniker, ColorTheme colorTheme, int width, int height)
        {
            var monikerAttributes = new MonikerAttributes(imageMoniker, colorTheme, _UIImageType.IT_Bitmap, width, height);
            var image = _imageMonikerFactory.GetImage(monikerAttributes);

            return image;
        }

        public StdPicture GetStandardPicture(ImageMoniker imageMoniker, ColorTheme colorTheme, int squareSize)
        {
            var monikerAttributes = new MonikerAttributes(imageMoniker, colorTheme, _UIImageType.IT_Bitmap, width: squareSize, height: squareSize);
            var standardPicture = _imageMonikerFactory.GetStandardPicture(monikerAttributes);

            return standardPicture;
        }

        public StdPicture GetStandardPicture(ImageMoniker imageMoniker, ColorTheme colorTheme, int width, int height)
        {
            var monikerAttributes = new MonikerAttributes(imageMoniker, colorTheme, _UIImageType.IT_Bitmap, width, height);
            var standardPicture = _imageMonikerFactory.GetStandardPicture(monikerAttributes);

            return standardPicture;
        }
    }
}
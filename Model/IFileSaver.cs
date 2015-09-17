using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

///----------------------------------------------------------------------------------------------------------------------------------------------------------------
///
///         IFileSaver.cs
///         Descritpion:  Interface to save image in a specific format
///                      
///         Remark: - The extension in class name correspond extension to use to save the image
///                 - Each class provide the Save method to properly save an image under Photoshop in the related format
///                 - Use the factory to get an IFileSaver instance
/// 
///----------------------------------------------------------------------------------------------------------------------------------------------------------------
namespace ImgProcPhotoshop.MyPhotoshop
{
    using ps = Photoshop;
    using PhotoshopTypeLibrary;
    using cst = PhotoshopTypeLibrary.PSConstants;

    /// <summary>
    /// IFileSaver Basic Interface
    /// </summary>
    public interface IFileSaver
    {
        /// <summary>
        /// Save the Photoshop document in the related format
        /// </summary>
        /// <param name="doc">Photoshop Document to save</param>
        /// <param name="file">File to save the Photoshop document in</param>
        void Save(ref ps.Document doc, FileInfo file);
    }

    /// <summary>
    /// Provider of IFileSaver instance
    /// </summary>
    /// <remarks>
    ///  Use this class to get an instance of IFileSaver
    /// </remarks>
    public static class IFileSaverFactory
    {
        /// <summary>
        /// Get the specified IFileSaver Instance
        /// </summary>
        /// <param name="ext"> Extension of the IFileSaver to get</param>
        /// <returns></returns>
        public static IFileSaver GetInstance(ImageExtensions ext)
        {
            switch (ext)
            {
                case ImageExtensions.ai:
                    return null;
                    //return new Ai();
                    break;
                case ImageExtensions.eps:
                    return new EpsSaver();
                    break;
                case ImageExtensions.gif:
                    return new GifSaver();
                    break;
                case ImageExtensions.jpeg:
                    return new JpgSaver();
                    break;
                case ImageExtensions.jpg:
                    return new JpgSaver();
                    break;
                case ImageExtensions.png:
                    return new PngSaver();
                    break;
                case ImageExtensions.tif:
                    return new TifSaver();
                    break;
                case ImageExtensions.unknown:
                    throw new ApplicationException("IFileSaver: Unknown Extension");
                    return null;
                default:
                    throw new ApplicationException("IFileSaver: Unknown Extension");
                    return null;
                    break;
            }
        }
    }

    /// <summary>
    /// PNG: IFileSaver
    /// </summary>
    public class PngSaver : IFileSaver
    {
        private const ps.PsSaveDocumentType _format = ps.PsSaveDocumentType.psPNGSave; // Save as png
        private const bool _interlaced = false;//Progressive download of the image
        private const bool _png8 = false; // Uses 24bits rather than 8bits 
        private const bool _transparency = true; // Enable transparency;
        private const int _quality = 70; // Percentage quality of the png generated


        public void Save(ref ps.Document doc, FileInfo file)
        {
            // Save png for web format
            ps.ExportOptionsSaveForWebClass SaveOpts = new ps.ExportOptionsSaveForWebClass();
            SaveOpts.Format = _format;
            SaveOpts.Interlaced = _interlaced;
            SaveOpts.PNG8 = _png8;
            SaveOpts.Transparency = _transparency;

            ps.RGBColorClass _rgbColor = new ps.RGBColorClass();
            _rgbColor.Red = 255;
            _rgbColor.Green = 255;
            _rgbColor.Blue = 255;

            SaveOpts.MatteColor = _rgbColor;

            doc.Export(file.FullName, ps.PsExportType.psSaveForWeb, SaveOpts);


        }
    }

    /// <summary>
    /// JPG: IFileSaver
    /// </summary>
    public class JpgSaver : IFileSaver
    {
        private const ps.PsSaveDocumentType _format = ps.PsSaveDocumentType.psJPEGSave; // Save as png
        private const bool _interlaced = false;//Progressive download of the image
        private const int _quality = 70; // Percentage quality of the png generated


        public void Save(ref ps.Document doc, FileInfo file)
        {
            // Save png for web format
            ps.ExportOptionsSaveForWebClass SaveOpts = new ps.ExportOptionsSaveForWebClass();
            SaveOpts.Format = _format;
            SaveOpts.Interlaced = _interlaced;


            ps.RGBColorClass _rgbColor = new ps.RGBColorClass();
            _rgbColor.Red = 255;
            _rgbColor.Green = 255;
            _rgbColor.Blue = 255;

            SaveOpts.MatteColor = _rgbColor;

            doc.Export(file.FullName, ps.PsExportType.psSaveForWeb, SaveOpts);

        }
    }

    /// <summary>
    /// TIF: IFileSaver
    /// </summary>
    public class TifSaver : IFileSaver
    {
        private const bool _alphaChannel = true;
        private const bool _annotations = true;
        private const ps.PsByteOrderType _byteOrder = ps.PsByteOrderType.psIBMByteOrder; // Windows byte order 
        private const int _jpgQuality = 7;
        private const ps.PsLayerCompressionType _layerCompression = ps.PsLayerCompressionType.psRLELayerCompression;
        private const bool _layers = true;
        private const bool _spotColors = true;
        private const bool _transparency = true;

        public void Save(ref ps.Document doc, FileInfo file)
        {
            // Save png for web format
            ps.TiffSaveOptionsClass SaveOpts = new ps.TiffSaveOptionsClass();
            SaveOpts.AlphaChannels = _alphaChannel;
            SaveOpts.Annotations = _annotations;
            SaveOpts.ByteOrder = _byteOrder;
            SaveOpts.JPEGQuality = _jpgQuality;
            SaveOpts.LayerCompression = _layerCompression;
            SaveOpts.Layers = _layers;
            SaveOpts.SpotColors = _spotColors;
            SaveOpts.Transparency = _transparency;

            doc.SaveAs(file.FullName, SaveOpts, true, ps.PsExtensionType.psLowercase);

        }
    }

    /// <summary>
    /// GIF: IFileSaver
    /// </summary>
    public class GifSaver : IFileSaver
    {
        private const ps.PsSaveDocumentType _format = ps.PsSaveDocumentType.psCompuServeGIFSave; // Save as png
        private const bool _interlaced = false;//Progressive download of the image
        private const int _quality = 70; // Percentage quality of the png generated

        public void Save(ref ps.Document doc, FileInfo file)
        {
            // Save png for web format
            ps.ExportOptionsSaveForWebClass SaveOpts = new ps.ExportOptionsSaveForWebClass();
            SaveOpts.Format = _format;
            SaveOpts.Interlaced = _interlaced;


            ps.RGBColorClass _rgbColor = new ps.RGBColorClass();
            _rgbColor.Red = 255;
            _rgbColor.Green = 255;
            _rgbColor.Blue = 255;

            SaveOpts.MatteColor = _rgbColor;

            doc.Export(file.FullName, ps.PsExportType.psSaveForWeb, SaveOpts);


        }
    }

    /// <summary>
    /// EPS: IFileSaver
    /// </summary>
    public class EpsSaver : IFileSaver
    {
        private const ps.PsPreviewType _preview = ps.PsPreviewType.psEightBitTIFF;
        private const bool _vectorData = true;

        public void Save(ref ps.Document doc, FileInfo file)
        {
            // Save png for web format
            ps.EPSSaveOptionsClass SaveOpts = new ps.EPSSaveOptionsClass();
            SaveOpts.Preview = _preview;
            SaveOpts.VectorData = _vectorData;
            SaveOpts.Interpolation = true;
            SaveOpts.Encoding = ps.PsSaveEncoding.psJPEGMedium;


            doc.SaveAs(file.FullName, SaveOpts, true, ps.PsExtensionType.psLowercase);

        }
    }

    /// <summary>
    /// AI: IFileSaver
    /// </summary>
    /// <remarks>
    /// NOT IMPLEMENTED
    /// </remarks>
    public class AiSaver : IFileSaver
    {
        private const ps.PsIllustratorPathType _path = ps.PsIllustratorPathType.psAllPaths;

        public AiSaver()
        {
            //TODO
            throw new ApplicationException("IFileSaver not implemented for AI");
        }

        public void Save(ref ps.Document doc, FileInfo file)
        {            
            ps.ExportOptionsIllustratorClass SaveOpts = new ps.ExportOptionsIllustratorClass();
            SaveOpts.Path = _path;

            doc.Export(file.FullName, ps.PsExportType.psIllustratorPaths, SaveOpts);

        }
    }




}

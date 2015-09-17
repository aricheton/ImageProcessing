using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

///----------------------------------------------------------------------------------------------------------------------------------------------------------------
///
///         IImageActions.cs
///         Descritpion:  Interface to proceed Photoshop actions depending on the source image format
///                      
///         Remark: - The extension in class name correspond to the source file extension to open
///                 - Use those classes to perform Photoshop actions
///                 - To add a method: 1) add it to the IImageActions interface
///                                    2) implement common behavior in DefaultImageAction Class
///                                    3) add method to each class that implement IImageActions and use DefaultImageAction if to use common behavior otherwise provide specific implementation
///                 - Use factory to get related instance
/// 
///----------------------------------------------------------------------------------------------------------------------------------------------------------------

namespace ImgProcPhotoshop.MyPhotoshop
{
    using ps = Photoshop;
    using PhotoshopTypeLibrary;
    using cst = PhotoshopTypeLibrary.PSConstants;

    /// <summary>
    /// Action pattern for each format
    /// </summary>
    public interface IImageActions
    {
        /// <summary>
        /// Set Photoshop unit
        /// </summary>
        /// <param name="app">Photoshop application object</param>
        /// <param name="doc">Photoshop document object</param>
        void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc);

        /// <summary>
        /// Open the the image with the best constraint based on dimension passed
        /// </summary>
        /// <param name="app">Photoshop application object</param>
        /// <param name="doc">Photoshop document object</param>
        /// <param name="file">File to open</param>
        /// <param name="width">Width constraint</param>
        /// <param name="height">Height constraint</param>
        void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height);

        /// <summary>
        /// Extend Canvas of document based on dimension given
        /// </summary>
        /// <param name="doc">Photoshop document object</param>
        /// <param name="width">Width canvas size</param>
        /// <param name="height">Height canvas size</param>
        /// <remarks>The doc size has to be smaller</remarks>
        void ExtendCanvas(ref ps.Document doc,double width, double height);

        /// <summary>
        /// Set the background of the given doc transparent
        /// </summary>
        /// <param name="app">Photoshop application object</param>
        /// <param name="doc">Photoshop document object</param>
        void SetBackGroundTransparent(ref ps.ApplicationClass app, ref ps.Document doc);

        /// <summary>
        /// Save the document
        /// </summary>
        /// <param name="doc">Photoshop document object</param>
        /// <param name="fileSaver">IFileSave format to save the file</param>
        /// <param name="file">File to save the image in</param>
        void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file);

        /// <summary>
        /// Close the passed document
        /// </summary>
        /// <param name="doc">Photoshop document object</param>
        void Close(ref ps.Document doc);
    }

    /// <summary>
    ///  Provider of IImageActions instance
    /// </summary>
    /// <remarks>
    /// Use this class to get related IImageActions derived class instance
    /// </remarks>
    public static class IImageActionsFactory
    {
        /// <summary>
        /// Get the instance of the related IImageActions
        /// </summary>
        /// <param name="ext"> Extension of the IImageAction instance to provide</param>
        /// <returns></returns>
        public static IImageActions GetInstance(ImageExtensions ext)
        {
            switch (ext)
            {
                case ImageExtensions.ai:
                    return new AiActions();
                    break;
                case ImageExtensions.eps:
                    return new EpsActions();
                    break;
                case ImageExtensions.gif:
                    return new GifActions();
                    break;
                case ImageExtensions.jpeg:
                    return new JpgActions();
                    break;
                case ImageExtensions.jpg:
                    return new JpgActions();
                    break;
                case ImageExtensions.png:
                    return new PngActions();
                    break;
                case ImageExtensions.tif:
                    return new TifActions();
                    break;
                case ImageExtensions.unknown:
                    throw new ApplicationException("IImageActionsFactory: Unknown Extension");
                    return null;
                default:
                    throw new ApplicationException("IImageActionsFactory: Unknown Extension");
                    return null;
                    break;
            }
        }
    }

    /// <summary>
    /// Default behavior for IImageActions
    /// </summary>
    /// <remarks>
    /// Implement common behavior between formats in this class
    /// Avoid redundont code/ Easier to maintain
    /// </remarks>
    public class DefaultImageActions : IImageActions
    {
        private ps.PsUnits _rulerUnits = ps.PsUnits.psPixels;
        private ps.PsTypeUnits _typeUnits = ps.PsTypeUnits.psTypePixels;

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {         
                app.Preferences.RulerUnits = _rulerUnits;
                app.Preferences.TypeUnits = _typeUnits;       
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            //Open any bitmap image
            doc = app.Open(file.FullName);
            ResizeImageWithConstraints(ref doc, width, height);

        }

        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {          
                doc.ResizeCanvas(width, height, ps.PsAnchorPosition.psMiddleCenter);     
        }
       
        public void SetBackGroundTransparent(ref ps.ApplicationClass app,ref ps.Document doc)
        {
            bool isTransparent=PsTools.IsTransparent(ref app, ref doc);
            
            if(!isTransparent){
                doc.ArtLayers[0].Name = "Background";
                doc.LayerComps[0].Name = "Background";
                PsTools.SetBckgTransparentAction(ref app);
            }
        }

        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            fileSaver.Save(ref doc, file);
        }

        public void Close(ref ps.Document doc)
        {
            doc.Close(ps.PsSaveOptions.psDoNotSaveChanges);
        }

        public void ResizeImageWithConstraints(ref ps.Document doc, double width, double height)
        {
            if (doc.Width/doc.Height<=width/height)
            {
                doc.ResizeImage(Height: height, ResampleMethod: ps.PsPDFResampleType.psPDFBicubic);
                    
            }
            else
            {
                doc.ResizeImage(Width: width, ResampleMethod: ps.PsPDFResampleType.psPDFBicubic);
            }
        }

    }

    /// <summary>
    /// EPS: IImageActions
    /// </summary>
    public class EpsActions : IImageActions
    {
        private IImageActions _def;

        private const bool _antiAlias= true;
        private const bool _constrainProportions = true;
        private const ps.PsOpenDocumentMode _mode = ps.PsOpenDocumentMode.psOpenRGB;
        private const int _resolution = 300;

        public EpsActions()
        {
            _def = new DefaultImageActions();
        }

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetUnits(ref app, ref doc);
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            // Set EPS openning options 
            // Open with Width constraint
            ps.EPSOpenOptionsClass EpsOpenOptsWidth = new ps.EPSOpenOptionsClass();
            EpsOpenOptsWidth.AntiAlias = _antiAlias;
            EpsOpenOptsWidth.ConstrainProportions = _constrainProportions;
            EpsOpenOptsWidth.Mode = _mode;
            EpsOpenOptsWidth.Resolution = _resolution;
            EpsOpenOptsWidth.Width = PsTools.GetPixelUnit(width,_resolution);

            // Open with Height constraint
            ps.EPSOpenOptionsClass EpsOpenOptsHeight = new ps.EPSOpenOptionsClass();
            EpsOpenOptsHeight.AntiAlias = _antiAlias;
            EpsOpenOptsHeight.ConstrainProportions = _constrainProportions;
            EpsOpenOptsHeight.Mode = _mode;
            EpsOpenOptsHeight.Resolution = _resolution;
            EpsOpenOptsHeight.Height = PsTools.GetPixelUnit(height, _resolution);


            doc = app.Open(file.FullName, EpsOpenOptsWidth);


            if (doc.Height > height)
            {
                Close(ref doc);
                doc = app.Open(file.FullName, EpsOpenOptsHeight);
            }

        }

        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {
           _def.ExtendCanvas(ref doc, width, height);
        }

        public void SetBackGroundTransparent(ref ps.ApplicationClass app,ref ps.Document doc)
        {
            _def.SetBackGroundTransparent(ref app, ref doc);
        }
               
        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            _def.Save(ref doc, fileSaver, file);
        }
               
        public void Close(ref ps.Document doc)
        {
            _def.Close(ref doc);
        }
        

    }

    /// <summary>
    /// JPG: IImageActions
    /// </summary>
    public class JpgActions : IImageActions
    {
        private IImageActions _def;

        public JpgActions()
        {
            _def = new DefaultImageActions();
        }

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetUnits(ref app, ref doc);
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            _def.OpenWithConstraints(ref app, ref doc, file, width, height);
        }

        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {
            _def.ExtendCanvas(ref doc, width, height);
        }

        public void SetBackGroundTransparent(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetBackGroundTransparent(ref app, ref doc);
        }

        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            _def.Save(ref doc, fileSaver, file);
        }

        public void Close(ref ps.Document doc)
        {
            _def.Close(ref doc);
        }
    }

    /// <summary>
    /// GIF: IImageActions
    /// </summary>
    public class GifActions : IImageActions
    {
        private IImageActions _def;

        public GifActions()
        {
            _def = new DefaultImageActions();
        }

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {
             _def.SetUnits(ref app, ref doc);
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            _def.OpenWithConstraints(ref app, ref doc, file, width, height);
        }

        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {
            _def.ExtendCanvas(ref doc, width, height);
        }

        public void SetBackGroundTransparent(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetBackGroundTransparent(ref app, ref doc);
        }

        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            _def.Save(ref doc, fileSaver, file);
        }

        public void Close(ref ps.Document doc)
        {
            _def.Close(ref doc);
        }
    }

    /// <summary>
    /// TIF: IImageActions
    /// </summary>
    public class TifActions : IImageActions
    {
        private IImageActions _def;

        public TifActions()
        {
            _def = new DefaultImageActions();
        }

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetUnits(ref app, ref doc);
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            _def.OpenWithConstraints(ref app, ref doc, file, width, height);
        }

        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {
            _def.ExtendCanvas(ref doc, width, height);
        }

        public void SetBackGroundTransparent(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetBackGroundTransparent(ref app, ref doc);
        }

        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            _def.Save(ref doc, fileSaver, file);
        }

        public void Close(ref ps.Document doc)
        {
            _def.Close(ref doc);
        }
    }

    /// <summary>
    /// PNG: IImageActions
    /// </summary>
    public class PngActions : IImageActions
    {
        private IImageActions _def;

        public PngActions()
        {
            _def = new DefaultImageActions();
        }

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetUnits(ref app, ref doc);
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            _def.OpenWithConstraints(ref app, ref doc, file, width, height);
        }

        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {
             _def.ExtendCanvas(ref doc, width, height);
        }

        public void SetBackGroundTransparent(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetBackGroundTransparent(ref app, ref doc);
        }

        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            _def.Save(ref doc, fileSaver, file);
        }

        public void Close(ref ps.Document doc)
        {
            _def.Close(ref doc);
        }
    }

    /// <summary>
    /// AI: IImageActions
    /// </summary>
    /// <remarks>
    /// NOT IMPLEMENTED
    /// </remarks>
    // Currently not opening .AI with Open dimensions set up  maybe resize the image
    public class AiActions : IImageActions
    {
        private IImageActions _def;

            private const bool _antiAlias = true;
            private const ps.PsBitsPerChannelType _bitsPerChannel = ps.PsBitsPerChannelType.psDocument8Bits;
            private const bool _constrainProportions = true;
            private const bool _suppressWarnings = true;
            private const ps.PsCropToType _cropPage = ps.PsCropToType.psMediaBox;
            private const ps.PsOpenDocumentMode _mode = ps.PsOpenDocumentMode.psOpenRGB;
            private const int _resolution = 300;
            private const bool _usePageNumber = true;
            private const int _page = 1;


        public AiActions()
        {
            throw new ApplicationException("AI: IImageActions not implemented");
            _def = new DefaultImageActions();
        }

        public void SetUnits(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetUnits(ref app, ref doc);
        }

        public void OpenWithConstraints(ref ps.ApplicationClass app, ref ps.Document doc, FileInfo file, double width, double height)
        {
            //Width Constraint
            ps.PDFOpenOptionsClass PdfOpenOptsWidth = new ps.PDFOpenOptionsClass();
            PdfOpenOptsWidth.AntiAlias = _antiAlias;
            PdfOpenOptsWidth.BitsPerChannel = _bitsPerChannel;
            PdfOpenOptsWidth.ConstrainProportions = _constrainProportions;
            PdfOpenOptsWidth.SuppressWarnings = _suppressWarnings;
            PdfOpenOptsWidth.CropPage = _cropPage;
            PdfOpenOptsWidth.Mode = _mode;
            PdfOpenOptsWidth.Page = _page;
            PdfOpenOptsWidth.UsePageNumber=_usePageNumber;
            PdfOpenOptsWidth.Resolution = _resolution;
            PdfOpenOptsWidth.Width = PsTools.GetPixelUnit(width, _resolution);

            //Height Constraint
            ps.PDFOpenOptionsClass PdfOpenOptsHeight = new ps.PDFOpenOptionsClass();
            PdfOpenOptsHeight.AntiAlias = _antiAlias;
            PdfOpenOptsHeight.BitsPerChannel = _bitsPerChannel;
            PdfOpenOptsHeight.ConstrainProportions = _constrainProportions;
            PdfOpenOptsHeight.SuppressWarnings = _suppressWarnings;
            PdfOpenOptsHeight.CropPage = _cropPage;
            PdfOpenOptsHeight.Mode = _mode;
            PdfOpenOptsWidth.UsePageNumber = _usePageNumber;
            PdfOpenOptsHeight.Page = _page;
            PdfOpenOptsHeight.Resolution = _resolution;
            PdfOpenOptsHeight.Height = PsTools.GetPixelUnit(height, _resolution);

            //Open the file

            doc = app.Open(file.FullName, PdfOpenOptsWidth);


            if (doc.Height > height)
            {
                Close(ref doc);
                //Open the file
                doc = app.Open(file.FullName, PdfOpenOptsHeight);
            }

        }
        
        public void ExtendCanvas(ref ps.Document doc, double width, double height)
        {
            _def.ExtendCanvas(ref doc, width, height);
        }
        
        public void SetBackGroundTransparent(ref ps.ApplicationClass app, ref ps.Document doc)
        {
            _def.SetBackGroundTransparent(ref app, ref doc);
        }
        
        public void Save(ref ps.Document doc, IFileSaver fileSaver, FileInfo file)
        {
            _def.Save(ref doc, fileSaver, file);
        }
       
        public void Close(ref ps.Document doc)
        {
            _def.Close(ref doc);
        }
    }

}

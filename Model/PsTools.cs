using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

///----------------------------------------------------------------------------------------------------------------------------------------------------------------
///
///         PsTools.cs
///         Descritpion: Static methods to proceed actions on photoshop
///                      
///         Remark: - Use PsTools staticly to proceed Photoshop actions
///                 - Convention: Those methods should be called by IImageAction objects
///                 - Some action are taken from ActionListener Plungin from Photoshop: https://helpx.adobe.com/photoshop/kb/plug-ins-photoshop-cs61.html#id_68969
///                 - See info: http://www.pcpix.com/Photoshop/
///                 
/// 
///----------------------------------------------------------------------------------------------------------------------------------------------------------------

namespace ImgProcPhotoshop.MyPhotoshop
{
    using ps = Photoshop;
    using PhotoshopTypeLibrary;
    using cst = PhotoshopTypeLibrary.PSConstants;

    //Enumeration of extension that can be process
    public enum ImageExtensions
    {
        png,
        jpg,
        jpeg,
        tif,
        gif,
        eps,
        ai,
        unknown,
    }

    public static class PsTools
    {
        // Action Script Name of the Action in Photoshop to set Background transparent
        // Becarefull this string is base on the name of the action inside of photoshop itself
        private const string  _transparentBckgActionName="White BG Removal - with Color";
        private const string _transparentBckgActionFolder = "Media Militia - Removal Techniques";

        /// <summary>
        /// Magic Wand selection
        /// </summary>
        /// <param name="app">Photoshop Application Object</param>
        /// <param name="x">x-Axis pixel source selection</param>
        /// <param name="y">y-Axis pixel source selection</param>
        /// <param name="tolerance"> Selection tolerance</param>
        /// <param name="antiAlias"> Anti-Alias option</param>
        /// <param name="contiguous"></param>
        /// <param name="mergedLayer"></param>
        public static void MagicWand(ps.ApplicationClass app,int x, int y, int tolerance = 32, bool antiAlias = true, bool contiguous = false, bool mergedLayer = false)
        {
                // Action recorded on photoshop and transcripted by the Action Listener script
                // This code is the result of the Action Listener script
                ps.ActionDescriptorClass desc = new ps.ActionDescriptorClass();
                ps.ActionReferenceClass refe = new ps.ActionReferenceClass();
                refe.PutProperty((int)cst.phClassChannel, (int)cst.phKeySelection);
                desc.PutReference((int)cst.phClassNull, refe);

                ps.ActionDescriptorClass positionDesc = new ps.ActionDescriptorClass();
                positionDesc.PutUnitDouble((int)cst.phEnumHorizontal, (int)cst.phUnitDistance, x);
                positionDesc.PutUnitDouble((int)cst.phEnumVertical, (int)cst.phUnitDistance, y);

                desc.PutObject((int)cst.phKey_Source, (int)cst.phClassPoint, positionDesc);
                desc.PutInteger((int)cst.phKeyTolerance, tolerance);// tolerance  
                desc.PutBoolean((int)cst.phEnumMerged, mergedLayer);// sample all layers  
                desc.PutBoolean((int)cst.phKeyContiguous, contiguous);//  contiguous  
                desc.PutBoolean((int)cst.phClassAntiAliasedPICTAcquire, antiAlias);// anti-alias  

                app.ExecuteAction((int)cst.phEventSet, desc, ps.PsDialogModes.psDisplayNoDialogs); 
        }

        /// <summary>
        /// Delete the current selection
        /// </summary>
        /// <param name="doc">Photoshop Document Object</param>
        public static void DeleteSelection(ref ps.Document doc)
        {            
            doc.Selection.Clear();
        }

        /// <summary>
        /// Determine if the background of the document is transparent
        /// </summary>
        /// <param name="app">Photoshop Application Object</param>
        /// <param name="doc">Photoshop Document Object</param>
        /// <returns>if document background is transparent </returns>
        public static bool IsTransparent(ref ps.ApplicationClass app,ref ps.Document doc)
        {
            //Make Transparent pixels as selection
            var desc = new ps.ActionDescriptorClass();
            var ref1 = new ps.ActionReferenceClass();
            ref1.PutProperty((int)cst.phClassChannel, (int)cst.phKeySelection);
            desc.PutReference((int)cst.phClassNull, ref1);
            var ref2 = new ps.ActionReferenceClass();
            ref2.PutEnumerated((int)cst.phClassChannel, (int)cst.phClassChannel, (int)cst.phEnumTransparency);
            desc.PutReference((int)cst.phKey_Source, ref2);
            app.ExecuteAction((int)cst.phEventSet, desc, ps.PsDialogModes.psDisplayNoDialogs);
                
            // Verify if the Selection is consistent
            if (doc.Selection.Solid)
            {
                Deselect(ref doc); ;
                return false;
            }
            else
            {
                Deselect(ref doc);
                return true;
            }


        }

        /// <summary>
        /// Deslect current document selection
        /// </summary>
        /// <param name="doc">Photoshop Document Object</param>
        public static void Deselect(ref ps.Document doc){

            doc.Selection.Deselect();            
        }

        /// <summary>
        /// Get Photoshop pixel unit value to use in Photoshop
        /// </summary>
        /// <param name="px">Number of pixels</param>
        /// <param name="resolution">Resolution of the document</param>
        /// <returns>Pixel value usable by photoshop</returns>
        public static double GetPixelUnit(double px, double resolution)
        {
            return (double)px * 72 / resolution;
        }

        /// <summary>
        /// Set a color sampler to the specified localization
        /// </summary>
        /// <param name="app">Photoshop Application object</param>
        /// <param name="doc">Photoshop Document object</param>
        /// <param name="x">X-Axis color sampler localization</param>
        /// <param name="y">Y-Axis color sampler localization</param>
        public static void SetColorSampler(ref ps.ApplicationClass app, ref ps.Document doc, int x, int y)
        {
                // Code from action listener plugin
                int idMk = (int)cst.phEventMake;
                var desc3 = new ps.ActionDescriptorClass();
                int idnull = (int)cst.phClassNull;
                var ref1 = new ps.ActionReferenceClass();
                int idClSm = (int)cst.phClassColorSampler;
                ref1.PutClass(idClSm);
                desc3.PutReference(idnull, ref1);
                int idPstn = (int)cst.phKeyPosition;
                var desc4 = new ps.ActionDescriptorClass();
                int idHrzn = (int)cst.phEnumHorizontal;
                int idPxl = (int)cst.phUnitPixels;
                desc4.PutUnitDouble(idHrzn, idPxl, (double)x);
                int idVrtc = (int)cst.phEnumVertical;
                desc4.PutUnitDouble(idVrtc, idPxl, (double)y);
                int idPnt = (int)cst.phClassPoint;
                desc3.PutObject(idPstn, idPnt, desc4);
                app.ExecuteAction(idMk, desc3, ps.PsDialogModes.psDisplayNoDialogs);

        }
        
        /// <summary>
        /// Set transparent the background of the active document
        /// </summary>
        /// <param name="app">Photoshop Application object</param>
        public static void SetBckgTransparentAction(ref ps.ApplicationClass app)
        {
               // Call action already stored in photoshop
               app.DoAction(_transparentBckgActionName,_transparentBckgActionFolder);
        }

        /// <summary>
        /// Get the related Extension enumeration
        /// </summary>
        /// <param name="extStr">Extension string to get the enumeration from</param>
        /// <returns>Related Image Extension Enumeration</returns>
        public static ImageExtensions GetExtensionsEnum(string extStr)
        {
            //Remove eventual extension dot
            string formatExt=extStr.Replace(".","");

            //Return appropriate enum extension
            switch (formatExt)
            {
                case "png":
                    return ImageExtensions.png;
                    break;
                case "jpg":
                    return ImageExtensions.jpg;
                    break;
                case "jpeg":
                    return ImageExtensions.jpeg;
                    break;
                case "tif":
                    return ImageExtensions.tif;
                    break;
                case "gif":
                    return ImageExtensions.gif;
                    break;
                case "eps":
                    return ImageExtensions.eps;
                    break;
                case "ai":
                    return ImageExtensions.ai;
                    break;
                default:
                    throw new ApplicationException("PsTools.GetExtensionEnum(): UnknownFormat");
                    return ImageExtensions.unknown;
            }
        }
        
        /// <summary>
        /// Get the related string Extension
        /// </summary>
        /// <param name="extEnum">Extension Enumeration to get the string from</param>
        /// <returns>Related Extension string</returns>
        public static string  GetExtensionsString(MyPhotoshop.ImageExtensions extEnum)
        {
            //Return appropriate string extension
            switch (extEnum)
            {
                case ImageExtensions.png:
                    return "png";
                    break;
                case ImageExtensions.jpg:
                    return "jpg";
                    break;
                case ImageExtensions.jpeg:
                    return  "jpeg";
                    break;
                case ImageExtensions.tif:
                    return "tif";
                    break;
                case ImageExtensions.gif:
                    return "gif";
                    break;
                case ImageExtensions.eps:
                    return "eps";
                    break;
                case ImageExtensions.ai:
                    return "ai";
                    break;
                case ImageExtensions.unknown:
                    throw new ApplicationException("PsTools.GetExtensionsString(): Unknown Extension");
                    return null;
                    break;
                default:
                    throw new ApplicationException("PsTools.GetExtensionsString(): Unknown Extension");
                    return null;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;

namespace OfficeToImagesConverter4dots
{
    public class WordImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();

        public List<FromToWordImage> ExtractedFromToWordImages = new List<FromToWordImage>();

        public string err = "";

        private object missing = System.Reflection.Missing.Value;
        private object yes = true;
        private object no = false;
        private object oDocuments = null;
        private object doc = null;
        private object Shapes = null;
        private object ShapesCount = null;
        private object Shape = null;


        private object Sections = null;
        private object Headers = null;
        private object HeaderShapes = null;

        public bool ExtractImages(string filepath, string slideranges)
        {
            err = "";

            Image image = null;
            object WordAppSelection = null;
            object HeaderRangeShape = null;
            int iHeaderRangeShapesCount = -1;
            object HeaderRangeShapesCount = null;
            object HeaderRangeShapes = null;
            object HeaderRange = null;
            object Header = null;

            try
            {
                string tempfile = System.IO.Path.GetTempFileName() + System.IO.Path.GetExtension(filepath);

                System.IO.File.Copy(filepath, tempfile, true);

                OfficeHelper.CreateWordApplication();

                oDocuments = OfficeHelper.WordApp.GetType().InvokeMember("Documents", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                OfficeHelper.WordApp.GetType().InvokeMember("ActivePrinter", BindingFlags.SetProperty, null, OfficeHelper.WordApp, new object[] { "Microsoft XPS Document Writer" });

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { tempfile });

                System.Threading.Thread.Sleep(200);

                OfficeHelper.WordApp.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, OfficeHelper.WordApp, new object[] { true });

                System.Threading.Thread.Sleep(200);

                OfficeHelper.WordApp.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, OfficeHelper.WordApp, null);

                System.Threading.Thread.Sleep(200);

                doc.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, doc, null);

                System.Threading.Thread.Sleep(200);

                object isMissing = Type.Missing;
                object isPrintToFile = (object)true;
                object oFalse = (object)false;

                string imgfp = tempfile + ".xps";

                object oRange = (object)0;
                object oCopies = (object)1;
                object oPageType = (object)0;
                object oItems = (object)0;
                object oPages = "";

                if (slideranges != string.Empty)
                {
                    oRange = (object)4;
                    oPages = slideranges;
                }


                object[] oParam = new object[] {

                    isMissing, false, oRange, imgfp, isMissing,isMissing, oItems, oCopies, oPages, oPageType, true, true,
                            isMissing, false, isMissing, isMissing, isMissing, isMissing};

                doc.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, doc, oParam);

                while (true)
                {
                    object obg = OfficeHelper.WordApp.GetType().InvokeMember("BackgroundPrintingStatus", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.WordApp, null);

                    if ((int)obg == 0)
                    {
                        break;
                    }

                    System.Windows.Forms.Application.DoEvents();
                }

                //OfficeHelper.WordApp.GetType().InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, OfficeHelper.WordApp, oParam);

                doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, null);

                System.Threading.Thread.Sleep(200);

                OfficeHelper.QuitWordApplication();

                System.Threading.Thread.Sleep(200);

                OfficeHelper.QuitOfficeApplications();

                System.Threading.Thread.Sleep(200);

                if (System.IO.File.Exists(imgfp))
                {
                    XPSToImageConverter.SaveXpsPageToBitmap(imgfp);

                    List<string> lst = XPSToImageConverter.lst;

                    for (int k = 0; k < lst.Count; k++)
                    {
                        //  string imagefp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, k + 1);
                        string imagefp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, ExtractedFilepaths.Count);

                        Bitmap bmp = new Bitmap(lst[k]);

                        //remove for normal

                        //if (frmAbout.LDT == string.Empty)
                        if (false)
                        {
                            using (Graphics g = Graphics.FromImage(bmp))
                            {
                                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                                Font fontp = new Font(frmMain.Instance.Font.FontFamily, 20, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);

                                SizeF sz = g.MeasureString(Module.ApplicationName + " - 4dots Software - Please Register", fontp);

                                g.DrawString(Module.ApplicationName + " - 4dots Software - Please Register", fontp, Brushes.DarkBlue,
                                    new PointF(bmp.Width - sz.Width, bmp.Height - sz.Height));

                                fontp.Dispose();
                            }
                        }

                        frmOptions.SaveImage(imagefp, bmp);

                        ExtractedFilepaths.Add(imagefp);
                        bmp.Dispose();
                        bmp = null;

                        System.IO.File.Delete(lst[k]);

                        if (frmMain.Instance.FirstOutputDocument == string.Empty)
                        {
                            frmMain.Instance.FirstOutputDocument = imagefp;
                        }

                        System.IO.FileInfo fi = new System.IO.FileInfo(filepath);
                        System.IO.FileInfo fi2 = new System.IO.FileInfo(imagefp);

                        if (Properties.Settings.Default.KeepCreationDate)
                        {
                            fi2.CreationTime = fi.CreationTime;
                        }

                        if (Properties.Settings.Default.KeepLastModificationDate)
                        {
                            fi2.LastWriteTime = fi.LastWriteTime;
                        }
                    }

                    //System.IO.File.Delete(imgfp);
                }
                else
                {
                    throw new Exception("XPS File missing !");
                }
                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Word File to Images") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }
    }

    public class FromToWordImage
    {
        public string WordFilepath = "";
        public string ImageFilepath = "";
        public int ShapeNr = -1;

        public FromToWordImageTypeEnum FromToWordImageType = FromToWordImageTypeEnum.DocumentInlineShape;

        public enum FromToWordImageTypeEnum
        {
            HeaderInlineShape,
            FooterInlineShape,
            DocumentInlineShape,
            DocumentShape
        }
    }
}

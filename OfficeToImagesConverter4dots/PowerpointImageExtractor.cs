using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Reflection;

namespace OfficeToImagesConverter4dots
{
    public class PowerpointImageExtractor
    {
        public List<string> ExtractedFilepaths = new List<string>();        

        public string err = "";

        public bool ExtractImages(string filepath,string slideranges)
        {
            err = "";
            
            object oDocuments = null;
            object doc = null;
            object Sections = null;            

            try
            {
                OfficeHelper.CreatePowerPointApplication();

                oDocuments = OfficeHelper.PPApp.GetType().InvokeMember("Presentations", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, OfficeHelper.PPApp, null);

                doc = oDocuments.GetType().InvokeMember("Open", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, oDocuments, new object[] { filepath });

                /*
                System.Threading.Thread.Sleep(100);

                OfficeHelper.PPApp.GetType().InvokeMember("Activate", BindingFlags.IgnoreReturn | BindingFlags.Public |
                BindingFlags.Static | BindingFlags.InvokeMethod, null, OfficeHelper.PPApp, null);
                */

                System.Threading.Thread.Sleep(200);

                Sections = doc.GetType().InvokeMember("Slides", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, null);

                object SectionsCount = Sections.GetType().InvokeMember("Count", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, Sections, null);
                int iSectionsCount = (int)SectionsCount;

                StringRange stringrange = new StringRange(slideranges);

                for (int m1 = 1; m1 <= iSectionsCount; m1++)
                {
                    try
                    {
                        if (stringrange.IsInRange(m1))
                        {
                            object oSlide = doc.GetType().InvokeMember("Slides", BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, doc, new object[] { m1 });

                            /// string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory,m1);

                            //  object[] oParam = new object[] { imgfp, frmOptions.GetImageExtension().ToUpper() };

                            string imgfp = frmOptions.GetSaveFilepath(filepath, Module.CurrentImagesDirectory, ExtractedFilepaths.Count);

                            string tmpfp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + frmOptions.GetImageExtension().ToUpper());

                            ExtractedFilepaths.Add(imgfp);

                            //object[] oParam = new object[] { imgfp, frmOptions.GetImageExtension().ToUpper() };

                            object[] oParam = new object[] { tmpfp, frmOptions.GetImageExtension().ToUpper() };

                            oSlide.GetType().InvokeMember("Export", BindingFlags.InvokeMethod, null, oSlide, oParam);

                            oSlide = null;

                            frmOptions.SaveImage(imgfp, (Bitmap)Image.FromFile(tmpfp));

                            if (frmMain.Instance.FirstOutputDocument == string.Empty)
                            {
                                frmMain.Instance.FirstOutputDocument = imgfp;
                            }

                            System.IO.FileInfo fi = new System.IO.FileInfo(filepath);
                            System.IO.FileInfo fi2 = new System.IO.FileInfo(imgfp);

                            if (Properties.Settings.Default.KeepCreationDate)
                            {
                                fi2.CreationTime = fi.CreationTime;
                            }

                            if (Properties.Settings.Default.KeepLastModificationDate)
                            {
                                fi2.LastWriteTime = fi.LastWriteTime;
                            }
                        }
                    }
                    catch (Exception exm)
                    {
                        err += "Error ! " + exm.Message + "\r\n";
                    }
                }

                oDocuments = null;
                doc = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                err += TranslateHelper.Translate("Error could not Convert Powerpoint to Images") + " : " + filepath + "\r\n" + ex.Message;
                return false;
            }

            return true;
        }                
    }
}

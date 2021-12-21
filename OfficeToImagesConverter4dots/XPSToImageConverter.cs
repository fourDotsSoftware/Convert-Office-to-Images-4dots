using System;
using System.Collections.Generic;
using System.Text;
using System.IO;  
using System.IO.Packaging; 
using System.Windows.Documents;  
using System.Windows.Xps.Packaging;  
using System.Windows.Media.Imaging;
using System.Threading;

namespace OfficeToImagesConverter4dots
{
    class XPSToImageConverter
    {
        public static List<string> lst = null;

        static public List<string> SaveXpsPageToBitmap(string xpsFileName)
        {
            lst = new List<string>();

            Thread thread = new Thread(new ParameterizedThreadStart(SaveXpsPageToBitmapFunction));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start(xpsFileName);
            thread.Join();

            return lst;
        }

        static public void SaveXpsPageToBitmapFunction(object oxpsFileName)
        {
            string xpsFileName = oxpsFileName.ToString();

            XpsDocument xpsDoc = new XpsDocument(xpsFileName, System.IO.FileAccess.Read);
            //FixedDocumentSequence docSeq = xpsDoc.GetFixedDocumentSequence();                       

            IDocumentPaginatorSource docSeq = xpsDoc.GetFixedDocumentSequence();                                              

            // You can get the total page count from docSeq.PageCount   
            for (int pageNum = 0; pageNum < docSeq.DocumentPaginator.PageCount; ++pageNum)
            {
                DocumentPage docPage = docSeq.DocumentPaginator.GetPage(pageNum);
                BitmapImage bitmap = new BitmapImage();
                RenderTargetBitmap renderTarget =
                    new RenderTargetBitmap((int)docPage.Size.Width,
                                            (int)docPage.Size.Height,
                                            96, // WPF (Avalon) units are 96dpi based   
                                            96,
                                            System.Windows.Media.PixelFormats.Default);

                renderTarget.Render(docPage.Visual);

                BitmapEncoder encoder = new BmpBitmapEncoder();  // Choose type here ie: JpegBitmapEncoder, etc  
                encoder.Frames.Add(BitmapFrame.Create(renderTarget));

                string imgfp = xpsFileName + ".Page" + pageNum + ".bmp";

                lst.Add(imgfp);

                FileStream pageOutStream = new FileStream(imgfp, FileMode.Create, FileAccess.Write);
                encoder.Save(pageOutStream);
                pageOutStream.Close();
            }

            return;
        }
   }
}

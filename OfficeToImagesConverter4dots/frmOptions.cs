using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO;

namespace OfficeToImagesConverter4dots
{
    public partial class frmOptions : OfficeToImagesConverter4dots.CustomForm
    {
        public frmOptions()
        {
            InitializeComponent();
        }

        private void btnInsertN_Click(object sender, EventArgs e)
        {
            txtFilename.Text += "[N]";
        }

        private void btnInsertDocFilename_Click(object sender, EventArgs e)
        {
            txtFilename.Text += "[DOC]";
        }

        private void btnInsertDate_Click(object sender, EventArgs e)
        {
            txtFilename.Text += "[DATE]";
        }

        private void btnInsertDateTime_Click(object sender, EventArgs e)
        {
            txtFilename.Text += "[DATETIME]";
        }

        private void frmOutputFilenamePattern_Load(object sender, EventArgs e)
        {
            txtFilename.Text = Properties.Settings.Default.FilenamePattern;

            chkOverwrite.Checked = Properties.Settings.Default.OverwriteExisting;

            rdPNG.Checked = Properties.Settings.Default.SaveImageFormat == 0;
            rdJPG.Checked = Properties.Settings.Default.SaveImageFormat == 1;
            rdGIF.Checked = Properties.Settings.Default.SaveImageFormat == 2;
            rdBMP.Checked = Properties.Settings.Default.SaveImageFormat == 3;

            nudJpegQuality.Value = Properties.Settings.Default.SaveJPEGQuality;

            if (!Properties.Settings.Default.ImageWidth &&
                !Properties.Settings.Default.ImageHeight)
            {
                chkImageSizeAsIs.Checked = true;
            }
            else
            {
                chkWidth.Checked = Properties.Settings.Default.ImageWidth;
                chkHeight.Checked = Properties.Settings.Default.ImageHeight;
            }

            nudWidth.Value = Properties.Settings.Default.ImageWidthValue;
            nudHeight.Value = Properties.Settings.Default.ImageHeightValue;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.FilenamePattern = txtFilename.Text;

            Properties.Settings.Default.OverwriteExisting = chkOverwrite.Checked;

            Properties.Settings.Default.SaveImageFormat = rdPNG.Checked ? 0 : rdJPG.Checked ? 1 : rdGIF.Checked ? 2 : rdBMP.Checked ? 3 : 0;

            Properties.Settings.Default.SaveJPEGQuality = (int)nudJpegQuality.Value;

            Properties.Settings.Default.ImageWidth = chkWidth.Checked;
            Properties.Settings.Default.ImageHeight = chkHeight.Checked;

            Properties.Settings.Default.ImageWidthValue = (int)nudWidth.Value;
            Properties.Settings.Default.ImageHeightValue = (int)nudHeight.Value;

            this.DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        public static string GetSaveFilepath(string docfilepath, string outdir, int m)
        {
            m = m + 1;

            string fp = Properties.Settings.Default.FilenamePattern + GetImageExtension();

            System.IO.FileInfo fi = new System.IO.FileInfo(docfilepath);

            string sDate = fi.LastWriteTime.Year.ToString("D4") + fi.LastWriteTime.Month.ToString("D2") + fi.LastWriteTime.Day.ToString("D2");

            string sDateTime = fi.LastWriteTime.Year.ToString("D4") + fi.LastWriteTime.Month.ToString("D2") + fi.LastWriteTime.Day.ToString("D2")
                + "_" + fi.LastWriteTime.Hour.ToString("D4") + fi.LastWriteTime.Minute.ToString("D2") + fi.LastWriteTime.Second.ToString("D2");

            fp = fp.Replace("[DATE]", sDate).Replace("[DATETIME]", sDateTime).Replace("[DOC]", System.IO.Path.GetFileName(docfilepath));

            int k = m;

            string outfp = "";

            if (m > 0)
            {
                outfp = System.IO.Path.Combine(outdir, fp.Replace("[N]", m.ToString()));
            }
            else
            {
                outfp = System.IO.Path.Combine(outdir, fp);
            }

            //if (Properties.Settings.Default.OverwriteExisting || !System.IO.File.Exists(outfp))            
            if (true)
            {
                return outfp;
            }
            else
            {
                while (true)
                {
                    k++;

                    outfp = System.IO.Path.Combine(outdir, fp.Replace("[N]", k.ToString()));

                    if (!System.IO.File.Exists(outfp))
                    {
                        return outfp;
                    }
                }
            }
        }

        /*
        public static string GetSaveFilepath(string docfilepath,string outdir,int slidenr)
        {
            string fp = Properties.Settings.Default.FilenamePattern + GetImageExtension();

            System.IO.FileInfo fi = new System.IO.FileInfo(docfilepath);

            string sDate = fi.LastWriteTime.Year.ToString("D4") + fi.LastWriteTime.Month.ToString("D2") + fi.LastWriteTime.Day.ToString("D2");

            string sDateTime = fi.LastWriteTime.Year.ToString("D4") + fi.LastWriteTime.Month.ToString("D2") + fi.LastWriteTime.Day.ToString("D2")
                + "_" + fi.LastWriteTime.Hour.ToString("D4") + fi.LastWriteTime.Minute.ToString("D2") + fi.LastWriteTime.Second.ToString("D2");

            fp = fp.Replace("[DATE]", sDate).Replace("[DATETIME]", sDateTime).Replace("[DOC]", System.IO.Path.GetFileName(docfilepath));

            fp = fp.Replace("[N]", slidenr.ToString());

            string outfp = System.IO.Path.Combine(outdir, fp);            

            int k = 1;            
            
            if (Properties.Settings.Default.OverwriteExisting || !System.IO.File.Exists(outfp))
            {
                return outfp;
            }
            else
            {
                while (true)
                {
                    k++;

                    string outfp2 = System.IO.Path.Combine(outdir, System.IO.Path.GetFileNameWithoutExtension(fp) + " - " + k.ToString() + System.IO.Path.GetExtension(fp));

                    if (!System.IO.File.Exists(outfp2))
                    {
                        return outfp2;
                    }
                }
            }                     
        }
        */
        private void nudJpegQuality_ValueChanged(object sender, EventArgs e)
        {
            tbJpegQuality.Value = (int)nudJpegQuality.Value;
        }

        private void tbJpegQuality_ValueChanged(object sender, EventArgs e)
        {
            nudJpegQuality.Value = tbJpegQuality.Value;
        }

        public static bool SaveImage(string filepath, Bitmap bmp1)
        {
            //Module.ShowMessage(filepath);

            try
            {
                if (System.IO.File.Exists(filepath))
                {
                    FileInfo fi = new FileInfo(filepath);
                    fi.Attributes = FileAttributes.Normal;
                    fi.Delete();
                }
            }
            catch { }

            int w = bmp1.Width;
            int h = bmp1.Height;
            decimal dw = (decimal)bmp1.Width;
            decimal dh = (decimal)bmp1.Height;

            decimal dnw = (decimal)Properties.Settings.Default.ImageWidthValue;
            decimal dnh = (decimal)Properties.Settings.Default.ImageHeightValue;

            if (Properties.Settings.Default.ImageWidth && !Properties.Settings.Default.ImageHeight)
            {
                // dw dh
                // dnw x

                dnh = (dnw * dh) / dw;

                w = (int)Properties.Settings.Default.ImageWidthValue;
                h = (int)dnh;
            }
            else if (!Properties.Settings.Default.ImageWidth && Properties.Settings.Default.ImageHeight)
            {
                // dw dh
                // x dnh

                dnw = (dnh * dw) / dh;

                w = (int)dnw;
                h = (int)Properties.Settings.Default.ImageHeightValue;
            }
            else if (Properties.Settings.Default.ImageWidth && Properties.Settings.Default.ImageHeight)
            {
                w = (int)Properties.Settings.Default.ImageWidthValue;
                h = (int)Properties.Settings.Default.ImageHeightValue;
            }
            else if (!Properties.Settings.Default.ImageWidth && !Properties.Settings.Default.ImageHeight)
            {
                w = bmp1.Width;
                h = bmp1.Height;
            }

            //Bitmap bmp = new Bitmap(bmp1.Width, bmp1.Height);

            Bitmap bmp = new Bitmap(w, h);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                bmp.SetResolution(bmp1.HorizontalResolution, bmp1.VerticalResolution);

                //g.DrawImage(bmp1, 0, 0, bmp1.Width, bmp1.Height);

                g.DrawImage(bmp1, new Rectangle(0, 0, w, h), new Rectangle(0, 0, bmp1.Width, bmp1.Height), GraphicsUnit.Pixel);
            }

            if (Properties.Settings.Default.SaveImageFormat == 0)
            {
                bmp.Save(filepath, System.Drawing.Imaging.ImageFormat.Png);
            }
            else if (Properties.Settings.Default.SaveImageFormat == 1)
            {
                ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);

                System.Drawing.Imaging.Encoder myEncoder =
                System.Drawing.Imaging.Encoder.Quality;

                // Create an EncoderParameters object.
                // An EncoderParameters object has an array of EncoderParameter
                // objects. In this case, there is only one
                // EncoderParameter object in the array.
                EncoderParameters myEncoderParameters = new EncoderParameters(1);

                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, (long)Properties.Settings.Default.SaveJPEGQuality);

                myEncoderParameters.Param[0] = myEncoderParameter;
                bmp.Save(filepath, jgpEncoder, myEncoderParameters);
            }
            else if (Properties.Settings.Default.SaveImageFormat == 2)
            {
                bmp.Save(filepath, System.Drawing.Imaging.ImageFormat.Gif);
            }
            else if (Properties.Settings.Default.SaveImageFormat == 3)
            {
                bmp.Save(filepath, System.Drawing.Imaging.ImageFormat.Bmp);
            }

            return true;
        }

        /*
        public static bool SaveImage(string filepath,Bitmap bmp)
        {
            if (Properties.Settings.Default.SaveImageFormat == 0)
            {
                bmp.Save(filepath, System.Drawing.Imaging.ImageFormat.Png);
            }
            else if (Properties.Settings.Default.SaveImageFormat == 1)
            {
                ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);

                System.Drawing.Imaging.Encoder myEncoder =
                System.Drawing.Imaging.Encoder.Quality;

                // Create an EncoderParameters object.
                // An EncoderParameters object has an array of EncoderParameter
                // objects. In this case, there is only one
                // EncoderParameter object in the array.
                EncoderParameters myEncoderParameters = new EncoderParameters(1);

                EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, (long)Properties.Settings.Default.SaveJPEGQuality);
                    
                myEncoderParameters.Param[0] = myEncoderParameter;
                bmp.Save(filepath, jgpEncoder, myEncoderParameters);                
            }
            else if (Properties.Settings.Default.SaveImageFormat == 2)
            {
                bmp.Save(filepath, System.Drawing.Imaging.ImageFormat.Gif);
            }
            else if (Properties.Settings.Default.SaveImageFormat == 3)
            {
                bmp.Save(filepath, System.Drawing.Imaging.ImageFormat.Bmp);
            }

            return true;
        }
        */
        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        public static string GetImageExtension()
        {
            if (Properties.Settings.Default.SaveImageFormat == 0)
            {
                return ".png";
            }
            else if (Properties.Settings.Default.SaveImageFormat == 1)
            {
                return ".jpg";
            }
            else if (Properties.Settings.Default.SaveImageFormat == 2)
            {
                return ".gif";
            }
            else if (Properties.Settings.Default.SaveImageFormat == 3)
            {
                return ".bmp";
            }
            else
            {
                return ".bmp";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtFilename.Text += "[SLIDENR]";
        }

        private void chkImageSizeAsIs_CheckedChanged(object sender, EventArgs e)
        {
            if (chkImageSizeAsIs.Checked)
            {
                chkWidth.Checked = false;
                chkHeight.Checked = false;
            }
        }

        private void chkWidth_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox chk = (CheckBox)sender;

            if (chk.Checked)
            {
                chkImageSizeAsIs.Checked = false;
            }
        }
    }
}

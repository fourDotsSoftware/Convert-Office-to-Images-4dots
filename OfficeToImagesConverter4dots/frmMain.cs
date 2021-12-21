using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;

using System.Text;

using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace OfficeToImagesConverter4dots
{
    public partial class frmMain : CustomForm
    {
        public static frmMain Instance = null;       

        public bool SilentAdd = false;
        public string SilentAddErr = "";

        public bool OperationStopped = false;
        public bool OperationPaused = false;

        public string Err = "";

        private string sOutputDir = "";
        private bool bKeepBackup = false;

        public string FirstOutputDocument = "";

        public int[] de = new int[5];

        public frmMain()
        {
            InitializeComponent();

            Instance = this;

            dt.Columns.Add("filename");
            dt.Columns.Add("slideranges");
            dt.Columns.Add("sizekb");
            dt.Columns.Add("fullfilepath");
            dt.Columns.Add("filedate");
            dt.Columns.Add("rootfolder");

            dgFiles.AutoGenerateColumns = false;

            for (int k = 0; k < de.Length; k++)
            {
                de[k] = 0;
            }
        }

        public DataTable dt = new DataTable("table");

        private bool _IsDirty = false;

        private bool IsDirty
        {
            get { return _IsDirty; }

            set
            {
                _IsDirty = value;

                lblTotal.Text = TranslateHelper.Translate("Total") + " : " + dt.Rows.Count + " " + TranslateHelper.Translate("Documents");
            }
        }


        private void tsdbAddFile_ButtonClick(object sender, EventArgs e)
        {
            openFileDialog1.Filter = Module.OpenFilesFilter;
            openFileDialog1.Multiselect = true;

            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SilentAddErr = "";

                try
                {
                    this.Cursor = Cursors.WaitCursor;

                    for (int k = 0; k < openFileDialog1.FileNames.Length; k++)
                    {
                        AddFile(openFileDialog1.FileNames[k]);
                        RecentFilesHelper.AddRecentFile(openFileDialog1.FileNames[k]);
                    }
                }
                finally
                {
                    this.Cursor = null;

                    if (SilentAddErr != string.Empty)
                    {
                        frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                        f.ShowDialog(this);
                    }
                }
            }
        }

        private void tsdbAddFile_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                SilentAddErr = "";

                AddFile(e.ClickedItem.Text);
                RecentFilesHelper.AddRecentFile(e.ClickedItem.Text);

            }
            finally
            {
                this.Cursor = null;

                if (SilentAddErr != string.Empty)
                {
                    frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                    f.ShowDialog(this);
                }
            }
        }

        private void tsbRemove_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedCellCollection cells = dgFiles.SelectedCells;
            List<DataGridViewRow> rows = new List<DataGridViewRow>();

            for (int k = 0; k < cells.Count; k++)
            {
                if (rows.IndexOf(dgFiles.Rows[cells[k].RowIndex]) < 0)
                {
                    rows.Add(dgFiles.Rows[cells[k].RowIndex]);
                }
            }

            for (int k = 0; k < rows.Count; k++)
            {
                dgFiles.Rows.Remove(rows[k]);
            }

            IsDirty = true;
        }

        private void tsbClear_Click(object sender, EventArgs e)
        {
            DialogResult dres = Module.ShowQuestionDialog(TranslateHelper.Translate("Are you sure that you want clear the added files ?"), TranslateHelper.Translate("Clear Added Files ?"));

            if (dres == DialogResult.Yes)
            {
                dt.Rows.Clear();
            }

            IsDirty = true;
        }

        private void tsdbAddFolder_ButtonClick(object sender, EventArgs e)
        {
            folderBrowserDialog1.SelectedPath = "";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                SilentAddErr = "";

                AddFolder(folderBrowserDialog1.SelectedPath);
                RecentFilesHelper.AddRecentFolder(folderBrowserDialog1.SelectedPath);

                if (SilentAddErr != string.Empty)
                {
                    frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                    f.ShowDialog(this);
                }
            }
        }

        private void tsdbAddFolder_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            SilentAddErr = "";

            AddFolder(e.ClickedItem.Text, "");
            RecentFilesHelper.AddRecentFolder(e.ClickedItem.Text);

            if (SilentAddErr != string.Empty)
            {
                frmError f = new frmError(TranslateHelper.Translate("Error"), SilentAddErr);
                f.ShowDialog(this);
            }
        }

        public void ImportList(string listfilepath)
        {
            string curdir = Environment.CurrentDirectory;

            try
            {
                SilentAdd = true;
                using (StreamReader sr = new StreamReader(listfilepath, Encoding.Default, true))
                {
                    string line = null;

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.StartsWith("#"))
                        {
                            continue;
                        }

                        string filepath = line;
                        string password = "";

                        try
                        {
                            if (line.StartsWith("\""))
                            {
                                int epos = line.IndexOf("\"", 1);

                                if (epos > 0)
                                {
                                    filepath = line.Substring(1, epos - 1);
                                }
                            }
                            else if (line.StartsWith("'"))
                            {
                                int epos = line.IndexOf("'", 1);

                                if (epos > 0)
                                {
                                    filepath = line.Substring(1, epos - 1);
                                }
                            }

                            int compos = line.IndexOf(",");

                            if (compos > 0)
                            {
                                password = line.Substring(compos + 1);

                                if (!line.StartsWith("\"") && !line.StartsWith("'"))
                                {
                                    filepath = line.Substring(0, compos);
                                }

                                if ((password.StartsWith("\"") && password.EndsWith("\""))
                                    || (password.StartsWith("'") && password.EndsWith("'")))
                                {
                                    if (password.Length == 2)
                                    {
                                        password = "";
                                    }
                                    else
                                    {
                                        password = password.Substring(1, password.Length - 2);
                                    }
                                }

                            }
                        }
                        catch (Exception exq)
                        {
                            SilentAddErr += TranslateHelper.Translate("Error while processing List !") + " " + line + " " + exq.Message + "\r\n";
                        }

                        line = filepath;

                        Environment.CurrentDirectory = System.IO.Path.GetDirectoryName(listfilepath);

                        line = System.IO.Path.GetFullPath(line);

                        if (System.IO.File.Exists(line))
                        {
                            AddFile(line, password);
                            /*
                            else
                            {
                                SilentAddErr += TranslateHelper.Translate("Error wrong file type !") + " " + line + "\r\n";
                            }*/
                        }
                        else if (System.IO.Directory.Exists(line))
                        {
                            AddFolder(line, password);
                        }
                        else
                        {
                            SilentAddErr += TranslateHelper.Translate("Error. File or Directory not found !") + " " + line + "\r\n";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SilentAddErr += TranslateHelper.Translate("Error could not read file !") + " " + ex.Message + "\r\n";
            }
            finally
            {
                Environment.CurrentDirectory = curdir;

                SilentAdd = false;
            }
        }

        private void tsdbImportList_ButtonClick(object sender, EventArgs e)
        {
            SilentAddErr = "";

            openFileDialog1.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.FileName = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ImportList(openFileDialog1.FileName);
                RecentFilesHelper.ImportListRecent(openFileDialog1.FileName);

                if (SilentAddErr != string.Empty)
                {
                    frmMessage f = new frmMessage();
                    f.txtMsg.Text = SilentAddErr;
                    f.ShowDialog();

                }
            }
        }

        private void tsdbImportList_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            SilentAddErr = "";

            ImportList(e.ClickedItem.Text);
            RecentFilesHelper.ImportListRecent(e.ClickedItem.Text);

            if (SilentAddErr != string.Empty)
            {
                frmMessage f = new frmMessage();
                f.txtMsg.Text = SilentAddErr;
                f.ShowDialog();

            }
        }

        #region Share

        private void tsiFacebook_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareFacebook();
        }

        private void tsiTwitter_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareTwitter();
        }

        private void tsiGooglePlus_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareGooglePlus();
        }

        private void tsiLinkedIn_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareLinkedIn();
        }

        private void tsiEmail_Click(object sender, EventArgs e)
        {
            ShareHelper.ShareEmail();
        }

        #endregion

        public bool AddFile(string filepath)
        {
            return AddFile(filepath, "", "");
        }

        public bool AddFile(string filepath, string password)
        {
            return AddFile(filepath, password, "");
        }

        public bool AddFile(string filepath, string password, string rootfolder)
        {
            string ext = "*" + System.IO.Path.GetExtension(filepath).ToLower() + ";";

            /*
            if (Module.AcceptableMediaInputPattern.IndexOf(ext) < 0)
            {
                SilentAddErr += filepath + "\n\n" + TranslateHelper.Translate("Please add only Office Files !") + "\n\n";

                return false;
            }
            */

            DataRow dr = dt.NewRow();

            FileInfo fi = new FileInfo(filepath);

            long sizekb = fi.Length / 1024;
            dr["filename"] = fi.Name;
            dr["fullfilepath"] = filepath;
            dr["sizekb"] = sizekb.ToString() + "KB";
            dr["filedate"] = fi.LastWriteTime.ToString();
            dr["rootfolder"] = rootfolder;

            dt.Rows.Add(dr);

            IsDirty = true;

            return true;
        }

        public void AddFolder(string folder_path)
        {
            AddFolder(folder_path, "");
        }

        public void AddFolder(string folder_path, string password)
        {
            string[] filez = null;

            if (!SilentAdd)
            {
                if (System.IO.Directory.GetDirectories(folder_path).Length > 0)
                {
                    DialogResult dres = Module.ShowQuestionDialog("Would you like to add also Subdirectories ?", TranslateHelper.Translate("Add Subdirectories ?"));

                    if (dres == DialogResult.Yes)
                    {
                        filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.AllDirectories);
                    }
                    else
                    {
                        filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.TopDirectoryOnly);
                    }
                }
                else
                {
                    filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.TopDirectoryOnly);
                }
            }
            else
            {
                // silent add for import list
                filez = System.IO.Directory.GetFiles(folder_path, "*.*", SearchOption.AllDirectories);
            }

            try
            {
                this.Cursor = Cursors.WaitCursor;

                for (int k = 0; k < filez.Length; k++)
                {
                    string filepath = filez[k];

                    //if (Module.IsOfficeDocument(filepath) || Module.IsPPDocument(filepath) || Module.IsOfficeDocument(filepath))
                    if (Module.IsOfficeDocument(filepath))
                    {
                        AddFile(filez[k], password, folder_path);
                    }
                }
            }
            finally
            {
                this.Cursor = null;
            }

        }

        private void SetupOutputFolders()
        {
            if (cmbOutputDir.Items.Count > 0) return;

            cmbOutputDir.Items.Add(TranslateHelper.Translate("Same Folder of Document"));
            //cmbOutputDir.Items.Add(TranslateHelper.Translate("Overwrite Document"));
            cmbOutputDir.Items.Add(TranslateHelper.Translate("Subfolder of Document"));
            cmbOutputDir.Items.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString());
            cmbOutputDir.Items.Add("---------------------------------------------------------------------------------------");

            OutputFolderHelper.LoadOutputFolders();
            cmbOutputDir.SelectedIndex = Properties.Settings.Default.OutputFolderIndex;            

            //4keepBackupOfChangedDocumentsToolStripMenuItem.Checked = Properties.Settings.Default.KeepBackup;

        }
        private void SetupOnLoad()
        {
            dgFiles.DataSource = dt;

            //3this.Icon = Properties.Resources.pdf_compress_48;

            this.Text = Module.ApplicationTitle;
            //this.Width = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width;
            //this.Left = 0;
            AddLanguageMenuItems();

            DownloadSuggestionsHelper ds = new DownloadSuggestionsHelper();
            ds.SetupDownloadMenuItems(downloadToolStripMenuItem);

            AdjustSizeLocation();

            SetupOutputFolders();

            keepFolderStructureToolStripMenuItem.Checked = Properties.Settings.Default.KeepFolderStructure;

            RecentFilesHelper.FillMenuRecentFile();
            RecentFilesHelper.FillMenuRecentFolder();
            RecentFilesHelper.FillMenuRecentImportList();            
            /*
            buyToolStripMenuItem.Visible = frmPurchase.RenMove;

            if (Properties.Settings.Default.Price != string.Empty && !buyApplicationToolStripMenuItem.Text.EndsWith(Properties.Settings.Default.Price))
            {
                buyApplicationToolStripMenuItem.Text = buyApplicationToolStripMenuItem.Text + " " + Properties.Settings.Default.Price;
            }
            */
            exploreFirstOutputDocumentToolStripMenuItem.Checked = Properties.Settings.Default.ExploreDocumentOnFinish;

            checkForNewVersionEachWeekToolStripMenuItem.Checked = Properties.Settings.Default.CheckWeek;

            keepCreationDateToolStripMenuItem.Checked = Properties.Settings.Default.KeepCreationDate;

            keepLastModificationDateToolStripMenuItem.Checked = Properties.Settings.Default.KeepLastModificationDate;

            showMessageOnSucessToolStripMenuItem.Checked = Properties.Settings.Default.ShowMessageOnSucess;

        }

        private void AdjustSizeLocation()
        {
            if (Properties.Settings.Default.Maximized)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {

                if (Properties.Settings.Default.Width == -1)
                {
                    this.CenterToScreen();
                    return;
                }
                else
                {
                    this.Width = Properties.Settings.Default.Width;
                }
                if (Properties.Settings.Default.Height != 0)
                {
                    this.Height = Properties.Settings.Default.Height;
                }

                if (Properties.Settings.Default.Left != -1)
                {
                    this.Left = Properties.Settings.Default.Left;
                }

                if (Properties.Settings.Default.Top != -1)
                {
                    this.Top = Properties.Settings.Default.Top;
                }

                if (this.Width < 300)
                {
                    this.Width = 300;
                }

                if (this.Height < 300)
                {
                    this.Height = 300;
                }

                if (this.Left < 0)
                {
                    this.Left = 0;
                }

                if (this.Top < 0)
                {
                    this.Top = 0;
                }
            }

        }

        private void SaveSizeLocation()
        {
            Properties.Settings.Default.Maximized = (this.WindowState == FormWindowState.Maximized);
            Properties.Settings.Default.Left = this.Left;
            Properties.Settings.Default.Top = this.Top;
            Properties.Settings.Default.Width = this.Width;
            Properties.Settings.Default.Height = this.Height;
            Properties.Settings.Default.Save();

        }

        #region Localization

        private void AddLanguageMenuItems()
        {
            for (int k = 0; k < frmLanguage.LangCodes.Count; k++)
            {
                ToolStripMenuItem ti = new ToolStripMenuItem();
                ti.Text = frmLanguage.LangDesc[k];
                ti.Tag = frmLanguage.LangCodes[k];
                ti.Image = frmLanguage.LangImg[k];

                if (Properties.Settings.Default.Language == frmLanguage.LangCodes[k])
                {
                    ti.Checked = true;
                }

                ti.Click += new EventHandler(tiLang_Click);

                if (k < 25)
                {
                    languages1ToolStripMenuItem.DropDownItems.Add(ti);
                }
                else
                {
                    languages2ToolStripMenuItem.DropDownItems.Add(ti);
                }

                //languageToolStripMenuItem.DropDownItems.Add(ti);
            }
        }

        void tiLang_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem ti = (ToolStripMenuItem)sender;
            string langcode = ti.Tag.ToString();
            ChangeLanguage(langcode);

            //for (int k = 0; k < languageToolStripMenuItem.DropDownItems.Count; k++)
            for (int k = 0; k < languages1ToolStripMenuItem.DropDownItems.Count; k++)
            {
                ToolStripMenuItem til = (ToolStripMenuItem)languages1ToolStripMenuItem.DropDownItems[k];
                if (til == ti)
                {
                    til.Checked = true;
                }
                else
                {
                    til.Checked = false;
                }
            }

            for (int k = 0; k < languages2ToolStripMenuItem.DropDownItems.Count; k++)
            {
                ToolStripMenuItem til = (ToolStripMenuItem)languages2ToolStripMenuItem.DropDownItems[k];
                if (til == ti)
                {
                    til.Checked = true;
                }
                else
                {
                    til.Checked = false;
                }
            }
        }

        private bool InChangeLanguage = false;

        private void ChangeLanguage(string language_code)
        {
            try
            {
                InChangeLanguage = true;

                Properties.Settings.Default.Language = language_code;
                frmLanguage.SetLanguage();

                bool maximized = (this.WindowState == FormWindowState.Maximized);
                this.WindowState = FormWindowState.Normal;

                /*
                RegistryKey key = Registry.CurrentUser;
                RegistryKey key2 = Registry.CurrentUser;

                try
                {
                    key = key.OpenSubKey("Software\\4dots Software", true);

                    if (key == null)
                    {
                        key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\4dots Software");
                    }

                    key2 = key.OpenSubKey(frmLanguage.RegKeyName, true);

                    if (key2 == null)
                    {
                        key2 = key.CreateSubKey(frmLanguage.RegKeyName);
                    }

                    key = key2;

                    //key.SetValue("Language", language_code);
                    key.SetValue("Menu Item Caption", TranslateHelper.Translate("Change PDF Properties"));
                }
                catch (Exception ex)
                {
                    Module.ShowError(ex);
                    return;
                }
                finally
                {
                    key.Close();
                    key2.Close();
                }
                */
                //1SaveSizeLocation();

                //3SavePositionSize();

                this.Controls.Clear();

                InitializeComponent();

                SetupOnLoad();

                if (maximized)
                {
                    this.WindowState = FormWindowState.Maximized;
                }

                this.ResumeLayout(true);
            }
            finally
            {
                InChangeLanguage = false;
            }
        }

        #endregion        

        private void buyApplicationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Module.BuyURL);
        }

        private void enterLicenseKeyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            SetupOnLoad();

            if (Properties.Settings.Default.CheckWeek)
            {
                UpdateHelper.InitializeCheckVersionWeek();
            }

            if (Module.args != null)
            {
                AddVisual(Module.args);
            }            
        }

        private void AddVisual(string[] argsvisual)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                //Module.ShowMessage("Is From Windows Explorer");                                

                for (int k = 0; k < argsvisual.Length; k++)
                {
                    if (System.IO.File.Exists(argsvisual[k]))
                    {
                        AddFile(argsvisual[k]);

                    }
                    else if (System.IO.Directory.Exists(argsvisual[k]))
                    {
                        AddFolder(argsvisual[k]);
                    }
                }
            }
            finally
            {
                this.Cursor = null;
            }
        }


        #region Help

        private void helpGuideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //System.Diagnostics.Process.Start(Application.StartupPath + "\\Video Cutter Joiner Expert - User's Manual.chm");
            System.Diagnostics.Process.Start(Module.HelpURL);
        }

        private void pleaseDonateToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/donate.php");
        }

        private void dotsSoftwarePRODUCTCATALOGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/downloads/4dots-Software-PRODUCT-CATALOG.pdf");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAbout f = new frmAbout();
            f.ShowDialog();
        }

        private void tiHelpFeedback_Click(object sender, EventArgs e)
        {
            /*
            frmUninstallQuestionnaire f = new frmUninstallQuestionnaire(false);
            f.ShowDialog();
            */

            System.Diagnostics.Process.Start("https://www.4dots-software.com/support/bugfeature.php?app=" + System.Web.HttpUtility.UrlEncode(Module.ShortApplicationTitle));
        }

        private void followUsOnTwitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("http://www.twitter.com/4dotsSoftware");
        }

        private void visit4dotsSoftwareWebsiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com");
        }

        private void checkForNewVersionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateHelper.CheckVersion(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                Application.Exit();
            }
            catch { }
        }

        #endregion

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {                       
            Properties.Settings.Default.ExploreDocumentOnFinish = exploreFirstOutputDocumentToolStripMenuItem.Checked;

            Properties.Settings.Default.OutputFolderIndex = cmbOutputDir.SelectedIndex;

            Properties.Settings.Default.CheckWeek = checkForNewVersionEachWeekToolStripMenuItem.Checked;

            Properties.Settings.Default.KeepCreationDate = keepCreationDateToolStripMenuItem.Checked;

            Properties.Settings.Default.KeepLastModificationDate = keepLastModificationDateToolStripMenuItem.Checked;

            Properties.Settings.Default.ShowMessageOnSucess = showMessageOnSucessToolStripMenuItem.Checked;

            Properties.Settings.Default.Save();
        }

        private void EnableDisableForm(bool enable)
        {
            foreach (Control co in this.Controls)
            {
                co.Enabled = enable;
            }
        }
        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < dgFiles.Rows.Count; k++)
            {
                dgFiles.Rows[k].Selected = true;
            }
        }

        private void seelctNoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < dgFiles.Rows.Count; k++)
            {
                dgFiles.Rows[k].Selected = false;
            }
        }

        private void invertSelectionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < dgFiles.Rows.Count; k++)
            {
                dgFiles.Rows[k].Selected = !dgFiles.Rows[k].Selected;
            }
        }
        
        private void tsbExtractImages_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.KeepCreationDate = keepCreationDateToolStripMenuItem.Checked;

            Properties.Settings.Default.KeepLastModificationDate = keepLastModificationDateToolStripMenuItem.Checked;

            Properties.Settings.Default.ShowMessageOnSucess = showMessageOnSucessToolStripMenuItem.Checked;

            OperationStopped = false;
            OperationPaused = false;

            //remove for normal

            //if (frmAbout.LDT == string.Empty)            

            try
            {
                dgFiles.EndEdit();

                sOutputDir = cmbOutputDir.Text;
                
                FirstOutputDocument = "";

                if (Properties.Settings.Default.MsgWordVisible)
                {
                    frmMsgWordVisible fmw = new frmMsgWordVisible();
                    fmw.ShowDialog();
                }

                frmProgress f = new frmProgress();
                f.progressBar1.Maximum = dt.Rows.Count;
                f.progressBar1.Value = 0;
                f.timTime.Enabled = true;

                Err = "";

                //this.Enabled = false;

                EnableDisableForm(false);

                f.Show(this);

                OfficeHelper.QuitExcelApplication();
                OfficeHelper.QuitWordApplication();
                OfficeHelper.QuitPowerPointApplication();                               
                OfficeHelper.QuitOfficeApplications();
                
                bwExtractImages.RunWorkerAsync();

                while (bwExtractImages.IsBusy)
                {
                    Application.DoEvents();
                }

            }
            finally
            {
                if (frmProgress.Instance != null && !frmProgress.Instance.IsDisposed)
                {
                    frmProgress.Instance.Close();
                }

                OfficeHelper.QuitExcelApplication();
                OfficeHelper.QuitWordApplication();
                OfficeHelper.QuitPowerPointApplication();                               
                OfficeHelper.QuitOfficeApplications();

                //this.Enabled = true;

                EnableDisableForm(true);

                //remove for normal

                //if (frmAbout.LDT == string.Empty)                

                if (OperationStopped)
                {
                    Module.ShowMessage("Operation Stopped !");
                }
                else if (Err != string.Empty)
                {
                    frmMessage f2 = new frmMessage();
                    f2.lblMsg.Text = TranslateHelper.Translate("Operation completed with Errors !");
                    f2.txtMsg.Text = Err;

                    f2.ShowDialog(this);
                }
                else if (Properties.Settings.Default.ShowMessageOnSucess)
                {
                    Module.ShowMessage("Operation completed successfully !");
                }

                if (exploreFirstOutputDocumentToolStripMenuItem.Checked)
                {
                    ExploreOnFinish();
                }

                //remove for normal

                //if (frmAbout.LDT == string.Empty)                
            }
        }

        private void ExploreOnFinish()
        {
            string filepath = FirstOutputDocument;

            if (FirstOutputDocument == string.Empty) return;

            string args = string.Format("/e, /select, \"{0}\"", filepath);

            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "explorer";
            info.UseShellExecute = true;
            info.Arguments = args;
            Process.Start(info);
        }
        
        private void bwExtractImages_DoWork(object sender, DoWorkEventArgs e)
        {           
            for (int k = 0; k < dt.Rows.Count; k++)
            {
                if (OperationStopped || bwExtractImages.CancellationPending)
                {
                    return;
                }

                while (OperationPaused)
                {
                    Application.DoEvents();
                }

                OfficeHelper.QuitExcelApplication();
                OfficeHelper.QuitWordApplication();
                OfficeHelper.QuitPowerPointApplication();                               
                OfficeHelper.QuitOfficeApplications();

                //remove for normal

                //if (frmAbout.LDT == string.Empty)
                if (false)
                {
                    bool found = false;

                    for (int m = 0; m < de.Length; m++)
                    {
                        if (de[m] == 0)
                        {
                            de[m] = 1;

                            if (m == de.Length - 1)
                            {
                                found = true;
                            }

                            break;
                        }
                    }

                    if (found)
                    {
                        return;
                    }
                }
                
                bwExtractImages.ReportProgress(-1, dt.Rows[k]["fullfilepath"].ToString());

                string filepath = dt.Rows[k]["fullfilepath"].ToString();
                string rootfolder = dt.Rows[k]["rootfolder"].ToString();
                string slideranges = dt.Rows[k]["slideranges"].ToString();

                string outfilepath = "";

                if (sOutputDir.Trim() == TranslateHelper.Translate("Same Folder of Document"))
                {
                    string dirpath = System.IO.Path.GetDirectoryName(filepath);

                    outfilepath = System.IO.Path.Combine(dirpath, System.IO.Path.GetFileNameWithoutExtension(filepath) + "_extracted" + System.IO.Path.GetExtension(filepath));

                }
                else if (sOutputDir.StartsWith(TranslateHelper.Translate("Subfolder") + " : "))
                {
                    int subfolderspos = (TranslateHelper.Translate("Subfolder") + " : ").Length;
                    string subfolder = sOutputDir.Substring(subfolderspos);

                    outfilepath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filepath) + "\\" + subfolder, System.IO.Path.GetFileName(filepath));
                }
                else if (sOutputDir.Trim() == TranslateHelper.Translate("Overwrite Document"))
                {
                    outfilepath = filepath;
                }
                else
                {
                    if (rootfolder != string.Empty && Properties.Settings.Default.KeepFolderStructure)
                    {
                        string dep = System.IO.Path.GetDirectoryName(filepath).Substring(rootfolder.Length);

                        string outdfp = sOutputDir + dep;

                        outfilepath = System.IO.Path.Combine(outdfp, System.IO.Path.GetFileName(filepath));

                        if (!System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(outfilepath)))
                        {
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outfilepath));
                        }
                    }
                    else
                    {
                        outfilepath = System.IO.Path.Combine(sOutputDir, System.IO.Path.GetFileName(filepath));
                    }
                }

                Module.CurrentImagesDirectory = System.IO.Path.GetDirectoryName(outfilepath);

                if (!System.IO.Directory.Exists(Module.CurrentImagesDirectory))
                {
                    System.IO.Directory.CreateDirectory(Module.CurrentImagesDirectory);
                }
                
                if (Module.IsWordDocument(filepath))
                {
                    WordImageExtractor wim = new WordImageExtractor();
                    wim.ExtractImages(filepath,slideranges);

                    Err += wim.err;

                    bwExtractImages.ReportProgress(0);
                }
                else if (Module.IsPPDocument(filepath))
                {
                    PowerpointImageExtractor wim = new PowerpointImageExtractor();
                    wim.ExtractImages(filepath, slideranges);

                    Err += wim.err;

                    bwExtractImages.ReportProgress(0);
                }
                else if (Module.IsExcelDocument(filepath))
                {
                    ExcelImageExtractor wim = new ExcelImageExtractor();
                    wim.ExtractImages(filepath, slideranges);

                    Err += wim.err;

                    bwExtractImages.ReportProgress(0);
                }                
            }                        
        }

        private void bwExtractImages_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                frmProgress.Instance.lblOutputFile.Text = System.IO.Path.GetFileName(e.UserState.ToString());
            }
            else
            {
                if (frmProgress.Instance.progressBar1.Value + 1 <= frmProgress.Instance.progressBar1.Maximum)
                {
                    frmProgress.Instance.progressBar1.Value++;
                }

                frmProgress.Instance.lblJoinNumber.Text = frmProgress.Instance.progressBar1.Value.ToString() + " / " +
                    frmProgress.Instance.progressBar1.Maximum.ToString();                                                    
            }
        }

        private bool IsFileTheSame(string filepath, string existingfp)
        {
            if (!System.IO.File.Exists(existingfp)) return false;

            FileInfo fi = new FileInfo(filepath);
            FileInfo fi2 = new FileInfo(existingfp);

            if (fi.Length != fi2.Length)
            {
                return false;
            }
            else
            {
                using (FileStream fs = File.OpenRead(filepath))
                {
                    using (FileStream fs2 = File.OpenRead(existingfp))
                    {
                        try
                        {
                            byte[] bytes = new byte[1024];
                            byte[] bytes2 = new byte[1024];

                            while (fs.Read(bytes, 0, bytes.Length) > 0)
                            {
                                if (fs2.Read(bytes2, 0, bytes2.Length) > 0)
                                {
                                    for (int m = 0; m < bytes.Length; m++)
                                    {
                                        if (bytes[m] != bytes2[m])
                                        {
                                            return false;
                                        }
                                    }
                                }
                                else
                                {
                                    return false;
                                }
                            }
                        }
                        catch
                        {
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        private void bwExtractImages_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        #region Grid Context menu

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgFiles.CurrentRow == null) return;

            DataRowView drv = (DataRowView)dgFiles.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            System.Diagnostics.Process.Start(filepath);
        }

        private void exploreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)dgFiles.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            string args = string.Format("/e, /select, \"{0}\"", filepath);

            ProcessStartInfo info = new ProcessStartInfo();
            info.FileName = "explorer";
            info.UseShellExecute = true;
            info.Arguments = args;
            Process.Start(info);
        }

        private void copyFullFilePathToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView drv = (DataRowView)dgFiles.CurrentRow.DataBoundItem;

            DataRow dr = drv.Row;

            string filepath = dr["fullfilepath"].ToString();

            Clipboard.Clear();

            Clipboard.SetText(filepath);
        }

        private void cmsFiles_Opening(object sender, CancelEventArgs e)
        {
            Point p = dgFiles.PointToClient(new Point(Control.MousePosition.X, Control.MousePosition.Y));
            DataGridView.HitTestInfo hit = dgFiles.HitTest(p.X, p.Y);

            if (hit.Type == DataGridViewHitTestType.Cell)
            {
                dgFiles.CurrentCell = dgFiles.Rows[hit.RowIndex].Cells[hit.ColumnIndex];
            }

            if (dgFiles.CurrentRow == null)
            {
                e.Cancel = true;
            }
        }
        #endregion

        private void cmbOutputDir_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if (cmbOutputDir.SelectedIndex == 3)
            {
                Module.ShowMessage("Please specify another option as the Output Folder !");
                cmbOutputDir.SelectedIndex = Properties.Settings.Default.OutputFolderIndex;
            }
            else if (cmbOutputDir.SelectedIndex == 1)
            {
                frmOutputSubFolder fob = new frmOutputSubFolder();

                if (fob.ShowDialog() == DialogResult.OK)
                {
                    OutputFolderHelper.SaveOutputFolder(TranslateHelper.Translate("Subfolder") + " : " + fob.txtSubfolder.Text);
                }
                else
                {
                    return;
                }
            }            
        }

        private void outputFilenamePatternToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmOptions f = new frmOptions();

            f.ShowDialog();
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            dgFiles.EndEdit();

            if (dt.Rows.Count == 0)
            {
                Module.ShowMessage("Please add a Document first !");

            }
            else
            {
                string dirpath = "";
                string filepath = "";
                string outfilepath = "";

                if (dgFiles.SelectedCells.Count == 0)
                {
                    filepath = dgFiles.Rows[0].Cells["colFullfilepath"].Value.ToString();
                }
                else
                {
                    filepath = dgFiles.SelectedCells[0].OwningRow.Cells["colFullfilepath"].Value.ToString();
                }

                if (frmMain.Instance.cmbOutputDir.Text.Trim() == TranslateHelper.Translate("Same Folder of Document"))
                {
                    dirpath = System.IO.Path.GetDirectoryName(filepath);

                    outfilepath = System.IO.Path.Combine(dirpath, System.IO.Path.GetFileNameWithoutExtension(filepath) + "_replaced" + System.IO.Path.GetExtension(filepath));

                }
                else if (frmMain.Instance.cmbOutputDir.Text.ToString().StartsWith(TranslateHelper.Translate("Subfolder") + " : "))
                {
                    int subfolderspos = (TranslateHelper.Translate("Subfolder") + " : ").Length;
                    string subfolder = frmMain.Instance.cmbOutputDir.Text.ToString().Substring(subfolderspos);

                    outfilepath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filepath) + "\\" + subfolder, System.IO.Path.GetFileName(filepath));
                }
                else if (frmMain.Instance.cmbOutputDir.Text.Trim() == TranslateHelper.Translate("Overwrite Document"))
                {
                    outfilepath = filepath;
                }
                else
                {
                    outfilepath = System.IO.Path.Combine(frmMain.Instance.cmbOutputDir.Text, System.IO.Path.GetFileName(filepath));
                }

                if (!System.IO.File.Exists(outfilepath))
                {
                    if (!System.IO.Directory.Exists(outfilepath))
                    {
                        outfilepath = System.IO.Path.GetDirectoryName(outfilepath);
                    }
                }

                string args = string.Format("/e, /select, \"{0}\"", outfilepath);
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "explorer";
                info.UseShellExecute = true;
                info.Arguments = args;
                Process.Start(info);
            }
        }

        private void btnChangeFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbr = new FolderBrowserDialog();

            if (cmbOutputDir.Text != string.Empty && System.IO.Directory.Exists(cmbOutputDir.Text))
            {
                fbr.SelectedPath = cmbOutputDir.Text;
            }
            if (fbr.ShowDialog() == DialogResult.OK)
            {
                OutputFolderHelper.SaveOutputFolder(fbr.SelectedPath);
            }

        }

        #region Drag and Drop

        private void dgFiles_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void dgFiles_DragOver(object sender, DragEventArgs e)
        {
            if ((e.AllowedEffect & DragDropEffects.Copy) == DragDropEffects.Copy)
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void dgFiles_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                string[] filez = (string[])e.Data.GetData(DataFormats.FileDrop);

                for (int k = 0; k < filez.Length; k++)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;

                        if (System.IO.File.Exists(filez[k]))
                        {
                            AddFile(filez[k]);
                        }
                        else if (System.IO.Directory.Exists(filez[k]))
                        {
                            AddFolder(filez[k]);
                        }
                    }
                    finally
                    {
                        this.Cursor = null;
                    }
                }
            }
        }

        #endregion

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OfficeToImagesConverter4dots
{
    public partial class frmAbout : CustomForm
    {
        public static string lblf = "";

        public static string[] bufluf = new string[5];
        public static List<string> lstbufluf = new List<string>();

        // license email
        public static string LDT = "";

        public frmAbout()
        {
            InitializeComponent();
        }
                     
        private void frmAbout_Load(object sender, EventArgs e)
        {
            lblAbout.Text = Module.ApplicationTitle + "\n\n" +
            "Developed by Alexander Triantafyllou\n" +
            "Copyright © 2018 - 4dots Software\n";

            ullProductWebpage.Text = Module.ProductWebpageURL;

            //3

            /*
            if (!SettingsHelper.RenMove && !frmPurchase.RenMove)
            {
                if (frmAbout.LDT == string.Empty)
                {
                    lblRegistered.Text = TranslateHelper.Translate("Trial Version - Unlicensed");
                }
                else                
                {
                    lblRegistered.Text = TranslateHelper.Translate("Business Version - Licensed");
                }                

                lblLicenseKey.Text = frmAbout.LDT;

                if (frmAbout.LDT != string.Empty)
                {
                    btnRegister.Visible = false;
                    btnBuy.Visible = false;
                }
            }
            else
            {
                lblRegistered.Text = TranslateHelper.Translate("Trial Version - Unlicensed");
            }*/                        
        }
                
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/support/");
        }

        private void btnBuy_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Module.BuyURL);
        }

        private void btnRegister_Click(object sender, EventArgs e)
        {
            
        }
    }
}
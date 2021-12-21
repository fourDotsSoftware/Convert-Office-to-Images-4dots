using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OfficeToImagesConverter4dots
{
    public partial class frmError : OfficeToImagesConverter4dots.CustomForm
    {
        public frmError(string lbl,string txt)
        {
            InitializeComponent();

            lblError.Text = lbl;
            txtError.Text = txt;

            this.BringToFront();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }
    }
}

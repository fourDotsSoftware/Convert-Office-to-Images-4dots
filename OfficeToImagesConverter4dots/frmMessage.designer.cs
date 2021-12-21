namespace OfficeToImagesConverter4dots
{
    partial class frmMessage
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMessage));
            this.lblMsg = new System.Windows.Forms.Label();
            this.txtMsg = new System.Windows.Forms.RichTextBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.Transparent;
            resources.ApplyResources(this.lblMsg, "lblMsg");
            this.lblMsg.Name = "lblMsg";
            // 
            // txtMsg
            // 
            resources.ApplyResources(this.txtMsg, "txtMsg");
            this.txtMsg.Name = "txtMsg";
            this.txtMsg.ReadOnly = true;
            this.txtMsg.TabStop = false;
            // 
            // btnOK
            // 
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Image = global::OfficeToImagesConverter4dots.Properties.Resources.check;
            this.btnOK.Name = "btnOK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // frmMessage
            // 
            this.AcceptButton = this.btnOK;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.txtMsg);
            this.Controls.Add(this.lblMsg);
            this.Name = "frmMessage";
            this.Load += new System.EventHandler(this.frmMessage_Load);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.Label lblMsg;
        public System.Windows.Forms.RichTextBox txtMsg;
        private System.Windows.Forms.Button btnOK;
    }
}

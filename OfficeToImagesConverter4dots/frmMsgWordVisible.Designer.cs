namespace OfficeToImagesConverter4dots
{
    partial class frmMsgWordVisible
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMsgWordVisible));
            this.btnOK = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.chkDoNotShowAgain = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Image = global::OfficeToImagesConverter4dots.Properties.Resources.check;
            this.btnOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOK.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnOK.Location = new System.Drawing.Point(231, 212);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(135, 29);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "&Close";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.DarkBlue;
            this.label1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label1.Location = new System.Drawing.Point(111, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(407, 83);
            this.label1.TabIndex = 4;
            this.label1.Text = "Microsoft Office Applications will now open.\r\nPlease do not interact with them an" +
    "d let them do itheir job.\r\nThey will close automatically.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkDoNotShowAgain
            // 
            this.chkDoNotShowAgain.AutoSize = true;
            this.chkDoNotShowAgain.BackColor = System.Drawing.Color.Transparent;
            this.chkDoNotShowAgain.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.chkDoNotShowAgain.ForeColor = System.Drawing.Color.DarkBlue;
            this.chkDoNotShowAgain.Location = new System.Drawing.Point(212, 180);
            this.chkDoNotShowAgain.Name = "chkDoNotShowAgain";
            this.chkDoNotShowAgain.Size = new System.Drawing.Size(154, 20);
            this.chkDoNotShowAgain.TabIndex = 5;
            this.chkDoNotShowAgain.Text = "Do not show again";
            this.chkDoNotShowAgain.UseVisualStyleBackColor = false;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.DarkBlue;
            this.label2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label2.Location = new System.Drawing.Point(111, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(407, 83);
            this.label2.TabIndex = 6;
            this.label2.Text = "If conversion fails, stop and try again !";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmMsgWordVisible
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.ClientSize = new System.Drawing.Size(644, 260);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chkDoNotShowAgain);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMsgWordVisible";
            this.Text = "Microsoft Office Applications will open. Do not interact with them.";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkDoNotShowAgain;
        private System.Windows.Forms.Label label2;
    }
}

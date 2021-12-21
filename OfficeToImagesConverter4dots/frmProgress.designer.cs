namespace OfficeToImagesConverter4dots
{
    partial class frmProgress
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmProgress));
            this.timTime = new System.Windows.Forms.Timer(this.components);
            this.btnStop = new System.Windows.Forms.Button();
            this.lblOutputFile = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblRemainingTime = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblElapsedTime = new System.Windows.Forms.Label();
            this.lblCaption = new System.Windows.Forms.Label();
            this.btnPause = new System.Windows.Forms.Button();
            this.progressBar1 = new OfficeToImagesConverter4dots.ucTextProgressBar();
            this.lblJoinNumber = new System.Windows.Forms.Label();
            this.lblJoinNumberCaption = new System.Windows.Forms.Label();
            this.pbarTotal = new OfficeToImagesConverter4dots.ucTextProgressBar();
            this.lblTotal = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // timTime
            // 
            this.timTime.Interval = 1000;
            this.timTime.Tick += new System.EventHandler(this.timTime_Tick);
            // 
            // btnStop
            // 
            this.btnStop.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnStop.Image = global::OfficeToImagesConverter4dots.Properties.Resources.media_stop;
            resources.ApplyResources(this.btnStop, "btnStop");
            this.btnStop.Name = "btnStop";
            this.btnStop.UseVisualStyleBackColor = true;
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            // 
            // lblOutputFile
            // 
            resources.ApplyResources(this.lblOutputFile, "lblOutputFile");
            this.lblOutputFile.BackColor = System.Drawing.Color.Transparent;
            this.lblOutputFile.Name = "lblOutputFile";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.ForeColor = System.Drawing.Color.DarkBlue;
            this.label3.Name = "label3";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.ForeColor = System.Drawing.Color.DarkBlue;
            this.label2.Name = "label2";
            // 
            // lblRemainingTime
            // 
            resources.ApplyResources(this.lblRemainingTime, "lblRemainingTime");
            this.lblRemainingTime.BackColor = System.Drawing.Color.Transparent;
            this.lblRemainingTime.Name = "lblRemainingTime";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.ForeColor = System.Drawing.Color.DarkBlue;
            this.label1.Name = "label1";
            // 
            // lblElapsedTime
            // 
            resources.ApplyResources(this.lblElapsedTime, "lblElapsedTime");
            this.lblElapsedTime.BackColor = System.Drawing.Color.Transparent;
            this.lblElapsedTime.Name = "lblElapsedTime";
            // 
            // lblCaption
            // 
            resources.ApplyResources(this.lblCaption, "lblCaption");
            this.lblCaption.BackColor = System.Drawing.Color.Transparent;
            this.lblCaption.Name = "lblCaption";
            // 
            // btnPause
            // 
            this.btnPause.Image = global::OfficeToImagesConverter4dots.Properties.Resources.media_pause;
            resources.ApplyResources(this.btnPause, "btnPause");
            this.btnPause.Name = "btnPause";
            this.btnPause.UseVisualStyleBackColor = true;
            this.btnPause.Click += new System.EventHandler(this.btnPause_Click);
            // 
            // progressBar1
            // 
            resources.ApplyResources(this.progressBar1, "progressBar1");
            this.progressBar1.Name = "progressBar1";
            // 
            // lblJoinNumber
            // 
            resources.ApplyResources(this.lblJoinNumber, "lblJoinNumber");
            this.lblJoinNumber.BackColor = System.Drawing.Color.Transparent;
            this.lblJoinNumber.Name = "lblJoinNumber";
            // 
            // lblJoinNumberCaption
            // 
            resources.ApplyResources(this.lblJoinNumberCaption, "lblJoinNumberCaption");
            this.lblJoinNumberCaption.BackColor = System.Drawing.Color.Transparent;
            this.lblJoinNumberCaption.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblJoinNumberCaption.Name = "lblJoinNumberCaption";
            // 
            // pbarTotal
            // 
            resources.ApplyResources(this.pbarTotal, "pbarTotal");
            this.pbarTotal.Name = "pbarTotal";
            // 
            // lblTotal
            // 
            resources.ApplyResources(this.lblTotal, "lblTotal");
            this.lblTotal.BackColor = System.Drawing.Color.Transparent;
            this.lblTotal.Name = "lblTotal";
            // 
            // frmProgress
            // 
            resources.ApplyResources(this, "$this");
            this.CancelButton = this.btnStop;
            this.Controls.Add(this.lblTotal);
            this.Controls.Add(this.pbarTotal);
            this.Controls.Add(this.lblJoinNumber);
            this.Controls.Add(this.lblJoinNumberCaption);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.lblOutputFile);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblRemainingTime);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblElapsedTime);
            this.Controls.Add(this.lblCaption);
            this.Controls.Add(this.btnStop);
            this.Controls.Add(this.btnPause);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "frmProgress";
            this.Load += new System.EventHandler(this.frmProgress_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Timer timTime;
        public System.Windows.Forms.Button btnPause;
        public System.Windows.Forms.Button btnStop;
        public System.Windows.Forms.Label lblCaption;
        public System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lblRemainingTime;
        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblElapsedTime;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label lblOutputFile;
        public ucTextProgressBar progressBar1;
        public System.Windows.Forms.Label lblJoinNumber;
        public System.Windows.Forms.Label lblJoinNumberCaption;
        public ucTextProgressBar pbarTotal;
        private System.Windows.Forms.Label lblTotal;
    }
}

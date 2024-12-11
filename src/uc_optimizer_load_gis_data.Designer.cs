namespace FIA_Biosum_Manager
{
    partial class uc_optimizer_load_gis_data
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblTitle = new System.Windows.Forms.Label();
            this.ckBackupData = new System.Windows.Forms.CheckBox();
            this.ckUpdateYardingDistance = new System.Windows.Forms.CheckBox();
            this.txtMaxOneWayHours = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnLoad = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtDataFile = new System.Windows.Forms.TextBox();
            this.lblLastUpdated = new System.Windows.Forms.Label();
            this.lblFileSize = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(715, 32);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Load GIS Data";
            // 
            // ckBackupData
            // 
            this.ckBackupData.AutoSize = true;
            this.ckBackupData.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ckBackupData.Location = new System.Drawing.Point(7, 151);
            this.ckBackupData.Name = "ckBackupData";
            this.ckBackupData.Size = new System.Drawing.Size(576, 21);
            this.ckBackupData.TabIndex = 29;
            this.ckBackupData.Text = "Create a backup copy of the existing database? The file name will include today\'s" +
    " date.";
            this.ckBackupData.UseVisualStyleBackColor = true;
            // 
            // ckUpdateYardingDistance
            // 
            this.ckUpdateYardingDistance.AutoSize = true;
            this.ckUpdateYardingDistance.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ckUpdateYardingDistance.Location = new System.Drawing.Point(7, 231);
            this.ckUpdateYardingDistance.Name = "ckUpdateYardingDistance";
            this.ckUpdateYardingDistance.Size = new System.Drawing.Size(612, 21);
            this.ckUpdateYardingDistance.TabIndex = 30;
            this.ckUpdateYardingDistance.Text = "Update the plot yarding distance from the plot_gis table in the master travel tim" +
    "es database?";
            this.ckUpdateYardingDistance.UseVisualStyleBackColor = true;
            // 
            // txtMaxOneWayHours
            // 
            this.txtMaxOneWayHours.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMaxOneWayHours.ForeColor = System.Drawing.Color.Black;
            this.txtMaxOneWayHours.Location = new System.Drawing.Point(7, 188);
            this.txtMaxOneWayHours.Name = "txtMaxOneWayHours";
            this.txtMaxOneWayHours.Size = new System.Drawing.Size(45, 22);
            this.txtMaxOneWayHours.TabIndex = 63;
            this.txtMaxOneWayHours.Text = "2.0";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(61, 191);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(257, 20);
            this.label1.TabIndex = 64;
            this.label1.Text = "Maximum hours from plot";
            // 
            // BtnLoad
            // 
            this.BtnLoad.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnLoad.Location = new System.Drawing.Point(470, 404);
            this.BtnLoad.Name = "BtnLoad";
            this.BtnLoad.Size = new System.Drawing.Size(165, 33);
            this.BtnLoad.TabIndex = 65;
            this.BtnLoad.Text = "Load Data";
            this.BtnLoad.UseVisualStyleBackColor = true;
            this.BtnLoad.Click += new System.EventHandler(this.BtnLoad_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnCancel.Location = new System.Drawing.Point(285, 404);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(165, 33);
            this.BtnCancel.TabIndex = 66;
            this.BtnCancel.Text = "Cancel";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(11, 400);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(115, 37);
            this.btnHelp.TabIndex = 67;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(4, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(257, 20);
            this.label2.TabIndex = 68;
            this.label2.Text = "Path to source GIS data file";
            // 
            // txtDataFile
            // 
            this.txtDataFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataFile.Location = new System.Drawing.Point(5, 66);
            this.txtDataFile.Name = "txtDataFile";
            this.txtDataFile.ReadOnly = true;
            this.txtDataFile.Size = new System.Drawing.Size(600, 24);
            this.txtDataFile.TabIndex = 69;
            // 
            // lblLastUpdated
            // 
            this.lblLastUpdated.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLastUpdated.Location = new System.Drawing.Point(4, 99);
            this.lblLastUpdated.Name = "lblLastUpdated";
            this.lblLastUpdated.Size = new System.Drawing.Size(359, 20);
            this.lblLastUpdated.TabIndex = 70;
            this.lblLastUpdated.Text = "Last updated";
            // 
            // lblFileSize
            // 
            this.lblFileSize.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileSize.Location = new System.Drawing.Point(369, 99);
            this.lblFileSize.Name = "lblFileSize";
            this.lblFileSize.Size = new System.Drawing.Size(236, 20);
            this.lblFileSize.TabIndex = 71;
            this.lblFileSize.Text = "File size";
            // 
            // uc_optimizer_load_gis_data
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblFileSize);
            this.Controls.Add(this.lblLastUpdated);
            this.Controls.Add(this.txtDataFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnLoad);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtMaxOneWayHours);
            this.Controls.Add(this.ckUpdateYardingDistance);
            this.Controls.Add(this.ckBackupData);
            this.Controls.Add(this.lblTitle);
            this.Name = "uc_optimizer_load_gis_data";
            this.Size = new System.Drawing.Size(715, 457);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.CheckBox ckBackupData;
        private System.Windows.Forms.CheckBox ckUpdateYardingDistance;
        private System.Windows.Forms.TextBox txtMaxOneWayHours;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnLoad;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button btnHelp;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtDataFile;
        private System.Windows.Forms.Label lblLastUpdated;
        private System.Windows.Forms.Label lblFileSize;
    }
}

namespace FIA_Biosum_Manager
{
    partial class uc_delete_packages
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_delete_packages));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpboxFilter = new System.Windows.Forms.GroupBox();
            this.btnFilterFinish = new System.Windows.Forms.Button();
            this.btnFilterHelp = new System.Windows.Forms.Button();
            this.btnFilterPrevious = new System.Windows.Forms.Button();
            this.btnFilterNext = new System.Windows.Forms.Button();
            this.btnFilterCancel = new System.Windows.Forms.Button();
            this.grpboxFilterOptions = new System.Windows.Forms.GroupBox();
            this.chkDeletesDisabled = new System.Windows.Forms.CheckBox();
            this.chkCreateLog = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnFilterByFileBrowse = new System.Windows.Forms.Button();
            this.txtFilterByFile = new System.Windows.Forms.TextBox();
            this.rdoFilterByFile = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.grpboxFilter.SuspendLayout();
            this.grpboxFilterOptions.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.grpboxFilter);
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(704, 423);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(698, 24);
            this.lblTitle.TabIndex = 32;
            this.lblTitle.Text = "Delete Packages";
            // 
            // grpboxFilter
            // 
            this.grpboxFilter.Controls.Add(this.btnFilterFinish);
            this.grpboxFilter.Controls.Add(this.btnFilterHelp);
            this.grpboxFilter.Controls.Add(this.btnFilterPrevious);
            this.grpboxFilter.Controls.Add(this.btnFilterNext);
            this.grpboxFilter.Controls.Add(this.btnFilterCancel);
            this.grpboxFilter.Controls.Add(this.grpboxFilterOptions);
            this.grpboxFilter.Location = new System.Drawing.Point(16, 46);
            this.grpboxFilter.Name = "grpboxFilter";
            this.grpboxFilter.Size = new System.Drawing.Size(672, 361);
            this.grpboxFilter.TabIndex = 31;
            this.grpboxFilter.TabStop = false;
            this.grpboxFilter.Text = "Filter Options";
            // 
            // btnFilterFinish
            // 
            this.btnFilterFinish.Enabled = false;
            this.btnFilterFinish.Location = new System.Drawing.Point(584, 326);
            this.btnFilterFinish.Name = "btnFilterFinish";
            this.btnFilterFinish.Size = new System.Drawing.Size(72, 24);
            this.btnFilterFinish.TabIndex = 5;
            this.btnFilterFinish.Text = "Delete";
            this.btnFilterFinish.Click += new System.EventHandler(this.btnFilterFinish_Click);
            // 
            // btnFilterHelp
            // 
            this.btnFilterHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterHelp.Location = new System.Drawing.Point(16, 326);
            this.btnFilterHelp.Name = "btnFilterHelp";
            this.btnFilterHelp.Size = new System.Drawing.Size(64, 24);
            this.btnFilterHelp.TabIndex = 2;
            this.btnFilterHelp.Text = "Help";
            this.btnFilterHelp.Click += new System.EventHandler(this.btnFilterHelp_Click);
            // 
            // btnFilterPrevious
            // 
            this.btnFilterPrevious.Enabled = false;
            this.btnFilterPrevious.Location = new System.Drawing.Point(424, 326);
            this.btnFilterPrevious.Name = "btnFilterPrevious";
            this.btnFilterPrevious.Size = new System.Drawing.Size(72, 24);
            this.btnFilterPrevious.TabIndex = 4;
            this.btnFilterPrevious.Text = "< Previous";
            // 
            // btnFilterNext
            // 
            this.btnFilterNext.Enabled = false;
            this.btnFilterNext.Location = new System.Drawing.Point(496, 326);
            this.btnFilterNext.Name = "btnFilterNext";
            this.btnFilterNext.Size = new System.Drawing.Size(72, 24);
            this.btnFilterNext.TabIndex = 5;
            this.btnFilterNext.Text = "Next >";
            // 
            // btnFilterCancel
            // 
            this.btnFilterCancel.Location = new System.Drawing.Point(336, 326);
            this.btnFilterCancel.Name = "btnFilterCancel";
            this.btnFilterCancel.Size = new System.Drawing.Size(64, 24);
            this.btnFilterCancel.TabIndex = 3;
            this.btnFilterCancel.Text = "Cancel";
            this.btnFilterCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // grpboxFilterOptions
            // 
            this.grpboxFilterOptions.Controls.Add(this.chkDeletesDisabled);
            this.grpboxFilterOptions.Controls.Add(this.chkCreateLog);
            this.grpboxFilterOptions.Controls.Add(this.label1);
            this.grpboxFilterOptions.Controls.Add(this.btnFilterByFileBrowse);
            this.grpboxFilterOptions.Controls.Add(this.txtFilterByFile);
            this.grpboxFilterOptions.Controls.Add(this.rdoFilterByFile);
            this.grpboxFilterOptions.Location = new System.Drawing.Point(85, 59);
            this.grpboxFilterOptions.Name = "grpboxFilterOptions";
            this.grpboxFilterOptions.Size = new System.Drawing.Size(519, 249);
            this.grpboxFilterOptions.TabIndex = 1;
            this.grpboxFilterOptions.TabStop = false;
            // 
            // chkDeletesDisabled
            // 
            this.chkDeletesDisabled.AutoSize = true;
            this.chkDeletesDisabled.Location = new System.Drawing.Point(154, 220);
            this.chkDeletesDisabled.Name = "chkDeletesDisabled";
            this.chkDeletesDisabled.Size = new System.Drawing.Size(232, 21);
            this.chkDeletesDisabled.TabIndex = 8;
            this.chkDeletesDisabled.Text = "Count Records Without Deleting";
            this.chkDeletesDisabled.UseVisualStyleBackColor = true;
            // 
            // chkCreateLog
            // 
            this.chkCreateLog.AutoSize = true;
            this.chkCreateLog.Checked = true;
            this.chkCreateLog.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCreateLog.Location = new System.Drawing.Point(51, 220);
            this.chkCreateLog.Name = "chkCreateLog";
            this.chkCreateLog.Size = new System.Drawing.Size(126, 21);
            this.chkCreateLog.TabIndex = 7;
            this.chkCreateLog.Text = "Create Log File";
            this.chkCreateLog.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(619, 25);
            this.label1.TabIndex = 6;
            this.label1.Text = "Warning: Deleting packages is irreversible after clicking Delete.";
            // 
            // btnFilterByFileBrowse
            // 
            this.btnFilterByFileBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnFilterByFileBrowse.Image")));
            this.btnFilterByFileBrowse.Location = new System.Drawing.Point(409, 120);
            this.btnFilterByFileBrowse.Name = "btnFilterByFileBrowse";
            this.btnFilterByFileBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnFilterByFileBrowse.TabIndex = 4;
            this.btnFilterByFileBrowse.Click += new System.EventHandler(this.btnFilterByFileBrowse_click);
            // 
            // txtFilterByFile
            // 
            this.txtFilterByFile.Location = new System.Drawing.Point(75, 127);
            this.txtFilterByFile.Name = "txtFilterByFile";
            this.txtFilterByFile.Size = new System.Drawing.Size(328, 22);
            this.txtFilterByFile.TabIndex = 3;
            this.txtFilterByFile.TextChanged += new System.EventHandler(this.txtFilterByFile_TextChanged);
            // 
            // rdoFilterByFile
            // 
            this.rdoFilterByFile.Checked = true;
            this.rdoFilterByFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByFile.Location = new System.Drawing.Point(51, 91);
            this.rdoFilterByFile.Name = "rdoFilterByFile";
            this.rdoFilterByFile.Size = new System.Drawing.Size(432, 32);
            this.rdoFilterByFile.TabIndex = 2;
            this.rdoFilterByFile.TabStop = true;
            this.rdoFilterByFile.Text = "Delete Packages with Text File of Package Numbers";
            this.rdoFilterByFile.Click += new System.EventHandler(this.rdoFilterByFile_Click);
            // 
            // uc_delete_packages
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_delete_packages";
            this.Size = new System.Drawing.Size(704, 427);
            this.groupBox1.ResumeLayout(false);
            this.grpboxFilter.ResumeLayout(false);
            this.grpboxFilterOptions.ResumeLayout(false);
            this.grpboxFilterOptions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox grpboxFilter;
        private System.Windows.Forms.Button btnFilterFinish;
        private System.Windows.Forms.Button btnFilterHelp;
        private System.Windows.Forms.Button btnFilterPrevious;
        private System.Windows.Forms.Button btnFilterNext;
        private System.Windows.Forms.Button btnFilterCancel;
        private System.Windows.Forms.GroupBox grpboxFilterOptions;
        private System.Windows.Forms.Button btnFilterByFileBrowse;
        private System.Windows.Forms.TextBox txtFilterByFile;
        private System.Windows.Forms.RadioButton rdoFilterByFile;
        public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkCreateLog;
        private System.Windows.Forms.CheckBox chkDeletesDisabled;
    }
}

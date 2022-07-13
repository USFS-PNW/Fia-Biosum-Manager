namespace FIA_Biosum_Manager
{
    partial class uc_processor_scenario_run
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnRun = new System.Windows.Forms.Button();
            this.cmbFilter = new System.Windows.Forms.ComboBox();
            this.pnlFileSizeMonitor = new System.Windows.Forms.Panel();
            this.uc_filesize_monitor3 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor2 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor1 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.btnRunOC7 = new System.Windows.Forms.Button();
            this.lblMsg = new System.Windows.Forms.Label();
            this.lstFvsOutput = new System.Windows.Forms.ListView();
            this.btnUncheckAll = new System.Windows.Forms.Button();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.pnlFileSizeMonitor.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.panel1);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(1007, 686);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.panel1.Controls.Add(this.btnRun);
            this.panel1.Controls.Add(this.cmbFilter);
            this.panel1.Controls.Add(this.pnlFileSizeMonitor);
            this.panel1.Controls.Add(this.btnRunOC7);
            this.panel1.Controls.Add(this.lblMsg);
            this.panel1.Controls.Add(this.lstFvsOutput);
            this.panel1.Controls.Add(this.btnUncheckAll);
            this.panel1.Controls.Add(this.btnChkAll);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(4, 58);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(999, 624);
            this.panel1.TabIndex = 32;
            this.panel1.Resize += new System.EventHandler(this.panel1_Resize);
            // 
            // btnRun
            // 
            this.btnRun.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRun.Location = new System.Drawing.Point(423, 436);
            this.btnRun.Margin = new System.Windows.Forms.Padding(4);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(168, 39);
            this.btnRun.TabIndex = 73;
            this.btnRun.Text = "Run";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // cmbFilter
            // 
            this.cmbFilter.FormattingEnabled = true;
            this.cmbFilter.Location = new System.Drawing.Point(25, 444);
            this.cmbFilter.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFilter.Name = "cmbFilter";
            this.cmbFilter.Size = new System.Drawing.Size(80, 24);
            this.cmbFilter.TabIndex = 69;
            // 
            // pnlFileSizeMonitor
            // 
            this.pnlFileSizeMonitor.AutoScroll = true;
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor3);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor2);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor1);
            this.pnlFileSizeMonitor.Location = new System.Drawing.Point(25, 482);
            this.pnlFileSizeMonitor.Margin = new System.Windows.Forms.Padding(4);
            this.pnlFileSizeMonitor.Name = "pnlFileSizeMonitor";
            this.pnlFileSizeMonitor.Size = new System.Drawing.Size(836, 122);
            this.pnlFileSizeMonitor.TabIndex = 68;
            // 
            // uc_filesize_monitor3
            // 
            this.uc_filesize_monitor3.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor3.Information = "";
            this.uc_filesize_monitor3.Location = new System.Drawing.Point(572, 12);
            this.uc_filesize_monitor3.Margin = new System.Windows.Forms.Padding(5);
            this.uc_filesize_monitor3.Name = "uc_filesize_monitor3";
            this.uc_filesize_monitor3.Size = new System.Drawing.Size(241, 94);
            this.uc_filesize_monitor3.TabIndex = 2;
            this.uc_filesize_monitor3.Visible = false;
            // 
            // uc_filesize_monitor2
            // 
            this.uc_filesize_monitor2.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor2.Information = "";
            this.uc_filesize_monitor2.Location = new System.Drawing.Point(275, 12);
            this.uc_filesize_monitor2.Margin = new System.Windows.Forms.Padding(5);
            this.uc_filesize_monitor2.Name = "uc_filesize_monitor2";
            this.uc_filesize_monitor2.Size = new System.Drawing.Size(241, 94);
            this.uc_filesize_monitor2.TabIndex = 1;
            this.uc_filesize_monitor2.Visible = false;
            // 
            // uc_filesize_monitor1
            // 
            this.uc_filesize_monitor1.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor1.Information = "";
            this.uc_filesize_monitor1.Location = new System.Drawing.Point(7, 12);
            this.uc_filesize_monitor1.Margin = new System.Windows.Forms.Padding(5);
            this.uc_filesize_monitor1.Name = "uc_filesize_monitor1";
            this.uc_filesize_monitor1.Size = new System.Drawing.Size(241, 94);
            this.uc_filesize_monitor1.TabIndex = 0;
            this.uc_filesize_monitor1.Visible = false;
            // 
            // btnRunOC7
            // 
            this.btnRunOC7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRunOC7.Location = new System.Drawing.Point(599, 436);
            this.btnRunOC7.Margin = new System.Windows.Forms.Padding(4);
            this.btnRunOC7.Name = "btnRunOC7";
            this.btnRunOC7.Size = new System.Drawing.Size(168, 39);
            this.btnRunOC7.TabIndex = 59;
            this.btnRunOC7.Text = "Run OC7";
            this.btnRunOC7.Visible = false;
            this.btnRunOC7.Click += new System.EventHandler(this.btnRunOC7_Click);
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.Location = new System.Drawing.Point(21, 402);
            this.lblMsg.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(193, 20);
            this.lblMsg.TabIndex = 58;
            this.lblMsg.Text = "Run Status Messages";
            this.lblMsg.Visible = false;
            // 
            // lstFvsOutput
            // 
            this.lstFvsOutput.CheckBoxes = true;
            this.lstFvsOutput.GridLines = true;
            this.lstFvsOutput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsOutput.HideSelection = false;
            this.lstFvsOutput.Location = new System.Drawing.Point(25, 16);
            this.lstFvsOutput.Margin = new System.Windows.Forms.Padding(4);
            this.lstFvsOutput.MultiSelect = false;
            this.lstFvsOutput.Name = "lstFvsOutput";
            this.lstFvsOutput.Size = new System.Drawing.Size(951, 381);
            this.lstFvsOutput.TabIndex = 57;
            this.lstFvsOutput.UseCompatibleStateImageBehavior = false;
            this.lstFvsOutput.View = System.Windows.Forms.View.Details;
            // 
            // btnUncheckAll
            // 
            this.btnUncheckAll.Location = new System.Drawing.Point(208, 436);
            this.btnUncheckAll.Margin = new System.Windows.Forms.Padding(4);
            this.btnUncheckAll.Name = "btnUncheckAll";
            this.btnUncheckAll.Size = new System.Drawing.Size(108, 39);
            this.btnUncheckAll.TabIndex = 55;
            this.btnUncheckAll.Text = "Uncheck All";
            this.btnUncheckAll.Click += new System.EventHandler(this.btnUncheckAll_Click);
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(115, 436);
            this.btnChkAll.Margin = new System.Windows.Forms.Padding(4);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(85, 39);
            this.btnChkAll.TabIndex = 54;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(4, 19);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(999, 39);
            this.lblTitle.TabIndex = 31;
            this.lblTitle.Text = "Run Processor Scenario";
            // 
            // uc_processor_scenario_run
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "uc_processor_scenario_run";
            this.Size = new System.Drawing.Size(1007, 686);
            this.groupBox1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.pnlFileSizeMonitor.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.Label lblTitle;
        public System.Windows.Forms.ListView lstFvsOutput;
        private System.Windows.Forms.Button btnUncheckAll;
        private System.Windows.Forms.Button btnChkAll;
        private System.Windows.Forms.Button btnRunOC7;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.Panel pnlFileSizeMonitor;
        private uc_filesize_monitor uc_filesize_monitor3;
        private uc_filesize_monitor uc_filesize_monitor2;
        private uc_filesize_monitor uc_filesize_monitor1;
        private System.Windows.Forms.ComboBox cmbFilter;
        private System.Windows.Forms.Button btnRun;
    }
}

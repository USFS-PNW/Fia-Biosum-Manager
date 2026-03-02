namespace FIA_Biosum_Manager
{
    partial class frmScanAndSynchronizeProjectFolderTool
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmScanAndSynchronizeProjectFolderTool));
            this.lvDatasources = new System.Windows.Forms.ListView();
            this.toolBar1 = new System.Windows.Forms.ToolBar();
            this.tlbBtnEdit = new System.Windows.Forms.ToolBarButton();
            this.tlbBtnRefresh = new System.Windows.Forms.ToolBarButton();
            this.tlbBtnSync = new System.Windows.Forms.ToolBarButton();
            this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
            this.tlbBtnHelp = new System.Windows.Forms.ToolBarButton();
            this.tlbBtnClose = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.txtReplace = new System.Windows.Forms.TextBox();
            this.lblCurrentProjectFolder = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.grpStatus = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lblFolderPathNotSynchronized = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblFolderPaths = new System.Windows.Forms.Label();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.grpStatus.SuspendLayout();
            this.SuspendLayout();
            // 
            // lvDatasources
            // 
            this.lvDatasources.CheckBoxes = true;
            this.lvDatasources.GridLines = true;
            this.lvDatasources.HideSelection = false;
            this.lvDatasources.Location = new System.Drawing.Point(12, 192);
            this.lvDatasources.MultiSelect = false;
            this.lvDatasources.Name = "lvDatasources";
            this.lvDatasources.Size = new System.Drawing.Size(543, 215);
            this.lvDatasources.TabIndex = 2;
            this.lvDatasources.UseCompatibleStateImageBehavior = false;
            this.lvDatasources.View = System.Windows.Forms.View.Details;
            // 
            // toolBar1
            // 
            this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.tlbBtnEdit,
            this.tlbBtnRefresh,
            this.tlbBtnSync,
            this.toolBarButton1,
            this.tlbBtnHelp,
            this.tlbBtnClose});
            this.toolBar1.ButtonSize = new System.Drawing.Size(45, 40);
            this.toolBar1.DropDownArrows = true;
            this.toolBar1.ImageList = this.imageList1;
            this.toolBar1.Location = new System.Drawing.Point(0, 0);
            this.toolBar1.Name = "toolBar1";
            this.toolBar1.ShowToolTips = true;
            this.toolBar1.Size = new System.Drawing.Size(784, 44);
            this.toolBar1.TabIndex = 4;
            this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
            // 
            // tlbBtnEdit
            // 
            this.tlbBtnEdit.ImageIndex = 0;
            this.tlbBtnEdit.Name = "tlbBtnEdit";
            this.tlbBtnEdit.Text = "Save";
            // 
            // tlbBtnRefresh
            // 
            this.tlbBtnRefresh.ImageIndex = 1;
            this.tlbBtnRefresh.Name = "tlbBtnRefresh";
            this.tlbBtnRefresh.Text = "Refresh";
            this.tlbBtnRefresh.ToolTipText = "Scan and identify out of sync folders";
            // 
            // tlbBtnSync
            // 
            this.tlbBtnSync.ImageIndex = 2;
            this.tlbBtnSync.Name = "tlbBtnSync";
            this.tlbBtnSync.Text = "Sync";
            this.tlbBtnSync.ToolTipText = "Scan and Replace";
            // 
            // toolBarButton1
            // 
            this.toolBarButton1.Name = "toolBarButton1";
            this.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // tlbBtnHelp
            // 
            this.tlbBtnHelp.ImageIndex = 4;
            this.tlbBtnHelp.Name = "tlbBtnHelp";
            this.tlbBtnHelp.Text = "Help";
            // 
            // tlbBtnClose
            // 
            this.tlbBtnClose.ImageIndex = 3;
            this.tlbBtnClose.Name = "tlbBtnClose";
            this.tlbBtnClose.Text = "Close";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "save.bmp");
            this.imageList1.Images.SetKeyName(1, "refresh-icon.png");
            this.imageList1.Images.SetKeyName(2, "Problem_16x16.ico");
            this.imageList1.Images.SetKeyName(3, "");
            this.imageList1.Images.SetKeyName(4, "");
            // 
            // txtReplace
            // 
            this.txtReplace.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtReplace.Location = new System.Drawing.Point(15, 79);
            this.txtReplace.Name = "txtReplace";
            this.txtReplace.Size = new System.Drawing.Size(540, 21);
            this.txtReplace.TabIndex = 5;
            // 
            // lblCurrentProjectRootFolder
            // 
            this.lblCurrentProjectFolder.AutoSize = true;
            this.lblCurrentProjectFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCurrentProjectFolder.Location = new System.Drawing.Point(12, 161);
            this.lblCurrentProjectFolder.Name = "lblCurrentProjectFolder";
            this.lblCurrentProjectFolder.Size = new System.Drawing.Size(196, 16);
            this.lblCurrentProjectFolder.TabIndex = 0;
            this.lblCurrentProjectFolder.Text = "Current Project Folder";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(291, 47);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 16);
            this.label1.TabIndex = 7;
            this.label1.Text = "REPLACE";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(187, 122);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(310, 16);
            this.label2.TabIndex = 8;
            this.label2.Text = "WITH CURRENT PROJECT FOLDER";
            // 
            // grpStatus
            // 
            this.grpStatus.Controls.Add(this.label4);
            this.grpStatus.Controls.Add(this.lblFolderPathNotSynchronized);
            this.grpStatus.Controls.Add(this.label3);
            this.grpStatus.Controls.Add(this.lblFolderPaths);
            this.grpStatus.Location = new System.Drawing.Point(561, 192);
            this.grpStatus.Name = "grpStatus";
            this.grpStatus.Size = new System.Drawing.Size(190, 197);
            this.grpStatus.TabIndex = 9;
            this.grpStatus.TabStop = false;
            this.grpStatus.Text = "Status";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 99);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(151, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "Folder Path Not Synchronized";
            // 
            // lblFolderPathNotSynchronized
            // 
            this.lblFolderPathNotSynchronized.BackColor = System.Drawing.Color.Beige;
            this.lblFolderPathNotSynchronized.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFolderPathNotSynchronized.Location = new System.Drawing.Point(19, 127);
            this.lblFolderPathNotSynchronized.Name = "lblFolderPathNotSynchronized";
            this.lblFolderPathNotSynchronized.Size = new System.Drawing.Size(142, 29);
            this.lblFolderPathNotSynchronized.TabIndex = 13;
            this.lblFolderPathNotSynchronized.Text = "0";
            this.lblFolderPathNotSynchronized.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(119, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = "Folder Paths Not Found";
            // 
            // lblFolderPaths
            // 
            this.lblFolderPaths.BackColor = System.Drawing.Color.Beige;
            this.lblFolderPaths.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFolderPaths.Location = new System.Drawing.Point(19, 48);
            this.lblFolderPaths.Name = "lblFolderPaths";
            this.lblFolderPaths.Size = new System.Drawing.Size(142, 29);
            this.lblFolderPaths.TabIndex = 11;
            this.lblFolderPaths.Text = "0";
            this.lblFolderPaths.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Enabled = false;
            this.btnAnalyze.Location = new System.Drawing.Point(561, 74);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(78, 30);
            this.btnAnalyze.TabIndex = 10;
            this.btnAnalyze.Text = "Analyze";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // frmScanAndSynchronizeProjectFolderTool
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(784, 419);
            this.Controls.Add(this.btnAnalyze);
            this.Controls.Add(this.grpStatus);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblCurrentProjectFolder);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtReplace);
            this.Controls.Add(this.toolBar1);
            this.Controls.Add(this.lvDatasources);
            this.Name = "frmScanAndSynchronizeProjectFolderTool";
            this.Text = "Scan and Synchronize Project Folder Tool";
            this.Resize += new System.EventHandler(this.frmScanAndSynchronizeProjectFolderTool_Resize);
            this.grpStatus.ResumeLayout(false);
            this.grpStatus.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ListView lvDatasources;
        private System.Windows.Forms.ToolBar toolBar1;
        private System.Windows.Forms.ToolBarButton tlbBtnSave;
        private System.Windows.Forms.ToolBarButton tlbBtnScan;
        private System.Windows.Forms.ToolBarButton toolBarButton1;
        private System.Windows.Forms.ToolBarButton tlbBtnHelp;
        private System.Windows.Forms.ToolBarButton tlbBtnClose;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.TextBox txtReplace;
        private System.Windows.Forms.Label lblCurrentProjectFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ToolBarButton tlbBtnSync;
        private System.Windows.Forms.ToolBarButton tlbBtnEdit;
        private System.Windows.Forms.ToolBarButton tlbBtnRefresh;
        private System.Windows.Forms.GroupBox grpStatus;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblFolderPathNotSynchronized;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblFolderPaths;
        private System.Windows.Forms.Button btnAnalyze;
    }
}
using System;
using System.Collections;
using SQLite.ADO;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;
using System.Linq;


namespace FIA_Biosum_Manager
{
	/// <summary> </summary>
	public class uc_fvs_output : System.Windows.Forms.UserControl
	{
		public Hashtable htSelectedRxFile=null;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnHelp;
        private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.Button btnRefresh;
		private System.Windows.Forms.Button btnClearAll;
		private System.Windows.Forms.Button btnChkAll;
		
		private System.Windows.Forms.TextBox txtOutDir;
		private System.Windows.Forms.Label label4;
		public int m_intError=0;
        public int m_intOverallError = 0;
		public string m_strError="";
		public int m_intWarning=0;
		public string m_strWarning="";
        private string m_strDateTimeCreated = "";
		public System.Windows.Forms.ListView lstFvsOutput;

		private const int COL_RX = -1;
		private const int COL_CHECKBOX=0;
		private const int COL_VARIANT = 1;
	    private const int COL_PACKAGE = 2;
		private const int COL_RUNSTATUS = 3;
		private const int COL_FOUND=4;
		private const int COL_SUMMARYCOUNT = 5;
		private const int COL_CUTCOUNT = 6;

		private string m_strFvsTreeTable="";
		private string m_strFVSSummaryAuditYearCountsTable="audit_fvs_summary_year_counts_table";
        private string m_strFVSSummaryAuditPrePostSeqNumTable = "fvs_summary_prepost_seqnum_matrix";
        private string m_strFVSSummaryAuditPrePostSeqNumCountsTable = "audit_fvs_summary_prepost_seqnum_counts_table";
        private List<string> m_strFVSPreAppendAuditTables = null;
        private List<string> m_strFVSPostAppendAuditTables = null;
		private string m_strProjDir="";
        private string m_strFvsTreeDb= frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
        private string m_strFvsOutDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile;
        private string m_strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile;
        private string m_dbConn = "";
        private string m_missingFvsOutDb = "This project has no valid /fvs/data/FVSOut.db. Multiple FVS Out functions will not work! Please use FVS to generate an /fvs/data/FVSOut.db file.";

        //POTFIRE BASE YEAR
        private string m_strPotFireTable = "FVS_POTFIRE_TEMP";
        private string m_strPotFireBaseYearAccessLinkedTableName = "FVS_POTFIRE_BASEYEAR";
        private string m_strPotFireStandardAccessLinkedTableName = "FVS_POTFIRE_STANDARD";
        private bool m_bPotFireBaseYearTableExist = true;

		string m_strLogFile;
		string m_strLogDate;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;

		private System.Threading.Thread m_thread;
		private FIA_Biosum_Manager.frmTherm m_frmTherm;
		private bool m_bDebug=true;
        private string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_fvsout_debug.txt";
		private System.Windows.Forms.Label lblRunStatus;
        private FIA_Biosum_Manager.utils m_oUtils = new FIA_Biosum_Manager.utils();
		private bool _bDisplayAuditMsg=true;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Button btnViewLogFile;
		private System.Windows.Forms.Button btnAuditDb;
		private ListViewAlternateBackgroundColors m_oLvAlternateColors = new ListViewAlternateBackgroundColors();
		
		
		Queries m_oQueries = new Queries();
		RxTools m_oRxTools = new RxTools();
        Tables m_oTables = new Tables();
		FIA_Biosum_Manager.RxPackageItem_Collection m_oRxPackageItem_Collection=null;
		FIA_Biosum_Manager.RxPackageItem m_oRxPackageItem=null;
		string m_strRxCycleList="";
		string[] m_strRxCycleArray=null;

        private Button btnViewPostLogFile;
        private Label lblMsg;

        int m_intProgressOverallTotalCount = 0;
        int m_intProgressStepTotalCount = 0;
        int m_intProgressOverallCurrentCount = 0;
        int m_intProgressStepCurrentCount = 0;

        private List<string> m_oFVSTables;
        private Button btnExecute;
        private ComboBox cmbStep;
        private Panel pnlFileSizeMonitor;
        private uc_filesize_monitor uc_filesize_monitor3;
        private uc_filesize_monitor uc_filesize_monitor2;
        private uc_filesize_monitor uc_filesize_monitor1;
        private ComboBox cmbFilter;

        private FVSPrePostSeqNumItem_Collection m_oFVSPrePostSeqNumItemCollection = null;
        private FVSPrePostSeqNumItem m_oFVSPrePostSeqNumItem = null;
        private DbFileItem_Collection m_oPrePostDbFileItem_Collection = null;
        private Button btnPostAppendAuditDb;

        // Mapping of sqlite column types to Access column names. Add new exception entries here.
        public static Dictionary<string, string> m_SqliteToAccessColTypes = new Dictionary<string, string>
        {
            { "SPECIESFIA", "INTEGER" }
        };
        public static string m_colBioSumAppend = "BIOSUM_Append_YN";

        private SQLite.ADO.DataMgr _SQLite = new SQLite.ADO.DataMgr();
        public SQLite.ADO.DataMgr SQLite
        {
            get { return _SQLite; }
            set { _SQLite = value; }
        }

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_fvs_output(string p_strProjDir)
		{

			InitializeComponent();
			this.m_strProjDir = p_strProjDir;

			this.txtOutDir.Text = this.m_strProjDir + "\\fvs\\data";

			this.m_oQueries = new Queries();
			m_oQueries.m_oFvs.LoadDatasource=true;
			m_oQueries.m_oFIAPlot.LoadDatasource=true;
			m_oQueries.LoadDatasources(true);
            this.m_oEnv = new env();

            this.m_bDebug = frmMain.g_bDebug;
			htSelectedRxFile=new Hashtable();
		}
       

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.lblTitle = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnPostAppendAuditDb = new System.Windows.Forms.Button();
            this.cmbFilter = new System.Windows.Forms.ComboBox();
            this.pnlFileSizeMonitor = new System.Windows.Forms.Panel();
            this.uc_filesize_monitor3 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor2 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor1 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.btnExecute = new System.Windows.Forms.Button();
            this.cmbStep = new System.Windows.Forms.ComboBox();
            this.lblMsg = new System.Windows.Forms.Label();
            this.btnViewPostLogFile = new System.Windows.Forms.Button();
            this.btnAuditDb = new System.Windows.Forms.Button();
            this.btnViewLogFile = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblRunStatus = new System.Windows.Forms.Label();
            this.lstFvsOutput = new System.Windows.Forms.ListView();
            this.txtOutDir = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.pnlFileSizeMonitor.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(714, 32);
            this.lblTitle.TabIndex = 2;
            this.lblTitle.Text = "Join And Append FVS Out Data ";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(8, 88);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(192, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Available FVS Out MDB Tables";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnPostAppendAuditDb);
            this.groupBox1.Controls.Add(this.cmbFilter);
            this.groupBox1.Controls.Add(this.pnlFileSizeMonitor);
            this.groupBox1.Controls.Add(this.btnExecute);
            this.groupBox1.Controls.Add(this.cmbStep);
            this.groupBox1.Controls.Add(this.lblMsg);
            this.groupBox1.Controls.Add(this.btnViewPostLogFile);
            this.groupBox1.Controls.Add(this.btnAuditDb);
            this.groupBox1.Controls.Add(this.btnViewLogFile);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.lblRunStatus);
            this.groupBox1.Controls.Add(this.lstFvsOutput);
            this.groupBox1.Controls.Add(this.txtOutDir);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.btnClearAll);
            this.groupBox1.Controls.Add(this.btnChkAll);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(720, 525);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // btnPostAppendAuditDb
            // 
            this.btnPostAppendAuditDb.Location = new System.Drawing.Point(427, 322);
            this.btnPostAppendAuditDb.Name = "btnPostAppendAuditDb";
            this.btnPostAppendAuditDb.Size = new System.Drawing.Size(159, 21);
            this.btnPostAppendAuditDb.TabIndex = 71;
            this.btnPostAppendAuditDb.Text = "POST-APPEND Audit Tables";
            this.btnPostAppendAuditDb.Click += new System.EventHandler(this.btnPostAppendAuditDb_Click);
            // 
            // cmbFilter
            // 
            this.cmbFilter.FormattingEnabled = true;
            this.cmbFilter.Location = new System.Drawing.Point(11, 305);
            this.cmbFilter.Name = "cmbFilter";
            this.cmbFilter.Size = new System.Drawing.Size(61, 24);
            this.cmbFilter.TabIndex = 70;
            // 
            // pnlFileSizeMonitor
            // 
            this.pnlFileSizeMonitor.AutoScroll = true;
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor3);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor2);
            this.pnlFileSizeMonitor.Controls.Add(this.uc_filesize_monitor1);
            this.pnlFileSizeMonitor.Location = new System.Drawing.Point(6, 364);
            this.pnlFileSizeMonitor.Name = "pnlFileSizeMonitor";
            this.pnlFileSizeMonitor.Size = new System.Drawing.Size(706, 99);
            this.pnlFileSizeMonitor.TabIndex = 67;
            // 
            // uc_filesize_monitor3
            // 
            this.uc_filesize_monitor3.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor3.Information = "";
            this.uc_filesize_monitor3.Location = new System.Drawing.Point(429, 10);
            this.uc_filesize_monitor3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uc_filesize_monitor3.Name = "uc_filesize_monitor3";
            this.uc_filesize_monitor3.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor3.TabIndex = 2;
            this.uc_filesize_monitor3.Visible = false;
            // 
            // uc_filesize_monitor2
            // 
            this.uc_filesize_monitor2.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor2.Information = "";
            this.uc_filesize_monitor2.Location = new System.Drawing.Point(206, 10);
            this.uc_filesize_monitor2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uc_filesize_monitor2.Name = "uc_filesize_monitor2";
            this.uc_filesize_monitor2.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor2.TabIndex = 1;
            this.uc_filesize_monitor2.Visible = false;
            // 
            // uc_filesize_monitor1
            // 
            this.uc_filesize_monitor1.ForeColor = System.Drawing.Color.Black;
            this.uc_filesize_monitor1.Information = "";
            this.uc_filesize_monitor1.Location = new System.Drawing.Point(5, 10);
            this.uc_filesize_monitor1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.uc_filesize_monitor1.Name = "uc_filesize_monitor1";
            this.uc_filesize_monitor1.Size = new System.Drawing.Size(181, 76);
            this.uc_filesize_monitor1.TabIndex = 0;
            this.uc_filesize_monitor1.Visible = false;
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(312, 336);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(112, 21);
            this.btnExecute.TabIndex = 66;
            this.btnExecute.Text = "Execute Step";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // cmbStep
            // 
            this.cmbStep.FormattingEnabled = true;
            this.cmbStep.Items.AddRange(new object[] {
            "Step 1 - Define PRE/POST Table SeqNum",
            "Step 2 - Pre-Processing Audit Check",
            "Step 3 - Append FVS Output Data",
            "Step 4 - Post-Processing Audit Check",
            "---Optional---",
            "Create FVSOut_BioSum.db",
            "Write FVS_InForest to FVSOUT_TREE_LIST.db"});
            this.cmbStep.Location = new System.Drawing.Point(8, 337);
            this.cmbStep.Name = "cmbStep";
            this.cmbStep.Size = new System.Drawing.Size(298, 24);
            this.cmbStep.TabIndex = 65;
            this.cmbStep.Text = "Step 1 - Define PRE/POST Table SeqNum";
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(8, 282);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(294, 17);
            this.lblMsg.TabIndex = 64;
            this.lblMsg.Text = "Tasks To Complete For Each Item: a=Append";
            // 
            // btnViewPostLogFile
            // 
            this.btnViewPostLogFile.Location = new System.Drawing.Point(588, 322);
            this.btnViewPostLogFile.Name = "btnViewPostLogFile";
            this.btnViewPostLogFile.Size = new System.Drawing.Size(124, 21);
            this.btnViewPostLogFile.TabIndex = 62;
            this.btnViewPostLogFile.Text = "Open Post Audit Log";
            this.btnViewPostLogFile.Click += new System.EventHandler(this.btnViewPostLogFile_Click);
            // 
            // btnAuditDb
            // 
            this.btnAuditDb.Location = new System.Drawing.Point(427, 298);
            this.btnAuditDb.Name = "btnAuditDb";
            this.btnAuditDb.Size = new System.Drawing.Size(159, 21);
            this.btnAuditDb.TabIndex = 60;
            this.btnAuditDb.Text = "PRE-APPEND Audit Tables";
            this.btnAuditDb.Click += new System.EventHandler(this.btnAuditDb_Click);
            // 
            // btnViewLogFile
            // 
            this.btnViewLogFile.Location = new System.Drawing.Point(588, 298);
            this.btnViewLogFile.Name = "btnViewLogFile";
            this.btnViewLogFile.Size = new System.Drawing.Size(124, 21);
            this.btnViewLogFile.TabIndex = 59;
            this.btnViewLogFile.Text = "Open Pre Audit Log";
            this.btnViewLogFile.Click += new System.EventHandler(this.btnViewLogFile_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(296, 299);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(80, 32);
            this.btnCancel.TabIndex = 58;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblRunStatus
            // 
            this.lblRunStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRunStatus.Location = new System.Drawing.Point(168, 466);
            this.lblRunStatus.Name = "lblRunStatus";
            this.lblRunStatus.Size = new System.Drawing.Size(352, 32);
            this.lblRunStatus.TabIndex = 54;
            this.lblRunStatus.Text = "Run Status";
            this.lblRunStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblRunStatus.Visible = false;
            // 
            // lstFvsOutput
            // 
            this.lstFvsOutput.CheckBoxes = true;
            this.lstFvsOutput.GridLines = true;
            this.lstFvsOutput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsOutput.HideSelection = false;
            this.lstFvsOutput.Location = new System.Drawing.Point(8, 74);
            this.lstFvsOutput.MultiSelect = false;
            this.lstFvsOutput.Name = "lstFvsOutput";
            this.lstFvsOutput.Size = new System.Drawing.Size(706, 205);
            this.lstFvsOutput.TabIndex = 53;
            this.lstFvsOutput.UseCompatibleStateImageBehavior = false;
            this.lstFvsOutput.View = System.Windows.Forms.View.Details;
            this.lstFvsOutput.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstFvsOutput_ItemCheck);
            this.lstFvsOutput.SelectedIndexChanged += new System.EventHandler(this.lstFvsOutput_SelectedIndexChanged);
            this.lstFvsOutput.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFvsOutput_MouseUp);
            // 
            // txtOutDir
            // 
            this.txtOutDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOutDir.Location = new System.Drawing.Point(118, 48);
            this.txtOutDir.Name = "txtOutDir";
            this.txtOutDir.Size = new System.Drawing.Size(599, 23);
            this.txtOutDir.TabIndex = 52;
            this.txtOutDir.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtOutDir_KeyPress);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(9, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(128, 16);
            this.label4.TabIndex = 51;
            this.label4.Text = "FVS Output Directory";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(213, 299);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(64, 32);
            this.btnRefresh.TabIndex = 49;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(143, 299);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(64, 32);
            this.btnClearAll.TabIndex = 48;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(73, 299);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(64, 32);
            this.btnChkAll.TabIndex = 47;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(616, 485);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 45;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(11, 485);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 44;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // uc_fvs_output
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_output";
            this.Size = new System.Drawing.Size(720, 525);
            this.Resize += new System.EventHandler(this.uc_fvs_output_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.pnlFileSizeMonitor.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}
		public void loadvalues()
		{
			int x,y;
			string strPackage="";
			string strVariant="";
            string strCurVariant = "";
            string strCurRunTitle = "";
			try
			{
                InitializeAuditLogTableArray();
                cmbFilter.Items.Clear();
                cmbFilter.Items.Add("All");
                cmbFilter.Text = "All";
				System.Windows.Forms.ListViewItem entryListItem=null;
				this.m_oLvAlternateColors.InitializeRowCollection();
				this.m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
				this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceForegroundColor=frmMain.g_oGridViewRowForegroundColor;
				this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
				this.m_oLvAlternateColors.ReferenceListView = lstFvsOutput;
				this.m_oLvAlternateColors.CustomFullRowSelect=true;
				if (frmMain.g_oGridViewFont!=null) this.lstFvsOutput.Font = frmMain.g_oGridViewFont;
				this.lstFvsOutput.Clear();
				this.lstFvsOutput.Columns.Add("",2,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("FVS Variant", 90, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("RxPackage", 90, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Run Status",190,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("FVSOut.db Recs", 120, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Summary Recs", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Cut Recs", 100, HorizontalAlignment.Left);
              
				this.lstFvsOutput.Columns[COL_CHECKBOX].Width = -2;

                //ABORT IF THERE IS NO FVSOUT.DB
                // Warning for older projects without FVSOut.db
                string dbConn = SQLite.GetConnectionString(m_strFvsOutDb);
                if (System.IO.File.Exists(m_strFvsOutDb))
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                    {
                        conn.Open();
                        if (! SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTableName))
                        {
                            MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                // Get variants/rxPackages in project
                IDictionary<string, RxPackageItem_Collection> dictFvsVariantPackage = this.m_oRxTools.GetFvsVariantPackageDictionary(m_oQueries);
                IList<string> lstRunTitles = new List<string>();
                foreach (string key in dictFvsVariantPackage.Keys)
                {
                    strVariant = key;
                    RxPackageItem_Collection oRxPackageItemCollection = dictFvsVariantPackage[key];
                    lstRunTitles.Clear();   //Eliminate duplicate entries on FVSOut screen
                    for (int i = 0; i < oRxPackageItemCollection.Count; i++)
                    {
                        RxPackageItem rxPackageItem = oRxPackageItemCollection.Item(i);
                        strPackage = rxPackageItem.RxPackageId;

                        // Example RunTitle: FVSOUT_WC_P999-999-999-999-999
                        strCurRunTitle = this.m_oRxTools.GetRxPackageRunTitle(strVariant, rxPackageItem);
                        lstRunTitles.Add(strCurRunTitle);

                        long lngPreSummaryRecords = 0;
                        string strSummaryConnect = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}{Tables.FVS.DefaultFVSOutPrePostDbFile}";
                        if (System.IO.File.Exists(strSummaryConnect))
                        {
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strSummaryConnect)))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, Tables.FVS.DefaultPreFVSSummaryTableName))
                                {
                                    string strSQL = $@"SELECT COUNT(*) FROM(SELECT biosum_cond_id FROM {Tables.FVS.DefaultPreFVSSummaryTableName} WHERE 
                                    FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}' LIMIT 1)";
                                    lngPreSummaryRecords = SQLite.getRecordCount(conn, strSQL, Tables.FVS.DefaultPreFVSSummaryTableName);
                                }
                            }
                        }

                        /************************************************************************
                        /**Check and Assign in the FVS_CASES whether the FVS output has been 
                         **appended to the fvs_cutTree table
                         ************************************************************************/
                        if (System.IO.File.Exists(m_strFvsOutDb) == true)
                        {
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTableName))
                                {
                                    if (!SQLite.ColumnExist(conn, Tables.FVS.DefaultFVSCasesTableName, m_colBioSumAppend))
                                    {
                                        SQLite.AddColumn(conn, Tables.FVS.DefaultFVSCasesTableName, m_colBioSumAppend, "TEXT", "1");
                                        SQLite.m_strSQL = $@"UPDATE {Tables.FVS.DefaultFVSCasesTableName} SET {m_colBioSumAppend} = 'N'";
                                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    }
                                    else
                                    {
                                        //check if records exist in the summary table for this variant/package
                                        if (lngPreSummaryRecords == 0)
                                        {
                                            SQLite.m_strSQL = $@"UPDATE {Tables.FVS.DefaultFVSCasesTableName} SET {m_colBioSumAppend}='N' WHERE RunTitle = '{strCurRunTitle}'";
                                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                        }
                                    }
                                }
                            }
                            if (strVariant != strCurVariant)
                            {
                                cmbFilter.Items.Add(strVariant);
                                strCurVariant = strVariant;
                            }
                        }   // END IF FVS_OUT.DB EXISTS                   
                    }
                    // Warning for older projects without FVSOut.db (and cutlist)
                    if (!FvsOutWithRequiredTable(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(), Tables.FVS.DefaultFVSCasesTableName))
                    {
                        MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    strVariant = "";
                    strCurVariant = "";
                    for (x = 0; x <= lstRunTitles.Count - 1; x++)
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            long memused = GC.GetTotalMemory(true);
                            this.WriteText(m_strDebugFile, "uc_fvs_output.loadvalues Memory Used: " + String.Format("{0:n0}" + " MB", memused));
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        strCurRunTitle = lstRunTitles.ElementAt(x);
                        string[] arrVariant = strCurRunTitle.Split('_');
                        if (arrVariant.Length == 3)
                        {
                            strVariant = arrVariant[1];
                            string[] arrRx = arrVariant[2].Split('-');
                            if (arrRx.Length == 5)
                            {
                                strPackage = arrRx[0].Substring(1, 3);
                            }
                        }

                        frmMain.g_sbpInfo.Text = $@"Loading FVS Variant RxPackage {strVariant} {strPackage} ...Stand By";
                        frmMain.g_sbpInfo.Parent.Refresh();

                        // Add a ListItem object to the ListView.
                        entryListItem = this.lstFvsOutput.Items.Add("");
                        entryListItem.UseItemStyleForSubItems = false;
                        this.m_oLvAlternateColors.AddRow();
                        this.m_oLvAlternateColors.AddColumns(lstFvsOutput.Items.Count - 1, lstFvsOutput.Columns.Count);
                        entryListItem.SubItems.Add(strVariant);
                        entryListItem.SubItems.Add(strPackage);
                        entryListItem.SubItems.Add(" ");  //file found
                        entryListItem.SubItems.Add(" ");  //summary record count
                        entryListItem.SubItems.Add(" ");  //tree cut list record count
                        entryListItem.SubItems.Add(" ");  //potential fire base year out file
                        this.m_oLvAlternateColors.ListViewItem(lstFvsOutput.Items[lstFvsOutput.Items.Count - 1]);


                        //FVS_SUMMARY
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                        {
                            conn.Open();
                            // VARIANT/PACKAGE FOUND IN FVS_CASES TABLE?
                            SQLite.m_strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSCasesTableName} WHERE RunTitle = '{strCurRunTitle}'";
                            long lngCasesCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCasesTableName);
                            if (lngCasesCount > 0)
                            {
                                entryListItem.SubItems[COL_FOUND].Text = "Yes";
                            }
                            else
                            {
                                entryListItem.SubItems[COL_FOUND].ForeColor = System.Drawing.Color.White;
                                entryListItem.SubItems[COL_FOUND].BackColor = System.Drawing.Color.Red;
                                this.m_oLvAlternateColors.m_oRowCollection.Item(this.lstFvsOutput.Items.Count - 1).m_oColumnCollection.Item(COL_FOUND).UpdateColumn = false;
                                entryListItem.SubItems[COL_FOUND].Text = "No";
                            }
                            if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSSummaryTableName))
                            {
                                if (!frmMain.g_bSuppressFVSOutputTableRowCount)
                                {
                                    SQLite.m_strSQL = $@"select count(*) from {Tables.FVS.DefaultFVSCasesTableName} c, {Tables.FVS.DefaultFVSSummaryTableName} s 
                                    WHERE c.CaseID = S.CaseID AND c.RunTitle = '{strCurRunTitle}'";
                                    entryListItem.SubItems[COL_SUMMARYCOUNT].Text = Convert.ToString(Convert.ToUInt32(SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSSummaryTableName)));
                                }
                                else
                                {
                                    SQLite.m_strSQL = $@"SELECT s.tpa FROM {Tables.FVS.DefaultFVSCasesTableName} c, {Tables.FVS.DefaultFVSSummaryTableName} s 
                                    WHERE c.CaseID = S.CaseID AND c.RunTitle = '{strCurRunTitle}' AND TPA > 0 LIMIT 1";
                                    long lngResult = SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSSummaryTableName);
                                    if (lngResult > 0)
                                    {
                                        entryListItem.SubItems[COL_SUMMARYCOUNT].Text = "Y";
                                    }
                                    else
                                    {
                                        entryListItem.SubItems[COL_SUMMARYCOUNT].Text = "N";
                                    }
                                }
                            }
                            else
                            {
                                entryListItem.SubItems[COL_SUMMARYCOUNT].Text = "0";
                            }

                            // SET APPEND FLAG
                            string strUpdateStatus = "";
                            SQLite.m_strSQL = $@"select count(*) from {Tables.FVS.DefaultFVSCasesTableName} WHERE {m_colBioSumAppend} ='N' 
                            AND RUNTITLE = '{strCurRunTitle}'";
                            if (Convert.ToUInt32(SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCasesTableName)) > 0)
                            {
                                strUpdateStatus = strUpdateStatus + "a";
                            }
                            if (strUpdateStatus.Trim().Length > 0)
                                entryListItem.Text = strUpdateStatus;

                            // FVS_CUTLIST
                            if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCutListTableName))
                            {
                                if (!frmMain.g_bSuppressFVSOutputTableRowCount)
                                {
                                    SQLite.m_strSQL = $@"select count(*) from {Tables.FVS.DefaultFVSCasesTableName} c, {Tables.FVS.DefaultFVSCutListTableName} s 
                                    WHERE c.CaseID = S.CaseID AND c.RunTitle = '{strCurRunTitle}'";
                                    entryListItem.SubItems[COL_CUTCOUNT].Text = Convert.ToString(Convert.ToUInt32(SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCutListTableName)));
                                }
                                else
                                {
                                    SQLite.m_strSQL = $@"SELECT t.HT from {Tables.FVS.DefaultFVSCasesTableName} c, {Tables.FVS.DefaultFVSCutListTableName} t WHERE 
                                    c.CaseID = t.CaseID AND c.RunTitle = '{strCurRunTitle}' AND HT > 0 LIMIT 1";
                                    long lngResult = SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCutListTableName);
                                    if (lngResult > 0)
                                    {
                                        entryListItem.SubItems[COL_CUTCOUNT].Text = "Y";
                                    }
                                    else
                                    {
                                        entryListItem.SubItems[COL_CUTCOUNT].Text = "N";
                                    }
                                }
                            }
                            else
                            {
                                if (!frmMain.g_bSuppressFVSOutputTableRowCount)
                                {
                                    entryListItem.SubItems[COL_CUTCOUNT].Text = "0";
                                }
                                else
                                {
                                    entryListItem.SubItems[COL_CUTCOUNT].Text = "N";
                                }
                            }
                        }
                    }
                }
			}
			catch (Exception e)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_output:loadvalues() \n" + 
					"Err Msg - " + e.Message,
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);

				this.m_intError=-1;
			}
			this.Refresh();

		}

		

		private void txtOutDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void textBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void btnRefresh_Click(object sender, System.EventArgs e)
		{
			loadvalues();
		}

		private void btnChkAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
			{
                if (cmbFilter.Text.Trim().ToUpper() == "ALL" ||
                    cmbFilter.Text.Trim().ToUpper() == lstFvsOutput.Items[x].SubItems[COL_VARIANT].Text.Trim().ToUpper())
				        this.lstFvsOutput.Items[x].Checked=true;
			}
		}

		private void btnClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0;x<=this.lstFvsOutput.Items.Count-1;x++)
			{
                if (cmbFilter.Text.Trim().ToUpper() == "ALL" ||
                    cmbFilter.Text.Trim().ToUpper() == lstFvsOutput.Items[x].SubItems[COL_VARIANT].Text.Trim().ToUpper())
				        this.lstFvsOutput.Items[x].Checked=false;
			}

		}

		private void uc_fvs_output_Resize(object sender, System.EventArgs e)
		{
            uc_fvs_output_Resize();
		}
        public void uc_fvs_output_Resize()
        {
            try
            {
                this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
                this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
                this.btnHelp.Top = this.btnClose.Top;

                this.lstFvsOutput.Left = 5;
                this.lstFvsOutput.Width = this.Width - 10;
                this.pnlFileSizeMonitor.Width = this.lstFvsOutput.Width;
                this.pnlFileSizeMonitor.Left = this.lstFvsOutput.Left;
                this.pnlFileSizeMonitor.Top = this.btnClose.Top - this.pnlFileSizeMonitor.Height - 2;
                this.btnExecute.Top = this.pnlFileSizeMonitor.Top - btnExecute.Height - 2;
                this.cmbStep.Top = btnExecute.Top;
                this.btnChkAll.Top = this.cmbStep.Top - this.btnChkAll.Height - 2;
                this.btnClearAll.Top = btnChkAll.Top;
                this.btnRefresh.Top = btnChkAll.Top;
                this.cmbFilter.Top = btnChkAll.Top;
                this.btnCancel.Top = this.btnChkAll.Top;
                this.btnCancel.Left = this.ClientSize.Width / 2 - (int)(btnCancel.Width * .5);
                this.btnViewLogFile.Top = this.btnChkAll.Top - 3;
                this.btnViewPostLogFile.Top = this.btnViewLogFile.Top + this.btnViewPostLogFile.Height + 1;
                //this.btnViewLogFile.Left = this.lstFvsOutput.ClientSize.Width - (this.lstFvsOutput.Left*2) - this.btnViewLogFile.Width;
                this.btnViewLogFile.Left = this.groupBox1.Width - this.btnViewLogFile.Width - 10;
                this.btnViewPostLogFile.Left = this.btnViewLogFile.Left;
                btnAuditDb.Top = this.btnViewLogFile.Top;
                btnPostAppendAuditDb.Top = btnViewPostLogFile.Top;
                this.btnAuditDb.Left = this.btnViewLogFile.Left - this.btnAuditDb.Width;
                this.btnPostAppendAuditDb.Left = this.btnAuditDb.Left;
                this.lblMsg.Top = btnChkAll.Top - this.lblMsg.Height - 2;
                if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width > uc_filesize_monitor1.Width)
                {
                    for (; ; )
                    {
                        uc_filesize_monitor1.Width = uc_filesize_monitor1.Width + 1;
                        if (uc_filesize_monitor1.lblMaxSize.Left + uc_filesize_monitor1.lblMaxSize.Width < uc_filesize_monitor1.Width)
                            break;

                    }
                }
                if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width > uc_filesize_monitor2.Width)
                {
                    for (; ; )
                    {
                        uc_filesize_monitor2.Width = uc_filesize_monitor2.Width + 1;
                        if (uc_filesize_monitor2.lblMaxSize.Left + uc_filesize_monitor2.lblMaxSize.Width < uc_filesize_monitor2.Width)
                            break;

                    }
                }
                if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width > uc_filesize_monitor3.Width)
                {
                    for (; ; )
                    {
                        uc_filesize_monitor3.Width = uc_filesize_monitor3.Width + 1;
                        if (uc_filesize_monitor3.lblMaxSize.Left + uc_filesize_monitor3.lblMaxSize.Width < uc_filesize_monitor3.Width)
                            break;

                    }
                }






                this.uc_filesize_monitor2.Left = this.uc_filesize_monitor1.Left + uc_filesize_monitor2.Width + 2;
                this.uc_filesize_monitor3.Left = this.uc_filesize_monitor2.Left + uc_filesize_monitor3.Width + 2;
               
                this.lstFvsOutput.Height = this.lblMsg.Top - this.lstFvsOutput.Top - 5;               
                this.txtOutDir.Width = this.Width - this.txtOutDir.Left - 10;               
                this.lblRunStatus.Left = (int)(this.groupBox1.Width / 2) - (int)(this.lblRunStatus.Width / 2);
                
            }
            catch
            {
            }
        }

        private void RunAppend_Start()
        {

            if (this.lstFvsOutput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            // Warning for older projects without FVSOut.db
            if (!System.IO.File.Exists(m_strFvsOutDb))
            {
                MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
                "FVS Output", "2");
            m_frmTherm.Visible = false;
            this.m_frmTherm.lblMsg.Text = "";
            this.m_frmTherm.TopMost = true;
            
            this.cmbStep.Enabled = false;
            this.btnExecute.Enabled = false;
            this.btnChkAll.Enabled = false;
            this.btnClearAll.Enabled = false;
            this.btnRefresh.Enabled = false;
            this.btnClose.Enabled = false;
            this.btnHelp.Enabled = false;
            this.btnCancel.Visible = false;
            this.btnViewLogFile.Enabled = false;
            this.btnViewPostLogFile.Enabled = false;
            this.btnAuditDb.Enabled = false;
            this.btnPostAppendAuditDb.Enabled = false;
            
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);
            m_frmTherm.Show((frmDialog)ParentForm);




            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(RunAppend_Main));
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();

        }
        /// <summary>
        /// Make sure required tables exist and that they have records
        /// </summary>
		private void val_data()
		{
            string strVariant = "";
            
            string strOutDir = (string)frmMain.g_oDelegate.GetControlPropertyValue(txtOutDir,"Text",false).ToString().Trim();
            string strFvsOutDb = System.IO.Path.GetFileName(Tables.FVS.DefaultFVSOutDbFile);
            string strSummaryCount = "";
            string strCutListCount = "";
            string strCurRunTitle = "";
            long lngPotFireBaseYrCount = 0;
            bool bPotFireBaseYear = true;
            string strRxPackage = "";
            DialogResult result;
           
           
            int intCount = 0;
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(this.lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;
            this.m_intError = 0;
            intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
			for (int x=0; x<=intCount-1;x++)
			{
                
                UpdateTherm(m_frmTherm.progressBar1,
                                        x,
                                        intCount);
                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false) == true)
				{
                    oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                    strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_VARIANT, "Text", false);
                    strRxPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_PACKAGE, "Text", false);
                    strSummaryCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_SUMMARYCOUNT, "Text", false).ToString().Trim();
                    strCutListCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_CUTCOUNT, "Text", false).ToString().Trim();
                    //find the package item in the package collection
                    var y = -1;
                    if (this.m_oRxPackageItem_Collection == null)
                    {
                        //load rxpackage properties
                        m_oRxPackageItem_Collection = new RxPackageItem_Collection();
                        this.m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oQueries, this.m_oRxPackageItem_Collection);
                    }
                    for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                    {
                        if (this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strRxPackage.Trim())
                            break;
                    }
                    if (y <= m_oRxPackageItem_Collection.Count - 1)
                    {
                        RxPackageItem tempRxPackageItem = new RxPackageItem();
                        tempRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), tempRxPackageItem);
                        //FVSOUT_NC_P002-002-999-999-002
                        strCurRunTitle = $@"FVSOUT_{strVariant}{tempRxPackageItem.RunTitleSuffix}";
                    }

                    GetPrePostSeqNumConfiguration("FVS_POTFIRE",strRxPackage);
                    //dont need to validate base year if baseyear is not being referenced
                    if (m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN=="Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN=="Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN=="Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN=="Y")
                         bPotFireBaseYear=true;
                    else 
                         bPotFireBaseYear=false;
                    if (bPotFireBaseYear)
                    {
                        string strFvsPotfireTableName = GetPotfireTableName(m_strFvsOutDb, strVariant);
                        if (m_bPotFireBaseYearTableExist)
                        {
                            lngPotFireBaseYrCount = 1;
                        }
                    }

					if (strSummaryCount.Length == 0 ||
						strSummaryCount == "0")
					{
                        // If we're counting and 0 records are found
                        if (frmMain.g_bSuppressFVSOutputTableRowCount == false)
                        {
                            this.m_intError = -1;
                        }
					}
                    else if (strSummaryCount.Equals("N"))
                    {
                        // If we're not counting
                        this.m_intError = -1;
                    }
                    if (m_intError != 0)
                    {
                        MessageBox.Show($@"!!Summary Table In {m_strFvsOutDb} Does Not Exist Or Has 0 Records!!",
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                        break;
                    }

                    if (strCutListCount.Length == 0 ||
                        strCutListCount == "0")
                    {
                        if (frmMain.g_bSuppressFVSOutputTableRowCount == false)
                        {
                            this.m_intError = -1;
                        }
                    }
                    else if (strCutListCount.Equals("N"))
                    {
                        this.m_intError = -1;
                    }
                    if (m_intError != 0)
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm, "Visible", false);
                        string strMessage = $@"No trees in FVS_CUTLIST for {strVariant}{strRxPackage}.";
                            result = MessageBox.Show("!!Warning!!\r\n-----------\r\n" + strMessage + "  Continue Processing?(Y/N)",
                                 "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo,
                                 System.Windows.Forms.MessageBoxIcon.Question);
                            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm, "Visible", true);
                            if (result == DialogResult.No)
                                break;
                            else
                            {
                                m_intError = 0;
                            }
                    }


                    if (lngPotFireBaseYrCount == 0 && bPotFireBaseYear)
                    {
                        MessageBox.Show("!!Potential Fire Base Year Table Does Not Exist or has 0 records!!",
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                        this.m_intError = -1;
                        break;
                    }
                     
				}
			}			
		}

        private void GetPrePostTableLinkItems(DbFileItem_Collection p_oCollection, 
                                       string p_strDbFileName,
                                       string p_strFVSOutTableName,
                                       ref string p_strSeqNumMtxTableLink,
                                       ref string p_strFVSSummarySeqNumMtxTableLink,
                                       ref string p_strFVSOutTableLink)
        {
             int x, y;
            p_strSeqNumMtxTableLink="";
            p_strFVSOutTableLink="";
            for (x = 0; x <= p_oCollection.Count - 1; x++)
            {
                if (p_oCollection.Item(x).DbFileName.Trim().ToUpper() == p_strDbFileName.Trim().ToUpper())
                {
                    for (y = 0; y <= p_oCollection.Item(x).TableLinkCollection.Count - 1; y++)
                    {
                        if (p_oCollection.Item(x).TableLinkCollection.Item(y).FVSOutputSeqNumMatrixTable &&
                            p_oCollection.Item(x).TableLinkCollection.Item(y).FVSOutputTableName=="FVS_SUMMARY")
                        {
                            p_strFVSSummarySeqNumMtxTableLink=p_oCollection.Item(x).TableLinkCollection.Item(y).LinkedTableName;
                        }
                        if (p_oCollection.Item(x).TableLinkCollection.Item(y).FVSOutputSeqNumMatrixTable &&
                            p_oCollection.Item(x).TableLinkCollection.Item(y).FVSOutputTableName!="FVS_SUMMARY")
                        {
                            p_strSeqNumMtxTableLink = p_oCollection.Item(x).TableLinkCollection.Item(y).LinkedTableName;
                        }
                        if (p_oCollection.Item(x).TableLinkCollection.Item(y).FVSOutputTable && 
                            p_oCollection.Item(x).TableLinkCollection.Item(y).FVSOutputTableName==p_strFVSOutTableName)
                        {
                            p_strFVSOutTableLink=p_oCollection.Item(x).TableLinkCollection.Item(y).LinkedTableName;
                        }

                    }
                    break;
                }
            }
        }
        private void GetTableLinkItems(DbFileItem_Collection p_oCollection, string p_strDbFileName,string p_strTableName, ref int p_intDbFileItem, ref int p_intTableLinkItem)
        {
            int x, y;
            
            for (x = 0; x <= p_oCollection.Count - 1; x++)
            {
                if (p_oCollection.Item(x).DbFileName.Trim().ToUpper() == p_strDbFileName.Trim().ToUpper())
                {
                    for (y = 0; y <= p_oCollection.Item(x).TableLinkCollection.Count - 1; y++)
                    {
                        if (p_oCollection.Item(x).TableLinkCollection.Item(y).TableName.Trim().ToUpper() == p_strTableName.Trim().ToUpper())
                        {
                            p_intDbFileItem = x;
                            p_intTableLinkItem = y;
                            return;
                        }
                    }
                }
            }
            p_intDbFileItem = -1;
            p_intTableLinkItem = -1;
        }

        private void RunAppend_UpdatePrePostTable(string p_strTempDb, string p_strPackage, string p_strVariant, string p_strRx1, string p_strRx2, string p_strRx3, string p_strRx4, bool p_bUpdatePreTableWithVariant,
            int p_intListViewItem, ref int p_intError, ref string p_strError, string p_strRunTitle)
        {

            int y, z, zz;

            string strSourceColumnsList = "";
            string[] strSourceColumnsArray = null;
            string strDestColumnsList = "";
            string[] strDestColumnsArray = null;
            System.Data.DataTable oDataTableSchema;
            bool bFound;
            string strRx = "";
            string strCycle = "";

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_UpdatePrePostTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + " Update PREPOST Tables");

            //
            // Create the host database if it doesn't exist
            //
            if (! System.IO.File.Exists(m_strFvsPrePostDb))
            {
                SQLite.CreateDbFile(m_strFvsPrePostDb);
            }

            //
            // Make a working copy of the FVS Prepost database in case something goes wrong
            //
            System.IO.File.Copy(m_strFvsPrePostDb, p_strTempDb, true);

            IList<string> lstTableNames = new List<string>();
            string strFvsOutConn = SQLite.GetConnectionString(m_strFvsOutDb);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strFvsOutConn))
            {
                conn.Open();
                string[] strSourceTableArray = GetValidFVSTables(conn, Tables.FVS.DefaultFVSCasesTempTableName, p_strRunTitle);
                //
                //make sure all the tables and columns exist
                //
                m_intProgressStepTotalCount = strSourceTableArray.Length;
                m_intProgressStepCurrentCount = 0;
                foreach (var strTable in strSourceTableArray)
                {
                    if (Tables.FVS.g_strFVSOutTablesArray.Contains(strTable.ToUpper()))
                    {
                        if (strTable.ToUpper() != Tables.FVS.DefaultFVSCasesTableName &&
                            strTable.ToUpper() != Tables.FVS.DefaultFVSTreeListTableName &&
                            strTable.ToUpper() != Tables.FVS.DefaultFVSCutListTableName &&
                            strTable.ToUpper() != Tables.FVS.DefaultFVSAtrtListTableName)
                        {
                            lstTableNames.Add(strTable);
                        }
                    }
                }
            }

            for (y = 0; y <= lstTableNames.Count() - 1; y++)
            {
                m_intProgressStepCurrentCount++;
                UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepCurrentCount,
                            m_intProgressStepTotalCount);

                if (uc_filesize_monitor1.File.Trim().Length == 0)
                {
                    uc_filesize_monitor1.BeginMonitoringFile(p_strTempDb, 2000000000, "2GB");
                    uc_filesize_monitor1.Information = "The temporary DB file listed above is a copy of the production DB file " + frmMain.g_oUtils.getFileName(m_strFvsPrePostDb);
                }
                //else if (uc_filesize_monitor2.File.Trim().Length == 0)
                //{
                //    uc_filesize_monitor2.BeginMonitoringFile(m_oPrePostDbFileItem_Collection.Item(y).FullPath.Trim(), 2000000000, "2GB");
                //    //uc_filesize_monitor2.Information = "Base year potential fire table for variant " + strVariant;
                //}

                string strConn = SQLite.GetConnectionString(p_strTempDb);
                string strFVSOutTable = lstTableNames[y].ToUpper();
                string strPreTable = $@"PRE_{strFVSOutTable}";
                string strPostTable = $@"POST_{strFVSOutTable}";
                string strFvsOutCustomSeqNumMatrixTable = $@"{strFVSOutTable}_PREPOST_SEQNUM_MATRIX";
                string strPreRunTitle = p_strRunTitle;
                GetPrePostSeqNumConfiguration(strFVSOutTable, p_strPackage);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + " Update PREPOST Tables:" + strFVSOutTable);

                // Because of BaseYr, POTFIRE output has a different name
                if (strFVSOutTable.ToUpper().IndexOf("POTFIRE") > -1)
                {
                    strFVSOutTable = strFVSOutTable + "_TEMP";
                }
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();

                    // Attach FVSOut.db
                    SQLite.m_strSQL = "ATTACH DATABASE '" + m_strFvsOutDb + "' AS FVSOUT";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                    // Attach FVS_AUDITS.db
                    string strAuditDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile;
                    if (!SQLite.DatabaseAttached(conn, strAuditDbFile))
                    {
                        SQLite.m_strSQL = "ATTACH DATABASE '" + strAuditDbFile + "' AS AUDITS";
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                    }

                    if (!SQLite.TableExist(conn, strPreTable))
                    {
                        //create the table
                        var sb = new System.Text.StringBuilder();
                        var strCol = "";
                        sb.Append($@"(biosum_cond_id CHAR(25), rxpackage CHAR(3), rx CHAR(3), rxcycle CHAR(1), fvs_variant CHAR(2),");
                        var strFields = "";
                        oDataTableSchema = SQLite.getTableSchema(conn, "SELECT * FROM " + strFVSOutTable);
                        strSourceColumnsList = SQLite.getFieldNames(conn, "SELECT * FROM " + strFVSOutTable);
                        strSourceColumnsArray = m_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                        SQLite.m_strSQL = "";

                        for (z = 0; z <= oDataTableSchema.Rows.Count - 1; z++)
                        {
                            var colName = oDataTableSchema.Rows[z]["columnname"].ToString().ToUpper();
                            var dataType = oDataTableSchema.Rows[z]["datatype"].ToString().ToUpper();
                            // This maps SPECIESFIA to SPECIES currently (and in the future should map any other column name differences between SQLite and target access DB
                            // Use converted name here. We want SPECIES for the access creation, and SPECIESFIA for the selects.
                            var convertedColName = utils.TranslateColumn(colName);
                            strCol = utils.WrapInBacktick(convertedColName) + " " + utils.DataTypeConvert(dataType, true);
                            if (strFields.Trim().Length == 0)
                            {
                                strFields = strCol;
                            }
                            else
                            {
                                strFields += "," + strCol;
                            }                   
                        }
                        // Add runtitle column
                        strCol = $@"RUNTITLE CHAR(255)";
                        strFields += "," + strCol;
                        sb.Append(strFields + ",");
                        if (sb.Length > 0)
                        {
                            SQLite.m_strSQL = sb.ToString();
                            string strPrimaryKey = $@" CONSTRAINT " + strPreTable + "_pk PRIMARY KEY(biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant))";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "CREATE TABLE " + strPreTable + " " + SQLite.m_strSQL + strPrimaryKey + "\r\n\r\n");
                            SQLite.SqlNonQuery(conn, "CREATE TABLE " + strPreTable + " " + SQLite.m_strSQL + strPrimaryKey);
                            strPrimaryKey = $@" CONSTRAINT " + strPostTable + "_pk PRIMARY KEY(biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant))";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "CREATE TABLE " + strPostTable + " " + SQLite.m_strSQL + strPrimaryKey + "\r\n\r\n");
                            // Create Post table
                            SQLite.SqlNonQuery(conn, "CREATE TABLE " + strPostTable + " " + SQLite.m_strSQL + strPrimaryKey);

                            // Indexes must be unique throughout an SQLite .db; Adding the table names to the index name
                            SQLite.AddIndex(conn, strPreTable, $@"biosumcondididx_{strPreTable}", "biosum_cond_id");
                            SQLite.AddIndex(conn, strPostTable, $@"biosumcondididx_{strPostTable}", "biosum_cond_id");
                    }
                }
                else
                {
                    //see if columns are the same
                    oDataTableSchema = SQLite.getTableSchema(conn, "SELECT * FROM " + strFVSOutTable);
                    strSourceColumnsList = SQLite.getFieldNames(conn, "SELECT * FROM " + strFVSOutTable);
                    strSourceColumnsArray = m_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                        for (int i = 0; i < strSourceColumnsArray.Length -1; i++)
                        {
                            strSourceColumnsArray[i] = utils.TranslateColumn(strSourceColumnsArray[i]);
                        }
                    strDestColumnsList = SQLite.getFieldNames(conn, "SELECT * FROM " + strPreTable);
                    strDestColumnsArray = m_oUtils.ConvertListToArray(strDestColumnsList, ",");
                        for (int i = 0; i < strDestColumnsArray.Length - 1; i++)
                        {
                            strDestColumnsArray[i] = utils.TranslateColumn(strDestColumnsArray[i]);
                        }

                        SQLite.m_strSQL = "";
                    for (z = 0; z <= oDataTableSchema.Rows.Count - 1; z++)
                    {
                        if (oDataTableSchema.Rows[z]["ColumnName"] != System.DBNull.Value)
                        {
                            bFound = false;
                            for (zz = 0; zz <= strDestColumnsArray.Length - 1; zz++)
                            {
                                if (oDataTableSchema.Rows[z]["ColumnName"].ToString().Trim().ToUpper() ==
                                    strDestColumnsArray[zz].Trim().ToUpper())
                                {
                                    bFound = true;
                                    break;
                                }
                            }
                            if (!bFound)
                            {
                                //column not found so let's add it
                                var colName = oDataTableSchema.Rows[z]["columnname"].ToString().ToUpper();
                                var dataType = oDataTableSchema.Rows[z]["datatype"].ToString().ToUpper();
                                    string strSize = "";
                                    if (oDataTableSchema.Rows[z]["ColumnSize"] != null)
                                        strSize = Convert.ToString(oDataTableSchema.Rows[z]["ColumnSize"]);
                                    SQLite.AddColumn(conn, strPreTable, utils.TranslateColumn(colName), utils.DataTypeConvert(dataType, true), strSize);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE: Added column " + utils.TranslateColumn(colName) + " " + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    SQLite.AddColumn(conn, strPostTable, utils.TranslateColumn(colName), utils.DataTypeConvert(dataType, true), strSize);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE: Added column " + utils.TranslateColumn(colName) + " " + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }
                        }
                    }
                        // Make sure runtitle field is there
                        if (!SQLite.ColumnExist(conn, strPreTable,"RUNTITLE"))
                        {
                            SQLite.AddColumn(conn, strPreTable, "RUNTITLE", "CHAR", "255");
                        }
                        if (!SQLite.ColumnExist(conn, strPostTable, "RUNTITLE"))
                        {
                            SQLite.AddColumn(conn, strPostTable, "RUNTITLE", "CHAR", "255");
                        }
                    }

                    // Create temp FVS Out table with subset for this variant package and primary key
                    const string TMPFVSOUT = "tmpFvsOut";
                    const string TMPBASEYROUT = "tmpBaseYrOut";
                    string strQueryPreTable = TMPFVSOUT;
                    string strQueryPostTable = strQueryPreTable;
                    RunCreateTmpFvsOutTable(conn, oDataTableSchema, strFVSOutTable, strPreRunTitle, strQueryPreTable);
                    // Create temp subset of POTFIRE BaseYr data in case we need it
                    if (strFVSOutTable.ToUpper().Equals(m_strPotFireTable))
                    {
                        RunCreateTmpFvsOutTable(conn, oDataTableSchema, strFVSOutTable, $@"FVSOUT_{ p_strVariant}_POTFIRE_BaseYr", 
                            TMPBASEYROUT);
                    }
 
                    oDataTableSchema.Dispose();

                if (SQLite.m_intError == 0)
                {
                    SQLite.m_strSQL = "DELETE FROM " + strPreTable + " " +
                                      "WHERE RXPACKAGE='" + p_strPackage.Trim() + "'" + " AND " +
                                       "FVS_VARIANT='" + p_strVariant.Trim() + "'";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                }

                if (SQLite.m_intError == 0)
                {
                        SQLite.m_strSQL = "DELETE FROM " + strPostTable + " " +
                                          "WHERE RXPACKAGE='" + p_strPackage.Trim() + "'" + " AND " +
                                          "FVS_VARIANT='" + p_strVariant.Trim() + "'";
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                }

                // Format the source and select column lists
                    string strFormattedSelectColumnList = "";
                    string strSourceColumnsFormattedList = "";
                for (zz = 0; zz <= strSourceColumnsArray.Length - 1; zz++)
                {
                    // This maps SPECIESFIA to SPECIES currently (and in the future should map any other column name differences between SQLite and target access DB
                    var convertedColName = utils.TranslateColumn(strSourceColumnsArray[zz].Trim());
                    strFormattedSelectColumnList = strFormattedSelectColumnList + "a." + convertedColName + ",";
                    strSourceColumnsFormattedList = strSourceColumnsFormattedList + $@"{convertedColName},";
                }
                strFormattedSelectColumnList = strFormattedSelectColumnList.Substring(0, strFormattedSelectColumnList.Length - 1);
                strSourceColumnsFormattedList = strSourceColumnsFormattedList.Substring(0, strSourceColumnsFormattedList.Length - 1);

                    //
                    //INSERT THE RECORDS BY CYCLE
                    //
                    for (z = 1; z <= 4; z++)
                    {
                        strCycle = z.ToString().Trim();
                        switch (strCycle)
                        {
                            case "1":
                                strRx = p_strRx1;
                                if (strFVSOutTable.ToUpper().Equals(m_strPotFireTable) &&
                                    m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN == "Y")
                                {
                                    strQueryPreTable = TMPBASEYROUT;
                                }
                                break;
                            case "2":
                                strRx = p_strRx2;
                                if (strFVSOutTable.ToUpper().Equals(m_strPotFireTable) &&
                                    m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN.Equals("Y"))
                                {
                                    strQueryPreTable = TMPBASEYROUT;
                                }
                                else
                                {
                                    strQueryPreTable = TMPFVSOUT;
                                }
                                break;
                            case "3":
                                strRx = p_strRx3;
                                if (strFVSOutTable.ToUpper().Equals(m_strPotFireTable) &&
                                    m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN.Equals("Y"))
                                {
                                    strQueryPreTable = TMPBASEYROUT;
                                }
                                else
                                {
                                    strQueryPreTable = TMPFVSOUT;
                                }
                                break;
                            case "4":
                                strRx = p_strRx4;
                                if (strFVSOutTable.ToUpper().Equals(m_strPotFireTable) &&
                                    m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN.Equals("Y"))
                                {
                                    strQueryPreTable = TMPBASEYROUT;
                                }
                                else
                                {
                                    strQueryPreTable = TMPFVSOUT;
                                }
                                break;
                        }


                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Rx:" + strRx.Trim() + " Cycle:" + strCycle + ": Get Pre And Post Treatment Years");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                        if (strFVSOutTable != "FVS_STRCLASS")
                        {
                            if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "Y")
                            {
                                SQLite.m_strSQL = "INSERT INTO " + strPreTable + " " +
                                    "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsFormattedList + ") " +
                                    "SELECT '" + p_strPackage + "' AS rxpackage," +
                                           "'" + strRx + "' AS rx," +
                                           "'" + strCycle + "' AS rxcycle," +
                                           "'" + p_strVariant + "' AS fvs_variant," +
                                          strFormattedSelectColumnList + " " +
                                   "FROM " + strQueryPreTable + " a," +
                                      "(SELECT standid,year " +
                                       "FROM " + m_strFVSSummaryAuditPrePostSeqNumTable + " " +
                                       "WHERE CYCLE" + strCycle + "_PRE_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  AS b " +
                                "WHERE a.standid=b.standid AND a.year=b.year ";
                            }
                            else
                            {
                                SQLite.m_strSQL = "INSERT INTO " + strPreTable + " " +
                                   "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsFormattedList + ") " +
                                   "SELECT '" + p_strPackage + "' AS rxpackage," +
                                          "'" + strRx + "' AS rx," +
                                          "'" + strCycle + "' AS rxcycle," +
                                          "'" + p_strVariant + "' AS fvs_variant," +
                                         strFormattedSelectColumnList + " " +
                                  "FROM " + strQueryPreTable + " a," +
                                     "(SELECT standid,year " +
                                      "FROM " + strFvsOutCustomSeqNumMatrixTable + " " +
                                      "WHERE CYCLE" + strCycle + "_PRE_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  AS b " +
                                    "WHERE a.standid=b.standid AND a.year=b.year";
                            }
                        }
                        else
                        {
                            SQLite.m_strSQL = "INSERT INTO " + strPreTable + " " +
                               "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsFormattedList + ") " +
                               "SELECT '" + p_strPackage + "' AS rxpackage," +
                                      "'" + strRx + "' AS rx," +
                                      "'" + strCycle + "' AS rxcycle," +
                                      "'" + p_strVariant + "' AS fvs_variant," +
                                     strFormattedSelectColumnList + " " +
                              "FROM " + strQueryPreTable + " a," +
                                 "(SELECT standid,year,removal_code " +
                                  "FROM " + strFvsOutCustomSeqNumMatrixTable + " " +
                                  "WHERE CYCLE" + strCycle + "_PRE_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  AS b " +
                              "WHERE a.standid=b.standid AND a.year=b.year AND a.removal_code=b.removal_code";
                           }
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    if (SQLite.m_intError == 0)
                    {
                        if (strFVSOutTable != "FVS_STRCLASS")
                        {
                                if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "Y")
                                {
                                    SQLite.m_strSQL = "INSERT INTO " + strPostTable + " " +
                                                "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsFormattedList + ") " +
                                                "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                "'" + strRx + "' AS rx," +
                                                "'" + strCycle + "' AS rxcycle," +
                                                "'" + p_strVariant + "' AS fvs_variant," +
                                                strFormattedSelectColumnList + " " +
                                               "FROM " + strQueryPostTable + " a," +
                                                  "(SELECT standid,year " +
                                                   "FROM " + m_strFVSSummaryAuditPrePostSeqNumTable + " " +
                                                   "WHERE CYCLE" + strCycle + "_POST_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  AS b " +
                                                "WHERE a.standid=b.standid AND a.year=b.year";
                                }
                                else
                                {
                                    SQLite.m_strSQL = "INSERT INTO " + strPostTable + " " +
                                       "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsFormattedList + ") " +
                                       "SELECT '" + p_strPackage + "' AS rxpackage," +
                                              "'" + strRx + "' AS rx," +
                                              "'" + strCycle + "' AS rxcycle," +
                                              "'" + p_strVariant + "' AS fvs_variant," +
                                             strFormattedSelectColumnList + " " +
                                      "FROM " + strQueryPostTable + " a," +
                                         "(SELECT standid,year " +
                                          "FROM " + strFvsOutCustomSeqNumMatrixTable + " " +
                                          "WHERE CYCLE" + strCycle + "_POST_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  AS b " +
                                          "WHERE a.standid=b.standid AND a.year=b.year";                                 }
                            }
                            else
                            {
                                SQLite.m_strSQL = "INSERT INTO " + strPostTable + " " +
                                              "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsFormattedList + ") " +
                                              "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                     "'" + strRx + "' AS rx," +
                                                     "'" + strCycle + "' AS rxcycle," +
                                                     "'" + p_strVariant + "' AS fvs_variant," +
                                                    strFormattedSelectColumnList + " " +
                                             "FROM " + strQueryPostTable + " a," +
                                                "(SELECT standid,year,removal_code " +
                                                 "FROM " + strFvsOutCustomSeqNumMatrixTable + " " +
                                                 "WHERE CYCLE" + strCycle + "_POST_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  AS b " +
                                              "WHERE a.standid=b.standid AND a.year=b.year AND a.removal_code=b.removal_code";
                            }
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        }

                    if (SQLite.m_intError == 0)
                    {
                        //update biosum_cond_id column and runtitle
                        SQLite.m_strSQL = $@"UPDATE {strPreTable} SET biosum_cond_id = 
                            CASE WHEN (biosum_cond_id IS NULL OR LENGTH(TRIM(biosum_cond_id))=0) AND (standid IS NOT NULL AND LENGTH(TRIM(standid)) = 25) 
                            THEN SUBSTR(standid,1,25) ELSE '' END, RUNTITLE = '{p_strRunTitle}' 
                            WHERE RXPACKAGE='{p_strPackage.Trim()}' AND RX='{strRx}' AND RXCYCLE='{strCycle}' AND FVS_VARIANT='{p_strVariant.Trim()}'";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    }


                    if (SQLite.m_intError == 0)
                    {
                            //update biosum_cond_id column and runtitle
                            SQLite.m_strSQL = $@"UPDATE {strPostTable} SET biosum_cond_id = 
                                CASE WHEN (biosum_cond_id IS NULL OR LENGTH(TRIM(biosum_cond_id))=0) AND (standid IS NOT NULL AND LENGTH(TRIM(standid)) = 25) 
                                THEN SUBSTR(standid,1,25) ELSE '' END, RUNTITLE = '{p_strRunTitle}' 
                                WHERE RXPACKAGE='{p_strPackage.Trim()}' AND RX='{strRx}' AND RXCYCLE='{strCycle}' AND FVS_VARIANT='{p_strVariant.Trim()}'";

                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    }
                    if (uc_filesize_monitor3.File.Length > 0) uc_filesize_monitor3.EndMonitoringFile();
                        else uc_filesize_monitor2.EndMonitoringFile();

                    if (SQLite.m_intError != 0) break;

                }
                if (SQLite.m_intError != 0) break;
                //Clean up temporary table(S)
                    if (SQLite.TableExist(conn, TMPFVSOUT))
                    {
                        SQLite.SqlNonQuery(conn, $@"DROP TABLE {TMPFVSOUT}");
                    }
                    if (SQLite.TableExist(conn, TMPBASEYROUT))
                    {
                        SQLite.SqlNonQuery(conn, $@"DROP TABLE {TMPBASEYROUT}");
                    }
                } // End using
        }  
            p_intError = SQLite.m_intError;
            p_strError = SQLite.m_strError;

            //copy the work db file over the production file
            if (p_intError == 0)
            {
                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                    this.WriteText(m_strDebugFile, "\r\nSTART:Copy work file to production file: Source File Name:" + p_strTempDb + " Destination File Name:" + m_strFvsPrePostDb + " " + System.DateTime.Now.ToString() + "\r\n");
                System.IO.File.Copy(p_strTempDb, m_strFvsPrePostDb, true);
                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                    this.WriteText(m_strDebugFile, "\r\nEND:Copy work file to production file: Source File Name:" + p_strTempDb + " Destination File Name:" + m_strFvsPrePostDb + " " + System.DateTime.Now.ToString() + "\r\n");

            }
        }

        private void RunCreateTmpFvsOutTable(System.Data.SQLite.SQLiteConnection conn, DataTable oDataTableSchema,
            string strFVSOutTable, string strRunTitle, string strTmpSubset)
        {
            int z = 0;
            // Create temp FVS Out table with subset for this variant package and primary key
            if (SQLite.TableExist(conn, strTmpSubset))
            {
                SQLite.SqlNonQuery(conn, "DROP TABLE " + strTmpSubset);
            }
            IList<string> lstSpecialColumns = new List<string>() { "CaseId".ToUpper(), "StandId".ToUpper(), "Year".ToUpper() };
            System.Text.StringBuilder sbCreate = new System.Text.StringBuilder();
            sbCreate.Append($@"CREATE TABLE {strTmpSubset} ");
            // t
            sbCreate.Append("( CaseID CHAR(36), StandId CHAR(25), Year INTEGER, ");
            if (strFVSOutTable.Trim().ToUpper().Equals("FVS_STRCLASS"))
            {
                sbCreate.Append("Removal_Code INTEGER, ");
                lstSpecialColumns.Add("Removal_Code".ToUpper());
            }
            for (z = 0; z <= oDataTableSchema.Rows.Count - 1; z++)
            {
                if (oDataTableSchema.Rows[z]["ColumnName"] != System.DBNull.Value)
                {
                    string strColumnName = Convert.ToString(oDataTableSchema.Rows[z]["ColumnName"]);
                    if (!lstSpecialColumns.Contains(strColumnName.ToUpper()))
                    {
                        var dataType = oDataTableSchema.Rows[z]["datatype"].ToString().ToUpper();
                        sbCreate.Append($@"{strColumnName} {utils.DataTypeConvert(dataType, true)},");
                    }
                }
            }
            // Create table with primary key for better performance
            if (strFVSOutTable.Trim().ToUpper().Equals("FVS_STRCLASS"))
            {
                sbCreate.Append(" PRIMARY KEY (CaseId,Year,Removal_Code))");
            }
            else
            {
                sbCreate.Append(" PRIMARY KEY (CaseId,Year))");
            }
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + sbCreate.ToString() + "\r\n");
            SQLite.SqlNonQuery(conn, sbCreate.ToString());
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
            if (SQLite.m_intError == 0)
            {
                SQLite.m_strSQL = $@"INSERT INTO {strTmpSubset} SELECT F.* FROM {strFVSOutTable} F, {Tables.FVS.DefaultFVSCasesTempTableName} C
                        WHERE F.CaseID = c.CaseID and c.RunTitle = '{strRunTitle}'";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
            }
        }
        private void RunAppend_UpdateFVSTreeTables(string p_strPackage,
                                          string p_strVariant,
                                          string p_strRx1,
                                          string p_strRx2,
                                          string p_strRx3,
                                          string p_strRx4,
                                          string strTreeTempDbFile,
                                          ref int p_intError,
                                          ref string p_strError)
        {
            int x;
            string strRx = "";
            string strCycle = "";
            string strFVSOutTableLink = Tables.FVS.DefaultFVSCutListTableName;
            bool bIdColumnExist = false;

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_UpdateFVSTreeTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr oDataMgr = new DataMgr();

            //
            //make sure all the tables and columns exist
            //

            string strFvsTreeTable = Tables.FVS.DefaultFVSCutTreeTableName;
            m_intProgressStepCurrentCount = 0;
            if (m_strRxCycleArray == null)
                m_strRxCycleArray = new string[1];
            m_intProgressStepTotalCount = 8 + (m_strRxCycleArray.Length * 4);

            /**************************************************************
             **delete records in the fvs_tree table that have the current
             **package
             *************************************************************/
             frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + ": Deleting Old Package+Variant " + strFvsTreeTable + " Table Records");
             frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Refresh");

             using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTreeTempDbFile)))
             {
                conn.Open();
                //
                //delete variant/package from fvs out tree table
                //
                oDataMgr.m_strSQL = "DELETE FROM " + strFvsTreeTable  +
                    " WHERE RXPACKAGE='" + p_strPackage.Trim() + "' AND FVS_VARIANT = '" + p_strVariant + "'";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
             }

             m_intProgressStepCurrentCount++;
             UpdateTherm(m_frmTherm.progressBar1,
                 m_intProgressStepCurrentCount,
                 m_intProgressStepTotalCount);

             if (p_intError == 0)
             {
                // Keep track of whether there are trees to calculate metrics
                bool bRunFcs = false;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTreeTempDbFile)))
                {
                    conn.Open();
                    //
                    //ATTACH TO AUDITS.DB for access to FVS_SUMMARY_PREPOST_SEQNUM_MATRIX
                    // Using FVS_SUMMARY_PREPOST_SEQNUM_MATRIX instead of FVS_CUTLIST_PREPOST_SEQNUM_MATRIX
                    // because not all rxs generate a cutlist. Sequence numbers should be the same for CUTLIST and SUMMARY
                    //
                    // This requires the existence of the FVS_SUMMARY_PREPOST_SEQNUM_MATRIX in the AUDITS.DB table which is currently
                    // Always the case. If this method is run without appending the prepost tables, it should be updated to build the sequence
                    // number tables in the temp database.
                    string strSeqNumMatrix = "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX";
                    string strAuditsFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile;
                    if (!oDataMgr.DatabaseAttached(conn, strAuditsFile))
                    {
                        oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strAuditsFile}' AS AUDITS";
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                    }
                    if (!oDataMgr.AttachedTableExist(conn, strSeqNumMatrix))
                    {
                        // Check for existence of seqnum_matrix
                        if (m_bDebug)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, $@"Append aborted due to missing {strSeqNumMatrix} table from {Tables.FVS.DefaultFVSAuditsDbFile}");
                            return;
                        }
                    }

                    if (oDataMgr.TableExist(conn, "cutlist_rowid_work_table"))
                        oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_rowid_work_table");

                    // Attach FVSOut.db
                    oDataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                        Tables.FVS.DefaultFVSOutDbFile + "' AS FVSOUT";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                                    
                    //check fvs version
                    bIdColumnExist = oDataMgr.ColumnExist(conn, "FVSOUT." + strFVSOutTableLink, "ID");
                    //create id column
                    if (!bIdColumnExist && (int)oDataMgr.getRecordCount(conn, "SELECT COUNT(*) FROM (SELECT t.standid FROM FVSOUT." + strFVSOutTableLink + " t LIMIT 1)", "FVSOUT." + strFVSOutTableLink) > 0)
                    {
                        //create structure; Includes autonumber column
                        frmMain.g_oTables.m_oFvs.CreateTreeListWorkTable(oDataMgr, conn, "cutlist_rowid_work_table");
                        //append all the cutlist records into the work table
                        oDataMgr.m_strSQL = "INSERT INTO cutlist_rowid_work_table SELECT null,caseid,standid,year,treeid,treeindex FROM FVSOUT." + strFVSOutTableLink;
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    }

                    //
                    //loop through all the rx cycles for this package and append them to the fvs tree
                    //
                    var runTitle = $@"FVSOUT_{p_strVariant}_P{p_strPackage}-{p_strRx1}-{p_strRx2}-{p_strRx3}-{p_strRx4}";
                    for (x = 0; x <= this.m_strRxCycleArray.Length - 1; x++)
                    {
                        m_intProgressStepCurrentCount++;
                        UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepCurrentCount,
                            m_intProgressStepTotalCount);

                        if (m_strRxCycleArray[x] == null ||
                            m_strRxCycleArray[x].Trim().Length == 0)
                        {
                        }
                        else
                        {
                            strCycle = m_strRxCycleArray[x].Trim();
                            switch (strCycle)
                            {
                                case "1":
                                    strRx = p_strRx1;
                                    break;
                                case "2":
                                    strRx = p_strRx2;
                                    break;
                                case "3":
                                    strRx = p_strRx3;
                                    break;
                                case "4":
                                    strRx = p_strRx4;
                                    break;
                        }

                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Rx:" + strRx.Trim() + " Cycle:" + strCycle + ": Insert " + strFvsTreeTable + " Tree Records");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                        if (oDataMgr.TableExist(conn, "cutlist_fia_trees_work_table"))
                            oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_fia_trees_work_table");

                        if (oDataMgr.TableExist(conn, "cutlist_fvs_created_seedlings_work_table"))
                            oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_fvs_created_seedlings_work_table");

                        //
                        //FIA TREES
                        //
                        //make sure there are records to insert
                        oDataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                            "(SELECT c.standid " +
                            "FROM FVSOUT." + Tables.FVS.DefaultFVSCasesTempTableName + " c, FVSOUT." + strFVSOutTableLink + " t " +
                            "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + runTitle + "' AND " +
                            "substr(t.treeid, 1, 2) NOT IN ('ES') LIMIT 1)";

                        if ((int)oDataMgr.getRecordCount(conn, oDataMgr.m_strSQL, "temp") > 0)
                        {
                            if (bRunFcs == false)
                            {
                                bRunFcs = true;
                            }
                            oDataMgr.m_strSQL = "CREATE TABLE cutlist_fia_trees_work_table AS " +
                                "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strPackage.Trim() + "' AS rxpackage," +
                                 "'" + strRx.Trim() + "' AS rx,'" + strCycle.Trim() + "' AS rxcycle," +
                                 "cast(t.year as text) as rxyear,'" +
                                 p_strVariant + "' AS fvs_variant, " +
                                 "Trim(t.treeid) AS fvs_tree_id," +
                                 "t.SpeciesFia AS fvs_species, t.TPA, ROUND(t.DBH,1) AS DBH , t.Ht, t.pctcr," +
                                 "t.treeval, t.mortpa, t.mdefect, t.bapctile, t.dg, t.htg, " +
                                 "'N' AS FvsCreatedTree_YN," +
                                 "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                 "FROM FVSOUT." + Tables.FVS.DefaultFVSCasesTempTableName + " c, FVSOUT." + strFVSOutTableLink + " t " +
                                 "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + runTitle + "' AND SUBSTR(t.treeid, 1, 2) NOT IN ('ES') ";

                                 if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                                 oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                 if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                 //insert into fvs tree table
                                 oDataMgr.m_strSQL = "INSERT INTO " + strFvsTreeTable + " " +
                                    "(biosum_cond_id, rxpackage,rx,rxcycle,rxyear,fvs_variant, fvs_tree_id," +
                                    "fvs_species, tpa, dbh, ht, pctcr," +
                                    "treeval, mortpa, mdefect, bapctile, dg, htg," +
                                    "FvsCreatedTree_YN,DateTimeCreated) " +
                                    "SELECT a.biosum_cond_id, a.rxpackage,a.rx,a.rxcycle,a.rxyear,a.fvs_variant," +
                                    "a.fvs_tree_id, a.fvs_species, a.tpa, a.dbh, a.ht, a.pctcr," +
                                    "a.treeval, a.mortpa, a.mdefect, a.bapctile, a.dg, a.htg," +
                                    "a.FvsCreatedTree_YN,a.DateTimeCreated  " +
                                    "FROM cutlist_fia_trees_work_table a," +
                                    "(SELECT standid, year FROM " + strSeqNumMatrix + 
                                    " WHERE CYCLE" + strCycle + "_PRE_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  b " +
                                    "WHERE TRIM(a.biosum_cond_id)=TRIM(b.standid) AND cast(a.rxyear as integer)=b.year";

                                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                            }
                                            p_intError = oDataMgr.m_intError;

                                            m_intProgressStepCurrentCount++;
                                            UpdateTherm(m_frmTherm.progressBar1,
                                                       m_intProgressStepCurrentCount,
                                                       m_intProgressStepTotalCount);

                                            //
                                            //FVS CREATED SEEDLING TREES
                                            //
                                            //make sure there are records to insert
                                            oDataMgr.m_strSQL = "SELECT COUNT(*) FROM " +
                                                "(SELECT c.standid " +
                                                "FROM FVSOUT." + Tables.FVS.DefaultFVSCasesTempTableName + " c, FVSOUT." + strFVSOutTableLink + " t " +
                                                "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + runTitle + "' AND " +
                                                "substr(t.treeid, 1, 2) = 'ES' LIMIT 1)";

                                            if ((int)oDataMgr.getRecordCount(conn, oDataMgr.m_strSQL, "temp") > 0)
                                            {
                                                if (bRunFcs == false)
                                                {
                                                    bRunFcs = true;
                                                }
                                                if (bIdColumnExist)
                                                {
                                                    //FVS CREATED SEEDLING TREES
                                                    oDataMgr.m_strSQL =
                                                       "CREATE TABLE cutlist_fvs_created_seedlings_work_table AS " +
                                                       "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strPackage.Trim() + "' AS rxpackage," +
                                                       "'" + strRx.Trim() + "' AS rx,'" + strCycle.Trim() + "' AS rxcycle," +
                                                        "cast(t.year as text) as rxyear,'" +
                                                       p_strVariant + "' AS fvs_variant, " +
                                                       "Trim(t.treeid) AS fvs_tree_id, " +
                                                       "t.SpeciesFia AS fvs_species, t.TPA, ROUND(t.DBH,1) AS dbh , t.Ht, t.pctcr, " +
                                                       "t.treeval, t.mortpa, t.mdefect, t.bapctile, t.dg, t.htg, " +
                                                       "CASE WHEN t.dbh < 1.0 AND t.TPA > 0 THEN 1 ELSE null END AS STATUSCD, " +  
                                                       "'Y' AS FvsCreatedTree_YN," +
                                                       "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                                       "FROM " + Tables.FVS.DefaultFVSCasesTempTableName + " c," + strFVSOutTableLink + " t " +
                                                       "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + runTitle + "' AND substr(t.treeid, 1, 2) = 'ES'";
                                                }
                                                else
                                                {
                                                    //FVS CREATED SEEDLING TREES 
                                                    oDataMgr.m_strSQL =
                                                       "CREATE TABLE cutlist_fvs_created_seedlings_work_table AS " +
                                                       "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strPackage.Trim() + "' AS rxpackage," +
                                                       "'" + strRx.Trim() + "' AS rx,'" + strCycle.Trim() + "' AS rxcycle," +
                                                       "cast(t.year as text) as rxyear,'" +
                                                       p_strVariant + "' AS fvs_variant, " +
                                                       "Trim(t.treeid) AS fvs_tree_id, " +
                                                       "t.SpeciesFia AS fvs_species, t.TPA, ROUND(t.DBH,1) AS dbh , t.Ht, t.pctcr, " +
                                                       "t.treeval, t.mortpa, t.mdefect, t.bapctile, t.dg, t.htg, " +
                                                       "CASE WHEN t.dbh < 1.0 AND t.TPA > 0 THEN 1 ELSE null END AS STATUSCD, " +  
                                                       "'Y' AS FvsCreatedTree_YN," +
                                                       "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                                       "FROM FVSOUT." + Tables.FVS.DefaultFVSCasesTempTableName + " c," + strFVSOutTableLink + " t " +
                                                       "WHERE c.CaseID = t.CaseID AND substr(t.treeid, 1, 2) = 'ES' AND " +
                                                       "c.RunTitle = '" + runTitle + "'";
                                                }

                                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                                oDataMgr.m_strSQL = "INSERT INTO " + strFvsTreeTable + " " +
                                                                     "(biosum_cond_id, rxpackage,rx,rxcycle,rxyear,fvs_variant, fvs_tree_id," +
                                                                      "fvs_species, tpa, dbh, ht, pctcr, " +
                                                                      "treeval, mortpa, mdefect, bapctile, dg, htg, " +
                                                                      "statuscd, FvsCreatedTree_YN,DateTimeCreated) " +
                                                                        "SELECT a.biosum_cond_id, a.rxpackage,a.rx,a.rxcycle,a.rxyear,a.fvs_variant," +
                                                                               "a.fvs_tree_id, a.fvs_species, a.tpa, a.dbh, a.ht, a.pctcr," +
                                                                               "a.treeval, a.mortpa, a.mdefect, a.bapctile, a.dg, a.htg, " +
                                                                               "a.statuscd, a.FvsCreatedTree_YN,a.DateTimeCreated  " +
                                                                        "FROM cutlist_fvs_created_seedlings_work_table a," +
                                                                        "(SELECT standid, year FROM " + strSeqNumMatrix +
                                                                        " WHERE CYCLE" + strCycle + "_PRE_YN='Y' AND FVS_VARIANT = '" + p_strVariant + "' AND RXPACKAGE = '" + p_strPackage + "')  b " +
                                                                        "WHERE TRIM(a.biosum_cond_id)=TRIM(b.standid) AND cast(a.rxyear as integer)=b.year";
                                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                            }

                                            if (oDataMgr.TableExist(conn, "cutlist_fia_trees_work_table"))
                                                oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_fia_trees_work_table");

                                            if (oDataMgr.TableExist(conn, "cutlist_fvs_created_seedlings_work_table"))
                                                oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_fvs_created_seedlings_work_table");

                                            if (oDataMgr.TableExist(conn, "cutlist_save_tree_species_work_table"))
                                                oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_save_tree_species_work_table");

                                            p_intError = oDataMgr.m_intError;
                                        }
                                        m_intProgressStepCurrentCount++;
                                        UpdateTherm(m_frmTherm.progressBar1,
                                                   m_intProgressStepCurrentCount,
                                                   m_intProgressStepTotalCount);

                                    }
                                    //
                                    //SWAP FVS-ONLY SPECIES CODES TO FIA SPECIES CODES IN ORDER FOR FIA TO CALCULATE VOLUMES
                                    //check if original spcd column exists
                                    //save the original spcd
                                    this.SaveTreeSpeciesWorkTable(conn, strFvsTreeTable, p_strPackage);

                                //
                                //UPDATE VOLUME COLUMNS
                                //
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1, m_intProgressStepCurrentCount, m_intProgressStepTotalCount);

                                //update growth projected trees with tree volumes
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Prepare " + strFvsTreeTable + " data for volume calculation");
                                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                                if (oDataMgr.TableExist(conn, Tables.VolumeAndBiomass.BiosumVolumesInputTable))
                                        oDataMgr.SqlNonQuery(conn, "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                frmMain.g_oTables.m_oFvs.CreateSQLiteInputBiosumVolumesTable(oDataMgr, conn, Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step1(
                                                   Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                   strFvsTreeTable, p_strPackage, p_strVariant);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");

                                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    oDataMgr.m_strSQL = 
                                        Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step1a(Tables.VolumeAndBiomass.BiosumVolumesInputTable, runTitle);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");

                                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                } // Closing SQLite connection in preparation to interact with Access tables                               

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);
                if (bRunFcs == true)
                {
                    System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTreeTempDbFile));

                    oConn.Open();
                    // Attach master.db
                    oDataMgr.m_strSQL = $@"attach '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree)}' as master";                    
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    // Attach biosum_ref.db
                    oDataMgr.m_strSQL = $@"attach '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.FiaTreeSpeciesReference)}' as ref";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    //join plot, cond, and tree table to oracle input tree volumes table.
                    //NOTE: this query handles existing FIADB trees that have been grown forward.
                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2(
                        Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                        m_oQueries.m_oFIAPlot.m_strTreeTable,
                        m_oQueries.m_oFIAPlot.m_strPlotTable,
                        m_oQueries.m_oFIAPlot.m_strCondTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                        m_intProgressStepCurrentCount,
                        m_intProgressStepTotalCount);

                    //Set DIAHTCD for FIADB Cycle<>1 trees to their Cycle=1 DIAHTCD values
                    //Set standing_dead_code from tree table on FIADB trees
                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFiadb(m_oQueries.m_oFIAPlot.m_strTreeTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    //Set DIAHTCD for FVS-Created trees using FIA_TREE_SPECIES_REF.WOODLAND_YN
                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFvs(Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    //Set STATUSCD for seedlings; It is populated from FCS for all other trees
                    oDataMgr.m_strSQL = $@"UPDATE {strFvsTreeTable} as b SET statuscd=t.statuscd
                        FROM {m_oQueries.m_oFIAPlot.m_strTreeTable} t WHERE t.biosum_cond_id=b.biosum_cond_id AND TRIM(t.fvs_tree_id)=b.fvs_tree_id
                        AND rxpackage='{p_strPackage.Trim()}' AND fvs_variant='{p_strVariant.Trim()}' and dbh < 1.0";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    //Update FVSCreatedTrees precipitation=plot.precipitation
                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step2a(
                        Tables.VolumeAndBiomass.BiosumVolumesInputTable, m_oQueries.m_oFIAPlot.m_strPlotTable,
                        m_oQueries.m_oFIAPlot.m_strCondTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    //join cond table to oracle input tree volumes table.
                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step3(
                        Tables.VolumeAndBiomass.BiosumVolumesInputTable, m_oQueries.m_oFIAPlot.m_strCondTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    //populate treeclcd column
                    if (oDataMgr.TableExist(oConn, "CULL_TOTAL_WORK_TABLE"))
                        oDataMgr.SqlNonQuery(oConn, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step4(
                        "cull_total_work_table", Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                        m_intProgressStepCurrentCount,
                        m_intProgressStepTotalCount);

                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                        "cull_total_work_table", Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                        m_intProgressStepCurrentCount,
                        m_intProgressStepTotalCount);

                    oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                        "cull_total_work_table", Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(oConn, oDataMgr.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                        m_intProgressStepCurrentCount,
                        m_intProgressStepTotalCount);

                    // CLOSE CONNECTION
                    if (oConn.State == ConnectionState.Open)
                    {
                        oConn.Close();
                    }

                    if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase) == false)
                    {
                        m_intError = -1;
                        m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase + " not found";
                    }
                    if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR") == false)
                    {
                        m_intError = -1;
                        m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR not found";
                    }
                    if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat") == false)
                    {
                        m_intError = -1;
                        m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat not found";
                    }
                    if (m_intError == 0)
                    {
                        //
                        //RE-CONNECT TO SQLITE AND REMOVE DATA FROM FCS SQLITE DB
                        //
                        oDataMgr.OpenConnection(false, 1, strTreeTempDbFile, "BIOSUM");
                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "ATTACH DATABASE '" + frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase +
                            "' AS FCS");
                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, $"DELETE FROM FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");
                        UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepTotalCount,
                            m_intProgressStepTotalCount);

                        //insert records from table biosum_volumes_input (BiosumVolumesInputTable)
                        //into table fcs_biosum_volumes_input (FcsBiosumVolumesInputTable)
                        oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteBiosumCalcTable_Step7(
                            Tables.VolumeAndBiomass.BiosumVolumesInputTable, $"FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        UpdateTherm(m_frmTherm.progressBar1, 1, 6);

                        // Delete work table
                        if (oDataMgr.TableExist(oDataMgr.m_Connection, "CULL_TOTAL_WORK_TABLE"))
                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                        //
                        //RUN JAVA APP TO CALCULATE VOLUME/BIOMASS
                        //
                        if (m_intError == 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Running volume calculations for " + strFvsTreeTable);
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            frmMain.g_oUtils.RunProcess(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum", "fcs_tree_calc.bat", "BAT");
                            if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt"))
                            {
                                // Read entire text file content in one string  
                                m_strError = System.IO.File.ReadAllText(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt");
                                if (m_strError.IndexOf("JAVA.EXE", 0) > 0)
                                {
                                    m_strError = "Problem detected running JAVA.EXE";
                                    m_intError = -2;
                                }
                            }
                        }
                        UpdateTherm(m_frmTherm.progressBar1, 2, 6);
                        //
                        //UPDATE OUTPUT TABLES WITH CALCULATED VALUES
                        //
                        if (m_intError == 0)
                        {
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n SQLite connection status: " + oDataMgr.m_Connection.State + "\r\n");
                            bool bAttached = oDataMgr.DatabaseAttached(oDataMgr.m_Connection,
                                frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase);
                            if (!bAttached)
                            {
                                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, "ATTACH DATABASE '" + frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase +
                                "' AS FCS");
                            }

                            UpdateTherm(m_frmTherm.progressBar1, 3, 6);

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Write calculation results to " + strFvsTreeTable);
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            //update calculated fields from biosum_calc table
                            //into table fvs_tree
                            oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step9(
                                strFvsTreeTable, Tables.VolumeAndBiomass.BiosumVolumeCalcTable, p_strVariant.Trim(), p_strPackage.Trim());
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step10(
                                strFvsTreeTable, Tables.VolumeAndBiomass.BiosumVolumeCalcTable, p_strVariant.Trim(), p_strPackage.Trim());
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            oDataMgr.m_strSQL = "DETACH DATABASE 'FCS'";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            if (oDataMgr.m_Connection.State == ConnectionState.Open)
                            {
                                oDataMgr.CloseAndDisposeConnection(oDataMgr.m_Connection, true);
                            }                            
                        }

                        UpdateTherm(m_frmTherm.progressBar1, 6, 6);
                                    }
                }

                m_intProgressStepCurrentCount++;
                UpdateTherm(m_frmTherm.progressBar1,
                    m_intProgressStepCurrentCount,
                    m_intProgressStepTotalCount);

                if (oDataMgr.m_intError == 0 && m_intError == 0)
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTreeTempDbFile)))
                    {
                        conn.Open();
                        if (oDataMgr.TableExist(conn, "cutlist_save_tree_species_work_table"))
                        {
                            //update fvs_species
                            oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step11(strFvsTreeTable, p_strPackage);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            oDataMgr.SqlNonQuery(conn, "DROP TABLE cutlist_save_tree_species_work_table");
                        }

                        // DELETE SEEDLINGS FROM CUTTREE TABLE; INFOREST TABLE TBD
                        if (strFvsTreeTable.Equals(Tables.FVS.DefaultFVSCutTreeTableName))
                        {
                            oDataMgr.m_strSQL = $@"DELETE FROM {strFvsTreeTable} WHERE DBH < 1.0" ;
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        }

                        // DELETE WORK TABLES
                        string[] arrDeleteTables = new string[] { Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                            "cutlist_rowid_work_table"};
                             foreach (var dTable in arrDeleteTables)
                             {
                                if (oDataMgr.TableExist(conn, dTable))
                                    {
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + "DROP TABLE " + dTable + "\r\n");
                                        oDataMgr.SqlNonQuery(conn, "DROP TABLE " + dTable);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    }
                             }
                    }
                }
            }
            p_intError = oDataMgr.m_intError;

            if (oDataMgr.m_intError != 0 || m_intError != 0)
            {
                if (oDataMgr.m_intError != 0)
                {
                    p_intError = oDataMgr.m_intError;
                    p_strError = oDataMgr.m_strError;
                }
                else
                {
                    if (m_intError != 0)
                    {
                        p_intError = m_intError;
                        p_strError = m_strError;
                    }
                }
            }
        }
		public void StopThread()
		{
			
			string strMsg="";
            if (m_frmTherm.Text.Trim() == "FVS Output")
                strMsg = "Do you wish to cancel appending and updating fvs out data (Y/N)?";
            else if (m_frmTherm.Text.Trim() == "Ready to create FVS_OutBioSum.db")
                strMsg = "Do you wish to cancel creating the FVSOut_BioSum.db (Y/N)?";
            else
                strMsg = "Do you wish to cancel audit (Y/N)?";

			frmMain.g_oDelegate.AbortProcessing("FIA Biosum",strMsg);

			if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
			{
				this.m_frmTherm.AbortProcess = true;
				frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,this.lstFvsOutput.SelectedItems[0].Index,COL_RUNSTATUS,"BackColor",Color.Red);
				frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsOutput,this.lstFvsOutput.SelectedItems[0].Index,COL_RUNSTATUS,"Text","Cancelled");
				frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,0,"Ready");
				this.ThreadCleanUp();
			}

			
		}
		public void FVSRecordsFinished()
		{
			if (this.m_frmTherm != null)
			{
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
				this.m_frmTherm = null;
			}
		}
		private void ThreadCleanUp()
		{
			try
			{
                //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpboxAppend, "Enabled", true);
                //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpboxAudit, "Enabled", true);
                //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpBoxPostAudit, "Enabled", true);
                //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpboxSpCdConvert, "Enabled", true);
                uc_filesize_monitor1.EndMonitoringFile();
                uc_filesize_monitor2.EndMonitoringFile();
                uc_filesize_monitor3.EndMonitoringFile();
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnExecute, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.ComboBox)cmbStep, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnChkAll, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnClearAll, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnRefresh, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnClose, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnHelp, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnViewLogFile, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnViewPostLogFile, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnAuditDb, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnPostAppendAuditDb, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.UserControl)this, "Enabled", true);
                frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Ready");
				this.ParentForm.Enabled = true;
				if (this.m_frmTherm != null)
				{
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");

					this.m_frmTherm = null;
				}
				this.m_thread = null;
			}
			catch
			{
			}

		}
		private void WriteText(string p_strTextFile,string p_strText)
		{
			System.IO.FileStream oTextFileStream;
			System.IO.StreamWriter oTextStreamWriter;

			if (!System.IO.File.Exists(p_strTextFile))
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Create, 
					System.IO.FileAccess.Write);
			}
			else
			{
				oTextFileStream = new System.IO.FileStream(p_strTextFile, System.IO.FileMode.Append, 
					System.IO.FileAccess.Write);
			}
			
			oTextStreamWriter = new System.IO.StreamWriter(oTextFileStream);
			oTextStreamWriter.Write(p_strText);
			oTextStreamWriter.Close();
			oTextFileStream.Close();
		}

		private void btnAudit_Click(object sender, System.EventArgs e)
		{
            RunPREAudit_Start();
		}

        private void RunPREAudit_Start()
        {
            if (this.lstFvsOutput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            this.DisplayAuditMessage = true;


            this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
                "FVS Output Audit", "2");
            this.m_frmTherm.TopMost = true;
            this.m_frmTherm.lblMsg.Text = "";
            this.cmbStep.Enabled = false;
            this.btnExecute.Enabled = false;
            this.btnChkAll.Enabled = false;
            this.btnClearAll.Enabled = false;
            this.btnRefresh.Enabled = false;
            this.btnClose.Enabled = false;
            this.btnHelp.Enabled = false;
            this.btnViewLogFile.Enabled = false;
            this.btnViewPostLogFile.Enabled = false;
            this.btnAuditDb.Enabled = false;
            this.btnPostAppendAuditDb.Enabled = false;


            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(RunPREAudit_Main));
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;

            frmMain.g_oDelegate.m_oThread.Start();

        }
		public void RunPREAudit_Main()
		{
			
			frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
			this.m_intError=0;
			int intCount=0;
			m_intProgressOverallTotalCount=0;
			m_intProgressStepCurrentCount=0;
			m_strError="";
			m_strWarning="";
			m_intWarning=0;
			m_intProgressOverallCurrentCount=0;

            string strPackage="";
			string strVariant="";
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(this.lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;

			Tables oTables = new Tables();
			

			
			

			string strOutDirAndFile;
			string strDbFile;
            string strAuditDbFile;
           
			string[] strSourceTableArray=null;			
			bool bSkip;
			bool bResult=false;
            bool bDisplay = false;

            if (m_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

			System.DateTime oDate = System.DateTime.Now;
			string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
			m_strLogDate = oDate.ToString(strDateFormat);
			m_strLogDate = m_strLogDate.Replace("/","_"); m_strLogDate=m_strLogDate.Replace(":","_");

            if (m_oFVSPrePostSeqNumItemCollection == null) m_oFVSPrePostSeqNumItemCollection = new FVSPrePostSeqNumItem_Collection();

            // Loads the sequence number configurations from db\fvsmaster.db\fvs_output_prepost_seqnum
            string strDbConn = SQLite.GetConnectionString($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");
            m_oRxTools.LoadFVSOutputPrePostRxCycleSeqNum(strDbConn, m_oFVSPrePostSeqNumItemCollection);

            int x,y;
			
			try
			{
				bSkip=false;
                intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);
                m_intProgressOverallTotalCount = intCount + 1;

                for (x=0;x<=intCount-1;x++)
				{
                    oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                    m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = true;
                    m_oLvAlternateColors.DelegateListViewSubItem(oLvItem,x, COL_RUNSTATUS);
                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLvItem.ListView, x, COL_RUNSTATUS, "Text", "");
					//see if checked
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x,"Checked",false))
                        m_intProgressOverallTotalCount++;
				}
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Maximum",100);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Minimum",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Value",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2,"Text","Overall Progress");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Visible",true);

                bDisplay = this.DisplayAuditMessage;
                this.DisplayAuditMessage = false;
                this.val_data();

                // We'll add the table links a single audit db rather than the fvs out .mdbs as we did previously
                strAuditDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile;
                strDbFile = System.IO.Path.GetFileName(Tables.FVS.DefaultFVSOutDbFile);
                if (!System.IO.File.Exists(strAuditDbFile))
                {
                    SQLite.CreateDbFile(strAuditDbFile);
                }
                if (uc_filesize_monitor1.File.Trim().Length > 0) uc_filesize_monitor1.EndMonitoringFile();
                uc_filesize_monitor1.BeginMonitoringFile(strAuditDbFile, 2000000000, "2GB");
                uc_filesize_monitor1.Information = "BIOSUM DB file containing PREPOST SEQNUM MATRIX and AUDIT tables for processing FVS Output DB file " + strDbFile;

                if (m_intError == 0)
                {
                    DisplayAuditMessage = bDisplay;
                    for (x = 0; x <= intCount - 1; x++)
                    {
                        oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                        this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
                        this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
                        m_intProgressOverallCurrentCount++;
                        UpdateTherm(m_frmTherm.progressBar2,
                            m_intProgressOverallCurrentCount,
                            m_intProgressOverallTotalCount);

                        if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false) == true)
                        {
                            m_bPotFireBaseYearTableExist = true;

                            m_intProgressStepTotalCount = 20;
                            m_intProgressStepCurrentCount = 0;

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);

                            string strItemDialogMsg = "";

                            int intItemError = 0;
                            string strItemError = "";
                            int intItemWarning = 0;
                            string strItemWarning = "";
                            bSkip = true;




                            this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = false;
                            frmMain.g_oDelegate.ExecuteControlMethodWithParam((System.Windows.Forms.Control)oLv, "EnsureVisible", new object[] { x });
                            frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Selected", true);
                            frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Focused", true);



                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkGoldenrod);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit");

                            //get the variant
                            strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_VARIANT, "Text", false);
                            strVariant = strVariant.Trim();

                            //get the package and treatments
                            strPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_PACKAGE, "Text", false);
                            strPackage = strPackage.Trim();

                            //find the package item in the package collection
                            for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                            {
                                if (this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
                                    break;


                            }
                            if (y <= m_oRxPackageItem_Collection.Count - 1)
                            {
                                this.m_oRxPackageItem = new RxPackageItem();
                                m_oRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), m_oRxPackageItem);
                            }
                            else
                            {
                                this.m_oRxPackageItem = null;
                            }


                            //get the list of treatment cycle year fields to reference for this package
                            this.m_strRxCycleList = "";
                            if (this.m_oRxPackageItem != null)
                            {

                            }
                            if (m_oRxPackageItem.SimulationYear1Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear1Rx.Trim() != "000") this.m_strRxCycleList = "1,";
                            if (m_oRxPackageItem.SimulationYear2Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear2Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                            if (m_oRxPackageItem.SimulationYear3Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear3Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                            if (m_oRxPackageItem.SimulationYear4Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear4Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                            if (this.m_strRxCycleList.Trim().Length > 0)
                                this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

                            this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                            strOutDirAndFile = strOutDirAndFile.Trim();
                            strOutDirAndFile = strOutDirAndFile + "\\" + Tables.FVS.DefaultFVSOutDbFile;

                            //strAuditDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                            //strAuditDbFile = strAuditDbFile.Trim();
                            //strAuditDbFile = strAuditDbFile + "\\" + Tables.FVS.DefaultFVSAuditsDbFile;


                            //m_oRxTools.CreateFVSOutputTableLinks(strAuditDbFile, strOutDirAndFile);

                            CreatePotFireTables(strAuditDbFile, m_strFvsOutDb, strVariant, strPackage);

                            //uc_filesize_monitor1.BeginMonitoringFile(
                            //strOutDirAndFile,
                            //2000000000, "2gb");
                            //uc_filesize_monitor1.Information = "FVS output file";


                            //oAdo.OpenConnection(oAdo.getMDBConnString(strAuditDbFile, "", ""));

                            m_strLogFile = $@"{this.txtOutDir.Text.Trim()}\FVSOUT_{strVariant}_P{strPackage}_Audit_{m_strLogDate.Replace(" ", "_")}.txt";
                            frmMain.g_oUtils.WriteText(m_strLogFile, "AUDIT LOG \r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "--------- \r\n\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Database File:" + strDbFile + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Variant:" + strVariant + " \r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Package:" + strPackage + " \r\n\r\n");

                            string strRunTitle = $@"FVSOUT_{strVariant}{m_oRxPackageItem.RunTitleSuffix}";
                            string strBaseYrRunTitle = $@"FVSOUT_{strVariant}_POTFIRE_BaseYr";

                            if (m_bDebug)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: Building FVSOut.db indexes" + System.DateTime.Now.ToString() + "\r\n");
                            m_oRxTools.CreateFvsOutDbIndexes(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile);
                            if (m_bDebug)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "END: Built FVSOut.db indexes" + System.DateTime.Now.ToString() + "\r\n");


                            string strConn = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile);
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                            {
                                conn.Open();
                                strSourceTableArray = GetValidFVSTables(conn, Tables.FVS.DefaultFVSCasesTableName, strRunTitle);

                                //
                                //check for duplicate standid+year records for tables that are 
                                //only supposed to have one stand+year combination represented
                                //
                                if (intItemError == 0)
                                {
                                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Check For duplicate standid+year");
                                    CheckForDuplicateStandIdandYear(conn, strSourceTableArray, strRunTitle, true, ref strItemError, ref intItemError);
                                }
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);
                           


                            if (intItemError==0) 
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Create SeqNum Matrix tables");
                                //CreatePrePostSeqNumMatrixTables(oAdo, oAdo.m_OleDbConnection,strPackage,true);
                                CreatePrePostSeqNumMatrixSqliteTables(strAuditDbFile, strRunTitle, strBaseYrRunTitle, true);
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            //check if summary table exists
                            if (intItemError == 0)
                            {
                                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                                {
                                    conn.Open();
                                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " FVS_SUMMARY");
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_SUMMARY-----\r\n");
                                    if (SQLite.TableExist(conn, "FVS_Summary"))
                                    {
                                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_SUMMARY");

                                        if (intItemError == 0)
                                        {
                                            //
                                            //get fvs_summary configuration
                                            //
                                            GetPrePostSeqNumConfiguration("FVS_SUMMARY", strPackage);

                                            // attach audit.db
                                            if (!SQLite.DatabaseAttached(conn, frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile))
                                            {
                                                SQLite.m_strSQL = $@"ATTACH DATABASE '{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile}' AS AUDIT";
                                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                            }

                                            //check pre and post-treatment seqnum assignments
                                            this.Validate_FvsSummaryPrePostSeqNum(conn, ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true,
                                                strVariant, strPackage);
                                            if (intItemError == 0 && intItemWarning == 0)
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                            }
                                            else if (intItemWarning != 0)
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                strItemWarning = "See Log File";
                                            }
                                        }
                                        if (intItemError != 0)
                                        {
                                            switch (intItemError)
                                            {
                                                case -2:
                                                    strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Summary table has no records\r\n";
                                                    break;
                                                case -3:
                                                    strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Summary table has pre-treatment designated sequence numbers that cannot be found\r\n";
                                                    break;
                                                case -4:
                                                    strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Summary table has post-treatment designated sequence numbers that cannot be found\r\n";
                                                    break;

                                            }

                                        }


                                    }
                                    else
                                    {
                                        intItemError = -1;
                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Summary table missing\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: FVS_Summary table missing\r\n\r\n");
                                    }
                                }
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);
                            if (intItemError == 0)
                            {
                               
                                
                                                                        
                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                                m_intProgressStepCurrentCount,
                                                m_intProgressStepTotalCount);

                                string strConnAudit = SQLite.GetConnectionString(strAuditDbFile);
                                for (y = 0; y <= strSourceTableArray.Length - 1; y++)
                                    {

                                        if (strSourceTableArray[y] == null) break;

                                        bSkip = !RxTools.ValidFVSTable(strSourceTableArray[y]);
                                        if (bSkip == false)
                                        {

                                            GetPrePostSeqNumConfiguration(strSourceTableArray[y].Trim().ToUpper(),strPackage);

                                            if (strSourceTableArray[y].Trim().ToUpper() == "FVS_SUMMARY" ||
                                                strSourceTableArray[y].Trim().ToUpper() == "FVS_CASES")
                                            {
                                            }
                                            else
                                               frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " " + strSourceTableArray[y]);

                                            if (strSourceTableArray[y].Trim().ToUpper() == "FVS_SUMMARY" ||
                                                strSourceTableArray[y].Trim().ToUpper() == "FVS_CASES")
                                            {
                                            }
                                            else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_TREELIST")
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_TREELIST-----\r\n");
                                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_Treelist");



                                                intItemWarning = 0;
                                                strItemWarning = "";

                                               

                                                this.Validate_TreeListTables(strConnAudit, "fvs_treelist", "fvs_summary", ref intItemError, ref strItemError, ref intItemWarning, 
                                                    ref strItemWarning, true, strRunTitle);

                                                if (intItemError == 0 && intItemWarning == 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                                }
                                                else if (intItemWarning != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                }



                                                if (intItemError != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                    if (intItemError == -3)
                                                    {
                                                        strItemError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
                                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
                                                    }
                                                    else if (intItemError == -4)
                                                    {
                                                        strItemError = "There are FVS_Treelist records whose standid and year are not found found in the FVS_Summary table (See Log File)";
                                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Treelist standid and year not found in FVS_Summary table\r\n";

                                                    }
                                                }


                                            }
                                            else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_CUTLIST")
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_CUTLIST-----\r\n");
                                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_Cutlist");

                                                intItemWarning = 0;
                                                strItemWarning = "";

                                                this.Validate_TreeListTables(strConnAudit, "fvs_cutlist", "fvs_summary", ref intItemError, ref strItemError, ref intItemWarning, 
                                                    ref strItemWarning, true, strRunTitle);





                                                if (intItemError != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                    if (intItemError == -3)
                                                    {
                                                        strItemError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
                                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
                                                    }
                                                    else if (intItemError == -4)
                                                    {
                                                        strItemError = "FVS_Cutlist Standid and year not found in the fvs_summary table";
                                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Cutlist standid and year not found in FVS_Summary table\r\n";

                                                    }
                                                }

                                                if (intItemError == 0)
                                                {
                                                this.Validate_FVSTreeId(Tables.FVS.DefaultFVSCasesTableName, Tables.FVS.DefaultFVSCutListTableName, strVariant, strPackage,
                                                        m_oRxPackageItem.SimulationYear1Rx, m_oRxPackageItem.SimulationYear2Rx,
                                                        m_oRxPackageItem.SimulationYear3Rx, m_oRxPackageItem.SimulationYear4Rx, 
                                                        strRunTitle, ref intItemWarning, ref strItemWarning, ref intItemError, ref strItemError);
                                                    if (intItemError != 0)
                                                    {
                                                        if (intItemError == -5)
                                                        {
                                                            frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                            strItemDialogMsg = strItemDialogMsg + strDbFile + " " + strItemError;
                                                        }

                                                    }
                                                   
                                                }

                                                if (intItemError == 0 && intItemWarning == 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                                }
                                                else if (intItemWarning != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                }

                                            }
                                            else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_ATRTLIST")
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_ATRTLIST-----\r\n");
                                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_ATRTList");

                                                intItemWarning = 0;
                                                strItemWarning = "";
                                                this.Validate_TreeListTables(strConnAudit, "fvs_atrtlist", "fvs_summary", ref intItemError, ref strItemError, ref intItemWarning, 
                                                    ref strItemWarning, true, strRunTitle);

                                                if (intItemError == 0 && intItemWarning == 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                                }
                                                else if (intItemWarning != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                }



                                                if (intItemError != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                    if (intItemError == -3)
                                                    {
                                                        strItemError = "FVS_Treelist Minimum treatment year not found in FVS_Summary table (See Log File)";
                                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Treelist Minimum treatment year not found in FVS_Summary table\r\n";
                                                    }
                                                    else if (intItemError == -4)
                                                    {
                                                        strItemError = "FVS_ATRTList Standid and year not found in the fvs_summary table";
                                                        strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_ATRTlist standid and year not found in FVS_Summary table\r\n";

                                                    }
                                                }


                                            }
                                            else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_POTFIRE")
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_POTFIRE-----\r\n");
                                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_Potfire");

                                                intItemWarning = 0;
                                                strItemWarning = "";

                                                
                                                this.Validate_PotFire(strConn, "FVS_POTFIRE", ref intItemError, ref strItemError, 
                                                    ref intItemWarning, ref strItemWarning, true, strRunTitle);

                                                if (intItemError == 0 && intItemWarning == 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                                }
                                                else if (intItemWarning != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                }

                                                if (intItemError != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                }

                                            }
                                            else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_STRCLASS")
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_STRCLASS-----\r\n");
                                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_StrClass");


                                                intItemWarning = 0;
                                                strItemWarning = "";

                                            this.Validate_FVSGenericTableSqlite(strConn, "FVS_STRCLASS", "FVS_STRCLASS", ref intItemError, ref strItemError, ref intItemWarning,
                                                ref strItemWarning, true, strRunTitle);

                                            if (intItemError == 0 && intItemWarning == 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                                }
                                                else if (intItemWarning != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                }

                                                if (intItemError != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                }

                                            }
                                            else
                                            {
                                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----" + strSourceTableArray[y].Trim().ToUpper() + "-----\r\n");
                                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit..." + strSourceTableArray[y].Trim());
                                                intItemWarning = 0;
                                                strItemWarning = "";

                                                this.Validate_FVSGenericTableSqlite(strConn, strSourceTableArray[y].Trim(), strSourceTableArray[y], ref intItemError, ref strItemError, ref intItemWarning, 
                                                    ref strItemWarning, true, strRunTitle);
                                                if (intItemError == 0 && intItemWarning == 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                                }
                                                else if (intItemWarning != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemWarning + "\r\n");
                                                }

                                                if (intItemError != 0)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, strItemError + "\r\n");
                                                }
                                            }
                                            if (intItemError != 0) break;
                                        }

                                    }
                                
                            }

                            // Delete temporary FVS_POTFIRE_TEMP
                            string strAudit = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile);
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strAudit))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, m_strPotFireTable))
                                {
                                    SQLite.m_strSQL = $@"DROP TABLE {m_strPotFireTable}";
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                            }

                            // Delete temporary FVS_CASES_TEMP
                                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTempTableName))
                                {
                                    SQLite.m_strSQL = $@"DROP TABLE {Tables.FVS.DefaultFVSCasesTempTableName}";
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                            }

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            if (intItemError == 0 && intItemWarning == 0)
                            {
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Green);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT: OK");
                            }
                            else if (intItemError != 0)
                            {
                                m_intError = intItemError;
                                if (strItemError.Trim().Length > 50)
                                {
                                    strItemError = strItemError.Substring(0, 45) + "....(See log file)";

                                }
                                if (strItemDialogMsg.Trim().Length > 0)
                                {
                                    m_strError = m_strError + strItemDialogMsg;
                                }
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Red);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT ERROR:" + strItemError.Replace("\r\n", " ").Replace("ERROR:", " "));
                            }
                            else if (intItemWarning != 0)
                            {
                                if (strItemWarning.Trim().Length > 50)
                                {
                                    strItemWarning = strItemWarning.Substring(0, 45) + "....(See log file)";

                                }
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkOrange);
                                if (strItemWarning.Substring(0, 8) == "WARNING:")
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT " + strItemWarning.Replace("\r\n", " ").Replace("WARNING:", " "));
                                else
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT WARNING:" + strItemWarning.Replace("\r\n", " ").Replace("WARNING:", " "));

                            }

                            frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "**EOF**");
                            frmMain.g_oDelegate.ExecuteListViewItemsMethod(oLv, "Refresh");

                            //compact and repair
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Compact and Repair");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            
                            UpdateTherm(m_frmTherm.progressBar2,
                                    m_intProgressOverallCurrentCount,
                                    m_intProgressOverallTotalCount);
                        }
                    }
                }
                UpdateTherm(m_frmTherm.progressBar1,
                                   m_intProgressStepTotalCount,
                                   m_intProgressStepTotalCount);
                UpdateTherm(m_frmTherm.progressBar2,
                                m_intProgressOverallTotalCount,
                                m_intProgressOverallTotalCount);
                
				System.Threading.Thread.Sleep(2000);
				this.FVSRecordsFinished();
			}
			catch (System.Threading.ThreadInterruptedException err)
			{
				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
				if (SQLite.m_Connection != null)
				{
					if (SQLite.m_Connection.State != System.Data.ConnectionState.Closed)
					{
						SQLite.CloseConnection(SQLite.m_Connection);
					}
				}
			    this.ThreadCleanUp();
				this.CleanupThread();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_fvs_output:Audit  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FVS Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}

			if (DisplayAuditMessage)
			{
                if (m_intError == 0) this.m_strError = m_strError + "Passed Audit";
                else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";
                //MessageBox.Show(m_strError,"FIA Biosum");
                FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
                frmTemp.Text = "FIA Biosum";
                frmTemp.AutoScroll = false;
                uc_textboxWithButtons uc_textbox1 = new uc_textboxWithButtons();
                frmTemp.Controls.Add(uc_textbox1);
                uc_textbox1.lblTitle.Text = "Audit Results";
                uc_textbox1.TextValue = m_strError;
                frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                frmTemp.ShowDialog();
			}

            if (m_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "****END*****" + System.DateTime.Now.ToString() + "\r\n");
			CleanupThread();

			frmMain.g_oDelegate.m_oEventThreadStopped.Set();
			this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

			
		
			
			
		}
        /// <summary>
        /// Check for duplicate standid+year values on FVS output tables 
        /// </summary>

        /// <param name="p_strFVSOutTables">tables in the fvsoutput file</param>
        /// <param name="p_strRxPackage">treatment package</param>
        /// <param name="p_bAudit">audit routine</param>
        /// <param name="p_strError"></param>
        /// <param name="p_intError"></param>
        private void CheckForDuplicateStandIdandYear(System.Data.SQLite.SQLiteConnection conn, string[] p_strFVSOutTables,string p_strRunTitle, bool p_bAudit,ref string p_strError, ref int p_intError)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CheckForDuplicateStandIdandYear\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int y;
            bool bSkip;
            p_strError = "";
            for (y = 0; y <= p_strFVSOutTables.Length - 1; y++)
            {

                if (p_strFVSOutTables[y] == null) break;

                bSkip = !RxTools.ValidFVSTable(p_strFVSOutTables[y]);
                if (!bSkip)
                {
                    IList<string> lstExcludedTables = new List<string>() {"FVS_TREELIST", "FVS_CUTLIST", "FVS_ATRTLIST",
                                                                          "FVS_STRCLASS", "FVS_CASES", "FVS_CANPROFILE"};
                    if (!lstExcludedTables.Contains(p_strFVSOutTables[y].Trim().ToUpper()))
                    {
                        SQLite.m_strSQL = "SELECT DISTINCT b.standid,b.year,b.rowcount " +
                                          "FROM " + p_strFVSOutTables[y] + " a," +
                                          "(SELECT count(*) as rowcount," + p_strFVSOutTables[y] + ".standid,year " +
                                          "FROM " + p_strFVSOutTables[y] + ", FVS_CASES C WHERE" +
                                          " c.StandID = " + p_strFVSOutTables[y] + ".StandID and c.caseid = " + p_strFVSOutTables[y] + ".caseid" +
                                          " and c.runtitle = '" + p_strRunTitle + "' GROUP BY " + p_strFVSOutTables[y] + ".STANDID,YEAR) b " +
                                          "WHERE a.standid=b.standid AND a.year=b.year AND b.rowcount > 1";
                        SQLite.SqlQueryReader(conn, SQLite.m_strSQL);
                        if (SQLite.m_DataReader.HasRows)
                        {
                            p_intError=-1;
                            while (SQLite.m_DataReader.Read())
                            {
                                if (SQLite.m_DataReader["standid"] != System.DBNull.Value &&
                                    SQLite.m_DataReader["year"] != System.DBNull.Value &&
                                    SQLite.m_DataReader["rowcount"] != System.DBNull.Value)
                                        p_strError = p_strError + "ERROR: [Duplicate StandId+Year] RunTitle: " + p_strRunTitle  + " Table:" + p_strFVSOutTables[y] + " StandId: " + SQLite.m_DataReader["standid"].ToString().Trim() + " Year: " + SQLite.m_DataReader["year"].ToString().Trim() + " Row Count:" + SQLite.m_DataReader["rowcount"].ToString().Trim() + "\r\n"; 
                            }
                        }
                        SQLite.m_DataReader.Close();
                    }
                }
            }
            if (p_intError != 0 && p_bAudit)
                frmMain.g_oUtils.WriteText(m_strLogFile, p_strError + "\r\n\r\n");
        }

        /// <summary>
        /// Create the FVS_POTFIRE table. If the configuration includes baseyear than combine the baseyear POTFIRE table to the 
        /// standard POTFIRE table.  If the baseyear POTFIRE table is not used then just use the standard POTFIRE table. Output
        /// is written to FVS_POTFIRE_TEMP table in AUDITS.DB
        /// </summary>
        /// <param name="p_strAuditDbFile">Audits.db file that contains audit results</param>
        /// <param name="p_strDbFile">The FVSOut.db file</param>
        /// <param name="p_strVariant">FVS variant</param>
        /// <param name="p_strRxPackage">RxPackage</param>
        private void CreatePotFireTables(string p_strAuditDbFile, string p_strFVSOutDbFile, string p_strFvsVariant, string p_strRxPackage)
        {

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// CreatePotFireTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x;
            string strPotFireTable = "";

            //
            //CHECK TO SEE IF THE BASEYEAR FILE AND TABLE EXIST
            //RunTitle: FVSOUT_SO_POTFIRE_BaseYr
            //

            string strRunTitle = $@"FVSOUT_{p_strFvsVariant}{m_oRxPackageItem.RunTitleSuffix}";
            string strBaseYrRunTitle = $@"FVSOUT_{p_strFvsVariant}_POTFIRE_BaseYr";
            strPotFireTable = GetPotfireTableName(p_strFVSOutDbFile, p_strFvsVariant);
            if (string.IsNullOrEmpty(strPotFireTable))
            {
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "FVS_POTFIRE/FVS_POTFIRE_EAST table does not exist. Potfire processing skipped\r\n");
                return;
            }
            //get the PREPOST SeqNum configuration for the POTFIRE table
            GetPrePostSeqNumConfiguration(strPotFireTable, p_strRxPackage);
            //
            //CHECK TO SEE IF THE CONFIGURATION INCLUDES BASEYEAR
            //
            if (m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN == "N" &&
                m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN == "N" &&
                m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN == "N" &&
                m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN == "N")
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(p_strAuditDbFile)))
                {
                    conn.Open();
                    if (SQLite.m_intError == 0)
                    {
                        if (SQLite.TableExist(conn, m_strPotFireTable))
                            SQLite.SqlNonQuery(conn, "DROP TABLE " + m_strPotFireTable);

                        // Attach FVSOut.db
                        SQLite.m_strSQL = $@"ATTACH DATABASE '{p_strFVSOutDbFile}' AS FVSOUT";
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        //@ToDo: Do we need to this? This is the path that executes of BaseYear is not used
                        // Historically this overwrote the contents of fvs_potfire with fvs_potfire temp in the output db
                        // For now I'm going to try to use fvs_potfire_temp if base year is used. Don't want to corrupt fvs_potfire in FVSOut.db
                        SQLite.m_strSQL = $@"CREATE TABLE {m_strPotFireTable} AS SELECT * FROM {strPotFireTable}";
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    }
                }
                return;
            }
            //create the new FVS_POTFIRE table by inserting the baseyear POTFIRE records
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(p_strAuditDbFile)))
            {
                conn.Open();
                if (SQLite.m_intError == 0)
                {
                    // Attach FVSOut.db
                    SQLite.m_strSQL = $@"ATTACH DATABASE '{p_strFVSOutDbFile}' AS FVSOUT";
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    if (SQLite.TableExist(conn, "tempBASEYEAR")) SQLite.SqlNonQuery(conn, "DROP TABLE TEMPBASEYEAR");
                    if (SQLite.TableExist(conn, "BASEYEAR")) SQLite.SqlNonQuery(conn, "DROP TABLE BASEYEAR");
                    if (SQLite.TableExist(conn, "NONBASEYEAR")) SQLite.SqlNonQuery(conn, "DROP TABLE NONBASEYEAR");
                    if (SQLite.TableExist(conn, m_strPotFireTable)) SQLite.SqlNonQuery(conn, "DROP TABLE " + m_strPotFireTable);
                    string[] strSQL = Queries.FVS.FVSOutputTable_SQlitePrePostPotFireBaseYearStep1SQL(strPotFireTable,
                        m_strPotFireTable, strBaseYrRunTitle, strRunTitle);

                    for (x = 0; x <= strSQL.Length - 1; x++)
                    {
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL[x] + "\r\n");
                        SQLite.SqlNonQuery(conn, strSQL[x]);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        if (SQLite.m_intError != 0) break;
                    }

                    // returns a comma-separated string
                    string strSourceFieldNames = SQLite.getFieldNames(conn, "SELECT * FROM NONBASEYEAR");
                    string[] arrFieldNames = frmMain.g_oUtils.ConvertListToArray(strSourceFieldNames, ",");
                    string[] arrRemoveFieldNames = new string[] { "NEWYEAR", "BASEYEAR_YN"};
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    for (int i = 0; i < arrFieldNames.Length; i++)
                    {
                        string strField = arrFieldNames[i];
                        if (!arrRemoveFieldNames.Contains(strField.ToUpper()))
                        {
                            sb.Append(strField);
                            sb.Append(",");
                        }
                    }
                    string strTargetFieldNames = sb.ToString();
                    if (strTargetFieldNames.Trim().Length > 0) strTargetFieldNames = strTargetFieldNames.Substring(0, strTargetFieldNames.Length - 1);
                    strSQL = Queries.FVS.FVSOutputTable_SQlitePrePostPotFireBaseYearStep2SQL(m_strPotFireTable,strTargetFieldNames);
                    for (x = 0; x <= strSQL.Length - 1; x++)
                    {
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL[x] + "\r\n");
                        SQLite.SqlNonQuery(conn, strSQL[x]);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        if (SQLite.m_intError != 0) break;
                    }
                }
            }
        }

        private string GetPotfireTableName(string p_strFVSOutDbFile, string p_strVariant)
        {
            string strBaseYrRunTitle = $@"FVSOUT_{p_strVariant}_POTFIRE_BaseYr";
            string strPotFireTableName = "";
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(p_strFVSOutDbFile)))
            {
                conn.Open();
                if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSPotFireTableName))
                {
                    strPotFireTableName = Tables.FVS.DefaultFVSPotFireTableName;
                }
                else if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSPotFireEastTableName))
                {
                    strPotFireTableName = Tables.FVS.DefaultFVSPotFireEastTableName;
                }
                long lngCount = SQLite.getRecordCount(conn, $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSCasesTableName} WHERE RUNTITLE = '{strBaseYrRunTitle}'", Tables.FVS.DefaultFVSCasesTableName);
                if (lngCount > 0)
                {
                    if (!string.IsNullOrEmpty(strPotFireTableName))
                    {
                        //see which potfire table
                        SQLite.m_strSQL = $@"SELECT YEAR FROM {strPotFireTableName}, {Tables.FVS.DefaultFVSCasesTableName} c where {strPotFireTableName}.CASEID = C.CASEID
                            AND RUNTITLE = '{strBaseYrRunTitle}' LIMIT 1";
                        lngCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, strPotFireTableName);
                        if (lngCount > 0)
                        {
                            m_bPotFireBaseYearTableExist = true;
                        }
                    }
                }
            }
            return strPotFireTableName;
        }
        // This is the function called by RunAppend_Main
        private void CreatePrePostSeqNumMatrixTables(string p_strDbFile,string p_strRxPackageId, bool p_bAudit)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            CreatePrePostSeqNumMatrixTables(oAdo, oAdo.m_OleDbConnection,p_strRxPackageId, p_bAudit);
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
        }
        private void CreatePrePostSeqNumMatrixTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string p_strRxPackageId,bool p_bAudit)
        {
           int z;
           string[] strSourceTableArray = p_oAdo.getTableNames(p_oAdo.m_OleDbConnection);
           CreateFVSPrePostSeqNumWorkTables(p_oAdo, p_oAdo.m_OleDbConnection, "FVS_SUMMARY", "FVS_SUMMARY",p_strRxPackageId, p_bAudit);
           CreateFVSPrePostSeqNumWorkTables(p_oAdo, p_oAdo.m_OleDbConnection, "FVS_CUTLIST", "FVS_CUTLIST",p_strRxPackageId, p_bAudit);
           CreateFVSPrePostSeqNumWorkTables(p_oAdo, p_oAdo.m_OleDbConnection, "FVS_POTFIRE", "FVS_POTFIRE",p_strRxPackageId, p_bAudit);
           
           for (z = 0; z <= strSourceTableArray.Length - 1; z++)
           {
               if (strSourceTableArray[z] == null) break;

               if (RxTools.ValidFVSTable(strSourceTableArray[z]))
               {
                   if (strSourceTableArray[z].Trim().ToUpper() != "FVS_SUMMARY" &&
                       strSourceTableArray[z].Trim().ToUpper() != "FVS_CUTLIST" &&
                       strSourceTableArray[z].Trim().ToUpper() != "FVS_POTFIRE")
                   {

                       CreateFVSPrePostSeqNumWorkTables(p_oAdo, p_oAdo.m_OleDbConnection, strSourceTableArray[z], strSourceTableArray[z], p_strRxPackageId, p_bAudit);

                   }
               }
           }

           
        }
        private void RunPOSTAudit_Main()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
			this.m_intError=0;
			int intCount=0;
			m_intProgressOverallTotalCount=1;
			m_intProgressStepCurrentCount=0;
			m_strError="";
			m_strWarning="";
			m_intWarning=0;
			m_intProgressOverallCurrentCount=1;
			string strPackage="";
			string strVariant="";
            string strSQL = "";
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(this.lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;
            string strAuditDbFile;
            string strRxPackageWorktable = "rxpackage_work_table";
            string strRxPackageWorktable2 = "rxpackage_work_table2";
            string strRxWorktable = "rx_work_table";
			
            bool bDisplay = false;

            if (m_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

			System.DateTime oDate = System.DateTime.Now;
			string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
			m_strLogDate = oDate.ToString(strDateFormat);
			m_strLogDate = m_strLogDate.Replace("/","_"); m_strLogDate=m_strLogDate.Replace(":","_");
			int x,y;

            try
            {
                intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);
                for (x = 0; x <= intCount - 1; x++)
                {
                    oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                    m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = true;
                    m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLvItem.ListView, x, COL_RUNSTATUS, "Text", "");
                    //see if checked
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false))
                        m_intProgressOverallTotalCount++;
                }
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);

                bDisplay = this.DisplayAuditMessage;
                this.DisplayAuditMessage = false;

                //total overall progress bar update
                UpdateTherm(m_frmTherm.progressBar2,
                        m_intProgressOverallCurrentCount,
                        m_intProgressOverallTotalCount);

                if (m_intError == 0)
                {
                    DisplayAuditMessage = bDisplay;
                    for (x = 0; x <= intCount - 1; x++)
                    {
                        oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                        this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
                        this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");

                        if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false) == true)
                        {
                            m_intProgressStepTotalCount = 30;
                            m_intProgressStepCurrentCount = 0;
                            m_intProgressOverallCurrentCount++;

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);

                            string strItemDialogMsg = "";

                            int intItemError = 0;
                            string strItemError = "";
                            int intItemWarning = 0;
                            string strItemWarning = "";

                            this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = false;
                            frmMain.g_oDelegate.ExecuteControlMethodWithParam((System.Windows.Forms.Control)oLv, "EnsureVisible", new object[] { x });
                            frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Selected", true);
                            frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Focused", true);



                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkGoldenrod);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Post-Processing Audit");



                            //get the variant
                            strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_VARIANT, "Text", false);
                            strVariant = strVariant.Trim();

                            //get the package and treatments
                            strPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_PACKAGE, "Text", false);
                            strPackage = strPackage.Trim();

                            //find the package item in the package collection
                            if (this.m_oRxPackageItem_Collection == null)
                            {
                                //load rxpackage properties
                                m_oRxPackageItem_Collection = new RxPackageItem_Collection();
                                this.m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oQueries, this.m_oRxPackageItem_Collection);
                            }
                            for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                            {
                                if (this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
                                    break;
                            }
                            if (y <= m_oRxPackageItem_Collection.Count - 1)
                            {
                                this.m_oRxPackageItem = new RxPackageItem();
                                m_oRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), m_oRxPackageItem);
                            }
                            else
                            {
                                this.m_oRxPackageItem = null;
                            }


                            //get the list of treatment cycle year fields to reference for this package
                            this.m_strRxCycleList = "";
                            if (m_oRxPackageItem.SimulationYear1Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear1Rx.Trim() != "000") this.m_strRxCycleList = "1,";
                            if (m_oRxPackageItem.SimulationYear2Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear2Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                            if (m_oRxPackageItem.SimulationYear3Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear3Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                            if (m_oRxPackageItem.SimulationYear4Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear4Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                            if (this.m_strRxCycleList.Trim().Length > 0)
                                this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

                            this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            // With the SQLite rewrite, we run the queries in a temp database instead of PostAudit.accdb
                            // The permanent audit tables are in the new SQLite audits.db
                            strAuditDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile;
                            string strTempDb = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                            SQLite.CreateDbFile(strTempDb);
                            string strTempCutListTable = "tmpCutTree";
                            int intTreeTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREE");
                            int intCondTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("CONDITION");
                            int intPlotTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("PLOT");

                            // Prepare tmpCutTree for next run
                            string strTempConn = SQLite.GetConnectionString(strTempDb);
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strTempConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, strTempCutListTable))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE " + strTempCutListTable);
                                }
                                frmMain.g_oTables.m_oFvs.CreateFVSOutTreeTable(SQLite, conn, strTempCutListTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Created table " + strTempCutListTable + "\r\n");
                                SQLite.SqlNonQuery(conn, $@"ATTACH DATABASE '{m_strFvsTreeDb}' AS TREELIST");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Attached database " + m_strFvsTreeDb + "\r\n");
                                SQLite.SqlNonQuery(conn, $@"INSERT INTO {strTempCutListTable} SELECT * FROM {Tables.FVS.DefaultFVSCutTreeTableName} 
                                                         WHERE FVS_VARIANT ='{strVariant}' AND RXPACKAGE = '{strPackage}'");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Populated table " + strTempCutListTable + "\r\n");

                                SQLite.SqlNonQuery(conn, $@"ATTACH DATABASE '{strAuditDbFile}' AS AUDITS");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Attached database " + strAuditDbFile + "\r\n");
                                if (SQLite.TableExist(conn, "audit_Post_SUMMARY") == false)
                                {
                                    SQLite.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistSUMMARYtableSQL("audit_Post_SUMMARY");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                else
                                {
                                    // Delete records for this variant package
                                    SQLite.m_strSQL = $@"DELETE FROM audit_Post_SUMMARY WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                }
                                if (SQLite.TableExist(conn, "audit_Post_NOVALUE_ERROR") == false)
                                {
                                    SQLite.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL("audit_Post_NOVALUE_ERROR");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                else
                                {
                                    // Delete records for this variant package
                                    SQLite.m_strSQL = $@"DELETE FROM audit_Post_NOVALUE_ERROR WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                if (SQLite.TableExist(conn, "audit_Post_VALUE_ERROR") == false)
                                {
                                    SQLite.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL("audit_Post_VALUE_ERROR");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                else
                                {
                                    // Delete records for this variant package
                                    SQLite.m_strSQL = $@"DELETE FROM audit_Post_VALUE_ERROR WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                if (SQLite.TableExist(conn, "audit_Post_NOTFOUND_ERROR") == false)
                                {
                                    SQLite.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL("audit_Post_NOTFOUND_ERROR");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                else
                                {
                                    // Delete records for this variant package
                                    SQLite.m_strSQL = $@"DELETE FROM audit_Post_NOTFOUND_ERROR WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                if (SQLite.TableExist(conn, "audit_Post_SPCDCHANGE_WARNING") == false)
                                {
                                    SQLite.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistFVSFIA_TREEMATCHINGtableSQL("audit_Post_SPCDCHANGE_WARNING", "WARNING_DESC");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                else
                                {
                                    // Delete records for this variant package
                                    SQLite.m_strSQL = $@"DELETE FROM audit_Post_SPCDCHANGE_WARNING WHERE FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                // Add for rxpackage_work_table; Converting rxpackage and rx to SQLite
                                string rxPackageDb = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.RxPackage);
                                SQLite.SqlNonQuery(conn, $@"ATTACH DATABASE '{rxPackageDb}' AS MASTER");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Attached database " + rxPackageDb + "\r\n");
                                ;
                                if (SQLite.TableExist(conn, strRxPackageWorktable) == true)
                                {
                                    SQLite.m_strSQL = $@"DROP TABLE {strRxPackageWorktable}";
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                int intRxPackageTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREATMENT PACKAGES");
                                string[] SQLiteArray = Queries.FVS.FVSOutputTable_RxPackageWorktable(strRxPackageWorktable,
                                    m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE]);
                                for (y = 0; y <= SQLiteArray.Length - 1; y++)
                                {
                                    strSQL = SQLiteArray[y];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                            }

                            uc_filesize_monitor1.BeginMonitoringFile(strAuditDbFile, 2000000000, "2GB");
                            uc_filesize_monitor1.Information = "BIOSUM DB file containing FVS OUTPUT PRE/POST AUDIT tables";

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            string[] sqlArray = null;
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strTempConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, strRxPackageWorktable2))
                                {
                                    SQLite.SqlNonQuery(conn, $@"DROP TABLE {strRxPackageWorktable2}");
                                }
                                if (SQLite.TableExist(conn, strRxWorktable))
                                {
                                    SQLite.SqlNonQuery(conn, $@"DROP TABLE {strRxWorktable}");
                                }
                                if (SQLite.TableExist(conn, "fvs_tree_unique_biosum_plot_id_work_table"))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE fvs_tree_unique_biosum_plot_id_work_table");
                                }
                                if (SQLite.TableExist(conn, "fvs_tree_biosum_plot_id_work_table"))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE fvs_tree_biosum_plot_id_work_table");
                                }
                                if (SQLite.TableExist(conn, "cond_biosum_cond_id_work_table"))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE cond_biosum_cond_id_work_table");
                                }
                                if (SQLite.TableExist(conn, "plot_biosum_plot_id_work_table"))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE plot_biosum_plot_id_work_table");
                                }
                                if (SQLite.TableExist(conn, "tree_fvs_tree_id_work_table"))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE tree_fvs_tree_id_work_table");
                                }
                                // Attach helper databases
                                string strMasterDb = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree);
                                SQLite.SqlNonQuery(conn, $@"ATTACH DATABASE '{strMasterDb}' AS master");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, $@"ATTACH DATABASE '{strAuditDbFile}' AS audit");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");

                                // audit_Post_SUMMARY
                                sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryFVS(
                                strRxWorktable,
                                strRxPackageWorktable2,
                                m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE],
                                m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE],
                                m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE],
                                "audit_Post_SUMMARY", strTempCutListTable,
                                strPackage, strVariant, strRxPackageWorktable);

                                for (y = 0; y <= sqlArray.Length - 1; y++)
                                {
                                    strSQL = sqlArray[y];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                                m_intProgressStepCurrentCount,
                                                m_intProgressStepTotalCount);

                                }
                            }

                            long lngRowCount = 0;
                            string strAuditConn = SQLite.GetConnectionString(strAuditDbFile);
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strAuditConn))
                            {
                                conn.Open();
                                // Attach FVSOUT_TREE_LIST.db for access to tmpTreeTable
                                SQLite.m_strSQL = $@"ATTACH DATABASE '{m_strFvsTreeDb}' AS TREELIST";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                // Attach temp database for access to worktables
                                SQLite.m_strSQL = $@"ATTACH DATABASE '{strTempDb}' AS TEMPDB";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                // Attach MASTER database for access to TREE table
                                SQLite.m_strSQL = $@"ATTACH DATABASE '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree)}' AS MASTER";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


                                //
                                //process NOVALUE_ERROR column
                                //
                                sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_NOVALUE_ERROR(
                                    "audit_Post_NOVALUE_ERROR",
                                    "audit_Post_SUMMARY",
                                    strTempCutListTable,
                                    strVariant, strPackage);

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any no value errors
                                lngRowCount = SQLite.getRecordCount(conn, "SELECT COUNT(*) FROM audit_Post_SUMMARY WHERE FVS_VARIANT = '" + strVariant + "' AND RXPACKAGE='" + strPackage + "' AND NOVALUE_ERROR IS NOT NULL AND LENGTH(TRIM(NOVALUE_ERROR)) > 0 AND TRIM(NOVALUE_ERROR) <> 'NA' AND CAST(NOVALUE_ERROR as int) > 0", "AUDIT_POST_SUMMARY");
                                if (lngRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = sqlArray[0];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);

                                //
                                //process VALUE_ERROR column
                                //
                                sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_VALUE_ERROR(
                                    "audit_Post_VALUE_ERROR",
                                    "audit_Post_SUMMARY",
                                    strTempCutListTable,
                                    strVariant, strPackage);

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any no value errors
                                lngRowCount = SQLite.getRecordCount(conn, "SELECT COUNT(*) FROM audit_Post_SUMMARY WHERE RXPACKAGE='" + strPackage + "' AND VALUE_ERROR IS NOT NULL AND LENGTH(TRIM(VALUE_ERROR)) > 0 AND TRIM(VALUE_ERROR) <> 'NA' AND CAST(VALUE_ERROR as INT) > 0 AND TRIM(COLUMN_NAME) <> 'DBH'", "AUDIT_POST_SUMMARY");
                                if (lngRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = sqlArray[0];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);

                                //
                                //PROCESS NOT FOUND IN TABLES ERRORS
                                //
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any not found in table errors
                                strSQL = "SELECT COUNT(*) FROM audit_Post_SUMMARY " +
                                         "WHERE rxpackage='" + strPackage + "' AND " +
                                               "(NF_IN_COND_TABLE_ERROR IS NOT NULL AND LENGTH(TRIM(NF_IN_COND_TABLE_ERROR)) > 0 AND TRIM(NF_IN_COND_TABLE_ERROR) <> 'NA' AND CAST(NF_IN_COND_TABLE_ERROR AS INTEGER) > 0) OR " +
                                               "(NF_IN_PLOT_TABLE_ERROR IS NOT NULL AND LENGTH(TRIM(NF_IN_PLOT_TABLE_ERROR)) > 0 AND TRIM(NF_IN_PLOT_TABLE_ERROR) <> 'NA' AND CAST(NF_IN_PLOT_TABLE_ERROR AS INTEGER) > 0) OR " +
                                               "(NF_IN_RX_TABLE_ERROR IS NOT NULL AND LENGTH(TRIM(NF_IN_RX_TABLE_ERROR)) > 0 AND TRIM(NF_IN_RX_TABLE_ERROR) <> 'NA' AND CAST(NF_IN_RX_TABLE_ERROR AS INTEGER) > 0) OR " +
                                               "(NF_IN_RXPACKAGE_TABLE_ERROR IS NOT NULL AND LENGTH(TRIM(NF_IN_RXPACKAGE_TABLE_ERROR)) > 0 AND TRIM(NF_IN_RXPACKAGE_TABLE_ERROR) <> 'NA' AND CAST(NF_IN_RXPACKAGE_TABLE_ERROR AS INTEGER) > 0) OR " +
                                               "(NF_RXPACKAGE_RXCYCLE_RX_ERROR IS NOT NULL AND LENGTH(TRIM(NF_RXPACKAGE_RXCYCLE_RX_ERROR)) > 0 AND TRIM(NF_RXPACKAGE_RXCYCLE_RX_ERROR) <> 'NA' AND CAST(NF_RXPACKAGE_RXCYCLE_RX_ERROR AS INTEGER) > 0) OR " +
                                               "(NF_IN_TREE_TABLE_ERROR IS NOT NULL AND LENGTH(TRIM(NF_IN_TREE_TABLE_ERROR)) > 0 AND TRIM(NF_IN_TREE_TABLE_ERROR) <> 'NA' AND CAST(NF_IN_TREE_TABLE_ERROR AS INTEGER) > 0)";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                lngRowCount = SQLite.getRecordCount(conn, strSQL, "audit_Post_SUMMARY");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                if (lngRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_NOTFOUND_ERROR(
                                    "audit_Post_NOTFOUND_ERROR", "audit_Post_SUMMARY", strTempCutListTable,
                                    m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE].Trim(),
                                    m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE].Trim(),
                                    m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE].Trim(),
                                    strRxWorktable, strRxPackageWorktable2, "rxpackage_work_table");
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //
                                //PROCESS SPCD CHANGE WARNINGS
                                //
                                sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_SPCDCHANGE_WARNING(
                                        "audit_Post_SPCDCHANGE_WARNING",
                                        "audit_Post_SUMMARY",
                                        strTempCutListTable,
                                        m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE].Trim(),
                                        strVariant, strPackage);

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                            //check if any no value errors
                            lngRowCount = SQLite.getRecordCount(conn, "SELECT COUNT(*) FROM audit_Post_SUMMARY WHERE RXPACKAGE='" + strPackage + "' AND TREE_SPECIES_CHANGE_WARNING IS NOT NULL AND LENGTH(TRIM(TREE_SPECIES_CHANGE_WARNING)) > 0 AND TRIM(TREE_SPECIES_CHANGE_WARNING) <> 'NA' AND CAST(TREE_SPECIES_CHANGE_WARNING AS INTEGER) > 0", "AUDIT_POST_SUMMARY");
                            if (lngRowCount > 0)
                            {
                                //insert the new audit records
                                strSQL = sqlArray[0];
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);
                        }
                                //
                                //process DATA CORRUPTION: TREES ARE MATCHED UP INCORRECTLY
                                // 22-NOV-2022: No longer running this audit. DBH will never match when we throw away cycle 1
                                //
                                //sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_TREEMATCH_ERROR(
                                //           "audit_Post_TREEMATCH_ERROR",
                                //           "audit_Post_SUMMARY",
                                //           strTempCutListTable,
                                //           m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE].Trim(),
                                //           strVariant, strPackage);

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                            //check if value errors for DBH
                            //strSQL = "SELECT COUNT(*) FROM audit_Post_SUMMARY " +
                            //         "WHERE RXPACKAGE='" + strPackage + "' AND " +
                            //               "TRIM(COLUMN_NAME)='DBH' AND " +
                            //               "VALUE_ERROR IS NOT NULL AND " +
                            //               "LEN(TRIM(VALUE_ERROR)) > 0 AND " +
                            //               "TRIM(VALUE_ERROR) <> 'NA' AND " + 
                            //               "VAL(VALUE_ERROR) > 0";
                            //if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            //    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                            //intRowCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, strSQL, "AUDIT_POST_SUMMARY");
                            //if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            //    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            //if (intRowCount > 0)
                            //    {
                            //        //insert the new audit records
                            //        strSQL = sqlArray[1];
                            //        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            //            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                            //        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                            //        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            //            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            //    }
                                //
                                //SUMMARIZE REPORT
                                //
                                if (m_strError.Trim().Length > 0) m_strError = m_strError + "\r\n\r\n==================================================================================================================\r\n\r\n";
                                m_strLogFile = m_strFvsTreeDb + "_Audit_" + m_strLogDate.Replace(" ", "_") + ".txt";
                                //if (strItemDialogMsg.Trim().Length > 0)
                                //{
                                //    strItemDialogMsg = strItemDialogMsg + "\r\n\r\n";
                                //    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\n");
                                //}
                                strItemDialogMsg = strItemDialogMsg + "POST-PROCESSING AUDIT LOG \r\n";
                                strItemDialogMsg = strItemDialogMsg + "-------------------------- \r\n\r\n";
                                strItemDialogMsg = strItemDialogMsg + "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n";
                                strItemDialogMsg = strItemDialogMsg + "Database File:" + frmMain.g_oUtils.getFileName(m_strFvsTreeDb) + "\r\n";
                                strItemDialogMsg = strItemDialogMsg + "Variant:" + strVariant + " \r\n";
                                strItemDialogMsg = strItemDialogMsg + "Package:" + strPackage + " \r\n\r\n";

                                frmMain.g_oUtils.WriteText(m_strLogFile, "POST-PROCESSING AUDIT LOG \r\n");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "-------------------------- \r\n\r\n");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "Database File:" + frmMain.g_oUtils.getFileName(m_strFvsTreeDb) + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "Variant:" + strVariant + " \r\n");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "Package:" + strPackage + " \r\n\r\n");
                            //NOVALUE ERRORS
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strAuditConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, "audit_Post_NOVALUE_ERROR"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_NOVALUE_ERROR\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_NOVALUE_ERROR\r\n---------------------------\r\n");

                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_NOVALUE_ERROR WHERE FVS_VARIANT = '" + strVariant + "' AND RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                                    SQLite.SqlQueryReader(conn, strSQL);
                                    if (SQLite.m_DataReader.HasRows)
                                    {
                                        intItemError = -1;
                                        strItemError = strItemError + "\r\n\r\naudit_Post_NOVALUE_ERROR\r\n---------------------------\r\n";
                                        while (SQLite.m_DataReader.Read())
                                        {

                                            strItemError = strItemError + "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    SQLite.m_DataReader.Close();
                                    SQLite.m_DataReader.Dispose();
                                }
                                //NOTFOUND errors
                                if (SQLite.TableExist(conn, "audit_Post_NOTFOUND_ERROR"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_NOTFOUND_ERROR\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_NOTFOUND_ERROR\r\n---------------------------\r\n");
                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_NOTFOUND_ERROR WHERE FVS_VARIANT = '" + strVariant + "' AND  RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                                    SQLite.SqlQueryReader(conn, strSQL);
                                    if (SQLite.m_DataReader.HasRows)
                                    {
                                        intItemError = -1;
                                        strItemError = strItemError + "\r\n\r\naudit_Post_NOTFOUND_ERROR\r\n---------------------------\r\n";

                                        while (SQLite.m_DataReader.Read())
                                        {

                                            strItemError = strItemError + "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    SQLite.m_DataReader.Close();
                                    SQLite.m_DataReader.Dispose();
                                }


                                //VALUE errors
                                if (SQLite.TableExist(conn, "audit_Post_VALUE_ERROR"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_VALUE_ERROR\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_VALUE_ERROR\r\n---------------------------\r\n");
                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_VALUE_ERROR WHERE FVS_VARIANT = '" + strVariant + "' AND RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                                    SQLite.SqlQueryReader(conn, strSQL);
                                    if (SQLite.m_DataReader.HasRows)
                                    {
                                        intItemError = -1;
                                        strItemError = strItemError + "\r\n\r\naudit_Post_VALUE_ERROR\r\n---------------------------\r\n";

                                        while (SQLite.m_DataReader.Read())
                                        {

                                            strItemError = strItemError + "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["ERROR_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    SQLite.m_DataReader.Close();
                                    SQLite.m_DataReader.Dispose();
                                }

                                //SPCD CHANGE WARNINGS
                                if (SQLite.TableExist(conn, "audit_Post_SPCDCHANGE_WARNING"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_SPCDCHANGE_WARNING\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_SPCDCHANGE_WARNING\r\n---------------------------\r\n");

                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,WARNING_DESC FROM audit_Post_SPCDCHANGE_WARNING WHERE FVS_VARIANT = '" + strVariant + "' AND RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,WARNING_DESC";
                                    SQLite.SqlQueryReader(conn, strSQL);
                                    if (SQLite.m_DataReader.HasRows)
                                    {
                                        intItemWarning = -1;
                                        strItemWarning = strItemWarning + "\r\n\r\naudit_Post_SPCDCHANGE_WARNING\r\n---------------------------\r\n";
                                        while (SQLite.m_DataReader.Read())
                                        {

                                            strItemWarning = strItemWarning + "WARNING: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["WARNING_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "WARNING: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["WARNING_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "WARNING: COLUMN:" + SQLite.m_DataReader["COLUMN_NAME"].ToString() + " MSG:" + SQLite.m_DataReader["WARNING_DESC"] + " Records:" + SQLite.m_DataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    SQLite.m_DataReader.Close();
                                    SQLite.m_DataReader.Dispose();
                                }

                                if (SQLite.TableExist(conn, strTempCutListTable))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE " + strTempCutListTable);
                                }
                            }

                            if (intItemError == 0 && intItemWarning == 0)
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\nPassed Audit";
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Green);
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT: OK");
                                }
                                else if (intItemError != 0)
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\nFailed Audit";
                                    m_intError = intItemError;
                                    if (strItemError.Trim().Length > 50)
                                    {
                                        strItemError = strItemError.Substring(0, 45) + "....(See log file)";

                                    }
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Red);
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT ERROR: See Log File");
                                }
                                else if (intItemWarning != 0)
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\nPassed Audit with Warning Message(s)";
                                    if (strItemWarning.Trim().Length > 50)
                                    {
                                        strItemWarning = strItemWarning.Substring(0, 45) + "....(See log file)";

                                    }
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkOrange);
                                    if (strItemWarning.Substring(0, 8) == "WARNING:")
                                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT See Log File");
                                    else
                                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "AUDIT WARNING: See Log File");

                                }
                                m_strError = m_strError + strItemDialogMsg;
                                
                                frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n\r\n");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "**EOF**");
                                
                                //detail progress bar
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                               

                                //total overall progress bar update
                                UpdateTherm(m_frmTherm.progressBar2,
                                        m_intProgressOverallCurrentCount,
                                        m_intProgressOverallTotalCount);
                                
                        }
                    }
                }
                UpdateTherm(m_frmTherm.progressBar1,
                                   m_intProgressStepTotalCount,
                                   m_intProgressStepTotalCount);
                UpdateTherm(m_frmTherm.progressBar2,
                                m_intProgressOverallTotalCount,
                                m_intProgressOverallTotalCount);

               this.FVSRecordsFinished();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                this.ThreadCleanUp();
                this.CleanupThread();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_output:Audit  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {
                if (DisplayAuditMessage)
                {
                    if (m_intError == 0) this.m_strError = m_strError + "\r\n\r\nOverall Rating: Passed Audit";
                    else m_strError = m_strError + "\r\n\r\n" + "Overall Rating: Failed Audit";
                    //MessageBox.Show(m_strError,"FIA Biosum");
                    FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
                    frmTemp.Text = "FIA Biosum";
                    frmTemp.AutoScroll = false;
                    uc_textboxWithButtons uc_textbox1 = new uc_textboxWithButtons();
                    frmTemp.Controls.Add(uc_textbox1);
                    uc_textbox1.AutoSize = true;
                    uc_textbox1.Dock = DockStyle.Fill;
                    uc_textbox1.lblTitle.Text = "Audit Results";
                    uc_textbox1.TextValue = m_strError;
                    frmTemp.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    frmTemp.ShowDialog();
                }

                if (m_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "****END*****" + System.DateTime.Now.ToString() + "\r\n");
                CleanupThread();

                frmMain.g_oDelegate.m_oEventThreadStopped.Set();
                this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
            }			
		}
		private void InitializeAuditLogTableArray(string[] p_strTableArray)
		{
			
            int y;

            m_strFVSPreAppendAuditTables = new List<string>();
            
			for (y=0;y<=p_strTableArray.Length-1;y++)
			{

                if (RxTools.ValidFVSTable(p_strTableArray[y]))
                {

                    m_strFVSPreAppendAuditTables.Add("audit_" + p_strTableArray[y].Trim() + "_prepost_seqnum_counts_table");
                    m_strFVSPreAppendAuditTables.Add(p_strTableArray[y].Trim() +  "_PREPOST_SEQNUM_MATRIX");
                }
                
			}

            m_strFVSPreAppendAuditTables.Add("audit_FVS_SUMMARY_year_counts_table");
            m_strFVSPreAppendAuditTables.Add("audit_fvs_tree_id");

            m_strFVSPostAppendAuditTables = new List<string>();
            m_strFVSPostAppendAuditTables.Add("audit_Post_SPCDCHANGE_WARNING");
            //m_strFVSPostAppendAuditTables.Add("audit_Post_TREEMATCH_ERROR");
            m_strFVSPostAppendAuditTables.Add("audit_Post_VALUE_ERROR");
            m_strFVSPostAppendAuditTables.Add("audit_Post_NOTFOUND_ERROR");
            m_strFVSPostAppendAuditTables.Add("audit_Post_NOVALUE_ERROR");
            m_strFVSPostAppendAuditTables.Add("audit_Post_SUMMARY");
            


			
		}
        private void InitializeAuditLogTableArray()
        {
            InitializeAuditLogTableArray(Tables.FVS.g_strFVSOutTablesArray);
        }
		
        private void CreateFVSPrePostSeqNumWorkTables(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strSourceTableName, string p_strSourceLinkedTableName,string p_strRxPackageId,bool p_bAudit)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateFVSPrePostSeqNumWorkTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Process Table:" + p_strSourceTableName + "\r\n\r\n");
            }
            

            if (p_strSourceTableName.Trim().ToUpper() == "FVS_CASES") return;
            int x;

            if (p_oAdo.TableExist(p_oConn, p_strSourceTableName))
            {

                GetPrePostSeqNumConfiguration(p_strSourceTableName, p_strRxPackageId);

                //m_oRxTools.CreateFVSPrePostSeqNumTables(p_oAdo, p_oConn, m_oFVSPrePostSeqNumItem, p_strSourceTableName, p_strSourceTableName, p_bAudit, m_bDebug, m_strDebugFile);
            }
            else
            {
                if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, p_strSourceTableName + " table does not exist.\r\n\r\n");
                }
            }

            


        }
        private void GetPrePostSeqNumConfiguration(string p_strFVSOutTable,string p_strRxPackageId)
        {
            int x,y;
            int intDefault = -1;
            int intCustom = -1;
            string strTable = p_strFVSOutTable.Trim().ToUpper();
            bool bDone = false;
           
            //the tree lists use the fvs_summary PREPOST CUSTOM or DEFAULT definition
            //the potential fire data uses the fvs_potfire PREPOST CUSTOM or DEFAULT definition
            //the other tables can either use FVS_SUMMARY PREPOST definition or 
            //use their own custom PREPOST definition.
            if (strTable == "FVS_ATRTLIST")
                strTable = "FVS_CUTLIST";
            else if (strTable == "FVS_TREELIST")
                strTable = "FVS_SUMMARY";

            //get the DEFAULT or CUSTOM configuration
            while (!bDone)
            {
                for (x = 0; x <= m_oFVSPrePostSeqNumItemCollection.Count - 1; x++)
                {

                    if (m_oFVSPrePostSeqNumItemCollection.Item(x).TableName.Trim().ToUpper() == strTable)
                    {
                        if (m_oFVSPrePostSeqNumItemCollection.Item(x).Type == "D")
                            intDefault = x;
                        else if (m_oFVSPrePostSeqNumItemCollection.Item(x).Type == "C")
                        {
                            for (y = 0; y <= m_oFVSPrePostSeqNumItemCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1.Count - 1; y++)
                            {
                                //CUSTOM definitions are used by specified package(s)
                                if (m_oFVSPrePostSeqNumItemCollection.Item(x).m_FVSPrePostSeqNumRxPackageAssgnItem_Collection1.Item(y).RxPackageId ==
                                    p_strRxPackageId)
                                {
                                    intCustom = x;
                                }
                            }
                        }

                    }

                }
                if (intCustom != -1 || intDefault != -1)
                    bDone = true;

                strTable = "FVS_SUMMARY";
            }
            
            if (intCustom != -1)
            {
                m_oFVSPrePostSeqNumItem = new FVSPrePostSeqNumItem();
                m_oFVSPrePostSeqNumItem.CopyProperties(m_oFVSPrePostSeqNumItemCollection.Item(intCustom), m_oFVSPrePostSeqNumItem);
            }
            else
            {
                m_oFVSPrePostSeqNumItem = new FVSPrePostSeqNumItem();
                m_oFVSPrePostSeqNumItem.CopyProperties(m_oFVSPrePostSeqNumItemCollection.Item(intDefault), m_oFVSPrePostSeqNumItem);

            }
        }		


        /// <summary>
        /// Validate that every SeqNum value the user specified for PRE-POST values is found in the FVS_SUMMARY table
        /// </summary>
        /// <param name="p_oConn"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings"></param>
        private void Validate_FvsSummaryPrePostSeqNum(System.Data.SQLite.SQLiteConnection p_oConn, ref int p_intItemError, ref string p_strItemError, ref int p_intItemWarning, 
            ref string p_strItemWarning, bool p_bDoWarnings, string p_strVariant, string p_strRxPackage)
        {           
            int intCycleLength = 10;
            //get the cycle length for this package
            if (this.m_oRxPackageItem != null) intCycleLength = this.m_oRxPackageItem.RxCycleLength;           
            Validate_SeqNumExistence("FVS_SUMMARY", p_oConn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning, p_bDoWarnings, p_strVariant, p_strRxPackage);			
        }

        /// <summary>
        /// Validate that the user specified SEQNUM exists in the fvs output table
        /// </summary>
        /// <param name="p_strFVSOutputTable"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings">AUDIT:true APPEND:false</param>
        private void Validate_SeqNumExistence(string p_strFVSOutputTable, System.Data.SQLite.SQLiteConnection conn, ref int p_intItemError, 
            ref string p_strItemError, ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings, string p_FvsVariant,
            string p_rxPackage)
        {
            string strCol = "";
            int x;

            // attach audit.db
            if (! SQLite.DatabaseAttached(conn, frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile))
            {
                SQLite.m_strSQL = $@"ATTACH DATABASE '{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile}' AS AUDIT";
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
            }

            // Add audit alias to table name
            string strAuditPrePostSeqNumCountsTable = "AUDIT.audit_" + p_strFVSOutputTable + "_prepost_seqnum_counts_table";

            SQLite.m_strSQL = $@"SELECT * FROM {strAuditPrePostSeqNumCountsTable} WHERE standid IS NOT NULL AND FVS_VARIANT = '{p_FvsVariant}' AND RXPACKAGE = '{p_rxPackage}'";
            SQLite.SqlQueryReader(conn, SQLite.m_strSQL);
            if (SQLite.m_intError == 0)
            {
                if (SQLite.m_DataReader.HasRows)
                {
                    if (p_bDoWarnings)
                    {
                        while (SQLite.m_DataReader.Read())
                        {
                            //PRE TREATMENT
                            for (x = 1; x <= 4; x++)
                            {
                                try
                                {
                                    strCol = "pre_cycle" + x.ToString().Trim() + "rows";
                                    if (SQLite.m_DataReader[strCol] != DBNull.Value &&
                                        Convert.ToInt32(SQLite.m_DataReader[strCol]) == 0)
                                    {
                                        switch (x)
                                        {
                                            case 1:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum + " for cycle 1 PRE treatment values.\r\n";
                                                break;
                                            case 2:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum + " for cycle 2 PRE treatment values.\r\n";
                                                break;
                                            case 3:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum + " for cycle 3 PRE treatment values.\r\n";
                                                break;
                                            case 4:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum + " for cycle 4 PRE treatment values.\r\n";
                                                break;
                                        }
                                        p_intItemWarning = -3;
                                        break;
                                    }
                                    else if (SQLite.m_DataReader[strCol] != DBNull.Value &&
                                       Convert.ToInt32(SQLite.m_DataReader[strCol]) > 1)
                                    {
                                        switch (x)
                                        {
                                            case 1:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum + " for cycle 1 PRE treatment values.\r\n";
                                                break;
                                            case 2:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum + " for cycle 2 PRE treatment values.\r\n";
                                                break;
                                            case 3:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum + " for cycle 3 PRE treatment values.\r\n";
                                                break;
                                            case 4:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum + " for cycle 4 PRE treatment values.\r\n";
                                                break;
                                        }
                                        p_intItemWarning = -4;
                                        break;
                                    }
                                }
                                catch
                                {
                                }
                            }
                            //POST TREATMENT
                            for (x = 1; x <= 4; x++)
                            {

                                try
                                {
                                    strCol = "post_cycle" + x.ToString().Trim() + "rows";
                                    if (SQLite.m_DataReader[strCol] != DBNull.Value &&
                                        Convert.ToInt32(SQLite.m_DataReader[strCol]) == 0)
                                    {
                                        switch (x)
                                        {
                                            case 1:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum + " for cycle 1 POST treatment values.\r\n";
                                                break;
                                            case 2:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum + " for cycle 2 POST treatment values.\r\n";
                                                break;
                                            case 3:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum + " for cycle 3 POST treatment values.\r\n";
                                                break;
                                            case 4:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum + " for cycle 4 POST treatment values.\r\n";
                                                break;
                                        }
                                        p_intItemWarning = -5;
                                        break;
                                    }
                                    else if (SQLite.m_DataReader[strCol] != DBNull.Value &&
                                       Convert.ToInt32(SQLite.m_DataReader[strCol]) > 1)
                                    {
                                        switch (x)
                                        {
                                            case 1:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum + " for cycle 1 POST treatment values.\r\n";
                                                break;
                                            case 2:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum + " for cycle 2 POST treatment values.\r\n";
                                                break;
                                            case 3:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum + " for cycle 3 POST treatment values.\r\n";
                                                break;
                                            case 4:
                                                p_strItemWarning = p_strItemWarning + "WARNING: Stand" + SQLite.m_DataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum + " for cycle 4 POST treatment values.\r\n";
                                                break;
                                        }
                                        p_intItemWarning = -6;
                                        break;
                                    }
                                }
                                catch
                                {
                                }
                            }
                        }
                    }
                }
                else
                {
                    p_intItemError = -2;
                    p_strItemError = p_strItemError + "ERROR: No records in " + strAuditPrePostSeqNumCountsTable + " table\r\n";
                }
                SQLite.m_DataReader.Close();
            }
            else
            {
                p_intItemError = SQLite.m_intError;
                p_strItemError = p_strItemError + "ERROR:" + SQLite.m_strError + "\r\n";
            }

        }

        /// <summary>
        /// Validation routine to check the FVS_TreeList,FVS_CutList,and FVS_ATRTLIST tables. Validate Checks:
        /// 1. Check for tree list records
        /// 2. Check if there are any standid,year records in the treelist table that 
        /// are not found in the fvs_summary table.  Every treelist standid,year combination should 
        /// be found in the fvs_summary table.
        /// </summary>
        /// <param name="p_strConnectionString"></param>
        /// <param name="p_strTreeListTableName"></param>
        /// <param name="p_strSummaryTableName"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings"></param>
        /// <param name="p_strRunTitle"></param>
        private void Validate_TreeListTables(string p_strConnectionString,
            string p_strTreeListTableName, string p_strSummaryTableName,
            ref int p_intItemError, ref string p_strItemError,
            ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings,
            string p_strRunTitle)
        {
            string strWorkListTable = "";
            int x;
            string strSQL = "";

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(p_strConnectionString))
            {
                conn.Open();

                SQLite.m_strSQL = $@"ATTACH DATABASE '{m_strFvsOutDb}' AS FVS";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                //ERRORS
                //
                //see if any records 
                //
                SQLite.m_strSQL = $@"SELECT count(*) FROM {m_strFVSSummaryAuditYearCountsTable} a WHERE a.{p_strTreeListTableName} > 0 and 
                    fvs_variant = '{p_strRunTitle.Substring(7, 2)}' and rxpackage = '{p_strRunTitle.Substring(11, 3)}' limit 1";
                if ((int)SQLite.getRecordCount(conn, SQLite.m_strSQL, strWorkListTable) == 0)
                {
                    p_intItemError = -2;
                    p_strItemError = p_strItemError + "ERROR: No trees in " + p_strTreeListTableName.Trim() + " table\r\n";
                    return;
                }
                //
                //ensure the cut list standid,treatment year exists in the fvs summary table
                //
                if (SQLite.TableExist(conn, "temp_treelist"))
                    SQLite.SqlNonQuery(conn, "DROP TABLE temp_treelist");

                if (SQLite.TableExist(conn, "temp_summary"))
                    SQLite.SqlNonQuery(conn, "DROP TABLE temp_summary");

                if (SQLite.TableExist(conn, "temp_missingrows"))
                    SQLite.SqlNonQuery(conn, "DROP TABLE temp_missingrows");

                string[] strSQLArray = Queries.FVS.FVSOutputTable_AuditSelectTreeListCycleYearExistInFVSSummaryTableSQL(
                                             "temp_treelist", "temp_summary", "temp_missingrows", p_strTreeListTableName, 
                                             m_strFVSSummaryAuditYearCountsTable, p_strRunTitle);

                for (x = 0; x <= strSQLArray.Length - 1; x++)
                {
                    strSQL = strSQLArray[x];
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                }


                SQLite.m_strSQL = "SELECT * FROM temp_missingrows";
                SQLite.SqlQueryReader(conn, SQLite.m_strSQL);

                if (SQLite.m_DataReader.HasRows)
                {
                    p_intItemError = -4;
                    while (SQLite.m_DataReader.Read())
                    {
                        p_strItemError = p_strItemError + "ERROR: STANDID:" + SQLite.m_DataReader["standid"].ToString().Trim() + " " +
                            "YEAR:" + SQLite.m_DataReader["year"].ToString().Trim() + " " +
                            "TREECOUNT:" + SQLite.m_DataReader["treecount"].ToString().Trim() + " Standid and year not found in the fvs_summary table\r\n";
                    }

                }
                SQLite.m_DataReader.Close();

                if (SQLite.TableExist(conn, "temp_missingrows"))
                    SQLite.SqlNonQuery(conn, "DROP TABLE temp_missingrows");

            }
            if (p_intItemError != 0) return;

        }
		/// <summary>
		/// The validation is two fold:
        /// 1. Validate that every tree, except FVS created trees, in the FVS_CUTLIST can be found in the FIA tree table
        /// 2. Check if a tree is cut multiple times
        /// An error is returned if cutlist trees can't be found in the tree table.
        /// A warning is returned if a tree is cut more than once.
        /// </summary>
		/// <param name="p_strFVSTreeTableNameToAudit"></param>
		/// <param name="p_strVariant"></param>
		/// <param name="p_strRxPackage"></param>
		/// <param name="p_strRx1"></param>
		/// <param name="p_strRx2"></param>
		/// <param name="p_strRx3"></param>
		/// <param name="p_strRx4"></param>
		/// <param name="p_bAudit"></param>
		/// <param name="p_intItemWarning"></param>
		/// <param name="p_strItemWarning"></param>
		/// <param name="p_intItemError"></param>
		/// <param name="p_strItemError"></param>
        private void Validate_FVSTreeId(string p_strFVSCasesTableName,
                                        string p_strFVSTreeTableNameToAudit,
                                        string p_strVariant,
                                        string p_strRxPackage,
                                        string p_strRx1,
                                        string p_strRx2,
                                        string p_strRx3,
                                        string p_strRx4,
                                        string p_strRunTitle,
                                        ref int p_intItemWarning, 
                                        ref string p_strItemWarning,
                                        ref int p_intItemError, 
                                        ref string p_strItemError)
                                        
        {
            int x,y;
            string strRxCycle;
            string strRx="";
            string strAuditTable = "audit_fvs_tree_id";
            string strSeqNumMatrixTable = "FVS_CUTLIST_PREPOST_SEQNUM_MATRIX";

            string strAuditConn = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile);
            using (System.Data.SQLite.SQLiteConnection auditConn = new System.Data.SQLite.SQLiteConnection(strAuditConn))
            {
                auditConn.Open();
                if (SQLite.TableExist(auditConn, strAuditTable))
                {
                    SQLite.m_strSQL = $@"DELETE FROM {strAuditTable} WHERE FVS_VARIANT = '{p_strVariant}' AND RXPACKAGE = '{p_strRxPackage}'";
                    SQLite.SqlNonQuery(auditConn, SQLite.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                }
                else
                {
                    //create audit_fvs_tree_id table
                    frmMain.g_oTables.m_oFvs.CreateSQLiteFVSTreeIdAudit(SQLite, auditConn, strAuditTable);
                }

                //Attach audits.db for access to seq num matrices
                SQLite.m_strSQL = $@"ATTACH DATABASE '{m_strFvsOutDb}' AS FVS";
                SQLite.SqlNonQuery(auditConn, SQLite.m_strSQL);

                // Insert records into audit_fvs_tree_id table
                for (x = 0; x <= this.m_strRxCycleArray.Length - 1; x++)
                {
                    if (m_strRxCycleArray[x] == null ||
                        m_strRxCycleArray[x].Trim().Length == 0)
                    {
                    }
                    else
                    {
                        strRxCycle = m_strRxCycleArray[x].Trim();
                        switch (strRxCycle)
                        {
                            case "1":
                                strRx = p_strRx1;
                                break;
                            case "2":
                                strRx = p_strRx2;
                                break;
                            case "3":
                                strRx = p_strRx3;
                                break;
                            case "4":
                                strRx = p_strRx4;
                                break;
                        }

                        SQLite.m_strSQL = Queries.FVS.FVSOutputTable_AuditFVSTreeId(
                                                    strAuditTable,
                                                        p_strFVSCasesTableName,
                                                        p_strFVSTreeTableNameToAudit,
                                                        strSeqNumMatrixTable,
                                                        p_strRxPackage,
                                                        strRx,
                                                        strRxCycle, p_strRunTitle);

                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(auditConn, SQLite.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    }
                }
            }

            //
            //loop through all the rx cycles for this package and append them to the fvs tree
            //
            using (var pConn = new System.Data.SQLite.SQLiteConnection(strAuditConn))
            {
                pConn.Open();
                SQLite.m_strSQL = $@"ATTACH DATABASE '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree)}' AS MASTER";
                SQLite.SqlNonQuery(pConn, SQLite.m_strSQL);
                SQLite.m_strSQL = "UPDATE audit_fvs_tree_id  i " +
                                          "INNER JOIN ((" + m_oQueries.m_oFIAPlot.m_strTreeTable + " t " +
                                                          "INNER JOIN " + m_oQueries.m_oFIAPlot.m_strCondTable + " c " +
                                                          "ON t.biosum_cond_id=c.biosum_cond_id) " +
                                                          "INNER JOIN " + m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                                          "ON p.biosum_plot_id=c.biosum_plot_id) " +
                                          "ON i.fvs_tree_id=trim(t.fvs_tree_id) AND i.biosum_cond_id=t.biosum_cond_id " +
                                          "SET i.FOUND_FvsTreeId_YN='Y'";
                SQLite.m_strSQL = $@"UPDATE audit_fvs_tree_id SET FOUND_FvsTreeId_YN = (SELECT 'Y' from {m_oQueries.m_oFIAPlot.m_strTreeTable} t
                    where t.biosum_cond_id = audit_fvs_tree_id.biosum_cond_id
                    and audit_fvs_tree_id.fvs_tree_id = trim(t.fvs_tree_id) AND FVS_VARIANT = '{p_strVariant}' AND RXPACKAGE = '{p_strRxPackage}')
                    where exists (select * from {m_oQueries.m_oFIAPlot.m_strTreeTable} t where t.biosum_cond_id = audit_fvs_tree_id.biosum_cond_id
                    and audit_fvs_tree_id.fvs_tree_id = trim(t.fvs_tree_id) AND FVS_VARIANT = '{p_strVariant}' AND RXPACKAGE = '{p_strRxPackage}')";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(pConn, SQLite.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
            }
            using (System.Data.SQLite.SQLiteConnection auditConn = new System.Data.SQLite.SQLiteConnection(strAuditConn))
            {
                auditConn.Open();
                int intCount = (int)SQLite.getSingleDoubleValueFromSQLQuery(auditConn, $@"SELECT COUNT(*) AS TTLCOUNT FROM audit_fvs_tree_id WHERE FOUND_FvsTreeId_YN='N' AND FVS_VARIANT = '{p_strVariant}' AND RXPACKAGE = '{p_strRxPackage}'", "TEMP");
                if (intCount > 0)
                {
                    p_strItemError = p_strItemError + "ERROR: " + intCount.ToString() + " fvs_cutlist trees for variant " + p_strVariant + " could not be found in the FIADB tree table by matching FVS_Tree_Id (See audit table audit_fvs_tree_id)\r\n";
                    p_intItemError = -5;
                }
            }
        }

        private void Validate_PotFire(string strConn, string p_strPotFireTableName,
            ref int p_intItemError, ref string p_strItemError,
            ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings, string p_strRunTitle)
        {

            //ERRORS
            if ((m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN == "Y" ||
                m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN == "Y" ||
                m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN == "Y" ||
                m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN == "Y") &&
                m_bPotFireBaseYearTableExist == false)
            {
                p_intItemError = -2;
                p_strItemError = "ERROR: POTFIRE Base year file and/or table missing";
                return;
            }

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                //
                //see if any records 
                //
                SQLite.m_strSQL = $@"SELECT count(*) FROM (select c.standid from {p_strPotFireTableName}, FVS_Cases c 
                    WHERE {p_strPotFireTableName}.standid = c.standid and runTitle = '{p_strRunTitle}' limit 1)";
                if ((int)SQLite.getRecordCount(conn, SQLite.m_strSQL, "FVS_POTFIRE") == 0)
                {
                    p_intItemError = -2;
                    p_strItemError = p_strItemError + "ERROR: No fvs potfire records\r\n";
                    return;
                }

                string strFvsVariant = p_strRunTitle.Substring(7, 2);
                string strRxPackageId = p_strRunTitle.Substring(11, 3);
                if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "N")
                {
                    Validate_SeqNumExistence(p_strPotFireTableName, conn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning,
                        p_bDoWarnings, strFvsVariant, strRxPackageId);
                }
            }
        }
        private void Validate_FVSGenericTableSqlite(string strConn, string p_strFvsTableName, string p_strFVSOutputTable, ref int p_intItemError, ref string p_strItemError,
            ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings, string p_strRunTitle)
        {
            //ERRORS
            //
            //see if any records 
            //

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                SQLite.m_strSQL = $@"SELECT count(*) FROM (select c.standid from {p_strFvsTableName}, FVS_Cases c 
                    WHERE {p_strFvsTableName}.standid = c.standid and runTitle = '{p_strRunTitle}' limit 1)";
                if ((int)SQLite.getRecordCount(conn, SQLite.m_strSQL, p_strFvsTableName) == 0)
                {
                    p_intItemError = -2;
                    p_strItemError = p_strItemError + "ERROR: No " + p_strFvsTableName + " records\r\n";
                    return;
                }

                string strFvsVariant = p_strRunTitle.Substring(7, 2);
                string strRxPackageId = p_strRunTitle.Substring(11, 3);
                if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "N")
                    Validate_SeqNumExistence(p_strFVSOutputTable, conn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning,
                        p_bDoWarnings, strFvsVariant, strRxPackageId);
            }


            if (p_intItemError != 0) return;
            if (p_bDoWarnings == false) return;
        }









        private void btnCancel_Click(object sender, System.EventArgs e)
		{
			CancelThread();
		}
		private void CancelThread()
		{
			bool bAbort=frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel processing (Y/N)?");
			if (bAbort)
			{
				if (frmMain.g_oDelegate.m_oThread.IsAlive)
				{
					frmMain.g_oDelegate.m_oThread.Join();
				}
				frmMain.g_oDelegate.StopThread();
				CleanupThread();
			}
		}
		private void CleanupThread()
		{
            uc_filesize_monitor1.EndMonitoringFile();
            uc_filesize_monitor2.EndMonitoringFile();
            uc_filesize_monitor3.EndMonitoringFile();
            //RunAppend_CloseDbConnections(m_oPrePostDbFileItem_Collection);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpboxAppend, "Enabled", true);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpboxAudit, "Enabled", true);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpBoxPostAudit, "Enabled", true);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.GroupBox)grpboxSpCdConvert, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.ComboBox)cmbStep, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnExecute, "Enabled", true);
			frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnChkAll,"Enabled",true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnClearAll, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnRefresh, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnClose, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnHelp, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnViewLogFile, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnViewPostLogFile, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnAuditDb, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnPostAppendAuditDb, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Button)btnCancel, "Enabled", false);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.UserControl)this, "Enabled", true);
           
//            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.UserControl)uc_filesize_monitor1, "Visible", false);
//            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.UserControl)uc_filesize_monitor2, "Visible", false);
//            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.UserControl)uc_filesize_monitor3, "Visible", false);
			this.ParentForm.Enabled=true;
			
		}

		private void lstFvsOutput_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			lstFvsOutput.Items[e.Index].Selected=true;
		}

		private void btnViewLogFile_Click(object sender, System.EventArgs e)
		{
            if (this.lstFvsOutput.SelectedItems.Count == 0)
            {
                MessageBox.Show("No Rows Are Selected", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            string strSearch = $@"FVSOUT_{lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim()}_P{lstFvsOutput.SelectedItems[0].SubItems[COL_PACKAGE].Text.Trim()}_AUDIT_*.txt";			
			string[] strFiles= System.IO.Directory.GetFiles(this.txtOutDir.Text.Trim(), strSearch);

			FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();

			oDlg.uc_select_list_item1.lblTitle.Text = "Open Audit Log File";
			oDlg.uc_select_list_item1.listBox1.Sorted = true;
			for (int x=0;x<=strFiles.Length - 1;x++)
			{
				oDlg.uc_select_list_item1.listBox1.Items.Add(strFiles[x].Substring(this.txtOutDir.Text.Trim().Length+1,strFiles[x].Length - this.txtOutDir.Text.Trim().Length - 1));
			}
			if (oDlg.uc_select_list_item1.listBox1.Items.Count > 0) oDlg.uc_select_list_item1.listBox1.SelectedIndex = oDlg.uc_select_list_item1.listBox1.Items.Count-1;
			oDlg.uc_select_list_item1.lblMsg.Text = "Log File Contents of " + this.txtOutDir.Text.Trim();
			oDlg.uc_select_list_item1.lblMsg.Show();

			oDlg.uc_select_list_item1.Show();

			DialogResult result = oDlg.ShowDialog();
			if (result==DialogResult.OK)
			{
				string strDirAndFile = this.txtOutDir.Text.Trim() + "\\" + oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim();
				System.Diagnostics.Process proc = new System.Diagnostics.Process();
				proc.StartInfo.UseShellExecute = true;
				try
				{
					proc.StartInfo.FileName = strDirAndFile;
				}
				catch
				{
				}
				try
				{
					proc.Start();
				}
				catch (Exception err)
				{
					MessageBox.Show("!!Error!! \n" + 
						"Module - uc_fvs_output:btnViewLogFile_Click \n" + 
						"Err Msg - " + err.Message,
						"View Script",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);
				}
				proc=null;
			}

						

			//strOutDirAndFile = strOutDirAndFile  + "\\" + strDbFile;



		}

		private void btnAuditDb_Click(object sender, System.EventArgs e)
		{
            if (this.lstFvsOutput.SelectedItems.Count == 0)
            {
                MessageBox.Show("No Rows Are Selected", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            string strConn = "";

            string strOutDirAndFile = this.m_strProjDir;
            strOutDirAndFile = strOutDirAndFile + "\\" + Tables.FVS.DefaultFVSAuditsDbFile;
            string strFVSVariant = lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
            string strRxPackage = lstFvsOutput.SelectedItems[0].SubItems[COL_PACKAGE].Text.Trim();
            frmGridView oFrm = new frmGridView();
            oFrm.Text = "Database: Browse (PRE-APPEND Audit Tables)";
            oFrm.UsingSQLite = true;
            if (System.IO.File.Exists(strOutDirAndFile))
            {
                strConn = SQLite.GetConnectionString(strOutDirAndFile);
                bool bAuditTablesForRxPkg = false;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    if (SQLite.TableExist(conn, this.m_strFVSSummaryAuditYearCountsTable))
                    {
                        SQLite.m_strSQL = $@"SELECT YEAR FROM {this.m_strFVSSummaryAuditYearCountsTable} WHERE fvs_variant = '{strFVSVariant}' and rxpackage = '{strRxPackage}' limit 1";
                        long lngCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, this.m_strFVSSummaryAuditYearCountsTable);
                        if (lngCount > 0)
                        {
                            bAuditTablesForRxPkg = true;
                        }
                    }
                    if (!bAuditTablesForRxPkg)
                    {
                        string strWarnMessage = "No PRE-APPEND audit tables exist in the file " + strOutDirAndFile + ". The PRE-APPEND Audit tables cannot be displayed.";
                        MessageBox.Show(strWarnMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                        return;
                    }
                    if (m_strFVSPreAppendAuditTables != null)
                    {
                        for (int x = 0; x <= m_strFVSPreAppendAuditTables.Count - 1; x++)
                        {
                            if (SQLite.TableExist(conn, m_strFVSPreAppendAuditTables[x].Trim()))
                            {
                                SQLite.m_strSQL = $@"SELECT * FROM {m_strFVSPreAppendAuditTables[x].Trim()} WHERE RXPACKAGE = '{strRxPackage}' AND FVS_VARIANT = '{strFVSVariant}'";
                                oFrm.LoadDataSet(strConn, SQLite.m_strSQL, m_strFVSPreAppendAuditTables[x].Trim());
                            }
                        }
                    }
                }

                oFrm.TileGridViews();
                oFrm.Show();
                oFrm.Focus();
            }
            else
            {
                MessageBox.Show("The file " + strOutDirAndFile + " does not exist");
            }
		}
		private void UpdateTherm(System.Windows.Forms.ProgressBar p_oPb,int p_intCurrentStep,int p_intTotalSteps)
		{
            int Percent = 0;
			if (p_oPb != null)
            {
                if (p_intCurrentStep == 0)
                {
                    Percent = 0;
                }
                else
                {
                    if (p_intCurrentStep > 0 && p_intCurrentStep < p_intTotalSteps)
                        Percent = (int)Math.Round((double)(100 * p_intCurrentStep) / p_intTotalSteps, 0);
                    else Percent = 100;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)p_oPb, "Value", Percent);


            }
		}

		private void lstFvsOutput_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			int x;
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.lstFvsOutput.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lstFvsOutput.Items[this.lstFvsOutput.TopItem.Index + (int)dblRow-1].Selected=true;
					this.m_oLvAlternateColors.DelegateListViewItem(lstFvsOutput.Items[this.lstFvsOutput.TopItem.Index + (int)dblRow-1]);
				}
			}
			catch 
			{
			}
		}

		private void lstFvsOutput_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            btnAuditDb.Enabled = false;
            btnPostAppendAuditDb.Enabled = false;
            if (this.lstFvsOutput.SelectedItems.Count > 0 && frmMain.g_oDelegate.CurrentThreadProcessIdle)
            {
                m_oLvAlternateColors.DelegateListViewItem(lstFvsOutput.SelectedItems[0]);

                //Enable/Disable PRE-APPEND Audit Tables; PRE-APPEND audit tables are included in the FVS_AUDITS.db
                string strAuditDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSAuditsDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strAuditDbPath)))
                {
                    if (System.IO.File.Exists(strAuditDbPath))
                    {
                        conn.Open();
                        if (SQLite.TableExist(conn, "audit_FVS_SUMMARY_year_counts_table"))
                        {
                            SQLite.m_strSQL = $@"SELECT YEAR FROM audit_FVS_SUMMARY_year_counts_table WHERE FVS_VARIANT = '{lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim()}' AND RXPACKAGE = '{lstFvsOutput.SelectedItems[0].SubItems[COL_PACKAGE].Text.Trim()}' LIMIT 1";
                            long lngCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, "audit_FVS_SUMMARY_year_counts_table");
                            if (lngCount > 0)
                            {
                                btnAuditDb.Enabled = true;
                            }
                        }
                        if (SQLite.TableExist(conn, "audit_Post_SUMMARY"))
                        {
                            SQLite.m_strSQL = $@"SELECT NOVALUE_ERROR FROM audit_Post_SUMMARY WHERE FVS_VARIANT = '{lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim()}' AND RXPACKAGE = '{lstFvsOutput.SelectedItems[0].SubItems[COL_PACKAGE].Text.Trim()}' LIMIT 1";
                            long lngCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, "audit_FVS_SUMMARY_year_counts_table");
                            if (lngCount > 0)
                            {
                                btnPostAppendAuditDb.Enabled = true;
                            }
                        }

                    }
                }

                    //Enable/Disable POST-APPEND Audit Tables; POST-APPEND audit tables are also included in the FVS_AUDITS.db
                    //Enable/Disable Open Pre Audit Log button
                    btnViewLogFile.Enabled = false;
                // FVSOUT_CA_P010_Audit_2023-01-18_13-53-58.txt
                string strDirectory = this.txtOutDir.Text.Trim();
                if (System.IO.Directory.Exists(strDirectory) == true)
                {
                    string strSearch = $@"FVSOUT_{lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim()}_P{lstFvsOutput.SelectedItems[0].SubItems[COL_PACKAGE].Text.Trim()}_AUDIT_*.txt";
                    string[] strFiles = System.IO.Directory.GetFiles(strDirectory, strSearch);
                    if (strFiles.Length > 0)
                        btnViewLogFile.Enabled = true;
                }

                //Enable/Disable Open Post Audit Log button
                btnViewPostLogFile.Enabled = false;
                if (System.IO.Directory.Exists(this.txtOutDir.Text.Trim()) == true)
                {
                    string strSearch = "FVSOUT_TREE_LIST.db_audit*.txt";
                    string[] strFiles = System.IO.Directory.GetFiles(this.txtOutDir.Text.Trim(), strSearch);
                    if (strFiles.Length > 0)
                        btnViewPostLogFile.Enabled = true;
                }
            }

		}

	
		public bool DisplayAuditMessage
		{
			get {return _bDisplayAuditMsg;}
			set {_bDisplayAuditMsg=value;}
		}

        private void btnPostAudit_Click(object sender, EventArgs e)
        {
            RunPOSTAudit_Start();
        }
        private void RunPOSTAudit_Start()
        {
            m_dbConn = SQLite.GetConnectionString(m_strFvsTreeDb);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_dbConn))
            {
                conn.Open();
                for (int i = 0; i < this.lstFvsOutput.Items.Count; i++)
                {
                    var lvItem = this.lstFvsOutput.Items[i];
                    if (lvItem.Checked)
                    {
                        string strVariant = lvItem.SubItems[COL_VARIANT].Text.Trim();
                        string strRxPackage = lvItem.SubItems[COL_PACKAGE].Text.Trim();
                        long lngTreeRecords = -1;
                        string strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSCutTreeTableName} 
                                           WHERE FVS_VARIANT = '{strVariant}' and RXPACKAGE = '{strRxPackage}'";
                        lngTreeRecords = SQLite.getRecordCount(conn, strSQL, Tables.FVS.DefaultFVSCutTreeTableName);

                        if (lngTreeRecords < 1)
                        {
                            DialogResult result = MessageBox.Show("!!Warning!!\r\n-----------\r\nNo trees in FVS_CUTTREE table for " + strVariant + strRxPackage + ". " +
                                "Continue Auditing? (Y/N)",
                                "FIA BioSum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result == DialogResult.No)
                                return;
                            else
                            {
                                // Uncheck the box for this variant package; Nothing to audit
                                lvItem.Checked = false;
                                m_intError = 0;
                            }
                        }
                    }
                }
            }
            // Check checked item count
            if (this.lstFvsOutput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked. The FVS_CutTree table may be empty!", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            this.DisplayAuditMessage = true;
            this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
                "FVS_TREE CUTLIST POST-PROCESSING Audit", "2");
            this.m_frmTherm.TopMost = true;
            this.m_frmTherm.lblMsg.Text = "";
            this.cmbStep.Enabled = false;
            this.btnExecute.Enabled = false;
            this.btnChkAll.Enabled = false;
            this.btnClearAll.Enabled = false;
            this.btnRefresh.Enabled = false;
            this.btnClose.Enabled = false;
            this.btnHelp.Enabled = false;
            this.btnViewLogFile.Enabled = false;
            this.btnViewPostLogFile.Enabled = false;
            this.btnAuditDb.Enabled = false;
            this.btnPostAppendAuditDb.Enabled = false;


            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(RunPOSTAudit_Main));
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;

            frmMain.g_oDelegate.m_oThread.Start();


        }

        private void btnViewPostLogFile_Click(object sender, EventArgs e)
        {
            if (this.lstFvsOutput.SelectedItems.Count == 0)
            {
                MessageBox.Show("No Rows Are Selected", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            
            string strSearch = "FVSOUT_TREE_LIST.db_audit*.txt";

            string[] strFiles = System.IO.Directory.GetFiles(this.txtOutDir.Text.Trim(), strSearch);

            FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();

            oDlg.uc_select_list_item1.lblTitle.Text = "Open Post Audit Log File";
            oDlg.uc_select_list_item1.listBox1.Sorted = true;
            for (int x = 0; x <= strFiles.Length - 1; x++)
            {
                oDlg.uc_select_list_item1.listBox1.Items.Add(strFiles[x].Substring(this.txtOutDir.Text.Trim().Length + 1, strFiles[x].Length - this.txtOutDir.Text.Trim().Length - 1));
            }
            if (oDlg.uc_select_list_item1.listBox1.Items.Count > 0) oDlg.uc_select_list_item1.listBox1.SelectedIndex = oDlg.uc_select_list_item1.listBox1.Items.Count - 1;
            oDlg.uc_select_list_item1.lblMsg.Text = "Log File Contents of " + this.txtOutDir.Text.Trim();
            oDlg.uc_select_list_item1.lblMsg.Show();

            oDlg.uc_select_list_item1.Show();

            DialogResult result = oDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                string strDirAndFile = this.txtOutDir.Text.Trim() + "\\" + oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim();
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo.UseShellExecute = true;
                try
                {
                    proc.StartInfo.FileName = strDirAndFile;
                }
                catch
                {
                }
                try
                {
                    proc.Start();
                }
                catch (Exception err)
                {
                    MessageBox.Show("!!Error!! \n" +
                        "Module - uc_fvs_output:btnViewLogFile_Click \n" +
                        "Err Msg - " + err.Message,
                        "View Script", System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                }
                proc = null;
            }
        }

        private void SaveTreeSpeciesWorkTable(System.Data.SQLite.SQLiteConnection conn, string strFvsTreeTable, string strRxPackage)
        {
            //
            //SWAP FVS-ONLY SPECIES CODES TO FIA SPECIES CODES IN ORDER FOR FIA TO CALCULATE VOLUMES
            //
            //check if original spcd column exists
            //save the original spcd
            if ((int)SQLite.getRecordCount(conn, "SELECT COUNT(*) AS ROWCOUNT FROM " + strFvsTreeTable + " WHERE TRIM(FVS_SPECIES) IN ('298','999')", strFvsTreeTable) > 0)
            {
                SQLite.m_strSQL = $@"CREATE TABLE cutlist_save_tree_species_work_table as SELECT BIOSUM_COND_ID, FVS_TREE_ID, FVS_SPECIES
                                    FROM {strFvsTreeTable} WHERE TRIM(FVS_SPECIES) IN('298', '999') AND RXPACKAGE = '{strRxPackage}'";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                //update the FVS_SPECIES with FIA SPCD for 298,999 trees
                SQLite.m_strSQL = "UPDATE " + strFvsTreeTable + " SET FVS_SPECIES = " +
                    "CASE WHEN FVS_SPECIES = '298' THEN '299' " +
                    "WHEN FVS_SPECIES = '999' THEN '998' END " +
                    "WHERE FVS_SPECIES IN ('298','999')";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
            }
        }
        
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            switch (cmbStep.Text.Trim())
            {
                case "Step 1 - Define PRE/POST Table SeqNum":
                    this.PREPOSTDefinition();
                    break;
                case "Step 2 - Pre-Processing Audit Check":
                    this.RunPREAudit_Start();
                    break;
                case "Step 3 - Append FVS Output Data":
                    this.RunAppend_Start();
                    break;
                case "Step 4 - Post-Processing Audit Check":
                    this.RunPOSTAudit_Start();
                    break;
                case "Create FVSOut_BioSum.db":
                    this.RunCreateFVSOut_BioSum_Start();
                    break;
                case "Write FVS_InForest to FVSOUT_TREE_LIST.db":
                    this.RunFVSInForestTable_Start();
                    break;
            }

        }
        private void PREPOSTDefinition()
        {
            frmDialog oDlg = new frmDialog();
            oDlg.Initialize_FVS_Output_PREPOST_SeqNum_User_Control();
            oDlg.DisposeOfFormWhenClosing = true;
            oDlg.ShowDialog();
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "FVS", "OUTPUT_DATA" });
        }

        private void btnPostAppendAuditDb_Click(object sender, EventArgs e)
        {
            if (this.lstFvsOutput.SelectedItems.Count == 0)
            {
                MessageBox.Show("No Rows Are Selected", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            string strConn = "";
            string strOutDirAndFile = this.m_strProjDir;
            strOutDirAndFile = strOutDirAndFile + "\\" + Tables.FVS.DefaultFVSAuditsDbFile;
            string strFVSVariant = lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
            string strRxPackage = lstFvsOutput.SelectedItems[0].SubItems[COL_PACKAGE].Text.Trim();

            if (System.IO.File.Exists(strOutDirAndFile))
            {
                strConn = SQLite.GetConnectionString(strOutDirAndFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    if (!SQLite.TableExist(conn, "audit_Post_SUMMARY"))
                    {
                        string strWarnMessage = "No POST-APPEND audit tables exist in the file " + strOutDirAndFile + ". The POST-APPEND Audit tables cannot be displayed.";
                        MessageBox.Show(strWarnMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                        return;
                    }
                    FIA_Biosum_Manager.frmGridView oFrm = new frmGridView();
                    oFrm.Text = "Database: Browse (POST-APPEND Audit Tables)";
                    oFrm.UsingSQLite = true;

                    if (m_strFVSPostAppendAuditTables != null)
                    {
                        for (int x = 0; x <= m_strFVSPostAppendAuditTables.Count - 1; x++)
                        {
                            if (SQLite.TableExist(conn, m_strFVSPostAppendAuditTables[x].Trim()))
                            {
                                SQLite.m_strSQL = $@"SELECT * FROM  {m_strFVSPostAppendAuditTables[x].Trim()} WHERE RXPACKAGE = '{strRxPackage}' AND FVS_VARIANT = '{strFVSVariant}'";
                                oFrm.LoadDataSet(strConn, SQLite.m_strSQL);
                            }
                        }
                    }
                    oFrm.TileGridViews();
                    oFrm.Show();
                    oFrm.Focus();
                }
            }
            else
            {
                MessageBox.Show("The file " + strOutDirAndFile + " does not exist");
            }
        }

        private bool FvsOutWithRequiredTable(string strProjectDir, string strRequiredTable)
        {
            if (System.IO.File.Exists(strProjectDir + Tables.FVS.DefaultFVSOutDbFile))
            {
                string strConn = SQLite.GetConnectionString(strProjectDir + Tables.FVS.DefaultFVSOutDbFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    if (SQLite.TableExist(conn, strRequiredTable))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void RunCreateFVSOut_BioSum_Start()
        {
            // Warning for older projects without FVSOut.db
            if (!System.IO.File.Exists(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile))
            {
                MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            string strBiosumDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutBiosumDbFile;
            if (System.IO.File.Exists(strBiosumDbPath))
            {
                var result = MessageBox.Show(strBiosumDbPath + " already exists. Do you wish to overwrite this file?",
                    "FIA Biosum", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    try
                    {
                        System.IO.File.Delete(strBiosumDbPath);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Unable to delete " + strBiosumDbPath + "! Please make sure the file is not open and try again.", "FIA Biosum");
                        return;
                    }
                }
            }

            this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
                "Create FVSOut_BioSum.db", "1");
            this.m_frmTherm.TopMost = true;
            this.m_frmTherm.lblMsg.Text = "";
            this.cmbStep.Enabled = false;
            this.btnExecute.Enabled = false;
            this.btnChkAll.Enabled = false;
            this.btnClearAll.Enabled = false;
            this.btnRefresh.Enabled = false;
            this.btnClose.Enabled = false;
            this.btnHelp.Enabled = false;
            this.btnViewLogFile.Enabled = false;
            this.btnViewPostLogFile.Enabled = false;
            this.btnAuditDb.Enabled = false;
            this.btnPostAppendAuditDb.Enabled = false;

            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(RunCreateFVSOut_BioSum_Main));
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;

            frmMain.g_oDelegate.m_oThread.Start();

        }

        private void RunCreateFVSOut_BioSum_Main()
        {

            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            this.m_intError = 0;
            m_intProgressOverallTotalCount = 0;
            m_intProgressStepCurrentCount = 0;
            m_strError = "";
            m_strWarning = "";
            m_intWarning = 0;
            m_intProgressOverallCurrentCount = 0;

            if (m_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

            System.DateTime oDate = System.DateTime.Now;
            string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
            m_strLogDate = oDate.ToString(strDateFormat);
            m_strLogDate = m_strLogDate.Replace("/", "_"); m_strLogDate = m_strLogDate.Replace(":", "_");

            try
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 10);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);

                if (m_intError == 0)
                {
                    {
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Ready to create FVS_OutBioSum.db");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");


                        m_strLogFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + $@"/fvs/data/FVSOut_BioSum_" + m_strLogDate.Replace(" ", "_") + ".txt";

                        frmMain.g_oUtils.WriteText(m_strLogFile, "CREATING FVSOUT_BIOSUM.DB \r\n");
                        frmMain.g_oUtils.WriteText(m_strLogFile, "--------- \r\n\r\n");
                        frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n");

                        string strBiosumDbPath = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutBiosumDbFile;

                        SQLite.CreateDbFile(strBiosumDbPath);
                        frmMain.g_oUtils.WriteText(m_strLogFile, "Created " + strBiosumDbPath + " file \r\n\r\n");

                        m_intProgressStepCurrentCount = 1;
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Querying table schemas from FVSOut.db");
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", m_intProgressStepCurrentCount);

                        string strConnection = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile;
                        Dictionary<string, string> dictCreateTableQueries = populateTableQueryDictionaries(SQLite.GetConnectionString(strConnection));
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", dictCreateTableQueries.Keys.Count + 5);

                        m_intProgressStepCurrentCount++;
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", m_intProgressStepCurrentCount);

                        using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strBiosumDbPath)))
                        {
                            con.Open();
                            string strSql = $@"ATTACH DATABASE '{strConnection}' AS SOURCE";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strLogFile, strSql + "\r\n");
                            SQLite.SqlNonQuery(con, strSql);
                            foreach (var tblName in dictCreateTableQueries.Keys)
                            {
                                strSql = dictCreateTableQueries[tblName];
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strLogFile, strSql + "\r\n");
                                SQLite.SqlNonQuery(con, dictCreateTableQueries[tblName]);
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Inserting records into " + tblName);
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", m_intProgressStepCurrentCount);
                                strSql = $@"INSERT INTO {tblName} SELECT * FROM SOURCE.{tblName}";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strLogFile, strSql + "\r\n");
                                m_intProgressStepCurrentCount++;
                                SQLite.SqlNonQuery(con, strSql);

                                if (tblName.ToUpper().Contains("FVS_") && !tblName.ToUpper().Equals("FVS_CASES"))
                                {
                                    if (! SQLite.ColumnExist(con, tblName, "RUNTITLE"))
                                    {
                                        string[] arrQueries = Queries.FVS.AppendRuntitleToFVSOut(tblName);
                                        for (int i = 0; i < arrQueries.Length; i++)
                                        {
                                            strSql = arrQueries[i];
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strLogFile, strSql + "\r\n");
                                            SQLite.SqlNonQuery(con, strSql);
                                        }
                                    }
                                }
                            }                            
                        }

                        m_oRxTools.CreateFvsOutDbIndexes(strBiosumDbPath);

                        frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n\r\n");
                        frmMain.g_oUtils.WriteText(m_strLogFile, "**EOF**");
                    }

                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 10);

                    this.FVSRecordsFinished();
                }
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                MessageBox.Show("Threading Abort Error " + err.Message.ToString());
                this.ThreadCleanUp();
                this.CleanupThread();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_output:RunCreateFVSOut_BioSum_Main  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }

            if (m_bDebug)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "****END*****" + System.DateTime.Now.ToString() + "\r\n");
            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }

        /// <summary>This method populates dictionaries storing table creation scripts and 
        /// for the new Access MDBs, and a dictionary of column definition information for each type of table
        /// within those Access MDBs. </summary>
        /// <param name="strConnection">The connection string for the source SQLite table.</param>
        private Dictionary<string, string> populateTableQueryDictionaries(string strConnection)
        {
            Dictionary<string, string> dictCreateTableQueries = new Dictionary<string, string>();
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                //getTableNames
                var tableNames = SQLite.getTableNames(con);
                //Build field list string to insert sql by matching up the column names in the biosum plot table and the fiadb plot table

                // Run this loop for each database we need to make.
                foreach (var tblName in tableNames)
                {
                    DataTable dtSourceSchema = SQLite.getTableSchema(con, $"select * from {tblName}");
                    var sb = new System.Text.StringBuilder();
                    var strCol = "";
                    sb.Append($@"CREATE TABLE {tblName} (");
                    var strFields = "";
                    // Iterate over table schema defined in Data Table format. Check debugger to see columns.
                    for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                    {
                        var colName = dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper();
                        var convertedColName = dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper();
                        var dataType = dtSourceSchema.Rows[y]["datatype"].ToString().ToUpper();
                        string convertedDataType = utils.DataTypeConvert(dataType, true);
                        if (m_SqliteToAccessColTypes.Keys.Contains(convertedColName))
                        {
                            convertedDataType = m_SqliteToAccessColTypes[convertedColName];
                        }                        
                        strCol = utils.WrapInBacktick(convertedColName) + " " + convertedDataType;
                        if (strFields.Trim().Length == 0)
                        {
                            strFields = strCol;
                        }
                        else
                        {
                            strFields += "," + strCol;
                        }
                    }

                    sb.Append(strFields + ") ");
                    // Populate the table create queries dict and the dictionary of output table columns and their access datatype.
                    dictCreateTableQueries[tblName] = sb.ToString();
                }
            }
            return dictCreateTableQueries;
        }

        private void RunFVSInForestTable_Start()
        {
            // Warning for older projects without FVSOut.db
            if (!System.IO.File.Exists(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile))
            {
                MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS IN FOREST TABLE",
                "Update FVS_InForest table", "2");
            m_frmTherm.Visible = false;
            this.m_frmTherm.lblMsg.Text = "";
            this.m_frmTherm.TopMost = true;

            this.cmbStep.Enabled = false;
            this.btnExecute.Enabled = false;
            this.btnChkAll.Enabled = false;
            this.btnClearAll.Enabled = false;
            this.btnRefresh.Enabled = false;
            this.btnClose.Enabled = false;
            this.btnHelp.Enabled = false;
            this.btnCancel.Visible = false;
            this.btnViewLogFile.Enabled = false;
            this.btnViewPostLogFile.Enabled = false;
            this.btnAuditDb.Enabled = false;
            this.btnPostAppendAuditDb.Enabled = false;

            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);
            m_frmTherm.Show((frmDialog)ParentForm);

            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(RunAppend_UpdateFVSInForestTable));
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void RunAppend_UpdateFVSInForestTable()
        {
            string strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_inforest_debug.txt";
            if (System.IO.File.Exists(strDebugFile))
                System.IO.File.Delete(strDebugFile);

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(strDebugFile, "//RunAppend_UpdateFVSInForestTable\r\n");
                frmMain.g_oUtils.WriteText(strDebugFile, "//\r\n");
            }

            try
            {

                this.m_strDateTimeCreated = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                bool bTreeList = false;
                // Keep track of whether there are trees to calculate metrics
                bool bRunFics = false;
                List<string> lstRunTitles = new List<string>();
                if (System.IO.File.Exists(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile))
            {
                string strConn = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    if (SQLite.TableExist(conn, "FVS_TREELIST"))
                    {
                        bTreeList = true;
                            SQLite.SqlQueryReader(conn, $@"select distinct(substr(RunTitle, 12, 3)) AS RXPACKAGE, RunTitle " +
                                "from FVS_SUMMARY, FVS_CASES where FVS_SUMMARY.CaseID = FVS_CASES.CaseID");
                            if (SQLite.m_DataReader.HasRows)
                            {
                                while (SQLite.m_DataReader.Read())
                                {
                                    lstRunTitles.Add(Convert.ToString(SQLite.m_DataReader["RunTitle"]));
                                }
                                lstRunTitles.Sort();
                            }
                            SQLite.m_DataReader.Close();
                        }
                }
            }
                // Only try to load if there is a tree list in the FVSOut.db
                m_intProgressOverallCurrentCount = 1;
                m_intProgressStepTotalCount = 9;
            if (bTreeList)
                {
                string strTreeTempDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                string strTreeListDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
                //copy the production file to the temp folder which will be used as the work db file.
                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                    this.WriteText(strDebugFile, "\r\nSTART:Copy production file to work file: Source File Name:" + strTreeListDbFile + " Destination File Name:" + strTreeTempDbFile + " " + System.DateTime.Now.ToString() + "\r\n");
                System.IO.File.Copy(strTreeListDbFile, strTreeTempDbFile, true);
                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                    this.WriteText(strDebugFile, "\r\nEND:Copy production file to work file: Source File Name:" + strTreeListDbFile + " Destination File Name:" + strTreeTempDbFile + " " + System.DateTime.Now.ToString() + "\r\n");
                    m_intProgressOverallTotalCount = lstRunTitles.Count + 2;
                    for (int i = 0; i < lstRunTitles.Count; i++)
                    {
                        // Example RunTitle: FVSOUT_BM_P001-101-101-101-101
                        string strFvsVariant = lstRunTitles[i].Substring(7, 2);
                        string strRxPackage = lstRunTitles[i].Substring(11, 3);
                        m_intProgressOverallCurrentCount++;
                        UpdateTherm(m_frmTherm.progressBar2, m_intProgressOverallCurrentCount, m_intProgressOverallTotalCount);
                        m_intProgressStepCurrentCount = 1;  // Reset current count for new package
                        UpdateTherm(m_frmTherm.progressBar1,
                                    m_intProgressStepCurrentCount,
                                    m_intProgressStepTotalCount);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"{strFvsVariant}{strRxPackage} Building sequence number matrix");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");


                        string strBaseYrRunTitle = $@"FVSOUT_{strFvsVariant}_POTFIRE_BaseYr";
                        CreatePrePostSeqNumMatrixSqliteTables(strTreeTempDbFile, lstRunTitles[i], strBaseYrRunTitle, false);

                        string strConn = SQLite.GetConnectionString(strTreeTempDbFile);
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            conn.Open();
                            // Attach FVSOut.db
                            SQLite.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                                Tables.FVS.DefaultFVSOutDbFile + "' AS fvsout";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                            if (SQLite.TableExist(conn, "cutlist_fia_trees_work_table"))
                                SQLite.SqlNonQuery(conn, "DROP TABLE cutlist_fia_trees_work_table");

                            if (SQLite.TableExist(conn, "cutlist_fvs_created_seedlings_work_table"))
                                SQLite.SqlNonQuery(conn, "DROP TABLE cutlist_fvs_created_seedlings_work_table");

                            if (!SQLite.TableExist(conn, Tables.FVS.DefaultFVSInForestTreeTableName))
                            {
                                frmMain.g_oTables.m_oFvs.CreateFVSInForestTable(SQLite, conn, Tables.FVS.DefaultFVSInForestTreeTableName);
                            }
                            else
                            {
                                SQLite.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultFVSInForestTreeTableName} WHERE FVS_VARIANT = '{strFvsVariant}' AND RXPACKAGE ='{strRxPackage}'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            }

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"{strFvsVariant}{strRxPackage} Gathering trees");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            //
                            //FIA TREES
                            //
                            //make sure there are records to insert
                            SQLite.m_strSQL = "SELECT COUNT(*) FROM " +
                                             "(SELECT c.standid " +
                                              "FROM fvsout.FVS_CASES c, fvsout.FVS_TreeList t " +
                                              "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + lstRunTitles[i] + "' AND " +
                                              "substr(t.treeid, 1, 2) NOT IN ('ES') LIMIT 1)";

                            string[] arrRx = lstRunTitles[i].Split('-');
                            string[] arrQueries = null;
                            if ((int)SQLite.getRecordCount(conn, SQLite.m_strSQL, "temp") > 0)
                            {
                                bRunFics = true;
                                SQLite.m_strSQL = "CREATE TABLE cutlist_fia_trees_work_table AS " +
                                    "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + strRxPackage + "' AS rxpackage," +
                                    "'' as rx, '' as rxcycle, '" + strFvsVariant + "' as fvs_variant, t.year as simYear," +
                                    "cast(t.year as text) as rxyear, Trim(t.treeid) AS fvs_tree_id," +
                                    "t.SpeciesFia AS fvs_species, t.TPA, t.DBH, t.Ht, t.pctcr," +
                                    "t.treeval, t.mortpa, t.mdefect, t.bapctile, t.dg, t.htg, " +
                                    "'N' AS FvsCreatedTree_YN," +
                                    "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                    "FROM fvsout.FVS_CASES c, fvsout.FVS_TreeList t " +
                                    "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + lstRunTitles[i] + "' AND SUBSTR(t.treeid, 1, 2) NOT IN ('ES') ";

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                                if (arrRx.Length == 5)
                                {
                                    arrQueries = Queries.VolumeAndBiomass.FVSOut.PopulateRxFields("cutlist_fia_trees_work_table",
                                        "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX", arrRx[1], arrRx[2], arrRx[3], arrRx[4]);
                                    if (arrQueries.Length == 2)
                                    {
                                        // Set RxCycle using FVS_SUMMARY_PREPOST_SEQNUM_MATRIX
                                        SQLite.m_strSQL = arrQueries[0];
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        //Set RX based on runtitle; Example: FVSOUT_BM_P001-101-101-101-101
                                        SQLite.m_strSQL = arrQueries[1];
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    }
                                }
                                //insert into fvs tree table
                                SQLite.m_strSQL = $@"INSERT INTO {Tables.FVS.DefaultFVSInForestTreeTableName} 
                                                     (biosum_cond_id, rxpackage,rx,rxcycle,rxyear, simYear,fvs_variant, fvs_tree_id,
                                                      fvs_species, tpa, dbh, ht, pctcr,
                                                      treeval, mortpa, mdefect, bapctile, dg, htg, FvsCreatedTree_YN,DateTimeCreated)
                                                      SELECT a.biosum_cond_id, a.rxpackage,a.rx,a.rxcycle,a.rxyear,a.simYear,a.fvs_variant,
                                                      a.fvs_tree_id, a.fvs_species, a.tpa, a.dbh, a.ht, a.pctcr,
                                                      a.treeval, a.mortpa, a.mdefect, a.bapctile, a.dg, a.htg,
                                                      a.FvsCreatedTree_YN,a.DateTimeCreated  
                                                      FROM cutlist_fia_trees_work_table a";

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }

                            //
                            //FVS CREATED SEEDLING TREES
                            //
                            //make sure there are records to insert
                            SQLite.m_strSQL = "SELECT COUNT(*) FROM " +
                                "(SELECT c.standid " +
                                "FROM fvsout.FVS_CASES c, fvsout.FVS_TreeList t " +
                                "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + lstRunTitles[i] + "' AND " +
                                "substr(t.treeid, 1, 2) = 'ES' LIMIT 1)";

                            if ((int)SQLite.getRecordCount(conn, SQLite.m_strSQL, "temp") > 0)
                            {
                                bRunFics = true;
                                //FVS CREATED SEEDLING TREES
                                SQLite.m_strSQL =
                                       "CREATE TABLE cutlist_fvs_created_seedlings_work_table AS " +
                                       "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + strRxPackage + "' AS rxpackage," +
                                       "'' AS rx,'' AS rxcycle, cast(t.year as text) as rxyear," +
                                        "t.year as simYear,'" +
                                       strFvsVariant + "' AS fvs_variant, " +
                                       "Trim(t.treeid) AS fvs_tree_id, " +
                                       "t.SpeciesFia AS fvs_species, t.TPA, t.DBH, t.Ht, t.pctcr, " +
                                       "t.treeval, t.mortpa, t.mdefect, t.bapctile, t.dg, t.htg, " +
                                       "'Y' AS FvsCreatedTree_YN," +
                                       "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                       "FROM fvsout.FVS_CASES c, fvsout.FVS_TreeList t " +
                                       "WHERE c.CaseID = t.CaseID AND c.RunTitle = '" + lstRunTitles[i] + "' AND substr(t.treeid, 1, 2) = 'ES'";

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                if (arrRx.Length == 5)
                                {
                                    arrQueries = Queries.VolumeAndBiomass.FVSOut.PopulateRxFields("cutlist_fvs_created_seedlings_work_table",
                                        "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX", arrRx[1], arrRx[2], arrRx[3], arrRx[4]);
                                    if (arrQueries.Length == 2)
                                    {
                                        // Set RxCycle using FVS_SUMMARY_PREPOST_SEQNUM_MATRIX
                                        SQLite.m_strSQL = arrQueries[0];
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        //Set RX based on runtitle; Example: FVSOUT_BM_P001-101-101-101-101
                                        SQLite.m_strSQL = arrQueries[1];
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    }
                                }

                                SQLite.m_strSQL = $@"INSERT INTO {Tables.FVS.DefaultFVSInForestTreeTableName} 
                                                    (biosum_cond_id, rxpackage,rx,rxcycle,rxyear,simYear,fvs_variant, fvs_tree_id,
                                                     fvs_species, tpa, dbh, ht, pctcr,
                                                     treeval, mortpa, mdefect, bapctile, dg, htg,
                                                     FvsCreatedTree_YN,DateTimeCreated) 
                                                     SELECT a.biosum_cond_id, a.rxpackage,a.rx,a.rxcycle,a.rxYear,a.simYear,a.fvs_variant,
                                                     a.fvs_tree_id, a.fvs_species, a.tpa, a.dbh, a.ht, a.pctcr,
                                                     a.treeval, a.mortpa, a.mdefect, a.bapctile, a.dg, a.htg, 
                                                     a.FvsCreatedTree_YN,a.DateTimeCreated  
                                                     FROM cutlist_fvs_created_seedlings_work_table a";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }

                            if (SQLite.TableExist(conn, "cutlist_fia_trees_work_table"))
                                SQLite.SqlNonQuery(conn, "DROP TABLE cutlist_fia_trees_work_table");

                            if (SQLite.TableExist(conn, "cutlist_fvs_created_seedlings_work_table"))
                                SQLite.SqlNonQuery(conn, "DROP TABLE cutlist_fvs_created_seedlings_work_table");

                            if (SQLite.TableExist(conn, "cutlist_save_tree_species_work_table"))
                                SQLite.SqlNonQuery(conn, "DROP TABLE cutlist_save_tree_species_work_table");

                            //
                            //SWAP FVS-ONLY SPECIES CODES TO FIA SPECIES CODES IN ORDER FOR FIA TO CALCULATE VOLUMES
                            //check if original spcd column exists
                            //save the original spcd
                            this.SaveTreeSpeciesWorkTable(conn, Tables.FVS.DefaultFVSInForestTreeTableName, strRxPackage);

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"{strFvsVariant}{strRxPackage} Updating values in biosum_volumes_input");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            if (SQLite.TableExist(conn, Tables.VolumeAndBiomass.BiosumVolumesInputTable))
                                SQLite.SqlNonQuery(conn, "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                            frmMain.g_oTables.m_oFvs.CreateSQLiteInputBiosumVolumesTable(SQLite, conn, Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step1(
                                               Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                               Tables.FVS.DefaultFVSInForestTreeTableName, strRxPackage, strFvsVariant);

                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            SQLite.m_strSQL =
                                Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step1a(Tables.VolumeAndBiomass.BiosumVolumesInputTable, lstRunTitles[i]);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        } // Closing SQLite connection in preparation to interact with Access tables

                        if (bRunFics == true)
                        {
                            System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strTreeTempDbFile));
                            oConn.Open();
                            // Attach master.db
                            SQLite.m_strSQL = $@"attach '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree)}' as master";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            // Attach biosum_ref.db
                            SQLite.m_strSQL = $@"attach '{m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.FiaTreeSpeciesReference)}' as ref";
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
              
                            //NOTE: this query handles existing FIADB trees that have been grown forward.
                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2(
                                               Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                               m_oQueries.m_oFIAPlot.m_strTreeTable,
                                               m_oQueries.m_oFIAPlot.m_strPlotTable,
                                               m_oQueries.m_oFIAPlot.m_strCondTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            //Set DIAHTCD for FIADB Cycle<>1 trees to their Cycle=1 DIAHTCD values
                            //Set standing_dead_code from tree table on FIADB trees
                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFiadb(m_oQueries.m_oFIAPlot.m_strTreeTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            //Set DIAHTCD for FVS-Created trees using FIA_TREE_SPECIES_REF.WOODLAND_YN
                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFvs(Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            //Update FVSCreatedTrees precipitation=plot.precipitation
                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step2a(
                                               Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                               m_oQueries.m_oFIAPlot.m_strPlotTable,
                                               m_oQueries.m_oFIAPlot.m_strCondTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            //Start cull calculations
                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step3(
                                              Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                              m_oQueries.m_oFIAPlot.m_strCondTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            //populate treeclcd column
                            if (SQLite.TableExist(oConn, "CULL_TOTAL_WORK_TABLE"))
                                SQLite.SqlNonQuery(oConn, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step4(
                                              "cull_total_work_table",
                                              Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                                                "cull_total_work_table",
                                                Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                                              "cull_total_work_table",
                                              Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(oConn, SQLite.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            // Delete work table
                            if (SQLite.TableExist(oConn, "CULL_TOTAL_WORK_TABLE"))
                                SQLite.SqlNonQuery(oConn, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepCurrentCount,
                                m_intProgressStepTotalCount);

                            // CLOSE CONNECTION
                            if (oConn.State == ConnectionState.Open)
                            {
                                oConn.Close();
                            }

                            if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase) == false)
                            {
                                m_intError = -1;
                                m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase + " not found";
                            }
                            if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR") == false)
                            {
                                m_intError = -1;
                                m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR not found";
                            }
                            if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat") == false)
                            {
                                m_intError = -1;
                                m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat not found";
                            }
                            if (m_intError == 0)
                            {
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"{strFvsVariant}{strRxPackage} Preparing to run FICS");
                                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                                //
                                //RE-CONNECT TO SQLITE AND REMOVE DATA FROM FCS SQLITE DB
                                //
                                SQLite.OpenConnection(false, 1, strTreeTempDbFile, "BIOSUM");
                                SQLite.SqlNonQuery(SQLite.m_Connection, "ATTACH DATABASE '" + frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase +
                                     "' AS FCS");
                                SQLite.SqlNonQuery(SQLite.m_Connection, $"DELETE FROM FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");
                                System.Threading.Thread.Sleep(2000);

                                //insert records 
                                //from 
                                //table biosum_volumes_input (BiosumVolumesInputTable)
                                //into 
                                //table fcs_biosum_volumes_input (FcsBiosumVolumesInputTable)
                                SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteBiosumCalcTable_Step7(
                                            Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                            $"FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(SQLite.m_Connection, SQLite.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //
                                //RUN JAVA APP TO SEND TO ORACLE AND CALCULATE VOLUME/BIOMASS
                                //
                                if (m_intError == 0)
                                {
                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                                m_intProgressStepCurrentCount,
                                                m_intProgressStepTotalCount);
                                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"{strFvsVariant}{strRxPackage} Running FICS");
                                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                                    frmMain.g_oUtils.RunProcess(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum", "fcs_tree_calc.bat", "BAT");
                                    if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt"))
                                    {
                                        // Read entire text file content in one string  
                                        m_strError = System.IO.File.ReadAllText(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_error_msg.txt");
                                        if (m_strError.IndexOf("JAVA.EXE", 0) > 0)
                                            m_strError = "Problem detected running JAVA.EXE";
                                        m_intError = -2;
                                    }
                                }
                                //
                                //UPDATE fvs_InForestTree table WITH CALCULATED VALUES
                                //
                                if (m_intError == 0)
                                {
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n SQLite connection status: " + SQLite.m_Connection.State + "\r\n");
                                    bool bAttached = SQLite.DatabaseAttached(SQLite.m_Connection,
                                        frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase);
                                    if (!bAttached)
                                    {
                                        SQLite.SqlNonQuery(SQLite.m_Connection, "ATTACH DATABASE '" + frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase +
                                            "' AS FCS");
                                    }

                                    //update calculated fields 
                                    //from 
                                    //biosum_calc table
                                    //into 
                                    //table fvs_tree
                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                                m_intProgressStepCurrentCount,
                                                m_intProgressStepTotalCount);
                                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"{strFvsVariant}{strRxPackage} Copying FICS values to FVS_InForest");
                                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                                    SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step9(
                                        Tables.FVS.DefaultFVSInForestTreeTableName, Tables.VolumeAndBiomass.BiosumVolumeCalcTable,
                                        strFvsVariant, strRxPackage);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(SQLite.m_Connection, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    // This query is customized because there are some fields missing from inForest that are in fvs_cutTree
                                    SQLite.m_strSQL = $@"UPDATE {Tables.FVS.DefaultFVSInForestTreeTableName} 
                                                        SET (volcfnet, volcfsnd, volcsgrs, voltsgrs) 
                                                        = (select volcfnet_calc, volcfsnd_calc, volcsgrs_calc, voltsgrs_calc
                                                        FROM FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}
                                                        WHERE {Tables.FVS.DefaultFVSInForestTreeTableName}.id = FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}.tree)
                                                        WHERE fvs_variant = '{strFvsVariant}' and rxpackage = '{strRxPackage}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(SQLite.m_Connection, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    SQLite.m_strSQL = "DETACH DATABASE 'FCS'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(SQLite.m_Connection, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    SQLite.CloseAndDisposeConnection(SQLite.m_Connection, true);
                                }
                            }
                        }
                        if (SQLite.m_intError == 0 && m_intError == 0)
                        {
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strTreeTempDbFile)))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, "cutlist_save_tree_species_work_table"))
                                {
                                    //update fvs_species
                                    SQLite.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step11(Tables.FVS.DefaultFVSInForestTreeTableName, strRxPackage);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    SQLite.SqlNonQuery(conn, "DROP TABLE cutlist_save_tree_species_work_table");
                                }

                                // DELETE WORK TABLES
                                string[] arrDeleteTables = new string[] { Tables.VolumeAndBiomass.BiosumVolumesInputTable, "cutlist_rowid_work_table"};
                                foreach (var dTable in arrDeleteTables)
                                {
                                    if (SQLite.TableExist(conn, dTable))
                                    {
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + "DROP TABLE " + dTable + "\r\n");
                                        SQLite.SqlNonQuery(conn, "DROP TABLE " + dTable);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    }
                                }

                                // DELETE SEQUENCE NUMBER MATRICES
                                string[] arrTables = SQLite.getTableNames(conn);
                                foreach (var tName in arrTables)
                                {
                                    if (tName.ToUpper().IndexOf("SEQNUM_MATRIX") > -1)
                                    {
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + "DROP TABLE " + tName + "\r\n");
                                        SQLite.SqlNonQuery(conn, "DROP TABLE " + tName);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    }
                                }
                            }
                        }
                    }

                    // end of work                               
                    //copy the work db file over the production file
                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                    this.WriteText(strDebugFile, "\r\nSTART:Copy work file to production file: Source File Name:" + strTreeTempDbFile + " Destination File Name:" + strTreeListDbFile + " " + System.DateTime.Now.ToString() + "\r\n");
                System.IO.File.Copy(strTreeTempDbFile, strTreeListDbFile, true);
                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                    this.WriteText(strDebugFile, "\r\nEND:Copy work file to production file: Source File Name:" + strTreeTempDbFile + " Destination File Name:" + strTreeListDbFile + " " + System.DateTime.Now.ToString() + "\r\n");

                    UpdateTherm(m_frmTherm.progressBar1, m_intProgressStepTotalCount, m_intProgressStepTotalCount);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", $@"Finalizing database");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                }
                else
                {
                    this.FVSRecordsFinished();
                    MessageBox.Show($@"The FVSOut.db for this project does not contain an FVS_TreeList table. The {Tables.FVS.DefaultFVSInForestTreeTableName} table cannot be created!", "FIA Biosum");
                }
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                this.ThreadCleanUp();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_output:RunAppend_UpdateFVSInForestTable  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }

            this.FVSRecordsFinished();
            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }

        private void CreatePrePostSeqNumMatrixSqliteTables(string p_strDbFile, string p_strRunTitle, string p_strBaseYrRunTitle, bool p_bAudit)
        {
            int z;
            string[] strSourceTableArray = new string[0];
            string strFvsOutDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strFvsOutDb)))
            {
                conn.Open();
                // Create index on FVS_CASES table
                if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTableName))
                {
                    if (!SQLite.IndexExist(conn, "idx_fvs_cases_bs"))
                    {
                        SQLite.AddIndex(conn, Tables.FVS.DefaultFVSCasesTableName, "idx_fvs_cases_bs", "runtitle, caseid, standid");
                    }
                    if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTempTableName))
                    {
                        // Delete current variant/package from table
                        SQLite.m_strSQL = $@"DROP TABLE {Tables.FVS.DefaultFVSCasesTempTableName}";
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    }

                    // Create temp FVS_CASES table that only contains records for current runTitle
                    // Do we have any records from a base year run to include?
                    SQLite.m_strSQL = $@"SELECT CAST(SamplingWt AS INT) FROM {Tables.FVS.DefaultFVSCasesTableName} WHERE RUNTITLE = '{p_strBaseYrRunTitle}' LIMIT 1";
                    long lngBaseYrCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCasesTableName);
                    string strRunTitleClause = $@"'{p_strRunTitle}'";
                    if (lngBaseYrCount > 0)
                    {
                        strRunTitleClause = $@"{strRunTitleClause},'{p_strBaseYrRunTitle}'";
                    }
                    SQLite.m_strSQL = $@"CREATE TABLE {Tables.FVS.DefaultFVSCasesTempTableName} AS SELECT *
                        FROM {Tables.FVS.DefaultFVSCasesTableName} WHERE RunTitle in ({strRunTitleClause})";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                    if (SQLite.m_intError == 0)
                    {
                        string strIndexName = $@"IDX_{Tables.FVS.DefaultFVSCasesTempTableName}_CaseId";
                        SQLite.AddIndex(conn, Tables.FVS.DefaultFVSCasesTempTableName, strIndexName, "CaseId");
                    }
                    strSourceTableArray = GetValidFVSTables(conn, Tables.FVS.DefaultFVSCasesTempTableName, p_strRunTitle);
                }

                CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, "FVS_SUMMARY", p_strRunTitle, p_strBaseYrRunTitle, p_bAudit);
                CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, "FVS_CUTLIST", p_strRunTitle, p_strBaseYrRunTitle, p_bAudit);
                // Attach the audits.db so we have access to the POTFIRE_TEMP table; supports BaseYr
                if (!SQLite.DatabaseAttached(conn, p_strDbFile))
                {
                    SQLite.m_strSQL = $@"ATTACH DATABASE '{p_strDbFile}' AS AUDITS";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                }
                CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, "FVS_POTFIRE_TEMP", p_strRunTitle, p_strBaseYrRunTitle, p_bAudit);

                for (z = 0; z <= strSourceTableArray.Length - 1; z++)
                {
                    if (strSourceTableArray[z] == null) break;
                    if (strSourceTableArray[z].Trim().ToUpper() != "FVS_SUMMARY" &&
                        strSourceTableArray[z].Trim().ToUpper() != "FVS_CUTLIST" &&
                        strSourceTableArray[z].Trim().ToUpper() != "FVS_POTFIRE")
                    {
                       string strRxPackage = p_strRunTitle.Substring(11, 3);
                       CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, strSourceTableArray[z], p_strRunTitle, p_strBaseYrRunTitle, p_bAudit);
                    }
                }
            }
        }

        private string[] GetValidFVSTables(System.Data.SQLite.SQLiteConnection p_oConn, string strFvsCasesTable, string strRunTitle)
        {
            // Make sure the source table is valid and has records for this case
            IList<string> lstValidSourceTables = new List<string>();
            string[] strSourceTableArray = SQLite.getTableNames(p_oConn);
            for (int z = 0; z <= strSourceTableArray.Length - 1; z++)
            {
                if (RxTools.ValidFVSTable(strSourceTableArray[z]))
                {
                    string strTable = strSourceTableArray[z].ToUpper();
                    if (!strTable.Equals(Tables.FVS.DefaultFVSCasesTableName))
                    {
                        if (SQLite.ColumnExist(p_oConn, strTable, "CaseId"))
                        {
                            string strSql = $@"SELECT s.year from {strTable} s, {strFvsCasesTable} c where s.CaseID = c.CaseID and s.standId = c.StandId 
                                AND RUNTITLE = '{strRunTitle}' LIMIT 1";
                            long lngCount = SQLite.getRecordCount(p_oConn, strSql, strTable);
                            if (lngCount > 0)
                            {
                                lstValidSourceTables.Add(strTable);
                            }
                        }
                    }
                }
            }
            strSourceTableArray = lstValidSourceTables.ToArray();
            return strSourceTableArray;
        }

        private void CreateFVSPrePostSeqNumWorkTables(System.Data.SQLite.SQLiteConnection conn, string p_strTreeTempDbFile, string p_strSourceTableName, 
            string p_strRunTitle, string p_strBaseYrRunTitle, bool p_bAudit)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateFVSPrePostSeqNumWorkTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Process Table:" + p_strSourceTableName + "\r\n\r\n");
            }

            if (p_strSourceTableName.Trim().ToUpper() == "FVS_CASES") return;

            if (SQLite.TableExist(conn, p_strSourceTableName) || SQLite.AttachedTableExist(conn,p_strSourceTableName))
            {
                if (m_oFVSPrePostSeqNumItemCollection == null) m_oFVSPrePostSeqNumItemCollection = new FVSPrePostSeqNumItem_Collection();
                string strParamConn = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile);
                m_oRxTools.LoadFVSOutputPrePostRxCycleSeqNum(strParamConn, m_oFVSPrePostSeqNumItemCollection);

                if (!p_strSourceTableName.Trim().ToUpper().Equals("FVS_POTFIRE_TEMP"))
                {
                    string strIdxName = $@"idx_fvs_{p_strSourceTableName}_bs";
                    if (!SQLite.IndexExist(conn, strIdxName))
                    {
                        if (p_strSourceTableName.Trim().ToUpper().Equals("FVS_STRCLASS"))
                        {
                            SQLite.AddIndex(conn, p_strSourceTableName, strIdxName, "caseid, standid, year, removal_code");
                        }
                        else
                        {
                            SQLite.AddIndex(conn, p_strSourceTableName, strIdxName, "caseid, standid, year");
                        }
                    }
                }

                string tmpTableName = "tmpSequence";
                if (SQLite.TableExist(conn, tmpTableName))
                {
                    SQLite.SqlNonQuery(conn, $@"DROP TABLE {tmpTableName}");
                }
                // Source table for tmpSequence has a variety of dependencies. FVS_CUTLIST always uses FVS_SUMMARY:
                string strRunTitleClause = $@"'{p_strRunTitle}'";
                string strActualPotfireTableName = GetPotfireTableName(m_strFvsOutDb, p_strRunTitle.Substring(7, 2));
                if (p_strSourceTableName.ToUpper().Equals("FVS_POTFIRE_TEMP"))
                {
                    GetPrePostSeqNumConfiguration(strActualPotfireTableName, p_strRunTitle.Substring(11, 3));
                    // Only include BaseYr in the runTitle clause if sequence number using BaseYr
                    if (m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN == "Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN == "Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN == "Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN == "Y")
                    {
                        strRunTitleClause = $@"{strRunTitleClause},'{p_strBaseYrRunTitle}'";
                    }
                }
                else
                {
                    GetPrePostSeqNumConfiguration(p_strSourceTableName, p_strRunTitle.Substring(11, 3));
                }
                
                string strSourceTable = Tables.FVS.DefaultFVSSummaryTableName;
                if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN != "Y")
                {
                    strSourceTable = p_strSourceTableName.Trim();
                }

                SQLite.m_strSQL = $@"CREATE TABLE {tmpTableName} AS select s.standid, year, '{p_strRunTitle.Substring(7, 2)}' as fvs_variant, '{p_strRunTitle.Substring(11, 3)}' as rxpackage from 
                    {strSourceTable} s, {Tables.FVS.DefaultFVSCasesTempTableName} c where s.CaseID = c.CaseID and runTitle in ({strRunTitleClause})";
                // Need to add removal_code if working with FVS_STRCLASS
                if (strSourceTable.ToUpper().Equals("FVS_STRCLASS"))
                {
                    SQLite.m_strSQL = $@"CREATE TABLE {tmpTableName} AS select s.standid, year, removal_code, '{p_strRunTitle.Substring(7, 2)}' as fvs_variant, '{p_strRunTitle.Substring(11, 3)}' as rxpackage from 
                    {strSourceTable} s, {Tables.FVS.DefaultFVSCasesTempTableName} c where s.CaseID = c.CaseID";
                }
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                m_oRxTools.CreateFVSPrePostSeqNumTables(p_strTreeTempDbFile, m_oFVSPrePostSeqNumItem, p_strSourceTableName, p_strSourceTableName, 
                    p_bAudit, m_bDebug, m_strDebugFile, p_strRunTitle, strActualPotfireTableName, tmpTableName);

                // Delete temp table
                SQLite.m_strSQL = $@"DROP TABLE {tmpTableName}";
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
            }
            else
            {
                if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, p_strSourceTableName + " table does not exist.\r\n\r\n");
                }
            }
        }

        private void RunAppend_Main()
        {
            string strAuditDbFile;
            string strCurVariant = "";
            bool bUpdateCondTable = false;
            int intCount;
            string strRx1 = "";
            string strRx2 = "";
            string strRx3 = "";
            string strRx4 = "";
            string strPackage = "";
            string strVariant = "";

            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;

            Tables oTables = new Tables();

            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            int x, y;
            m_intProgressOverallTotalCount = 0;
            m_intProgressStepCurrentCount = 0;
            m_intProgressOverallCurrentCount = 0;

            try
            {
                if (System.IO.File.Exists(m_strDebugFile))
                    System.IO.File.Delete(m_strDebugFile);

                if (m_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

                intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);

                //inititalize the run status column for each row to blank in the list view
                for (x = 0; x <= intCount - 1; x++)
                {
                    oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                    //alternate the row color in the list view
                    m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = true;
                    m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                    //inititalize the run status column for each row to blank in the list view
                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
                    //see if checked
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false))
                        this.m_intProgressOverallTotalCount++;

                }

                if (m_oFVSPrePostSeqNumItemCollection == null) m_oFVSPrePostSeqNumItemCollection = new FVSPrePostSeqNumItem_Collection();
                string strDbConn = SQLite.GetConnectionString($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");
                m_oRxTools.LoadFVSOutputPrePostRxCycleSeqNum(strDbConn, m_oFVSPrePostSeqNumItemCollection);

                m_intProgressOverallCurrentCount = 0;
                this.m_intProgressOverallTotalCount += 2;

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);

                this.val_data();

                // Ensure that indexes are on the FVSOut.db
                m_oRxTools.CreateFvsOutDbIndexes(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile);

                if (m_intError == 0)
                {

                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "check point 1 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "checkpoint 2 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                    m_intProgressStepTotalCount = Tables.FVS.g_strFVSOutTablesArray.Length;
                    m_intProgressStepCurrentCount = 0;
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);

                    //
                    //backup prepost files
                    //
                    System.DateTime oDate = System.DateTime.Now;
                    string strDateFormat = "yyyy-MM-dd_HH-mm-ss";
                    string strFileDate = oDate.ToString(strDateFormat);
                    strFileDate = strFileDate.Replace("/", "_"); strFileDate = strFileDate.Replace(":", "_");
                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepCurrentCount,
                                m_intProgressStepTotalCount);

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepTotalCount,
                                m_intProgressStepTotalCount);

                    m_intProgressOverallCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar2,
                                m_intProgressOverallCurrentCount,
                                m_intProgressOverallTotalCount);

                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "checkpoint 5 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                    this.m_strDateTimeCreated = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                    m_intProgressStepTotalCount = intCount;

                    string strTempDb = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                    for (x = 0; x <= intCount - 1; x++)
                    {
                        m_intProgressStepCurrentCount = 0;
                        oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                        this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);

                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);

                        if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false) == true)
                        {
                            int intItemError = 0;
                            string strItemError = "";

                            m_intProgressStepTotalCount = 20;

                            //make sure the list view item is selected and visible to the user
                            frmMain.g_oDelegate.ExecuteControlMethodWithParam((System.Windows.Forms.Control)oLv, "EnsureVisible", new object[] { x });
                            frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Selected", true);
                            frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Focused", true);

                            this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = false;
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkGoldenrod);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing");

                            //get the variant
                            strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_VARIANT, "Text", false);
                            strVariant = strVariant.Trim();

                            //get the package and treatments
                            strPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_PACKAGE, "Text", false);
                            strPackage = strPackage.Trim();

                            //find the package item in the package collection
                            for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                            {
                                if (this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
                                    break;
                            }
                            if (y <= m_oRxPackageItem_Collection.Count - 1)
                            {
                                this.m_oRxPackageItem = new RxPackageItem();
                                m_oRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), m_oRxPackageItem);
                            }
                            else
                            {
                                this.m_oRxPackageItem = null;
                            }
                            //get the list of treatment cycle year fields to reference for this package
                            this.m_strRxCycleList = "";
                            if (m_oRxPackageItem.SimulationYear1Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear1Rx.Trim() != "000") this.m_strRxCycleList = "1,";
                            if (m_oRxPackageItem.SimulationYear2Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear2Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                            if (m_oRxPackageItem.SimulationYear3Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear3Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                            if (m_oRxPackageItem.SimulationYear4Rx.Trim().Length > 0 && m_oRxPackageItem.SimulationYear4Rx.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                            if (this.m_strRxCycleList.Trim().Length > 0)
                                this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

                            this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");
                            strRx1 = m_oRxPackageItem.SimulationYear1Rx.Trim();
                            strRx2 = m_oRxPackageItem.SimulationYear2Rx.Trim();
                            strRx3 = m_oRxPackageItem.SimulationYear3Rx.Trim();
                            strRx4 = m_oRxPackageItem.SimulationYear4Rx.Trim();

                            //see if this is a different variant
                            //only update the pre-treatment tables when the variant changes
                            if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
                            {
                                bUpdateCondTable = true;
                                strCurVariant = strVariant;
                            }

                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "strOutDirAndFile=" + m_strFvsOutDb + "  \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Processing " + m_strFvsOutDb + "...Stand By");
                            strAuditDbFile = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}{Tables.FVS.DefaultFVSAuditsDbFile}";

                            //
                            //CREATE PREPOST SEQNUM MATRIX TABLES
                            //
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepCurrentCount,
                                         m_intProgressStepTotalCount);

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create POTFIRE table");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nSTART:Create POTFIRE tables " + System.DateTime.Now.ToString() + "\r\n");
                            CreatePotFireTables(strAuditDbFile, m_strFvsOutDb, strVariant, strPackage);
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepCurrentCount,
                                         m_intProgressStepTotalCount);
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nEND:Create POTFIRE tables " + System.DateTime.Now.ToString() + "\r\n");

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create PREPOST SeqNum Matrix Tables");
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nSTART:Create PrePostSeqNumMatrixTables " + System.DateTime.Now.ToString() + "\r\n");
                            string strRunTitle = $@"FVSOUT_{strVariant}{m_oRxPackageItem.RunTitleSuffix}";
                            string strBaseYrRunTitle = $@"FVSOUT_{strVariant}_POTFIRE_BaseYr";
                            CreatePrePostSeqNumMatrixSqliteTables(strAuditDbFile, strRunTitle, strBaseYrRunTitle, false);
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nEND:Create PrePostSeqNumMatrixTables " + System.DateTime.Now.ToString() + "\r\n");
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepCurrentCount,
                                         m_intProgressStepTotalCount);

                            //create treelist DB if it does not exist
                            m_oRxTools.CheckCutListDbExist();
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepTotalCount,
                                         m_intProgressStepTotalCount);

                            //
                            //SHOW FILE MONITORS
                            //

                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "checkpoint 6 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                            if (uc_filesize_monitor1.File.Trim().Length == 0)
                            {
                                uc_filesize_monitor1.BeginMonitoringFile(strTempDb, 2000000000, "2GB");
                                uc_filesize_monitor1.Information = strTempDb;
                            }

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            //
                            //validation
                            //
                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Validate data");
                                m_intProgressStepCurrentCount = 0;
                            }

                            m_intProgressStepCurrentCount = 0;
                            //
                            //Table Structure Checks And Edits
                            //
                            if (intItemError == 0)
                            {

                                m_intProgressStepCurrentCount = 0;
                                RunAppend_UpdatePrePostTable(
                                    strTempDb, strPackage, strVariant, strRx1, strRx2, strRx3, strRx4,
                                    bUpdateCondTable, x,
                                    ref intItemError, ref strItemError, strRunTitle);

                                if (intItemError == 0)
                                {
                                    intItemError = SQLite.m_intError;
                                    strItemError = SQLite.m_strError;
                                }


                            }

                            //
                            //update FVSTree table
                            //
                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Update " + m_strFvsTreeTable + " table");

                                // Only try to load if there is a cut list in the FVSOut.db
                                if (FvsOutWithRequiredTable(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(), "FVS_CutList"))
                                {
                                    string strTreeTempDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
                                    string strTreeListDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
                                    //copy the production file to the temp folder which will be used as the work db file.
                                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                        this.WriteText(m_strDebugFile, "\r\nSTART:Copy production file to work file: Source File Name:" + strTreeListDbFile + " Destination File Name:" + strTreeTempDbFile + " " + System.DateTime.Now.ToString() + "\r\n");
                                    System.IO.File.Copy(strTreeListDbFile, strTreeTempDbFile, true);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                        this.WriteText(m_strDebugFile, "\r\nEND:Copy production file to work file: Source File Name:" + strTreeListDbFile + " Destination File Name:" + strTreeTempDbFile + " " + System.DateTime.Now.ToString() + "\r\n");
                                    RunAppend_UpdateFVSTreeTables(strPackage, strVariant, strRx1, strRx2, strRx3, strRx4, strTreeTempDbFile, ref intItemError, ref strItemError);
                                    //copy the work db file over the production file
                                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                        this.WriteText(m_strDebugFile, "\r\nSTART:Copy work file to production file: Source File Name:" + strTreeTempDbFile + " Destination File Name:" + strTreeListDbFile + " " + System.DateTime.Now.ToString() + "\r\n");
                                    System.IO.File.Copy(strTreeTempDbFile, strTreeListDbFile, true);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                        this.WriteText(m_strDebugFile, "\r\nEND:Copy work file to production file: Source File Name:" + strTreeTempDbFile + " Destination File Name:" + strTreeListDbFile + " " + System.DateTime.Now.ToString() + "\r\n");

                                }
                                UpdateTherm(m_frmTherm.progressBar1,
                                                m_intProgressStepTotalCount,
                                                m_intProgressStepTotalCount);
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Finalizing updates");
                                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                                System.Threading.Thread.Sleep(1000);
                            }

                            //
                            //clean up for this list item
                            //
                            if (intItemError == 0)
                            {
                                string dbConn = SQLite.GetConnectionString(m_strFvsOutDb);
                                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                                {
                                    conn.Open();
                                    SQLite.m_strSQL = $@"UPDATE FVS_CASES SET {m_colBioSumAppend} ='Y' WHERE RunTitle = '{strRunTitle}'";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                frmMain.g_oDelegate.SetListViewTextValue(
                                    oLv, x, COL_CHECKBOX, Convert.ToString(frmMain.g_oDelegate.GetListViewTextValue(oLv, x, COL_CHECKBOX, false).Replace("a", "")));
                                frmMain.g_oDelegate.SetListViewTextValue(
                                    oLv, x, COL_CHECKBOX, Convert.ToString(frmMain.g_oDelegate.GetListViewTextValue(oLv, x, COL_CHECKBOX, false).Replace("p", "")));
                            }

                            UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepTotalCount,
                                           m_intProgressStepTotalCount);


                            if (intItemError == 0)
                            {
                                intItemError = SQLite.m_intError;
                                strItemError = SQLite.m_strError;
                            }

                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkGreen);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Completed");
                            }
                            else if (intItemError != 0)
                            {
                                m_intError = intItemError;
                                m_intOverallError = m_intOverallError += 1;
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Red);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "ERROR:" + strItemError);
                            }
                            m_intProgressOverallCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar2,
                                    m_intProgressOverallCurrentCount,
                                    m_intProgressOverallTotalCount);
                        }

                    }
                }

                string strFvsOutConn = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile);
                // Delete temporary FVS_CASES_TEMP
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strFvsOutConn))
                {
                    conn.Open();
                    if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTempTableName))
                    {
                        SQLite.m_strSQL = $@"DROP TABLE {Tables.FVS.DefaultFVSCasesTempTableName}";
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    }
                }


                UpdateTherm(m_frmTherm.progressBar1,
                           m_intProgressStepTotalCount,
                           m_intProgressStepTotalCount);
                UpdateTherm(m_frmTherm.progressBar2,
                           m_intProgressOverallTotalCount,
                          m_intProgressOverallTotalCount);
                System.Threading.Thread.Sleep(2000);
                if (m_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");
                this.FVSRecordsFinished();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {

                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                this.ThreadCleanUp();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_output:RunAppend_Main  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Ready");

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

        }
    }
}

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
		public run_PostFvsForeFrcs procPreFrcs=null;
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
		public string m_strError="";
		public int m_intWarning=0;
		public string m_strWarning="";
        private string m_strDateTimeCreated = "";
		public System.Windows.Forms.ListView lstFvsOutput;

		private const int COL_RX = -1;
		private const int COL_CHECKBOX=0;
		private const int COL_VARIANT = 1;
	    private const int COL_PACKAGE = 2;
		private const int COL_RXCYCLE1=3;
		private const int COL_RXCYCLE2=4;
		private const int COL_RXCYCLE3=5;
		private const int COL_RXCYCLE4=6;
		private const int COL_RUNSTATUS = 7;
		private const int COL_MDBOUT = 8;
		private const int COL_FOUND=9;
		private const int COL_SUMMARYCOUNT = 10;
		//private const int COL_SUMMARYPREYEAR=6;
		//private const int COL_SUMMARYPOSTYEAR=7;
		private const int COL_CUTCOUNT = 11;
		private const int COL_LEFTCOUNT = 12;
		private const int COL_POTFIRECOUNT = 13;
        private const int COL_POTFIREMDBOUT = 14;
        private const int COL_POTFIREMDBFOUND = 15;
        private const int COL_POTFIREBASEYEARCOUNT = 16;

		//parse FVSOUT file name
		const int VARIANT_POS = 7;
		const int PACKAGE_POS = 11;
		const int RX1_POS = 15;
		const int RX2_POS = 19;
		const int RX3_POS = 23;
		const int RX4_POS = 27;


		private string m_strFFETable="";
		//private string m_strVariant;
		private string m_strOutMDBFile="";
        private string m_strAuditDbFile = "";
		private string m_strRxTable="";
		private string m_strPlotTable="";
		private string m_strCondTable="";
		private string m_strFvsTreeTable="";
		private string m_strFVSSummaryAuditYearCountsTable="audit_fvs_summary_year_counts_table";
        private string m_strFVSSummaryAuditPrePostSeqNumTable = "fvs_summary_prepost_seqnum_matrix";
        private string m_strFVSSummaryAuditPrePostSeqNumCountsTable = "audit_fvs_summary_prepost_seqnum_counts_table";
        private string m_strFVSTreeIdAuditTable = "audit_fvs_tree_id";
        private string m_strFVSTreeMissingVolumeBiomassValuesTable = "audit_fvs_tree_missing_volume_biomass_values_table";
		private string[] m_strFVSPrePostYearAuditTablesArray = null;
        private string[] m_strFVSPrePostSeqNumTablesArray = null;
        private List<string> m_strFVSPreAppendAuditTables = null;
        private List<string> m_strFVSPostAppendAuditTables = null;
		private Datasource m_DataSource;
		private ado_data_access m_ado;
		private dao_data_access m_dao;
		private string m_strTempMDBFileConnectionString="";
		private string m_strTempMDBFile="";
		private string m_strProjDir="";
        private string m_strFvsTreeDb= frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
        private string m_strFvsOutDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile;
        private string m_dbConn = "";
        private string m_missingFvsOutDb = "This project has no valid /fvs/data/FVSOut.db. Multiple FVS Out functions will not work! Please use FVS to generate an /fvs/data/FVSOut.db file.";

        //POTFIRE BASE YEAR
        private string m_strOutPotFireBaseYearMDBFile = "";
        private string m_strPotFireBaseYearFile = "";
        private string m_strPotFireStandardFile = "";
        private string m_strPotFireBaseYearLinkedTableName = "FVS_POTFIRE_BASEYEAR";
        private string m_strPotFireStandardLinkedTableName = "FVS_POTFIRE_STANDARD";
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
        private FIA_Biosum_Manager.ado_data_access _MSAccess;
        private Label lblFSNetwork;
        private Label label16;
    
        public FIA_Biosum_Manager.ado_data_access MSAccess
        {
            get { return _MSAccess; }
            set { _MSAccess = value; }
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

            if (m_oQueries.m_oFvs.m_strFvsTreeTable.Trim().Length == 0)
            {
                m_oQueries.m_oFvs.m_strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
                m_strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
            }

			
			this.m_ado = new ado_data_access();
			this.m_dao = new dao_data_access();
			this.m_strTempMDBFileConnectionString = this.m_ado.getMDBConnString(this.m_oQueries.m_strTempDbFile,"","");
			this.m_ado.OpenConnection(this.m_strTempMDBFileConnectionString);
			if (this.m_ado.m_intError != 0)
			{
				this.m_intError = this.m_ado.m_intError;
				this.m_ado = null;
				return ;

			}

            this.m_oEnv = new env();

            this.m_bDebug = frmMain.g_bDebug;
			htSelectedRxFile=new Hashtable();

            lblFSNetwork.Text = "FCS_TREE.DB BIOSUM_CALC";
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
            this.lblFSNetwork = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.btnPostAppendAuditDb = new System.Windows.Forms.Button();
            this.cmbFilter = new System.Windows.Forms.ComboBox();
            this.pnlFileSizeMonitor = new System.Windows.Forms.Panel();
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
            this.uc_filesize_monitor3 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor2 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.uc_filesize_monitor1 = new FIA_Biosum_Manager.uc_filesize_monitor();
            this.groupBox1.SuspendLayout();
            this.pnlFileSizeMonitor.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 22);
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
            this.groupBox1.Controls.Add(this.lblFSNetwork);
            this.groupBox1.Controls.Add(this.label16);
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
            // lblFSNetwork
            // 
            this.lblFSNetwork.AutoSize = true;
            this.lblFSNetwork.Location = new System.Drawing.Point(612, 22);
            this.lblFSNetwork.Name = "lblFSNetwork";
            this.lblFSNetwork.Size = new System.Drawing.Size(31, 20);
            this.lblFSNetwork.TabIndex = 73;
            this.lblFSNetwork.Text = "NA";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(381, 22);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(257, 20);
            this.label16.TabIndex = 72;
            this.label16.Text = "FCS Volume and Biomass Method:";
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
            this.cmbFilter.Size = new System.Drawing.Size(61, 28);
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
            "Write FVS_InForest to FVSOUT_TREE_LIST.db"
            });
            this.cmbStep.Location = new System.Drawing.Point(8, 337);
            this.cmbStep.Name = "cmbStep";
            this.cmbStep.Size = new System.Drawing.Size(298, 28);
            this.cmbStep.TabIndex = 65;
            this.cmbStep.Text = "Step 1 - Define PRE/POST Table SeqNum";
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(8, 282);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(332, 20);
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
            this.txtOutDir.Size = new System.Drawing.Size(599, 26);
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

		#region Private Methods
		//((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory
		private void HandleDisplayRxSelection(object sender, EventArgs e)
		{                         
			if(htSelectedRxFile!=null)
			{
				htSelectedRxFile.Clear();
				htSelectedRxFile=procPreFrcs.FioIsRxExist(@"C:\FiaBiosumProjects\fia_biosum\orca_converted_data\fvs\db\out");          
			}
			Array aryT=Array.CreateInstance(Type.GetType("System.String"), htSelectedRxFile.Keys.Count);
			htSelectedRxFile.Keys.CopyTo(aryT, 0);
			Array.Sort(aryT);

		}
	  
      
		private void HandleSelectedRx(object sender, EventArgs e)
		{
			if(htSelectedRxFile!=null && htSelectedRxFile.Count >1)
			{
				//foreach(string item in listBox1.SelectedItems)
				//{
				// procPreFrcs.WfSelectedRxList.Add((string)htSelectedRxFile[item.Trim()]); //adds to the selected list of RX 
				//}
				//UnitTest();
			}
			else
				return;
		}

		#region UNIT TEST
		private void UnitTest()
		{
			try
			{
				/*//1) UpdateCondPreValues
				string strCIDpotFireStandId="4047C01";
				string strCIDmaster="2001006000000700000075061";
				procPreFrcs.DbRecsGetCondTabPreValues(strCIDmaster, strCIDpotFireStandId, procPreFrcs.WfSelectedRxList[0].ToString());                 
				*/
            
				//2) DbRecsTdgTsg2ProcDbCopy
				//procPreFrcs.DbRecsTdgTsg2ProcDbCopy(true);   

				//3) DbRecsFinalizeFvsTree
				//procPreFrcs.DbRecsFinalizeFvsTree(procPreFrcs.WfSelectedRxList[0].ToString());

				//4) DbRecsUpdateFfeFromPotFire
				procPreFrcs.DbRecsUpdateFfeFromPotFire(procPreFrcs.WfSelectedRxList[0].ToString());
            
			}
			catch(Exception e)
			{
				MessageBox.Show(e.Message,"UnitTest");
			}
		}



		#endregion

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		#endregion

		public void loadvalues()
		{
			
			
			
			string strOutDirAndFile;
			int x,y;
			string strRx1="";
			string strRx2="";
			string strRx3="";
			string strRx4="";
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
				ado_data_access oAdo = new ado_data_access();
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
				this.lstFvsOutput.Columns.Add("FVS Variant", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("RxPackage", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Cycle1Rx", 120, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Cycle2Rx", 120, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Cycle3Rx", 120, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Cycle4Rx", 120, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Run Status",250,HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Output File", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("File Found", 80, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Summary Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Cut List Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Tree Standing Count", 100, HorizontalAlignment.Left);
				this.lstFvsOutput.Columns.Add("Potential Fire Count", 100, HorizontalAlignment.Left);
                this.lstFvsOutput.Columns.Add("Potential Fire Base Yr Output File", 100, HorizontalAlignment.Left);
                this.lstFvsOutput.Columns.Add("Potential Fire Base Yr File Found", 100, HorizontalAlignment.Left);
                this.lstFvsOutput.Columns.Add("Potential Fire Base Yr Count", 100, HorizontalAlignment.Left);

               

				this.lstFvsOutput.Columns[COL_CHECKBOX].Width = -2;
				
                //load rxpackage properties
				m_oRxPackageItem_Collection = new RxPackageItem_Collection();
				this.m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_ado,m_ado.m_OleDbConnection,m_oQueries,this.m_oRxPackageItem_Collection);


                //@ToDo: Can I delete this dao?
                string strTempDbFile="";
				dao_data_access oDao = new dao_data_access();
				strTempDbFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
				oDao.CreateMDB(strTempDbFile);
				oDao.OpenDb(strTempDbFile);


                // Get variants/rxPackages in project
                m_ado.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(this.m_oQueries.m_oFIAPlot.m_strPlotTable,this.m_oQueries.m_oFvs.m_strRxPackageTable);				
				this.m_ado.SqlQueryReader(this.m_ado.m_OleDbConnection,this.m_ado.m_strSQL);
                IList<string> lstRunTitles = new List<string>();
				while (this.m_ado.m_OleDbDataReader.Read())
				{
                    strVariant = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                    strPackage = this.m_ado.m_OleDbDataReader["RxPackage"].ToString().Trim();                  
                    this.m_strOutMDBFile = this.m_oRxTools.GetRxPackageFvsOutDbFileName(m_ado.m_OleDbDataReader);
                    // This is the Access DB for this variant/package ie: FVSOUT_CA_P001-001-001-001-001.MDB
                    strOutDirAndFile = this.txtOutDir.Text.Trim() + "\\" +
                           this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "\\" +
                            this.m_strOutMDBFile.Trim();
                    // Example RunTitle: FVSOUT_WC_P999-999-999-999-999
                    strCurRunTitle = this.m_oRxTools.GetRxPackageRunTitle(m_ado.m_OleDbDataReader);

                    long lngPreSummaryRecords = 0;
                    string strSummaryConnect = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.FVS.DefaultPreFVSSummaryDbFile}";
                    if (System.IO.File.Exists(strSummaryConnect))
                    {
                        using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(oAdo.getMDBConnString(strSummaryConnect, "", "")))
                        {
                            conn.Open();
                            if (oAdo.TableExist(conn, Tables.FVS.DefaultPreFVSSummaryTableName))
                            {
                                string strSQL = $@"SELECT COUNT(*) FROM(SELECT TOP 1 biosum_cond_id FROM {Tables.FVS.DefaultPreFVSSummaryTableName} WHERE 
                                    FVS_VARIANT = '{strVariant}' AND RXPACKAGE = '{strPackage}')";
                                lngPreSummaryRecords = (long)oAdo.getRecordCount(conn, strSQL, Tables.FVS.DefaultPreFVSSummaryTableName);
                            }
                        }
                    }
                    /************************************************************************
                    /**Check and Assign in the FVS_CASES whether the FVS output has been 
                     **appended to the fvs_cutTree table
                     ************************************************************************/
                    if (System.IO.File.Exists(m_strFvsOutDb) == true)
					{
                        string dbConn = SQLite.GetConnectionString(m_strFvsOutDb);
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                        {
                            conn.Open();
                            if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSCasesTableName))
                            {
                                if (! SQLite.ColumnExist(conn, Tables.FVS.DefaultFVSCasesTableName, m_colBioSumAppend))
                                {
                                    SQLite.AddColumn(conn, Tables.FVS.DefaultFVSCasesTableName, m_colBioSumAppend, "TEXT", "1");
                                    SQLite.m_strSQL = $@"UPDATE {Tables.FVS.DefaultFVSCasesTableName} SET {m_colBioSumAppend} = 'N'";
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                                else
                                {
                                    //check if records exist in the cut list for this variant/package
                                    if (lngPreSummaryRecords == 0)
                                    {
                                        SQLite.m_strSQL = $@"UPDATE {Tables.FVS.DefaultFVSCasesTableName} SET {m_colBioSumAppend}='N' WHERE RunTitle = '{strCurRunTitle}'";
                                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                    }
                                }
                                SQLite.m_strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSCasesTableName} WHERE RunTitle = '{strCurRunTitle}'";
                                long lngCasesCount = SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCasesTableName);
                                if (lngCasesCount > 0)
                                {
                                    lstRunTitles.Add(strCurRunTitle);
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
                if (!FvsOutWithCutList())
                {
                    MessageBox.Show(m_missingFvsOutDb, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

				oDao.m_DaoDatabase.Close();
				oDao.m_DaoWorkspace.Close();

                oDao.m_DaoDatabase = null;
                oDao.m_DaoWorkspace = null;

				m_ado.m_OleDbDataReader.Close();

				

				oAdo.OpenConnection(oAdo.getMDBConnString(strTempDbFile,"",""));

                strVariant="";
                strCurVariant="";

               
				for (x=0;x<=lstRunTitles.Count -1;x++)
				{
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel  >2)
                    {
                        long memused = GC.GetTotalMemory(true);
                        this.WriteText(m_strDebugFile, "uc_fvs_output.loadvalues Memory Used: " + String.Format("{0:n0}" + " MB",memused));
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    oAdo.DisplayErrors = false;

                    strCurRunTitle = lstRunTitles.ElementAt(x);
                    string[] arrVariant = strCurRunTitle.Split('_');
                    if (arrVariant.Length == 3)
                    {
                        strVariant = arrVariant[1];
                        string[] arrRx = arrVariant[2].Split('-');
                        if (arrRx.Length == 5)
                        {
                            strPackage = arrRx[0].Substring(1, 3);
                            strRx1 = arrRx[1];
                            strRx2 = arrRx[2];
                            strRx3 = arrRx[3];
                            strRx4 = arrRx[4];
                        }
                    }

                    frmMain.g_sbpInfo.Text = $@"Loading FVS Variant RxPackage {strVariant} {strPackage} ...Stand By";
				    frmMain.g_sbpInfo.Parent.Refresh();

					// Add a ListItem object to the ListView.
					entryListItem = this.lstFvsOutput.Items.Add("");					
					entryListItem.UseItemStyleForSubItems=false;
					this.m_oLvAlternateColors.AddRow();
					this.m_oLvAlternateColors.AddColumns(lstFvsOutput.Items.Count-1,lstFvsOutput.Columns.Count);
					entryListItem.SubItems.Add(strVariant);
					entryListItem.SubItems.Add(strPackage);
					entryListItem.SubItems.Add(strRx1);
					entryListItem.SubItems.Add(strRx2);
					entryListItem.SubItems.Add(strRx3);
					entryListItem.SubItems.Add(strRx4);
					entryListItem.SubItems.Add(" ");  //out mdb file
					entryListItem.SubItems.Add(" ");  //file found
					entryListItem.SubItems.Add(" ");  //summary record count
					entryListItem.SubItems.Add(" ");  //tree cut list record count
					entryListItem.SubItems.Add(" ");  //tree standing record count
					entryListItem.SubItems.Add(" ");  //potential fire record count
                    entryListItem.SubItems.Add(" ");  //potential fire base year out file
                    entryListItem.SubItems.Add(" ");  //file found
                    entryListItem.SubItems.Add(" ");  //potential fire base year record count
				    entryListItem.SubItems.Add(" ");  //run status

                    //FVS_SUMMARY
                    string dbConn = SQLite.GetConnectionString(m_strFvsOutDb);
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                    {
                        conn.Open();
                        if (SQLite.TableExist(conn, Tables.FVS.DefaultFVSSummaryTableName))
                        {
                            if (!frmMain.g_bSuppressFVSOutputTableRowCount)
                            {
                                SQLite.m_strSQL = $@"select count(*) from {Tables.FVS.DefaultFVSSummaryTableName}";
                                entryListItem.SubItems[COL_SUMMARYCOUNT].Text = Convert.ToString(Convert.ToUInt32(SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSSummaryTableName)));
                            }
                            else
                            {
                                SQLite.m_strSQL = $@"SELECT TPA FROM {Tables.FVS.DefaultFVSSummaryTableName} WHERE TPA > 0 LIMIT 1";
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
                            entryListItem.SubItems[COL_SUMMARYCOUNT].Text = "N";
                        }

                        // SET APPEND FLAG
                        string strUpdateStatus = "";
                        SQLite.m_strSQL = $@"select count(*) from {Tables.FVS.DefaultFVSCasesTableName} WHERE {m_colBioSumAppend} ='N'";
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
                                SQLite.m_strSQL = $@"select count(*) from {Tables.FVS.DefaultFVSCutListTableName}";
                                entryListItem.SubItems[COL_CUTCOUNT].Text = Convert.ToString(Convert.ToUInt32(SQLite.getRecordCount(conn, SQLite.m_strSQL, Tables.FVS.DefaultFVSCutListTableName)));
                            }
                            else
                            {
                                SQLite.m_strSQL = $@"SELECT HT FROM {Tables.FVS.DefaultFVSCutListTableName} WHERE HT > 0 LIMIT 1";
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
                    }
                       				
				}
			    oAdo.CloseAndDisposeConnection(oAdo.m_OleDbConnection,true);
                

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

                this.label16.Left = (int)(lblTitle.Width / 2) + 5;
                this.lblFSNetwork.Left = this.label16.Left + this.label16.Width + 2;

               
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
        /// Make sure MDB files exist and that they have records
        /// </summary>
		private void val_data()
		{
            ado_data_access oAdoStandard = null;
            ado_data_access oAdoPotFireBaseYr = null;
            string strOutDirAndFile="";
            string strVariant = "";
            string strCurVariant = "";
            string strStandardMDB = "";
            string strStandardMDBFound = "";
            
            string strPotFireBaseYrMDB = "";
            string strPotFireBaseYrMDBFound = "";
            string strOutDir = (string)frmMain.g_oDelegate.GetControlPropertyValue(txtOutDir,"Text",false).ToString().Trim();
            string strSummaryCount = "";
            string strCutListCount = "";
            string strTreeListCount = "";
            string strPotFireCount = "";
            string strPotFireBaseYrCount = "";
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
                    strStandardMDB = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_MDBOUT, "Text", false).ToString().Trim();
                    strStandardMDBFound = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_FOUND, "Text", false).ToString().Trim();
                    strPotFireBaseYrMDB = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_POTFIREMDBOUT, "Text", false).ToString().Trim();
                    strPotFireBaseYrMDBFound = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_POTFIREMDBFOUND, "Text", false).ToString().Trim();
                    strSummaryCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_SUMMARYCOUNT, "Text", false).ToString().Trim();
                    strPotFireCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_POTFIRECOUNT, "Text", false).ToString().Trim();
                    strCutListCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_CUTCOUNT, "Text", false).ToString().Trim();
                    strTreeListCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_LEFTCOUNT, "Text", false).ToString().Trim();
                    strPotFireBaseYrCount = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_POTFIREBASEYEARCOUNT, "Text", false).ToString().Trim();




                    GetPrePostSeqNumConfiguration("FVS_POTFIRE",strRxPackage);
                    //dont need to validate base year if baseyear is not being referenced
                    if (m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN=="Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN=="Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN=="Y" ||
                        m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN=="Y")
                         bPotFireBaseYear=true;
                    else 
                         bPotFireBaseYear=false;

                    if (strStandardMDBFound == "No")
					{
						MessageBox.Show("!!File " + 
						    strStandardMDB + " Not Found!!",
							"FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						break;
					}
                    if (strPotFireBaseYrMDBFound == "No" && bPotFireBaseYear)
                    {
                        MessageBox.Show("!!File " +
                            strPotFireBaseYrMDB + " Not Found!!",
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                        this.m_intError = -1;
                        break;
                    }
                    //open file if suppressing record counts is checked
                    if (frmMain.g_bSuppressFVSOutputTableRowCount)
                    {
                        if (oAdoStandard != null &&
                             oAdoStandard.m_OleDbConnection.State == System.Data.ConnectionState.Open)
                        {
                            oAdoStandard.CloseConnection(oAdoStandard.m_OleDbConnection);
                            oAdoStandard = null;
                        }
                        strOutDirAndFile = strOutDir + "\\" +
                             strVariant + "\\" + strStandardMDB;
                        oAdoStandard = new ado_data_access();
                        oAdoStandard.OpenConnection(oAdoStandard.getMDBConnString(strOutDirAndFile, "", ""));
                    }
					if (strSummaryCount.Length  == 0 ||
						strSummaryCount == "0")
					{
                        if (frmMain.g_bSuppressFVSOutputTableRowCount == false)
                        {
                            this.m_intError = -1;
                        }
                        else
                        {
                            if (oAdoStandard.TableExist(oAdoStandard.m_OleDbConnection, "FVS_SUMMARY") == false)
                            {
                                m_intError = -1;
                            }
                            else
                            {
                                oAdoStandard.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM FVS_SUMMARY)";
                                if ((int)oAdoStandard.getRecordCount(oAdoStandard.m_OleDbConnection, oAdoStandard.m_strSQL, "FVS_SUMMARY") == 0)
                                {
                                    m_intError = -1;
                                }

                            }

                        }
                        if (m_intError != 0)
                        {
                            MessageBox.Show("!!Summary Table In File  " + strStandardMDB + " " +
                                " Does Not Exist Or Has 0 Records!!",
                                "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Exclamation);
                            break;
                        }
					}
					
					if (strCutListCount.Length  == 0)
					{
                        
                        if (frmMain.g_bSuppressFVSOutputTableRowCount == false)
                        {
                            this.m_intError = -1;
                        }
                        else
                        {
                            if (oAdoStandard.TableExist(oAdoStandard.m_OleDbConnection, "FVS_CUTLIST") == false)
                            {
                                m_intError = -1;
                            }
                            else
                            {
                                oAdoStandard.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM FVS_CUTLIST)";
                                if ((int)oAdoStandard.getRecordCount(oAdoStandard.m_OleDbConnection, oAdoStandard.m_strSQL, "FVS_CUTLIST") == 0)
                                {
                                    m_intError = -1;
                                }

                            }

                        }
                        if (m_intError != 0)
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm, "Visible", false);
                            result = MessageBox.Show("!!Warning!!\r\n-----------\r\nCut Tree Table In File  " + strStandardMDB + " " +
                                 " Does Not Exist. Continue Processing?(Y/N)",
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
					}
					if (strPotFireCount.Length  == 0)
					{
                        
                        if (frmMain.g_bSuppressFVSOutputTableRowCount == false)
                        {
                            this.m_intError = -1;
                        }
                        else
                        {
                            if (oAdoStandard.TableExist(oAdoStandard.m_OleDbConnection, "FVS_POTFIRE") == false)
                            {
                                m_intError = -1;
                            }
                            else
                            {
                                oAdoStandard.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM FVS_POTFIRE)";
                                if ((int)oAdoStandard.getRecordCount(oAdoStandard.m_OleDbConnection, oAdoStandard.m_strSQL, "FVS_POTFIRE") == 0)
                                {
                                    m_intError = -1;
                                }

                            }

                        }
                        if (m_intError != 0)
                        {
                            MessageBox.Show("!!Potential Fire Table In File  " + strStandardMDB + " " +
                                 " Does Not Exist!!",
                                 "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                 System.Windows.Forms.MessageBoxIcon.Exclamation);
                            break;
                        }
					}
                    
                    if (strPotFireBaseYrCount.Length == 0 && bPotFireBaseYear)
                    {

                        if (frmMain.g_bSuppressFVSOutputTableRowCount == false)
                        {
                            MessageBox.Show("!!Potential Fire Base Year Table In File  " + strPotFireBaseYrMDB + " " +
                                " Does Not Exist!!",
                                "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Exclamation);
                            this.m_intError = -1;
                            break;
                        }
                        else
                        {
                            if (strVariant != strCurVariant)
                            {
                                strCurVariant = strVariant;
                                if (oAdoPotFireBaseYr != null &&
                                    oAdoPotFireBaseYr.m_OleDbConnection.State == System.Data.ConnectionState.Open)
                                {
                                    oAdoPotFireBaseYr.CloseConnection(oAdoPotFireBaseYr.m_OleDbConnection);
                                    oAdoPotFireBaseYr = null;
                                }
                                strOutDirAndFile = this.txtOutDir.Text.Trim() + "\\" +
                                     strVariant + "\\FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                                oAdoPotFireBaseYr = new ado_data_access();
                                oAdoPotFireBaseYr.OpenConnection(oAdoPotFireBaseYr.getMDBConnString(strOutDirAndFile, "", ""));
                                if (oAdoPotFireBaseYr.TableExist(oAdoPotFireBaseYr.m_OleDbConnection, "FVS_POTFIRE") == false)
                                {
                                    MessageBox.Show("!!Potential Fire Base Year Table In File  " + strPotFireBaseYrMDB + " " +
                                                   " Does Not Exist!!",
                                                   "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                                   System.Windows.Forms.MessageBoxIcon.Exclamation);
                                                        this.m_intError = -1;
                                     break;
                                }
                                else
                                {
                                    oAdoPotFireBaseYr.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM FVS_POTFIRE)";
                                    if ((int)oAdoPotFireBaseYr.getRecordCount(oAdoPotFireBaseYr.m_OleDbConnection, oAdoPotFireBaseYr.m_strSQL, "FVS_POTFIRE") == 0)
                                    {
                                        m_intError = -1;
                                        MessageBox.Show("!!Potential Fire Base Year Table In File  " + strPotFireBaseYrMDB + " " +
                                                        " Has No Records!!",
                                                        "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                                                        System.Windows.Forms.MessageBoxIcon.Exclamation);
                                        break;
                                    }

                                }
                            }
                        }
                    }
                     
				}
			}
            if (oAdoPotFireBaseYr != null)
            {
                oAdoPotFireBaseYr.CloseConnection(oAdoPotFireBaseYr.m_OleDbConnection);
                oAdoPotFireBaseYr = null;
            }
            if (oAdoStandard != null)
            {
                oAdoStandard.CloseConnection(oAdoStandard.m_OleDbConnection);
                oAdoStandard = null;
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

        private void RunAppend_UpdatePrePostTable(string p_strPackage, string p_strVariant, string p_strRx1, string p_strRx2, string p_strRx3, string p_strRx4, bool p_bUpdatePreTableWithVariant,
            int p_intListViewItem,ref int p_intError, ref string p_strError)
        {

            int x, y, z, zz;
           
            string strSourceColumnsList = "";
            string strSourceColumnsReservedWordFormattedList = "";
            string[] strSourceColumnsArray = null;
            string strDestColumnsList = "";
            string[] strDestColumnsArray = null;
            System.Data.DataTable oDataTableSchema;
            bool bFound;
            string strRx = "";
            string strCycle = "";
            
            System.Data.OleDb.OleDbConnection oConn;
            string strDbFile = "";
            string strFVSOutTable="";
            string strFVSOutTableLink = "";
            string strFVSSummarySeqNumMtxTableLink="";
            string strFVSOutSeqNumMatrixTableLink="";

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_UpdatePrePostTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            ado_data_access oAdo = new ado_data_access();

            oAdo.m_intError = 0;
            oAdo.m_strError = "";
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + " Update PREPOST Tables");

            //
            //make sure all the tables and columns exist
            //
            oAdo.m_strSQL = "";
            m_intProgressStepTotalCount = m_oPrePostDbFileItem_Collection.Count;
            m_intProgressStepCurrentCount = 0;
            for (y = 0; y <= m_oPrePostDbFileItem_Collection.Count - 1; y++)
            {
                m_intProgressStepCurrentCount++;
                UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepCurrentCount,
                            m_intProgressStepTotalCount);

                strDbFile = m_oPrePostDbFileItem_Collection.Item(y).DbFileName.Trim().ToUpper();
                if (strDbFile != "PREPOST_FVS_CASES.ACCDB" &&
                    strDbFile != "PREPOST_FVS_TREELIST.ACCDB" &&
                    strDbFile != "PREPOST_FVS_CUTLIST.ACCDB" &&
                    strDbFile != "PREPOST_FVS_ATRTLIST.ACCDB")
                          
                {

                    for (x = 0; x <= m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Count - 1; x++)
                    {
                        TableLinkItem oTableLinkItem = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(x);
                        if (m_oPrePostDbFileItem_Collection.Item(y).TableType.Trim().ToUpper() == 
                            oTableLinkItem.FVSOutputTableName.Trim().ToUpper() && oTableLinkItem.FVSOutputTable)
                        {
                            

                            strFVSOutTable = oTableLinkItem.FVSOutputTableName.Trim().ToUpper();
                            strFVSOutTableLink = oTableLinkItem.LinkedTableName.Trim().ToUpper();

                            oConn = m_oPrePostDbFileItem_Collection.Item(y).Connection;

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + " Update PREPOST Tables:" + strFVSOutTable);

                            if (uc_filesize_monitor1.File.Trim().Length == 0)
                            {
                                uc_filesize_monitor1.BeginMonitoringFile(m_oPrePostDbFileItem_Collection.Item(y).FullPath.Trim(), 2000000000, "2GB");
                                //uc_filesize_monitor1.Information = "The temporary DB file listed above is a copy of the production DB file " + frmMain.g_oUtils.getFileName(strFvsTreeFile);
                            }
                            else if (uc_filesize_monitor2.File.Trim().Length == 0)
                            {
                                uc_filesize_monitor2.BeginMonitoringFile(m_oPrePostDbFileItem_Collection.Item(y).FullPath.Trim(), 2000000000, "2GB");
                                //uc_filesize_monitor2.Information = "Base year potential fire table for variant " + strVariant;
                            }

                            if (!oAdo.TableExist(oConn, "PRE_" + strFVSOutTable))
                            {
                                //create the table
                                oDataTableSchema = oAdo.getTableSchema(oConn, "SELECT * FROM " + strFVSOutTableLink);
                                strSourceColumnsList = oAdo.getFieldNames(oConn, "SELECT * FROM " + strFVSOutTableLink);
                                strSourceColumnsReservedWordFormattedList = oAdo.FormatReservedWordsInColumnNameList(strSourceColumnsList, ",");
                                strSourceColumnsArray = m_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                                oAdo.m_strSQL = "";
                                for (z = 0; z <= oDataTableSchema.Rows.Count - 1; z++)
                                {
                                    oAdo.m_strSQL = oAdo.m_strSQL +
                                        oAdo.FormatCreateTableSqlFieldItem(oDataTableSchema.Rows[z]) + ",";
                                }

                                if (oAdo.m_strSQL.Trim().Length > 0)
                                {
                                    oAdo.m_strSQL = oAdo.m_strSQL.Substring(0, oAdo.m_strSQL.Length - 1);
                                    oAdo.m_strSQL = strFVSOutTable + " (biosum_cond_id text(25),rxpackage text(3),rx text(3), rxcycle text(1), fvs_variant text(2)," +
                                        oAdo.m_strSQL + ")";

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "CREATE TABLE PRE_" + oAdo.m_strSQL + "\r\n\r\n");
                                    oAdo.SqlNonQuery(oConn, "CREATE TABLE PRE_" + oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "CREATE TABLE POST_" + oAdo.m_strSQL + "\r\n\r\n");
                                    oAdo.SqlNonQuery(oConn, "CREATE TABLE POST_" + oAdo.m_strSQL);




                                    oAdo.m_strSQL = "CREATE INDEX biosumcondididx_pre ON PRE_" + strFVSOutTable + " (biosum_cond_id)";
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                    oAdo.m_strSQL = "CREATE INDEX biosumcondididx_post ON POST_" + strFVSOutTable + " (biosum_cond_id)";
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                    oAdo.m_strSQL = "CREATE INDEX biosumcondidrxidx_pre ON POST_" + strFVSOutTable + " (biosum_cond_id,rxpackage,rx,rxcycle)";
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                    oAdo.m_strSQL = "CREATE INDEX biosumcondidrxidx_post ON POST_" + strFVSOutTable + " (biosum_cond_id,rxpackage,rx,rxcycle)";
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                }
                            }
                            else
                            {
                                //see if columns are the same
                                oDataTableSchema = oAdo.getTableSchema(oConn, "SELECT * FROM " + strFVSOutTableLink);
                                strSourceColumnsList = oAdo.getFieldNames(oConn, "SELECT * FROM " + strFVSOutTableLink);
                                strSourceColumnsReservedWordFormattedList = oAdo.FormatReservedWordsInColumnNameList(strSourceColumnsList, ",");
                                strSourceColumnsArray = m_oUtils.ConvertListToArray(strSourceColumnsList, ",");
                                strDestColumnsList = oAdo.getFieldNames(oConn, "SELECT * FROM PRE_" + strFVSOutTable);
                                strDestColumnsArray = m_oUtils.ConvertListToArray(strDestColumnsList, ",");

                                oAdo.m_strSQL = "";
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
                                            oAdo.m_strSQL = oAdo.FormatCreateTableSqlFieldItem(oDataTableSchema.Rows[z]);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + "ALTER TABLE PRE_" + strFVSOutTable + " " +
                                                "ADD COLUMN " + oAdo.m_strSQL + "\r\n");
                                            oAdo.SqlNonQuery(oConn, "ALTER TABLE PRE_" + strFVSOutTable + " " +
                                                "ADD COLUMN " + oAdo.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + "ALTER TABLE POST_" + strFVSOutTable + " " +
                                                "ADD COLUMN " + oAdo.m_strSQL + "\r\n");
                                            oAdo.SqlNonQuery(oConn, "ALTER TABLE POST_" + strFVSOutTable + " " +
                                                "ADD COLUMN " + oAdo.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        }
                                    }

                                }


                            }
                            oDataTableSchema.Dispose();

                            if (oAdo.m_intError == 0)
                            {
                                oAdo.m_strSQL = "DELETE FROM PRE_" + strFVSOutTable + " " +
                                    "WHERE RXPACKAGE='" + p_strPackage.Trim() + "'" + " AND " +
                                          "FVS_VARIANT='" + p_strVariant.Trim() + "'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }

                            if (oAdo.m_intError == 0)
                            {
                                oAdo.m_strSQL = "DELETE FROM POST_" + strFVSOutTable + " " +
                                    "WHERE RXPACKAGE='" + p_strPackage.Trim() + "'" + " AND " +
                                          "FVS_VARIANT='" + p_strVariant.Trim() + "'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }

                            if (oAdo.m_intError == 0)
                            {
                                GetPrePostSeqNumConfiguration(strFVSOutTable, p_strPackage);
                            }
                            //
                            //GET THE SEQNUM MATRIX TABLE
                            //
                            GetPrePostTableLinkItems(
                                    m_oPrePostDbFileItem_Collection,
                                    "PREPOST_" + strFVSOutTable + ".ACCDB",
                                    strFVSOutTable,
                                    ref strFVSOutSeqNumMatrixTableLink,
                                    ref strFVSSummarySeqNumMtxTableLink,
                                    ref strFVSOutTableLink);

                            //
                            //INSERT THE RECORDS BY CYCLE
                            //
                            //for (z = 0; z <= this.m_strRxCycleArray.Length - 1; z++)
                            for (z=1;z<=4;z++)
                            {
                                //strCycle = m_strRxCycleArray[z].Trim();
                                strCycle = z.ToString().Trim();
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


                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Rx:" + strRx.Trim() + " Cycle:" + strCycle + ": Get Pre And Post Treatment Years");
                                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");



                                string strFormattedSelectColumnList = "";
                                for (zz = 0; zz <= strSourceColumnsArray.Length - 1; zz++)
                                {
                                    if (strSourceColumnsArray[zz].Substring(0,1) != "_")
                                        strFormattedSelectColumnList = strFormattedSelectColumnList + "a." +  strSourceColumnsArray[zz].Trim() + ",";
                                    else
                                        strFormattedSelectColumnList = strFormattedSelectColumnList + "a.[" + strSourceColumnsArray[zz].Trim() + "],";
                                }

                                strFormattedSelectColumnList = strFormattedSelectColumnList.Substring(0, strFormattedSelectColumnList.Length - 1);

                                if (strFVSOutTable != "FVS_STRCLASS")
                                {
                                    if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "Y")
                                    {
                                        oAdo.m_strSQL = "INSERT INTO PRE_" + strFVSOutTable + " " +
                                            "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " +
                                            "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                   "'" + strRx + "' AS rx," +
                                                   "'" + strCycle + "' AS rxcycle," +
                                                   "'" + p_strVariant + "' AS fvs_variant," +
                                                  strFormattedSelectColumnList + " " +
                                           "FROM " + strFVSOutTableLink + " a," +
                                              "(SELECT standid,year " +
                                               "FROM " + strFVSSummarySeqNumMtxTableLink + " " +
                                               "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  AS b " +
                                           "WHERE a.standid=b.standid AND a.year=b.year";
                                    }
                                    else
                                    {
                                        oAdo.m_strSQL = "INSERT INTO PRE_" + strFVSOutTable + " " +
                                           "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " +
                                           "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                  "'" + strRx + "' AS rx," +
                                                  "'" + strCycle + "' AS rxcycle," +
                                                  "'" + p_strVariant + "' AS fvs_variant," +
                                                 strFormattedSelectColumnList + " " +
                                          "FROM " + strFVSOutTableLink + " a," +
                                             "(SELECT standid,year " +
                                              "FROM " + strFVSOutSeqNumMatrixTableLink + " " +
                                              "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  AS b " +
                                          "WHERE a.standid=b.standid AND a.year=b.year";
                                    }
                                }
                                else
                                {

                                    oAdo.m_strSQL = "INSERT INTO PRE_" + strFVSOutTable + " " +
                                       "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " +
                                       "SELECT '" + p_strPackage + "' AS rxpackage," +
                                              "'" + strRx + "' AS rx," +
                                              "'" + strCycle + "' AS rxcycle," +
                                              "'" + p_strVariant + "' AS fvs_variant," +
                                             strFormattedSelectColumnList + " " +
                                      "FROM " + strFVSOutTableLink + " a," +
                                         "(SELECT standid,year,removal_code " +
                                          "FROM " + strFVSOutSeqNumMatrixTableLink + " " +
                                          "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  AS b " +
                                      "WHERE a.standid=b.standid AND a.year=b.year AND a.removal_code=b.removal_code";

                                }
                                
                                




                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


                                if (oAdo.m_intError == 0)
                                {
                                    if (strFVSOutTable != "FVS_STRCLASS")
                                    {
                                        if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "Y")
                                        {
                                            oAdo.m_strSQL = "INSERT INTO POST_" + strFVSOutTable + " " +
                                                "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " +
                                                "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                       "'" + strRx + "' AS rx," +
                                                       "'" + strCycle + "' AS rxcycle," +
                                                       "'" + p_strVariant + "' AS fvs_variant," +
                                                      strFormattedSelectColumnList + " " +
                                               "FROM " + strFVSOutTableLink + " a," +
                                                  "(SELECT standid,year " +
                                                   "FROM " + strFVSSummarySeqNumMtxTableLink + " " +
                                                   "WHERE CYCLE" + strCycle + "_POST_YN='Y')  AS b " +
                                               "WHERE a.standid=b.standid AND a.year=b.year";
                                        }
                                        else
                                        {
                                            oAdo.m_strSQL = "INSERT INTO POST_" + strFVSOutTable + " " +
                                               "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " +
                                               "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                      "'" + strRx + "' AS rx," +
                                                      "'" + strCycle + "' AS rxcycle," +
                                                      "'" + p_strVariant + "' AS fvs_variant," +
                                                     strFormattedSelectColumnList + " " +
                                              "FROM " + strFVSOutTableLink + " a," +
                                                 "(SELECT standid,year " +
                                                  "FROM " + strFVSOutSeqNumMatrixTableLink + " " +
                                                  "WHERE CYCLE" + strCycle + "_POST_YN='Y')  AS b " +
                                              "WHERE a.standid=b.standid AND a.year=b.year";
                                        }
                                    }
                                    else
                                    {
                                        oAdo.m_strSQL = "INSERT INTO POST_" + strFVSOutTable + " " +
                                              "(rxpackage,rx,rxcycle,fvs_variant," + strSourceColumnsReservedWordFormattedList + ") " +
                                              "SELECT '" + p_strPackage + "' AS rxpackage," +
                                                     "'" + strRx + "' AS rx," +
                                                     "'" + strCycle + "' AS rxcycle," +
                                                     "'" + p_strVariant + "' AS fvs_variant," +
                                                    strFormattedSelectColumnList + " " +
                                             "FROM " + strFVSOutTableLink + " a," +
                                                "(SELECT standid,year,removal_code " +
                                                 "FROM " + strFVSOutSeqNumMatrixTableLink + " " +
                                                 "WHERE CYCLE" + strCycle + "_POST_YN='Y')  AS b " +
                                             "WHERE a.standid=b.standid AND a.year=b.year AND a.removal_code=b.removal_code";
                                    }

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                }



                                if (oAdo.m_intError == 0)
                                {
                                    //update biosum_cond_id column
                                    oAdo.m_strSQL = "UPDATE PRE_" + strFVSOutTable + " " +
                                        "SET biosum_cond_id = IIF((biosum_cond_id IS NULL OR LEN(TRIM(biosum_cond_id))=0) AND (standid IS NOT NULL AND LEN(TRIM(standid)) = 25),MID(standid,1,25),'') " +
                                        "WHERE RXPACKAGE='" + p_strPackage.Trim() + "' AND " +
                                              "RX='" + strRx.Trim() + "' AND " +
                                              "RXCYCLE='" + strCycle + "' AND " +
                                              "FVS_VARIANT='" + p_strVariant.Trim() + "'";

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }



                                if (oAdo.m_intError == 0)
                                {
                                    //update biosum_cond_id column
                                    oAdo.m_strSQL = "UPDATE POST_" + strFVSOutTable + " " +
                                        "SET biosum_cond_id = IIF((biosum_cond_id IS NULL OR LEN(TRIM(biosum_cond_id))=0) AND (standid IS NOT NULL AND LEN(TRIM(standid)) = 25),MID(standid,1,25),'') " +
                                        "WHERE RXPACKAGE='" + p_strPackage.Trim() + "' AND " +
                                        "RX='" + strRx.Trim() + "' AND " +
                                        "RXCYCLE='" + strCycle + "' AND " +
                                        "FVS_VARIANT='" + p_strVariant.Trim() + "'";

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                if (uc_filesize_monitor3.File.Length > 0) uc_filesize_monitor3.EndMonitoringFile();
                                else uc_filesize_monitor2.EndMonitoringFile();

                                if (oAdo.m_intError != 0) break;

                            }
                        }

                        
                        if (oAdo.m_intError != 0) break;

                    }
                    if (oAdo.m_intError != 0) break;
                }
            }
            p_intError=oAdo.m_intError;
            p_strError = oAdo.m_strError;
            oAdo=null;
        }
		private void RunAppend_Main()
		{
			string strOutDirAndFile;
            string strAuditDbFile;
			string strCurVariant="";
			
			bool bUpdateCondTable=false;
			
            string strFVSOutPrePostPathAndDbFile = "";

           
			
			
			
			ado_data_access oAdo = new ado_data_access();
			
			
			
			
			int intCount;
			string strRx1="";
			string strRx2="";
			string strRx3="";
			string strRx4="";
			string strPackage="";
			string strVariant="";
           
            string strPotFireBaseYrDbFile = "";
            

			
			
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;
            

			Tables oTables = new Tables();
			
			frmMain.g_oDelegate.CurrentThreadProcessStarted=true;
			int x,y;
            oAdo.DisplayErrors = false;
			m_intProgressOverallTotalCount=0;
			m_intProgressStepCurrentCount=0;
			m_intProgressOverallCurrentCount=0;

			try
			{
                if (System.IO.File.Exists(m_strDebugFile))
                    System.IO.File.Delete(m_strDebugFile);

                m_dao.DisplayErrors = false;

                System.Threading.Thread.Sleep(2000);

                if (m_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");

				intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv,"Count",false);

				//inititalize the run status column for each row to blank in the list view
				for (x=0;x<=intCount-1;x++)
				{
                    oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
                    //alternate the row color in the list view
                    m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = true;
                    m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
                    //inititalize the run status column for each row to blank in the list view
                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x,COL_RUNSTATUS, "Text", "");
                    //see if checked
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked",false))
                        this.m_intProgressOverallTotalCount++;
			
				}

                if (m_oFVSPrePostSeqNumItemCollection == null) m_oFVSPrePostSeqNumItemCollection = new FVSPrePostSeqNumItem_Collection();
                if (m_ado.m_OleDbConnection.State == ConnectionState.Closed)
                {
                    m_ado.OpenConnection(m_strTempMDBFileConnectionString);

                }
                m_oRxTools.LoadFVSOutputPrePostRxCycleSeqNum(m_ado, m_ado.m_OleDbConnection, m_oFVSPrePostSeqNumItemCollection);
            

                m_intProgressOverallCurrentCount = 0;
                this.m_intProgressOverallTotalCount++;

				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Maximum",100);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Minimum",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2,"Value",0);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2,"Text","Overall Progress");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Text","");
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg,"Visible",true);

                this.val_data();

                // Ensure that indexes are on the FVSOut.db
                m_oRxTools.CreateFvsOutDbIndexes(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +Tables.FVS.DefaultFVSOutDbFile);

                if (m_intError == 0)
                {

                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "check point 1 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                    

                    //create a link to each of the selected fvs out files in the temp mdb file
                    //close the current ado oledb connection
                    if (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
                        this.m_ado.CloseConnection(m_ado.m_OleDbConnection);

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
                    for (x = 0; x <= Tables.FVS.g_strFVSOutTablesArray.Length - 1; x++)
                    {
                        m_intProgressStepCurrentCount++;
                        UpdateTherm(m_frmTherm.progressBar1,
                                    m_intProgressStepCurrentCount,
                                    m_intProgressStepTotalCount);
                        strFVSOutPrePostPathAndDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\PREPOST_" + Tables.FVS.g_strFVSOutTablesArray[x].Trim() + ".ACCDB";
                        if (System.IO.File.Exists(strFVSOutPrePostPathAndDbFile))
                        {
                            System.IO.File.Copy(strFVSOutPrePostPathAndDbFile, strFVSOutPrePostPathAndDbFile + "_" + strFileDate, true);
                        }
                    }

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
                    m_intProgressStepCurrentCount = 0;

                    
                   
                    for (x = 0; x <= intCount - 1; x++)
                    {

                        oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);

                        this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);


                        frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);




                        if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false) == true)
                        {
                            m_bPotFireBaseYearTableExist = true;
                            m_strPotFireBaseYearLinkedTableName = "FVS_POTFIRE_BASEYEAR";
                            m_strPotFireStandardLinkedTableName = "FVS_POTFIRE_STANDARD";
                            m_oPrePostDbFileItem_Collection = new DbFileItem_Collection();
                            

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

                            strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE1, "Text", false);
                            strRx1 = strRx1.Trim();

                            strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE2, "Text", false);
                            strRx2 = strRx2.Trim();

                            strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE3, "Text", false);
                            strRx3 = strRx3.Trim();

                            strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE4, "Text", false);
                            strRx4 = strRx4.Trim();

                            //find the package item in the package collection
                            for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                            {
                                if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim() == strRx1.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim() == strRx2.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim() == strRx3.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim() == strRx4.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
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
                            if (strRx1.Trim().Length > 0 && strRx1.Trim() != "000") this.m_strRxCycleList = "1,";
                            if (strRx2.Trim().Length > 0 && strRx2.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                            if (strRx3.Trim().Length > 0 && strRx3.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                            if (strRx4.Trim().Length > 0 && strRx4.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                            if (this.m_strRxCycleList.Trim().Length > 0)
                                this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

                            this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");
                           

                            //see if this is a different variant
                            //only update the pre-treatment tables when the variant changes
                            if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
                            {
                                bUpdateCondTable = true;
                                strCurVariant = strVariant;
                            }
                            //
                            //FVSOUT_P000-000-000-000-000.MDB
                            //
                            //get the fvs output file. 
                            strOutDirAndFile = this.txtOutDir.Text.Trim() + "\\" + strVariant + "\\" +
                                Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_MDBOUT, "Text", false)).Trim();
                            this.m_strOutMDBFile = Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_MDBOUT, "Text", false)).Trim();

                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "strOutDirAndFile=" + strOutDirAndFile + "  \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "m_strOutMDBFile=" + m_strOutMDBFile + "  \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Processing " + m_strOutMDBFile + "...Stand By");
                            //
                            //FVSOUT_P000-000-000-000-000_BIOSUM.ACCDB
                            //
                            strAuditDbFile = strOutDirAndFile.Replace(".MDB", "_BIOSUM.ACCDB");
                            //
                            //CREATE PREPOST SEQNUM MATRIX TABLES
                            //
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepCurrentCount,
                                         m_intProgressStepTotalCount);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create links to audit db file");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nSTART:Create DAO Audit Table Links " + System.DateTime.Now.ToString() + "\r\n");
                            m_oRxTools.CreateFVSOutputTableLinks(strAuditDbFile, strOutDirAndFile);
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepCurrentCount,
                                         m_intProgressStepTotalCount);
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nEND::Create DAO Audit Table Links " + System.DateTime.Now.ToString() + "\r\n");


                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create POTFIRE table");
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                               this.WriteText(m_strDebugFile, "\r\nSTART:Create POTFIRE tables " + System.DateTime.Now.ToString() + "\r\n");
                            CreatePotFireTables(strAuditDbFile, strOutDirAndFile, strVariant, strPackage);
                             m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                         m_intProgressStepCurrentCount,
                                         m_intProgressStepTotalCount);
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nEND:Create POTFIRE tables " + System.DateTime.Now.ToString() + "\r\n");

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create PREPOST SeqNum Matrix Tables");
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nSTART:Create PrePostSeqNumMatrixTables " + System.DateTime.Now.ToString() + "\r\n");
                            CreatePrePostSeqNumMatrixTables(strAuditDbFile,strPackage, false);
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
                            //
                            //DAO ROUTINES
                            //
                            //
                            //CREATE PREPOST ACCDB AND AND LINKS
                            //
                            
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nSTART: Create PREPOST DbFile table links " + System.DateTime.Now.ToString() + "\r\n");
                            RunAppend_CreatePREPOSTDbFileAndTableLinks(strOutDirAndFile, strAuditDbFile, strVariant);
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "\r\nEND: Create PREPOST DbFile table links " + System.DateTime.Now.ToString() + "\r\n");
                           
                            intItemError = m_dao.m_intError;
                            if (intItemError == 0)
                            {
                                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                    this.WriteText(m_strDebugFile, "\r\nSTART: Open PREPOST DbFile Connections " + System.DateTime.Now.ToString() + "\r\n");
                                RunAppend_OpenDbConnections(m_oPrePostDbFileItem_Collection);
                                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                    this.WriteText(m_strDebugFile, "\r\nEND: Open PREPOST DbFile Connections " + System.DateTime.Now.ToString() + "\r\n");

                               
                            }
                            //m_intProgressStepCurrentCount++;
                            //UpdateTherm(m_frmTherm.progressBar1,
                            //        m_intProgressStepCurrentCount,
                            //        m_intProgressStepTotalCount);
                            //intItemError = m_intError;

                            
                            


                            intItemError = m_dao.m_intError;
                            strItemError = m_dao.m_strError;
                           
                            


                            //
                            //SHOW FILE MONITORS
                            //

                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "checkpoint 6 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                            if (uc_filesize_monitor1.File.Trim().Length == 0)
                            {
                                uc_filesize_monitor1.BeginMonitoringFile(strPotFireBaseYrDbFile, 2000000000, "2GB");
                                uc_filesize_monitor1.Information = "Base year potential fire table for variant " + strVariant;
                            }

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            //UpdateTherm(m_frmTherm.progressBar1,
                            //             m_intProgressStepTotalCount,
                            //             m_intProgressStepTotalCount);
                            
                            //
                            //validation
                            //
                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Validate data");
                                m_intProgressStepCurrentCount = 0;
                                //RunAppend_Validation(x, strVariant, strPackage, strRx1, strRx2, strRx3, strRx4, false, ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning);
                                if (intItemError == 0)
                                {
                                    intItemError = oAdo.m_intError;
                                    strItemError = oAdo.m_strError;
                                }


                            }

                            

                            //UpdateTherm(m_frmTherm.progressBar1,
                            //                m_intProgressStepTotalCount,
                            //                m_intProgressStepTotalCount);

                            System.Threading.Thread.Sleep(1000);

                            m_intProgressStepCurrentCount = 0;
                            //
                            //Table Structure Checks And Edits
                            //
                            if (intItemError == 0)
                            {
                                
                                m_intProgressStepCurrentCount = 0;
                                RunAppend_UpdatePrePostTable(
                                    strPackage, strVariant, strRx1, strRx2, strRx3, strRx4,
                                    bUpdateCondTable, x,
                                    ref intItemError, ref strItemError);

                              

                                if (intItemError == 0)
                                {
                                    intItemError = oAdo.m_intError;
                                    strItemError = oAdo.m_strError;
                                }


                            }
                            UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepTotalCount,
                                            m_intProgressStepTotalCount);
                            System.Threading.Thread.Sleep(1000);

                            //
                            //update FVSTree table
                            //
                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Update " + m_strFvsTreeTable + " table");
                                //RunAppend_UpdateFVSTreeTable(strPackage, strVariant, strRx1, strRx2, strRx3, strRx4, m_strFvsTreeTable, ref intItemError, ref strItemError);

                                if (intItemError == 0)
                                {
                                    intItemError = oAdo.m_intError;
                                    strItemError = oAdo.m_strError;
                                }

                                // Only try to load if there is a cut list in the FVSOut.db
                                if (FvsOutWithCutList())
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
                                ado_data_access oAdoTemp = new ado_data_access();
                                oAdoTemp.OpenConnection(oAdoTemp.getMDBConnString(strOutDirAndFile, "", ""));
                                oAdoTemp.m_strSQL = $@"UPDATE FVS_CASES SET {m_colBioSumAppend} ='Y'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdoTemp.m_strSQL + "\r\n");
                                oAdoTemp.SqlNonQuery(oAdoTemp.m_OleDbConnection, oAdoTemp.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                oAdoTemp.CloseConnection(oAdoTemp.m_OleDbConnection);
                                frmMain.g_oDelegate.SetListViewTextValue(
                                    oLv, x, COL_CHECKBOX, Convert.ToString(frmMain.g_oDelegate.GetListViewTextValue(oLv, x, COL_CHECKBOX, false).Replace("a", "")));
                                frmMain.g_oDelegate.SetListViewTextValue(
                                    oLv, x, COL_CHECKBOX, Convert.ToString(frmMain.g_oDelegate.GetListViewTextValue(oLv, x, COL_CHECKBOX, false).Replace("p", "")));
                            }
                            else
                            {
                                //variant+package processing
                                //if an error for the variant+package combination than delete
                                //all records with this combination
                                RunAppend_DeleteVariantRxPackageFromPrePostTables(strVariant, strPackage);
                                //variant only processing
                                //look to see if the next item in the list is the current variant, if not,
                                //check if any post-treatment records with the current variant, if not,
                                //delete the variant from the pre-treatment tables
                                //if (!bGoodVariant) DeleteVariantFromPreTables(oAdo,oAdo.m_OleDbConnection,x,strSourceTableArray,strVariant);

                            }
                            //close the ado connection in order to use dao to delete table links

                            //oAdo.CloseConnection(oAdo.m_OleDbConnection);

                            RunAppend_CloseDbConnections(m_oPrePostDbFileItem_Collection);

                            UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepTotalCount,
                                           m_intProgressStepTotalCount);

                            System.Threading.Thread.Sleep(1000);
                            
                            //compact and repair
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Compact and Repair");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            System.Threading.Thread.Sleep(5000);
                            m_dao.DisplayErrors = false;
                            for (y = 0; y <= m_oPrePostDbFileItem_Collection.Count - 1; y++)
                            {
                                if (m_oPrePostDbFileItem_Collection.Item(y).DbFileName.Trim().Length > 0)
                                {
                                    if (m_oPrePostDbFileItem_Collection.Item(y).Connection != null &&
                                        m_oPrePostDbFileItem_Collection.Item(y).Connection.State != ConnectionState.Closed)
                                    {
                                        m_ado.CloseConnection(m_oPrePostDbFileItem_Collection.Item(y).Connection);
                                        System.Threading.Thread.Sleep(5000);
                                    }
                                   
                                        m_dao.CompactMDB(m_oPrePostDbFileItem_Collection.Item(y).FullPath.Trim());
                                    
                                        
                                        m_dao.m_intError = 0;
                                        m_dao.m_strError = "";
                                    
                                }
                            }
                            m_dao.DisplayErrors = true;
                            System.Threading.Thread.Sleep(5000);

                            if (intItemError == 0)
                            {
                                intItemError = m_dao.m_intError;
                                strItemError = m_dao.m_strError;
                            }

                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkGreen);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Completed");
                            }
                            else if (intItemError != 0)
                            {
                                m_intError = intItemError;
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Red);
                                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "ERROR:" + strItemError);
                                //MessageBox.Show("ERROR:" + strItemError,"FIA Biosum");

                            }
                            m_intProgressOverallCurrentCount++;
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
                if (m_bDebug)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");
				this.FVSRecordsFinished();
			}
			catch (System.Threading.ThreadInterruptedException err)
			{

				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
				this.ThreadCleanUp();
				
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_output:RunAppend_Main  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FVS Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,1,"Ready");

			CleanupThread();

			frmMain.g_oDelegate.m_oEventThreadStopped.Set();
			this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

		}
        private void RunAppend_OpenDbConnections(DbFileItem_Collection p_oDbFileCollection)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_OpenDbConnections\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Open PREPOST Database Connections");
            m_intProgressStepCurrentCount = 0;
            m_intProgressStepTotalCount = p_oDbFileCollection.Count;
            ado_data_access oAdo = new ado_data_access();
            for (int x = 0; x <= p_oDbFileCollection.Count - 1; x++)
            {
                UpdateTherm(m_frmTherm.progressBar1,
                                   m_intProgressStepCurrentCount,
                                   m_intProgressStepTotalCount);

                if (p_oDbFileCollection.Item(x).FullPath.Trim().Length > 0 &&
                    System.IO.File.Exists(p_oDbFileCollection.Item(x).FullPath))
                {
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nOpening Db Connection:" + p_oDbFileCollection.Item(x).DbFileName + "\r\n\r\n");
                    }
                    p_oDbFileCollection.Item(x).OpenConnection(oAdo);
                    if (oAdo.m_intError != 0) break;

                    
                }
            }
            m_intError = oAdo.m_intError;
            oAdo = null;
        }
        private void RunAppend_CloseDbConnections(DbFileItem_Collection p_oDbFileCollection)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// RunAppend_CloseDbConnections\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            bool bTherm=false;
            if (p_oDbFileCollection != null)
            {
                if (m_frmTherm != null && m_frmTherm.Visible)
                {
                    bTherm = true;
                    m_intProgressStepCurrentCount=0;
                    m_intProgressStepTotalCount = p_oDbFileCollection.Count;
                   
                }
                ado_data_access oAdo = new ado_data_access();
                for (int x = 0; x <= p_oDbFileCollection.Count - 1; x++)
                {
                    if (bTherm)
                    {
                        UpdateTherm(m_frmTherm.progressBar1,
                                    m_intProgressStepCurrentCount,
                                    m_intProgressStepTotalCount);
                    }
                    if (p_oDbFileCollection.Item(x).FullPath.Trim().Length > 0 &&
                        System.IO.File.Exists(p_oDbFileCollection.Item(x).FullPath))
                    {
                        if (p_oDbFileCollection.Item(x).Connection.State == ConnectionState.Open)
                        {
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\nClosing Db Connection:" + p_oDbFileCollection.Item(x).DbFileName + "\r\n\r\n");
                            }
                            oAdo.CloseConnection(p_oDbFileCollection.Item(x).Connection);
                        }

                    }
                }
            }
        }
       
        private void RunAppend_CreatePREPOSTDbFileAndTableLinks(string p_strFVSOutDbFile, 
                                                      string p_strAuditDbFile,
                                                      string p_strVariant)
        {

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_CreatePREPOSTDbFileAndTableLinks\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x,y,z;
            string strLink;
            
            string strTable;
            string strDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db";
                                
            string[] strTempArray = null;
            string[] strSeqNumArray = null;
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                this.WriteText(m_strDebugFile, "checkpoint 7 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

            this.m_dao.getTableNames(p_strFVSOutDbFile, ref strTempArray, false);

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                this.WriteText(m_strDebugFile, "checkpoint 8 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create Table Links");
            m_intProgressStepTotalCount = strTempArray.Length;
            m_intProgressStepCurrentCount = 0;
            //create temporary links to the tables in the fvsoutput db file (FVSOUT_P000-000-000-000-000.MDB)
            for (y = 0; y <= strTempArray.Length - 1; y++)
            {
                m_intProgressStepCurrentCount++;
                UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepCurrentCount,
                            m_intProgressStepTotalCount);
                if (strTempArray == null) break;
                if (strTempArray[y] != null && strTempArray[y].Trim().Length > 3)
                {
                    if (RxTools.ValidFVSTable(strTempArray[y]))
                    {
                            strTable = strTempArray[y].Trim().ToUpper();
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Create Table Links:" + strTempArray[y].Trim());
                       
                            //
                            //CREATE PREPOST_FVS_TABLENAME FILE
                            //
                            if (!System.IO.File.Exists(strDir + "\\PREPOST_" + strTable + ".ACCDB"))
                            {
                                m_dao.CreateMDB(strDir + "\\PREPOST_" + strTable + ".ACCDB");
                            }
                            //
                            //CREATE DBFILEITEM
                            //
                            DbFileItem oDbFileItem = new DbFileItem();
                            oDbFileItem.DbFileName="PREPOST_" + strTable + ".ACCDB";
                            oDbFileItem.Directory = strDir;
                            oDbFileItem.TableType = strTable;
                            //
                            //CREATE TABLE LINK TO FVS OUTPUT TABLE
                            //
                            //create table link to fvs output table
                            TableLinkItem oTableLinkItem = new TableLinkItem();
                            oTableLinkItem.TableName = strTable;
                            oTableLinkItem.FVSOutputTableName = strTable;
                            oTableLinkItem.FVSOutputTable = true;
                            if ((strTable == "FVS_POTFIRE" || 
                                strTable == "FVS_POTFIRE_EAST") &&
                                m_bPotFireBaseYearTableExist ==true)
                            {
                                oTableLinkItem.DbFileName = frmMain.g_oUtils.getFileName(p_strAuditDbFile.Trim());
                                oTableLinkItem.Directory = frmMain.g_oUtils.getDirectory(p_strAuditDbFile.Trim());
                                strLink = oTableLinkItem.DbFileName.Trim() + "_" + strTable;
                            }
                            else
                            {
                               

                                oTableLinkItem.DbFileName = frmMain.g_oUtils.getFileName(p_strFVSOutDbFile.Trim());
                                oTableLinkItem.Directory = frmMain.g_oUtils.getDirectory(p_strFVSOutDbFile.Trim());
                                strLink = oTableLinkItem.DbFileName + "_" + strTable;
                            }
                            strLink = strLink.Replace(".", "_");
                            strLink = strLink.Replace("-", "_");
                            oTableLinkItem.LinkedTableName = strLink;

                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "PREPOSTFile:" + oDbFileItem.DbFileName + " PREPOSTTableLink:" + oTableLinkItem.LinkedTableName + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");
                           
                            m_dao.CreateTableLink(oDbFileItem.FullPath, 
                                                 oTableLinkItem.LinkedTableName,
                                                 oTableLinkItem.FullPath,
                                                 oTableLinkItem.TableName,
                                                 true);
                            oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                            //
                            //CREATE TABLE LINK TO FVS_SUMMARY TABLE
                            //
                            //create a table link to the summary table
                            if (m_dao.m_intError == 0)
                            {
                                if (strTable != "FVS_SUMMARY")
                                {
                                    oTableLinkItem = new TableLinkItem();
                                    oTableLinkItem.DbFileName = frmMain.g_oUtils.getFileName(p_strFVSOutDbFile.Trim());
                                    oTableLinkItem.Directory = frmMain.g_oUtils.getDirectory(p_strFVSOutDbFile.Trim());
                                    strLink = oTableLinkItem.DbFileName + "_FVS_SUMMARY";
                                    strLink = strLink.Replace(".", "_");
                                    strLink = strLink.Replace("-", "_");
                                    oTableLinkItem.LinkedTableName = strLink;
                                    oTableLinkItem.TableName = "FVS_SUMMARY";
                                    oTableLinkItem.FVSOutputTableName = "FVS_SUMMARY";
                                    oTableLinkItem.FVSOutputTable = true;

                                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                        this.WriteText(m_strDebugFile, "PREPOSTFile:" + oDbFileItem.DbFileName + " PREPOSTTableLink:" + oTableLinkItem.LinkedTableName + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                                    m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                         oTableLinkItem.LinkedTableName,
                                                         oTableLinkItem.FullPath,
                                                         oTableLinkItem.TableName,
                                                         true);
                                    oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                                }
                            }
                            //
                            //CREATE TABLE LINK TO FVS OUTPUT PREPOST_SEQNUM_MATRIX TABLES
                            //
                            //create a link to the sequence number matrix table
                            if (m_dao.m_intError == 0)
                            {
                                this.m_dao.getTableNames(p_strAuditDbFile, ref strSeqNumArray, false);
                                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                    this.WriteText(m_strDebugFile, "checkpoint 8 \r\n START: " + System.DateTime.Now.ToString() + "\r\n");


                                for (x = 0; x <= strSeqNumArray.Length - 1; x++)
                                {
                                    if (strSeqNumArray == null) break;
                                    if (strSeqNumArray[x] != null && strSeqNumArray[x].Trim().Length > 3)
                                    {
                                        if (strSeqNumArray[x].Trim().ToUpper().IndexOf("PREPOST_SEQNUM_MATRIX", 0) > 0 &&
                                            strSeqNumArray[x].Trim().ToUpper().IndexOf(strTable, 0) >= 0)
                                        {
                                            oTableLinkItem = new TableLinkItem();
                                            oTableLinkItem.DbFileName = frmMain.g_oUtils.getFileName(p_strAuditDbFile.Trim());
                                            oTableLinkItem.Directory = frmMain.g_oUtils.getDirectory(p_strAuditDbFile.Trim());
                                            strLink = strSeqNumArray[x].Trim() + "_TableLink";
                                            strLink = strLink.Replace(".", "_");
                                            strLink = strLink.Replace("-", "_");
                                            oTableLinkItem.LinkedTableName = strLink;
                                            oTableLinkItem.TableName = strSeqNumArray[x].Trim();
                                            oTableLinkItem.FVSOutputSeqNumMatrixTable = true;
                                            oTableLinkItem.FVSOutputTableName = strTable;

                                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                                this.WriteText(m_strDebugFile, "PREPOSTFile:" + oDbFileItem.DbFileName + " PREPOSTTableLink:" + oTableLinkItem.LinkedTableName + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                                            m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                                    oTableLinkItem.LinkedTableName,
                                                                    oTableLinkItem.FullPath,
                                                                    oTableLinkItem.TableName,
                                                                    true);
                                            oDbFileItem.TableLinkCollection.Add(oTableLinkItem);

                                        }
                                        //create a link to the fvs summary seqnuM matrix table
                                        if (strSeqNumArray[x].Trim().ToUpper() != "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX")
                                        {
                                            oTableLinkItem = new TableLinkItem();
                                            oTableLinkItem.DbFileName = frmMain.g_oUtils.getFileName(p_strAuditDbFile.Trim());
                                            oTableLinkItem.Directory = frmMain.g_oUtils.getDirectory(p_strAuditDbFile.Trim());
                                            strLink = "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX_TableLink";
                                            strLink = strLink.Replace(".", "_");
                                            strLink = strLink.Replace("-", "_");
                                            oTableLinkItem.LinkedTableName = strLink;
                                            oTableLinkItem.TableName = "FVS_SUMMARY_PREPOST_SEQNUM_MATRIX";
                                            oTableLinkItem.FVSOutputSeqNumMatrixTable = true;
                                            oTableLinkItem.FVSOutputTableName = "FVS_SUMMARY";

                                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                                this.WriteText(m_strDebugFile, "PREPOSTFile:" + oDbFileItem.DbFileName + " PREPOSTTableLink:" + oTableLinkItem.LinkedTableName + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                                            m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                                    oTableLinkItem.LinkedTableName,
                                                                    oTableLinkItem.FullPath,
                                                                    oTableLinkItem.TableName,
                                                                    true);
                                            oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                                        }
                                    }
                                    if (m_dao.m_intError > 0) break;
                                }
                            }
                            //
                            //CREATE A TABLE LINK TO THE FVS_TREE TABLE
                            //
                            //fvs_tree table link
                            if (m_dao.m_intError == 0)
                            {
                                //FVS Output CutList Table
                                if (strTable == "FVS_CUTLIST" || strTable=="FVS_TREELIST" || strTable=="FVS_ATRTLIST")
                                {
                                    //FVS_Cases table
                                    if (m_dao.m_intError == 0)
                                    {
                                        oTableLinkItem = new TableLinkItem();
                                        oTableLinkItem.DbFileName = frmMain.g_oUtils.getFileName(p_strFVSOutDbFile.Trim());
                                        oTableLinkItem.Directory = frmMain.g_oUtils.getDirectory(p_strFVSOutDbFile.Trim());
                                        strLink = oTableLinkItem.DbFileName + "_FVS_CASES";
                                        strLink = strLink.Replace(".", "_");
                                        strLink = strLink.Replace("-", "_");
                                        oTableLinkItem.LinkedTableName = strLink;
                                        oTableLinkItem.TableName = "FVS_CASES";
                                        oTableLinkItem.FVSOutputTableName = "FVS_CASES";
                                        oTableLinkItem.FVSOutputTable = true;

                                        if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                            this.WriteText(m_strDebugFile, "PREPOSTFile:" + oDbFileItem.DbFileName + " PREPOSTTableLink:" + oTableLinkItem.LinkedTableName + " \r\n START: " + System.DateTime.Now.ToString() + "\r\n");

                                        m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                             oTableLinkItem.LinkedTableName,
                                                             oTableLinkItem.FullPath,
                                                             oTableLinkItem.TableName,
                                                             true);
                                        oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                                    }
                                    //FIA Tree Table
                                    if (m_dao.m_intError == 0)
                                    {
                                        z = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREE");
                                        oTableLinkItem = new TableLinkItem();
                                        oTableLinkItem.TableName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.TABLE];
                                        oTableLinkItem.DbFileName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.MDBFILE].Trim();
                                        oTableLinkItem.Directory = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.PATH].Trim();

                                        m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                                oTableLinkItem.TableName,
                                                                oTableLinkItem.FullPath,
                                                                oTableLinkItem.TableName,
                                                                true);
                                        oDbFileItem.TableLinkCollection.Add(oTableLinkItem);

                                    }

                                    //
                                    //CREATE A TABLE LINK TO THE REF_SPECIES TABLE
                                    //
                                    //ref_species table link
                                    if (m_dao.m_intError == 0)
                                    {
                                        z = m_oQueries.m_oDataSource.getDataSourceTableNameRow("FIA Tree Species Reference");
                                        oTableLinkItem = new TableLinkItem();
                                        oTableLinkItem.TableName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.TABLE];
                                        oTableLinkItem.LinkedTableName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.TABLE];
                                        oTableLinkItem.DbFileName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.MDBFILE].Trim();
                                        oTableLinkItem.Directory = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.PATH].Trim();
                                        macrosubst oMacroSub = new macrosubst();
                                        oMacroSub.ReferenceGeneralMacroSubstitutionVariableCollection = frmMain.g_oGeneralMacroSubstitutionVariable_Collection;
                                        var strAppDataPath = oMacroSub.GeneralTranslateVariableSubstitution(oTableLinkItem.FullPath);
                                        m_dao.CreateTableLink(oDbFileItem.FullPath, oTableLinkItem.LinkedTableName, strAppDataPath, oTableLinkItem.TableName, true);
                                        oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                                    }
                                }
                               
                            }
                            //CREATE A TABLE LINK TO THE CONDITION TABLE
                            //
                            //create a link to the condition table
                            //create a temporary link to the FIA CONDITION table
                            
                            //condition table link
                            if (m_dao.m_intError == 0)
                            {
                                z = m_oQueries.m_oDataSource.getDataSourceTableNameRow("CONDITION");
                                oTableLinkItem = new TableLinkItem();
                                oTableLinkItem.TableName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.TABLE];
                                oTableLinkItem.DbFileName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.MDBFILE].Trim();
                                oTableLinkItem.Directory = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.PATH].Trim();

                                m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                        oTableLinkItem.TableName,
                                                        oTableLinkItem.FullPath,
                                                        oTableLinkItem.TableName,
                                                        true);
                                oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                            }
                            //
                            //CREATE A TABLE LINK TO THE PLOT TABLE
                            //
                            //plot table link
                            if (m_dao.m_intError == 0)
                            {
                                z = m_oQueries.m_oDataSource.getDataSourceTableNameRow("PLOT");
                                oTableLinkItem = new TableLinkItem();
                                oTableLinkItem.TableName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.TABLE];
                                oTableLinkItem.DbFileName = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.MDBFILE].Trim();
                                oTableLinkItem.Directory = m_oQueries.m_oDataSource.m_strDataSource[z, Datasource.PATH].Trim();
                                m_dao.CreateTableLink(oDbFileItem.FullPath,
                                                        oTableLinkItem.TableName,
                                                        oTableLinkItem.FullPath,
                                                        oTableLinkItem.TableName,
                                                        true);
                                oDbFileItem.TableLinkCollection.Add(oTableLinkItem);
                            }

                            if (m_dao.m_intError == 0)
                            {
                                m_oPrePostDbFileItem_Collection.Add(oDbFileItem);
                            }
                    }
                    
                    
                }
            }
            UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepTotalCount,
                            m_intProgressStepTotalCount);
        }
		private void RunAppend_DeleteVariantRxPackageFromPrePostTables(string p_strVariant, string p_strPackage)
		{

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_DeleteVariantRxPackageFromPrePostTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
             int x, y;
           
           
            System.Data.OleDb.OleDbConnection oConn;
            string strDbFile = "";
            string strFVSOutTable="";
            string strFVSOutTableLink = "";
           
            ado_data_access oAdo = new ado_data_access();

            oAdo.m_intError = 0;
            oAdo.m_strError = "";
            //
            //make sure all the tables and columns exist
            //
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + " Rolling back variant and package records");
            oAdo.m_strSQL = "";
            m_intProgressStepTotalCount = m_oPrePostDbFileItem_Collection.Count;
            m_intProgressStepCurrentCount = 0;
            for (y = 0; y <= m_oPrePostDbFileItem_Collection.Count - 1; y++)
            {
                m_intProgressStepCurrentCount++;
                UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepCurrentCount,
                            m_intProgressStepTotalCount);

                strDbFile = m_oPrePostDbFileItem_Collection.Item(y).DbFileName.Trim().ToUpper();
                if (strDbFile != "PREPOST_FVS_CASES.ACCDB" &&
                    strDbFile != "PREPOST_FVS_TREELIST.ACCDB" &&
                    strDbFile != "PREPOST_FVS_CUTLIST.ACCDB" &&
                    strDbFile != "PREPOST_FVS_ATRTLIST.ACCDB")
                {

                    for (x = 0; x <= m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Count - 1; x++)
                    {
                        TableLinkItem oTableLinkItem = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(x);
                        if (m_oPrePostDbFileItem_Collection.Item(y).TableType.Trim().ToUpper() ==
                            oTableLinkItem.FVSOutputTableName.Trim().ToUpper() && oTableLinkItem.FVSOutputTable)
                        {
                            strFVSOutTable = oTableLinkItem.FVSOutputTableName.Trim().ToUpper();
                            strFVSOutTableLink = oTableLinkItem.LinkedTableName.Trim().ToUpper();

                            oConn = m_oPrePostDbFileItem_Collection.Item(y).Connection;
                            if (oAdo.TableExist(oConn, "POST_" + oTableLinkItem.FVSOutputTableName))
                            {
                                oAdo.m_strSQL = "DELETE FROM POST_" + oTableLinkItem.FVSOutputTableName + " " +
                                    "WHERE FVS_VARIANT='" + p_strVariant.Trim() + "' AND " +
                                    "RXPACKAGE='" + p_strPackage + "'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }
                            if (oAdo.TableExist(oConn, "PRE_" + oTableLinkItem.FVSOutputTableName))
                            {
                                oAdo.m_strSQL = "DELETE FROM PRE_" + oTableLinkItem.FVSOutputTableName + " " +
                                    "WHERE FVS_VARIANT='" + p_strVariant.Trim() + "' AND " +
                                    "RXPACKAGE='" + p_strPackage + "'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }
                        }
                    }
                }
            }
            
		}
		private void DeleteVariantFromPreTables(FIA_Biosum_Manager.ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,int p_intListViewItem,string[] p_strTableArray,string p_strVariant)
		{
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//DeleteVariantFromPreTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
		    int y,z;

			//check to see if the current variant has been saved
			bool bDeleteVariant=true;
			//check if the variant changes for the next item
			for (y=p_intListViewItem+1;y<=this.lstFvsOutput.Items.Count-1;y++)
			{

				if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(this.lstFvsOutput,y,"Checked",false)==true)
				{
					//check if the next item is the same variant
					if (Convert.ToString(frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsOutput,y,COL_VARIANT,"Text",false)).Trim().ToUpper() ==
						p_strVariant.Trim().ToUpper())
					{
						bDeleteVariant=false;
						break;
					}
					else break;
				}
			}
				//delete all records for this variant
			if (bDeleteVariant)
			{
				for (z=0;z<=p_strTableArray.Length-1;z++)
				{
					if (p_strTableArray[z] == null) break;
					if (p_strTableArray[z].Trim().ToUpper() != "FVS_CASES" &&
						p_strTableArray[z].Trim().ToUpper() != "FVS_TREELIST" &&
						p_strTableArray[z].Trim().ToUpper() != "FVS_CUTLIST" &&
						p_strTableArray[z].Trim().ToUpper() != "FVS_ATRTLIST")
					{
						//check to see if there are some valid post treatment records.
						//if valid post-treatment records exist for the current variant 
						//then do not delete the pretreatment records for the same variant.
						//post-treatment records are dependent on pre-treatment records
						if (p_oAdo.TableExist(p_oConn,"POST_" + p_strTableArray[z].Trim()))
						{
							if (p_oAdo.ValuesExistEqualToTargetValue(p_oConn,"POST_" + p_strTableArray[z].Trim(),"fvs_variant",p_strVariant.Trim(),false))
								bDeleteVariant=false;
						}
						if (bDeleteVariant)
						{
							if (p_oAdo.TableExist(p_oConn,"PRE_" + p_strTableArray[z].Trim()))
							{
								//check to see if there are other post treatment records

								//delete any rows with the current variant
								p_oAdo.m_strSQL = "DELETE FROM PRE_" + p_strTableArray[z].Trim() + " " + 
									"WHERE FVS_VARIANT='" + p_strVariant.Trim() + "'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
								p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
							}
						}
					}
										
				}
			}
			
		}
        private void PotentialFireBaseYear(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,
                                        List<string> p_strFVSOutLinkedTableNames,
                                        ref int p_intError, ref string p_strError)
        {
            string strCasesTableName = "";
            string strPotFireTableName = "";
            int y;
            bool bCASEIDNumeric=true;

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//PotentialFireBaseYear\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            try
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Process Potential Fire Base Year");
                m_intProgressStepCurrentCount = 0;
                m_intProgressStepTotalCount = 6;
                strCasesTableName = "";
                strPotFireTableName = "";
                //find the cases table link
                for (y = 0; y <= p_strFVSOutLinkedTableNames.Count - 1; y++)
                {
                    if (p_strFVSOutLinkedTableNames[y] == null) break;
                    if (p_strFVSOutLinkedTableNames[y].ToUpper().IndexOf("FVS_CASES", 0) > 0)
                    {
                        strCasesTableName = p_strFVSOutLinkedTableNames[y].Trim();

                    }
                    else if (p_strFVSOutLinkedTableNames[y].ToUpper().IndexOf("FVS_POTFIRE", 0) > 0)
                    {
                        strPotFireTableName = p_strFVSOutLinkedTableNames[y].Trim();

                    }
                    if (strCasesTableName != "" && strPotFireTableName != "") break;

                }
                //
                //test if caseid is numeric or text
                //
                System.Data.DataTable dtCaseIdSchema = p_oAdo.getTableSchema(p_oAdo.m_OleDbConnection, "SELECT TOP 1 CASEID FROM " + strCasesTableName);
                if (p_oAdo.m_intError == 0)
                {
                    if ((string)p_oAdo.getIsTheFieldAStringDataType(dtCaseIdSchema.Rows[0]["caseid"].ToString()) == "N")
                        bCASEIDNumeric = true;
                    else
                        bCASEIDNumeric = false;
                    if (bCASEIDNumeric)
                    {
                        if (m_bDebug && frmMain.g_intDebugLevel > 1)
                            this.WriteText(m_strDebugFile, "--copy each potfire base year record into potfire_baseyr_work_table--\r\n");
                        //copy distinct potfire base year records to work table
                        //create the work table of the last caseid (last fvs run) for each stand
                        p_oAdo.m_strSQL = "SELECT a.* INTO potfire_baseyr_work_table FROM PotFire_BaseYr a, " +
                                            "(SELECT MAX(caseid) AS maxcaseid,standid " +
                                             "FROM PotFire_BaseYr GROUP BY standid) b " +
                                       "WHERE a.caseid=b.maxcaseid AND a.standid=b.standid ";
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    }
                    else
                    {
                        if (m_bDebug && frmMain.g_intDebugLevel > 1)
                            this.WriteText(m_strDebugFile, "--copy each potfire base year record into potfire_baseyr_work_table--\r\n");
                        //copy distinct potfire base year records to work table
                        //create the work table of the last rundatetime (last fvs run) for each stand
                        p_oAdo.m_strSQL = "SELECT  a.* INTO potfire_baseyr_work_table FROM PotFire_BaseYr a," +
                                             "(SELECT MAX(RUNDATETIME) AS LASTRUN, STANDID " +
                                                    "FROM " + strCasesTableName + " " +
                                                    "GROUP BY STANDID)  b," +
                                             "(SELECT * " +
                                          "FROM " + strCasesTableName + ")  c " +
                                          "WHERE a.STANDID=b.standid AND " +
                                                "a.standid = c.standid AND " +
                                                "b.standid=c.standid AND " +
                                                "c.rundatetime = b.lastrun AND " +
                                                "a.caseid = c.caseid";
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    }
                }

                if (p_oAdo.m_intError == 0)
                {
                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepCurrentCount,
                                m_intProgressStepTotalCount);
                    //remove stands from the potfire base year work table that do not exist in the FVS_POTFIRE table
                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "--remove stands from the potfire base year work table that do not exist in the FVS_POTFIRE table--\r\n");
                    p_oAdo.m_strSQL = "SELECT a.* INTO potfire_baseyr_work_table2 FROM potfire_baseyr_work_table a," +
                                        "(SELECT MIN(year) AS minyear,standid " +
                                         "FROM potfire_baseyr_work_table GROUP BY standid) b," +
                                            "(SELECT DISTINCT standid FROM " + strPotFireTableName + ") c " +
                                     "WHERE a.standid=b.standid AND a.year=b.minyear AND c.standid=a.standid";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                }

                if (p_oAdo.m_intError == 0)
                {
                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepCurrentCount,
                                m_intProgressStepTotalCount);

                    //update the work table with the destination potfire case id
                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "--update the work table with the destination potfire case id--\r\n");
                    p_oAdo.m_strSQL = "UPDATE potfire_baseyr_work_table2 a " +
                                    "INNER JOIN " + strPotFireTableName + " b " +
                                    "ON a.standid = b.standid " +
                                    "SET a.caseid=b.caseid";
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                }

                if (p_oAdo.m_intError == 0)
                {

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepCurrentCount,
                                m_intProgressStepTotalCount);

                    //initialize for any failed transactions
                    if (m_bDebug && frmMain.g_intDebugLevel > 1)
                        this.WriteText(m_strDebugFile, "--initialize for any failed transactions--\r\n");
                    p_oAdo.m_strSQL = "UPDATE " + strPotFireTableName + " p " +
                                        "INNER JOIN " + strCasesTableName + " c " +
                                        "ON p.caseID = c.caseID AND " +
                                        "   p.standID = c.standID " +
                                        "SET p.year = p.year - 1," +
                                            "c.BIOSUM_PotFireOneYearAdded_YN='N' " + 
                                      "WHERE c.BIOSUM_PotFireOneYearAdded_YN='*'";

                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                    m_intProgressStepCurrentCount++;
                    UpdateTherm(m_frmTherm.progressBar1,
                                m_intProgressStepCurrentCount,
                                m_intProgressStepTotalCount);
                }
                if (p_oAdo.m_intError == 0)
                {

                    //see if any stands to add year + 1
                    p_oAdo.m_strSQL = "SELECT COUNT(*) as rowcount FROM " +
                         strCasesTableName + " WHERE BIOSUM_PotFireOneYearAdded_YN='N'";
                    if ((int)p_oAdo.getRecordCount(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL, "FVS_CASES") > 0)
                    {
                        //add 1 to the year
                        if (m_bDebug && frmMain.g_intDebugLevel > 1)
                            this.WriteText(m_strDebugFile, "--add 1 to the year--\r\n");
                        p_oAdo.m_strSQL = "UPDATE " + strPotFireTableName + " p " +
                                            "INNER JOIN " + strCasesTableName + " c " +
                                            "ON p.caseID = c.caseID AND " +
                                            "   p.standID = c.standID " +
                                            "SET p.year = p.year + 1," +
                                                "c.BIOSUM_PotFireOneYearAdded_YN='*' " + 
                                          "WHERE c.BIOSUM_PotFireOneYearAdded_YN='N'";

                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                        if (p_oAdo.m_intError == 0)
                        {
                            //COLUMN ID processing
                            //FVS changed the CASEID datatype at the same time removing the ID column
                            if (bCASEIDNumeric)
                            {
                                //since the id is a unique key field the 
                                //pot fire base yr id will need to change
                                //so that every id column will be unique in the FVS_POTFIRE table.
                                //First get the MAX id in the FVS_POTFIRE table
                                //and add 1 to every row of records for a 
                                //sequential count of records.
                                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                    this.WriteText(m_strDebugFile, "--calculate unique values for ID field--\r\n");
                                p_oAdo.m_strSQL = "SELECT  (SELECT COUNT(a.ID) AS COUNTID " +
                                                           "FROM potfire_baseyr_work_table2 a " +
                                                           "WHERE a.id<=b.id) AS sequential_count, " +
                                                           "b.id, c.maxid + sequential_count AS new_id," +
                                                           "b.standid,b.year " +
                                                  "INTO potfire_baseyr_newid_value_work_table FROM potfire_baseyr_work_table2 b," +
                                                    "(SELECT MAX(id) AS maxid FROM " + strPotFireTableName + ") c " +
                                                  "ORDER BY b.ID";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                                p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                            }
                           

                        }
                        if (p_oAdo.m_intError == 0)
                        {
                             //COLUMN ID processing
                            //FVS changed the CASEID datatype at the same time removing the ID column
                            if (bCASEIDNumeric)
                            {
                                //assign the new id values
                                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                    this.WriteText(m_strDebugFile, "--assign the new ID values--\r\n");
                                p_oAdo.m_strSQL = "UPDATE potfire_baseyr_work_table2 a " +
                                                  "INNER JOIN potfire_baseyr_newid_value_work_table b " +
                                                  "ON a.standid=b.standid AND a.year=b.year " +
                                                  "SET a.id=b.new_id";

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                                p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                            }

                        }

                        if (p_oAdo.m_intError == 0)
                        {
                             //COLUMN ID processing
                            //FVS changed the CASEID datatype at the same time removing the ID column
                           
                                //insert the potfire work table records
                                if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                    this.WriteText(m_strDebugFile, "--insert the potfire work table records--\r\n");
                                p_oAdo.m_strSQL = "INSERT INTO " + strPotFireTableName + " " +
                                                "SELECT a.* FROM potfire_baseyr_work_table2 a " +
                                                "INNER JOIN " + strCasesTableName + " b " +
                                                "ON a.caseid=b.caseid AND a.standid=b.standid " +
                                                "WHERE b.BIOSUM_PotFireOneYearAdded_YN='*'";
                            
                           
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        }


                        if (p_oAdo.m_intError==0)
                        {
                            if (m_bDebug && frmMain.g_intDebugLevel > 1)
                                this.WriteText(m_strDebugFile, "--finalize the transaction--\r\n");
                            p_oAdo.m_strSQL = "UPDATE " + strCasesTableName + " " +
                                              "SET BIOSUM_PotFireOneYearAdded_YN='Y' " +
                                            "WHERE BIOSUM_PotFireOneYearAdded_YN='*'";

                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                        }

                        m_intProgressStepCurrentCount++;
                        UpdateTherm(m_frmTherm.progressBar1,
                                    m_intProgressStepCurrentCount,
                                    m_intProgressStepTotalCount);
                    }




                    else
                    {
                        //update the potfire records
                        string strColumns = p_oAdo.getFieldNames(p_oAdo.m_OleDbConnection, "SELECT * FROM " + strPotFireTableName);
                        string[] strColumnsArray = frmMain.g_oUtils.ConvertListToArray(strColumns, ",");
                        strColumns = "";
                        for (y = 0; y <= strColumnsArray.Length - 1; y++)
                        {
                            if (strColumnsArray[y] != null &&
                                strColumnsArray[y].Trim().Length > 0)
                            {
                                if (strColumnsArray[y].Trim().ToUpper() != "CASEID" &&
                                    strColumnsArray[y].Trim().ToUpper() != "STANDID" &&
                                    strColumnsArray[y].Trim().ToUpper() != "YEAR" && 
                                    strColumnsArray[y].Trim().ToUpper() != "ID")
                                {
                                    strColumns = strColumns + "a." + strColumnsArray[y].Trim() + "=" +
                                                              "b." + strColumnsArray[y].Trim() + ",";

                                }
                            }
                        }
                        if (m_bDebug && frmMain.g_intDebugLevel > 1)
                            this.WriteText(m_strDebugFile, "--update the potfire records with the potfire base year records--\r\n");
                        strColumns = strColumns.Substring(0, strColumns.Length - 1);
                        p_oAdo.m_strSQL = "UPDATE " + strPotFireTableName + " a " +
                                        "INNER JOIN potfire_baseyr_work_table2 b " +
                                        "ON a.caseid=b.caseid AND a.standid=b.standid AND a.year=b.year " +
                                        "SET " + strColumns;
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                        m_intProgressStepCurrentCount++;
                        UpdateTherm(m_frmTherm.progressBar1,
                                    m_intProgressStepCurrentCount,
                                    m_intProgressStepTotalCount);


                    }
                }
                p_intError = p_oAdo.m_intError;
                m_intProgressStepCurrentCount++;
                UpdateTherm(m_frmTherm.progressBar1,
                            m_intProgressStepTotalCount,
                            m_intProgressStepTotalCount);
            }
            catch (Exception err)
            {
                p_intError = -1;
                p_strError = err.Message;

            }
           

        }
        private void RunAppend_UpdateFVSTreeTable(string p_strPackage,
                                                  string p_strVariant,
                                                  string p_strRx1,
                                                  string p_strRx2,
                                                  string p_strRx3,
                                                  string p_strRx4,
                                                  string strFvsTreeTable,
                                                  ref int p_intError,
                                                  ref string p_strError)
        {
            int x, y, z, xx, yy;







            string strColumns = "";
            string strValues = "";
            string strRx = "";
            string strCycle = "";

            System.Data.OleDb.OleDbConnection oConn;
            string strDbFile = "";
            string strFVSOutTable = "";
            string strFVSOutTableLink = "";
            string strFvsOutSummaryLink = "";
            string strFiaTreeSpeciesRefTableLink = "";
            string strFVSSummarySeqNumMtxTableLink = "";
            string strFVSOutSeqNumMatrixTableLink = "";
            string strCasesTable = "";
            string strConn = "";
            bool bIdColumnExist = false;
            

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_UpdateFVSTreeTable\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            ado_data_access oAdo = new ado_data_access();

            oAdo.m_intError = 0;
            oAdo.m_strError = "";
            //
            //make sure all the tables and columns exist
            //
            oAdo.m_strSQL = "";

            for (y = 0; y <= m_oPrePostDbFileItem_Collection.Count - 1; y++)
            {
                strDbFile = m_oPrePostDbFileItem_Collection.Item(y).DbFileName.Trim().ToUpper();
                if (strDbFile == "PREPOST_FVS_CUTLIST.ACCDB")
                {
                    for (x = 0; x <= m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Count - 1; x++)
                    {
                        TableLinkItem oTableLinkItem = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(x);
                        if (m_oPrePostDbFileItem_Collection.Item(y).TableType.Trim().ToUpper() ==
                            oTableLinkItem.FVSOutputTableName.Trim().ToUpper() && oTableLinkItem.FVSOutputTable)
                        {
                            strFVSOutTable = oTableLinkItem.FVSOutputTableName.Trim().ToUpper();
                            strFVSOutTableLink = oTableLinkItem.LinkedTableName.Trim().ToUpper();

                            oConn = m_oPrePostDbFileItem_Collection.Item(y).Connection;

                            m_intProgressStepCurrentCount = 0;
                            if (m_strRxCycleArray == null)
                                m_strRxCycleArray = new string[1];

                            m_intProgressStepTotalCount = 8 + (m_strRxCycleArray.Length * 5);
                            /**************************************************************
                             **delete records in the fvs_tree table that have the current
                             **package
                             **************************************************************/
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Variant:" + p_strVariant.Trim() + " Package:" + p_strPackage.Trim() + ": Deleting Old Package+Variant FVS Tree Table Records");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Refresh");
                            //
                            //delete package from fvs out tree table
                            //
                            oAdo.m_strSQL = "DELETE FROM fvs_tree " +
                                "WHERE RXPACKAGE='" + p_strPackage.Trim() + "'";

                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + m_ado.m_strSQL + "\r\n");

                            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                this.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            if (p_intError == 0)
                            {

                                for (z = 0; z <= m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Count - 1; z++)
                                {
                                    if (m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).TableName == "FVS_CASES")
                                    {
                                        strCasesTable = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).LinkedTableName;
                                    }
                                    else if (m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).TableName == "FVS_SUMMARY")
                                    {
                                        strFvsOutSummaryLink = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).LinkedTableName;
                                    }
                                    else if (m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).TableName ==
                                             m_oQueries.m_oDataSource.getValidDataSourceTableName(Datasource.TableTypes.FiaTreeSpeciesReference))
                                    {
                                        strFiaTreeSpeciesRefTableLink = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).LinkedTableName;
                                    }
                                }

                                if (oAdo.m_intError == 0)
                                {
                                    GetPrePostSeqNumConfiguration(strFVSOutTable, p_strPackage);
                                }
                                //
                                //GET THE SEQNUM MATRIX TABLE
                                //
                                GetPrePostTableLinkItems(
                                        m_oPrePostDbFileItem_Collection,
                                        "PREPOST_" + strFVSOutTable + ".ACCDB",
                                        strFVSOutTable,
                                        ref strFVSOutSeqNumMatrixTableLink,
                                        ref strFVSSummarySeqNumMtxTableLink,
                                        ref strFVSOutTableLink);

                                if (oAdo.TableExist(oConn, "cutlist_rowid_work_table"))
                                    oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_rowid_work_table");


                                //check fvs version
                                bIdColumnExist = oAdo.ColumnExist(oConn, strFVSOutTableLink, "ID");
                                //create id column
                                if (!bIdColumnExist && (int)oAdo.getRecordCount(oConn, "SELECT COUNT(*) FROM (SELECT TOP 1 t.standid FROM " + strFVSOutTableLink + " t)", strFVSOutTableLink) > 0)
                                {
                                    //create structure
                                    oAdo.m_strSQL = "SELECT TOP 1 caseid,standid,year,treeid,treeindex INTO cutlist_rowid_work_table FROM " + strFVSOutTableLink;
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    //delete the one record
                                    oAdo.m_strSQL = "DELETE FROM cutlist_rowid_work_table";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    //create autonumber column
                                    oAdo.AddColumn(oConn, "cutlist_rowid_work_table", "rowid", "LONG", "");
                                    oAdo.AddAutoNumber(oConn, "cutlist_rowid_work_table", "rowid");
                                    //append all the cutlist records into the work table
                                    oAdo.m_strSQL = "INSERT INTO cutlist_rowid_work_table SELECT caseid,standid,year,treeid,treeindex FROM " + strFVSOutTableLink;
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                }


                                //
                                //loop through all the rx cycles for this package and append them to the fvs tree
                                //
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


                                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Rx:" + strRx.Trim() + " Cycle:" + strCycle + ": Insert Cutlist Tree Records");
                                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                                        if (oAdo.TableExist(oConn, "cutlist_fia_trees_work_table"))
                                            oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_fia_trees_work_table");

                                        if (oAdo.TableExist(oConn, "cutlist_fvs_created_seedlings_work_table"))
                                            oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_fvs_created_seedlings_work_table");

                                        //
                                        //FIA TREES
                                        //
                                        //make sure there are records to insert
                                        oAdo.m_strSQL = "SELECT COUNT(*) FROM " +
                                                         "(SELECT TOP 1 c.standid " +
                                                          "FROM " + strCasesTable + " c," + strFVSOutTableLink + " t " +
                                                          "WHERE c.CaseID = t.CaseID AND MID(t.treeid, 1, 2) NOT IN ('ES','CM'))";


                                        if ((int)oAdo.getRecordCount(oConn, oAdo.m_strSQL, "temp") > 0)
                                        {
                                            oAdo.m_strSQL = "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strPackage.Trim() + "' AS rxpackage," +
                                                "'" + strRx.Trim() + "' AS rx,'" + strCycle.Trim() + "' AS rxcycle," +
                                                "CSTR(t.year) AS rxyear," +
                                                "c.Variant AS fvs_variant, " +
                                                "Trim(t.treeid) AS fvs_tree_id," +
                                                "t.Species AS fvs_species, t.TPA, ROUND(t.DBH,1) AS DBH , t.Ht,t.estht,t.pctcr,t.TCuFt,'N' AS FvsCreatedTree_YN," +
                                                "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                                "INTO cutlist_fia_trees_work_table " +
                                                "FROM " + strCasesTable + " c," + strFVSOutTableLink + " t " +
                                                "WHERE c.CaseID = t.CaseID AND MID(t.treeid, 1, 2) NOT IN ('ES','CM') ";

                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                            //insert into fvs tree table
                                            oAdo.m_strSQL = "INSERT INTO " + strFvsTreeTable + " " +
                                                                 "(biosum_cond_id, rxpackage,rx,rxcycle,rxyear,fvs_variant, fvs_tree_id," +
                                                                  "fvs_species, tpa, dbh, ht, estht,pctcr,FvsCreatedTree_YN,DateTimeCreated) " +
                                                                    "SELECT a.biosum_cond_id, a.rxpackage,a.rx,a.rxcycle,a.rxyear,a.fvs_variant," +
                                                                           "a.fvs_tree_id, a.fvs_species, a.tpa, a.dbh, a.ht, a.estht,a.pctcr," +
                                                                           "a.FvsCreatedTree_YN,a.DateTimeCreated  " +
                                                                    "FROM cutlist_fia_trees_work_table a," +
                                                                        "(SELECT standid,year " +
                                                                         "FROM " + strFVSOutSeqNumMatrixTableLink + " " +
                                                                         "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  b " +
                                                                    "WHERE TRIM(a.biosum_cond_id)=TRIM(b.standid) AND CINT(a.rxyear)=b.year";

                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        }
                                        p_intError = oAdo.m_intError;

                                        m_intProgressStepCurrentCount++;
                                        UpdateTherm(m_frmTherm.progressBar1,
                                                   m_intProgressStepCurrentCount,
                                                   m_intProgressStepTotalCount);

                                        //
                                        //FVS CREATED SEEDLING TREES
                                        //
                                        //make sure there are records to insert
                                        oAdo.m_strSQL = "SELECT COUNT(*) FROM " +
                                                         "(SELECT TOP 1 c.standid " +
                                                          "FROM " + strCasesTable + " c," + strFVSOutTableLink + " t " +
                                                          "WHERE c.CaseID = t.CaseID AND MID(t.treeid, 1, 2) = 'ES' )";
                                        if ((int)oAdo.getRecordCount(oConn, oAdo.m_strSQL, "temp") > 0)
                                        {
                                            if (bIdColumnExist)
                                            {
                                                //FVS CREATED SEEDLING TREES 
                                                oAdo.m_strSQL =
                                                   "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strPackage.Trim() + "' AS rxpackage," +
                                                   "'" + strRx.Trim() + "' AS rx,'" + strCycle.Trim() + "' AS rxcycle," +
                                                   "CSTR(t.year) AS rxyear," +
                                                   "c.Variant AS fvs_variant, IIf(LEN(TRIM(CSTR(t.id)))=1," +
                                                   "c.variant+'ES00000'+TRIM(CSTR(t.id)),IIf(LEN(TRIM(CSTR(t.id)))=2," +
                                                   "c.variant+'ES0000'+TRIM(CSTR(t.id)),IIf(LEN(TRIM(CSTR(t.id)))=3," +
                                                   "c.variant+'ES000'+TRIM(CSTR(t.id)),IIf(LEN(TRIM(CSTR(t.id)))=4," +
                                                   "c.variant+'ES00'+TRIM(CSTR(t.id)),IIf(LEN(TRIM(CSTR(t.id)))=5," +
                                                   "c.variant+'ES0'+TRIM(CSTR(t.id)),c.variant+'ES'+TRIM(CSTR(t.id))))))) AS fvs_tree_id," +
                                                   "t.Species AS fvs_species, t.TPA, ROUND(t.DBH,1) AS dbh , t.Ht,t.estht,t.pctcr,t.TCuFt,'Y' AS FvsCreatedTree_YN," +
                                                   "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                                   "INTO cutlist_fvs_created_seedlings_work_table " +
                                                   "FROM " + strCasesTable + " c," + strFVSOutTableLink + " t " +
                                                   "WHERE c.CaseID = t.CaseID AND MID(t.treeid, 1, 2) = 'ES' ";
                                            }
                                            else
                                            {
                                                //FVS CREATED SEEDLING TREES 
                                                oAdo.m_strSQL =
                                                   "SELECT DISTINCT c.StandID AS biosum_cond_id,'" + p_strPackage.Trim() + "' AS rxpackage," +
                                                   "'" + strRx.Trim() + "' AS rx,'" + strCycle.Trim() + "' AS rxcycle," +
                                                   "CSTR(t.year) AS rxyear," +
                                                   "c.Variant AS fvs_variant, IIf(LEN(TRIM(CSTR(r.rowid)))=1," +
                                                   "c.variant+'ES00000'+TRIM(CSTR(r.rowid)),IIf(LEN(TRIM(CSTR(r.rowid)))=2," +
                                                   "c.variant+'ES0000'+TRIM(CSTR(r.rowid)),IIf(LEN(TRIM(CSTR(r.rowid)))=3," +
                                                   "c.variant+'ES000'+TRIM(CSTR(r.rowid)),IIf(LEN(TRIM(CSTR(r.rowid)))=4," +
                                                   "c.variant+'ES00'+TRIM(CSTR(r.rowid)),IIf(LEN(TRIM(CSTR(r.rowid)))=5," +
                                                   "c.variant+'ES0'+TRIM(CSTR(r.rowid)),c.variant+'ES'+TRIM(CSTR(r.rowid))))))) AS fvs_tree_id," +
                                                   "t.Species AS fvs_species, t.TPA, ROUND(t.DBH,1) AS dbh , t.Ht,t.estht,t.pctcr,t.TCuFt,'Y' AS FvsCreatedTree_YN," +
                                                   "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                                   "INTO cutlist_fvs_created_seedlings_work_table " +
                                                   "FROM " + strCasesTable + " c," + strFVSOutTableLink + " t,cutlist_rowid_work_table r " +
                                                   "WHERE c.CaseID = t.CaseID AND MID(t.treeid, 1, 2) = 'ES' AND " +
                                                         "(r.CaseId = t.CaseId AND r.StandId = t.StandId AND r.year = t.year AND r.treeid = t.treeid AND r.treeindex = t.treeindex) AND " +
                                                         "(r.CaseId = c.CaseId)";
                                            }

                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                            oAdo.m_strSQL = "INSERT INTO " + strFvsTreeTable + " " +
                                                                 "(biosum_cond_id, rxpackage,rx,rxcycle,rxyear,fvs_variant, fvs_tree_id," +
                                                                  "fvs_species, tpa, dbh, ht, estht,pctcr,FvsCreatedTree_YN,DateTimeCreated) " +
                                                                    "SELECT a.biosum_cond_id, a.rxpackage,a.rx,a.rxcycle,a.rxyear,a.fvs_variant," +
                                                                           "a.fvs_tree_id, a.fvs_species, a.tpa, a.dbh, a.ht, a.estht,a.pctcr," +
                                                                           "a.FvsCreatedTree_YN,a.DateTimeCreated  " +
                                                                    "FROM cutlist_fvs_created_seedlings_work_table a," +
                                                                        "(SELECT standid,year " +
                                                                         "FROM " + strFVSOutSeqNumMatrixTableLink + " " +
                                                                         "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  b " +
                                                                    "WHERE TRIM(a.biosum_cond_id)=TRIM(b.standid) AND CINT(a.rxyear)=b.year";
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        }

                                        if (oAdo.TableExist(oConn, "cutlist_fia_trees_work_table"))
                                            oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_fia_trees_work_table");

                                        if (oAdo.TableExist(oConn, "cutlist_fvs_created_seedlings_work_table"))
                                            oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_fvs_created_seedlings_work_table");

                                        if (oAdo.TableExist(oConn, "cutlist_save_tree_species_work_table"))
                                            oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_save_tree_species_work_table");

                                        p_intError = oAdo.m_intError;


                                    }
                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                               m_intProgressStepCurrentCount,
                                               m_intProgressStepTotalCount);

                                }
                                //
                                //SWAP FVS-ONLY SPECIES CODES TO FIA SPECIES CODES IN ORDER FOR FIA TO CALCULATE VOLUMES
                                //
                                //check if original spcd column exists
                                //save the original spcd
                                if ((int)oAdo.getRecordCount(oConn, "SELECT COUNT(*) AS ROWCOUNT FROM " + strFvsTreeTable + " WHERE TRIM(FVS_SPECIES) IN ('298','999')", strFvsTreeTable) > 0)
                                {
                                    oAdo.m_strSQL = "SELECT BIOSUM_COND_ID, FVS_TREE_ID, FVS_SPECIES INTO cutlist_save_tree_species_work_table FROM " + strFvsTreeTable + " WHERE TRIM(FVS_SPECIES) IN ('298','999')";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    //update the FVS_SPECIES with FIA SPCD for 298,999 trees
                                    oAdo.m_strSQL = "UPDATE " + strFvsTreeTable + " " +
                                                    "SET FVS_SPECIES = IIF(FVS_SPECIES='298','299'," +
                                                                      "IIF(FVS_SPECIES='999','998',FVS_SPECIES)) " +
                                                    "WHERE FVS_SPECIES IN ('298','999')";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }

                                //
                                //UPDATE VOLUME COLUMNS
                                //
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1, m_intProgressStepCurrentCount, m_intProgressStepTotalCount);

                                //update growth projected trees with tree volumes
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Prepare data for volume calculation");
                                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                                if (oAdo.TableExist(oConn, Tables.VolumeAndBiomass.BiosumVolumesInputTable))
                                    oAdo.SqlNonQuery(oConn, "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(oAdo, oConn, Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step1(
                                                   Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                   strFvsTreeTable, p_strPackage);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);

                                //join plot, cond, and tree table to oracle input tree volumes table.
                                //NOTE: this query handles existing FIADB trees that have been grown forward.
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2(
                                                   Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                   m_oQueries.m_oFIAPlot.m_strTreeTable,
                                                   m_oQueries.m_oFIAPlot.m_strPlotTable,
                                                   m_oQueries.m_oFIAPlot.m_strCondTable);





                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);


                                //Set DIAHTCD for FIADB Cycle<>1 trees to their Cycle=1 DIAHTCD values
                                oAdo.m_strSQL = $@"UPDATE {Tables.VolumeAndBiomass.BiosumVolumesInputTable} b 
                                                   INNER JOIN {m_oQueries.m_oFIAPlot.m_strTreeTable} t 
                                                   ON t.biosum_cond_id=b.biosum_cond_id AND t.fvs_tree_id=b.fvs_tree_id
                                                   SET b.diahtcd=t.diahtcd WHERE b.fvscreatedtree_yn='N'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //Set DIAHTCD for FVS-Created trees using FIA_TREE_SPECIES_REF.WOODLAND_YN
                                oAdo.m_strSQL = $@"UPDATE {Tables.VolumeAndBiomass.BiosumVolumesInputTable} b 
                                                   INNER JOIN {strFiaTreeSpeciesRefTableLink} ref ON cint(b.spcd)=ref.spcd
                                                   SET b.diahtcd=IIF(ref.woodland_yn='N', 1, 2) WHERE b.fvscreatedtree_yn='Y'";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //Update FVSCreatedTrees balive=fvs_summary.BA
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2a(
                                                   Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                   m_oQueries.m_oFIAPlot.m_strPlotTable,
                                                   m_oQueries.m_oFIAPlot.m_strCondTable, 
                                                   strFvsOutSummaryLink);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //join cond table to oracle input tree volumes table.
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step3(
                                                  Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                  m_oQueries.m_oFIAPlot.m_strCondTable);



                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");



                                //populate treeclcd column
                                if (oAdo.TableExist(oConn, "CULL_TOTAL_WORK_TABLE"))
                                    oAdo.SqlNonQuery(oConn, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step4(
                                                  "cull_total_work_table",
                                                  Tables.VolumeAndBiomass.BiosumVolumesInputTable);


                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);


                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                                                    "cull_total_work_table",
                                                    Tables.VolumeAndBiomass.BiosumVolumesInputTable);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);

                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                                                  "cull_total_work_table",
                                                  Tables.VolumeAndBiomass.BiosumVolumesInputTable);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);
                                
                                int intTotalRecs = (int)oAdo.getSingleDoubleValueFromSQLQuery(oConn, "SELECT COUNT(*) AS TTLCOUNT FROM " + Tables.VolumeAndBiomass.BiosumVolumesInputTable, "temp");
                                    if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase) == false)
                                    {
                                        m_intError = -1;
                                        m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase + " not found";
                                    }
                                    if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR")==false)
                                    {
                                        m_intError=-1;
                                        m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\BioSumComps.JAR not found";
                                    }
                                    if (m_intError == 0 && System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat")==false)
                                    {
                                        m_intError=-1;
                                        m_strError = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\fcs_tree_calc.bat not found";
                                    }
                                    if (m_intError == 0)
                                    {
                                        int COUNT = 0;
                                        SQLite.ADO.DataMgr oSQLite = new SQLite.ADO.DataMgr();

                                        _SQLite = oSQLite;
                                        _MSAccess = oAdo;
                                        var columnsAndDataTypes = Tables.VolumeAndBiomass.ColumnsAndDataTypes;
                                        strColumns = string.Join(",", columnsAndDataTypes.Select(item => item.Item1));
                                        //
                                        //CONNECT TO SQLITE AND REMOVE DATA FROM SQLITE DB
                                        //
                                        oSQLite.OpenConnection(false, 1, frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase, "BIOSUM");

                                        oSQLite.SqlNonQuery(oSQLite.m_Connection, $"DELETE FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");

                                        if (oAdo.TableExist(oConn, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable))
                                            oAdo.SqlNonQuery(oConn, "DROP TABLE " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);

                                        frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(oAdo, oConn, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);



                                        UpdateTherm(m_frmTherm.progressBar1,
                                                    m_intProgressStepTotalCount,
                                                    m_intProgressStepTotalCount);

                                        System.Threading.Thread.Sleep(2000);

                                        //insert records 
                                        //from 
                                        //table biosum_volumes_input (BiosumVolumesInputTable)
                                        //into 
                                        //table fcs_biosum_volumes_input (FcsBiosumVolumesInputTable)
                                        oAdo.m_strSQL =
                                            Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step7(
                                                    Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                    Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);

                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        //
                                        //INSERT ALL THE MS ACCESS DATA INTO SQLITE
                                        //

                                        oAdo.SqlQueryReader(oConn, "SELECT * FROM " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                                        if (oAdo.m_OleDbDataReader.HasRows)
                                        {
                                            System.Data.SQLite.SQLiteTransaction transaction;
                                            System.Data.SQLite.SQLiteCommand command = oSQLite.m_Connection.CreateCommand();

                                            // Start a local transaction
                                            transaction = oSQLite.m_Connection.BeginTransaction(IsolationLevel.ReadCommitted);
                                            // Assign transaction object for a pending local transaction
                                            command.Transaction = transaction;
                                            try
                                            {
                                                COUNT = 0;
                                                UpdateTherm(m_frmTherm.progressBar1, COUNT, intTotalRecs);
                                                while (oAdo.m_OleDbDataReader.Read())
                                                {
                                                    COUNT++;
                                                    strValues = utils.GetParsedInsertValues(oAdo.m_OleDbDataReader, columnsAndDataTypes);
                                                    command.CommandText = $"INSERT INTO {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} ({strColumns}) VALUES ({strValues})";
                                                    command.ExecuteNonQuery();
                                                    //if (COUNT == 100) break;
                                                    //frmMain.g_oDelegate.SetControlPropertyValue((Control)lblSQLite2Msg, "Text", "INSERT DATA: " + COUNT.ToString() + " of " + intTotalCount.ToString());
                                                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", 
                                                        "Package:" + p_strPackage.Trim() + " Prepare Tree Data for Input...Stand By [" + COUNT.ToString() + "/" + intTotalRecs.ToString() + "]");
                                                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                                                    UpdateTherm(m_frmTherm.progressBar1,
                                                   COUNT, intTotalRecs);
                                                }
                                                transaction.Commit();
                                            }
                                            catch (Exception err)
                                            {
                                                m_intError = -1;
                                                MessageBox.Show(err.Message);
                                                transaction.Rollback();
                                            }
                                            transaction.Dispose();
                                            transaction = null;
                                        }
                                        strConn = oConn.ConnectionString;
                                        oAdo.m_OleDbDataReader.Close(); oAdo.m_OleDbDataReader.Dispose();
                                        _MSAccess.CloseConnection(oConn);
                                        oConn.Dispose();
                                        //oAdo = null;
                                        UpdateTherm(m_frmTherm.progressBar1, 1, 6);
                                        //
                                        //RUN JAVA APP TO SEND TO ORACLE AND CALCULATE VOLUME/BIOMASS
                                        //
                                        if (m_intError == 0)
                                        {
                                            oSQLite.CloseAndDisposeConnection(oSQLite.m_Connection, true);
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
                                        UpdateTherm(m_frmTherm.progressBar1, 2, 6);
                                        //
                                        //UPDATE MSACCESS WITH CALCULATED VALUES
                                        //
                                        if (m_intError == 0)
                                        {
                                            //oAdo = new ado_data_access();

                                            _MSAccess = oAdo;

                                            _SQLite.OpenConnection(false, 1, frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase, "BIOSUM");

                                            //COUNT = Convert.ToInt32(_oSQLite.getSingleDoubleValueFromSQLQuery(_oSQLite.m_Connection,"SELECT COUNT(*) AS ROWCOUNT FROM BIOSUM_CALC","biosum_volume"));

                                            _MSAccess.OpenConnection(strConn, ref oConn);

                                            if (_MSAccess.TableExist(oConn, Tables.VolumeAndBiomass.BiosumCalcOutputTable))
                                                _MSAccess.SqlNonQuery(oConn, $"DROP TABLE {Tables.VolumeAndBiomass.BiosumCalcOutputTable}");

                                            System.Threading.Thread.Sleep(3000);

                                            _MSAccess.SqlNonQuery(oConn, $"SELECT * INTO {Tables.VolumeAndBiomass.BiosumCalcOutputTable} FROM {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable} WHERE 1=2");


                                            //MSAccessBeginTransaction("BIOSUM_VOLUME_INPUT", "TRE_CN,VOLCSGRS_CALC,VOLCFGRS_CALC,VOLCFNET_CALC,DRYBIOT_CALC,DRYBIOM_CALC,VOLTSGRS_CALC", "TRE_CN",COUNT , "");

                                            intTotalRecs = Convert.ToInt32(_SQLite.getSingleDoubleValueFromSQLQuery( _SQLite.m_Connection,
                                                $"SELECT COUNT(*) AS ROWCOUNT FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} WHERE DRYBIOT_CALC IS NOT NULL OR VOLTSGRS_CALC IS NOT NULL",
                                                Tables.VolumeAndBiomass.BiosumVolumeCalcTable));

                                            UpdateTherm(m_frmTherm.progressBar1, 3, 6);

                                            oSQLite.SqlQueryReader(oSQLite.m_Connection, $"SELECT * FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} WHERE DRYBIOT_CALC IS NOT NULL OR VOLTSGRS_CALC IS NOT NULL");
                                            if (oSQLite.m_DataReader.HasRows)
                                            {
                                                System.Data.OleDb.OleDbTransaction transaction;
                                                System.Data.OleDb.OleDbCommand command = oConn.CreateCommand();
                                                // Start a local transaction
                                                transaction = oConn.BeginTransaction(IsolationLevel.ReadCommitted);
                                                // Assign transaction object for a pending local transaction
                                                command.Transaction = transaction;
                                                try
                                                {
                                                    COUNT = 0;
                                                    UpdateTherm(m_frmTherm.progressBar1, COUNT, intTotalRecs);
                                                    while (oSQLite.m_DataReader.Read())
                                                    {
                                                        COUNT++;
                                                        if (oSQLite.m_DataReader["TRE_CN"] != DBNull.Value && Convert.ToString(oSQLite.m_DataReader["TRE_CN"]).Trim().Length > 0)
                                                        {
                                                            strValues = utils.GetParsedInsertValues(oSQLite.m_DataReader, columnsAndDataTypes);
                                                            command.CommandText = $"INSERT INTO {Tables.VolumeAndBiomass.BiosumCalcOutputTable} ({strColumns}) VALUES ({strValues})";
                                                            command.ExecuteNonQuery();
                                                            UpdateTherm(m_frmTherm.progressBar1, COUNT, intTotalRecs);
                                                        }

                                                        frmMain.g_oDelegate.SetControlPropertyValue(
                                                            (System.Windows.Forms.Control) m_frmTherm.lblMsg, "Text",
                                                            "Package:" + p_strPackage.Trim() + "...Stand By [" +
                                                            COUNT.ToString() + "/" + intTotalRecs.ToString() + "]");
                                                        frmMain.g_oDelegate.ExecuteControlMethod(
                                                            (System.Windows.Forms.Control) this.m_frmTherm.lblMsg,
                                                            "Refresh");
                                                    }

                                                    transaction.Commit();
                                                }
                                                catch (Exception err)
                                                {
                                                    m_intError = -1;
                                                    MessageBox.Show(err.Message);
                                                    transaction.Rollback();
                                                }
                                                finally
                                                {
                                                transaction.Dispose();
                                                transaction = null;
                                                if (m_intError == 0)
                                                    {
                                                        //oAdo.m_strSQL = "UPDATE BIOSUM_VOLUME_INPUT i INNER JOIN BIOSUM_VOLUME_OUTPUT o ON i.TRE_CN = o.TRE_CN SET i.VOLCSGRS_CALC=o.VOLCSGRS_CALC,i.VOLCFGRS_CALC=o.VOLCFGRS_CALC,i.VOLCFNET_CALC=o.VOLCFNET_CALC,i.DRYBIOT_CALC=o.DRYBIOT_CALC,i.DRYBIOM_CALC=o.DRYBIOM_CALC,i.VOLTSGRS_CALC=o.VOLTSGRS_CALC";
                                                        //oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                                        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Update FVS_TREE table with compiled values...stand by");
                                                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                                                        System.Threading.Thread.Sleep(2000);
                                                        strConn = oConn.ConnectionString;
                                                        oAdo.CloseConnection(oConn);

                                                        oAdo.OpenConnection(strConn);
                                                        m_oPrePostDbFileItem_Collection.Item(y).Connection = oAdo.m_OleDbConnection;
                                                        oConn = oAdo.m_OleDbConnection;
                                                        //insert all records 
                                                        //from 
                                                        //oracle table fcs_biosum_volume 
                                                        //into 
                                                        //table fvs_tree
                                                        oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step9(
                                                                           strFvsTreeTable, Tables.VolumeAndBiomass.BiosumCalcOutputTable);

                                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                                        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                                    }
                                                }
                                                //intChanges = MSAccessCommitChanges("BIOSUM_VOLUME_INPUT");
                                            }
                                            //_MSAccess.CloseConnection(_MSAccess.m_OleDbConnection);
                                            //_MSAccess.m_OleDbConnection.Dispose();
                                            //oAdo = null;
                                            oSQLite.CloseAndDisposeConnection(oSQLite.m_Connection, true);
                                            oSQLite = null;

                                        }

                                        UpdateTherm(m_frmTherm.progressBar1, 6, 6);

                                       
                                    }

                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                           m_intProgressStepCurrentCount,
                                           m_intProgressStepTotalCount);

                                if (oAdo.m_intError == 0 && m_intError==0)
                                {
                                    if (oAdo.TableExist(oConn, "cutlist_save_tree_species_work_table"))
                                    {
                                        //update fvs_species
                                        oAdo.m_strSQL = "UPDATE " + strFvsTreeTable + " fvs INNER JOIN cutlist_save_tree_species_work_table w " +
                                                        "ON fvs.FVS_TREE_ID = TRIM(w.FVS_TREE_ID) AND fvs.biosum_cond_id = w.biosum_cond_id " +
                                                        "SET fvs.FVS_SPECIES = TRIM(w.fvs_species) " +
                                                        "WHERE  w.fvs_species IS NOT NULL AND " +
                                                        "LEN(TRIM(w.fvs_species)) > 0 AND " +
                                                        "TRIM(fvs.fvs_species) <> TRIM(w.fvs_species)";
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        oAdo.SqlNonQuery(oConn, "DROP TABLE cutlist_save_tree_species_work_table");
                                    }

                                }
                                p_intError = oAdo.m_intError;





                            }
                            p_intError = oAdo.m_intError;
                        }



                        if (oAdo.m_intError != 0 || m_intError != 0) break;

                    }
                    if (oAdo.m_intError != 0 || m_intError != 0) break;
                }
            }
            if (oAdo.m_intError != 0 || m_intError != 0)
            {
                if (oAdo.m_intError != 0)
                {
                    p_intError = oAdo.m_intError;
                    p_strError = oAdo.m_strError;
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

            oAdo = null;
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
            int x, y, z;
            string strRx = "";
            string strCycle = "";

            System.Data.OleDb.OleDbConnection oConn;
            string strDbFile = "";
            string strFVSOutTable = "";
            string strFVSOutTableLink = "";
            string strFiaTreeSpeciesRefTableLink = "";
            string strCasesTable = "FVS_Cases";
            bool bIdColumnExist = false;
            string[] arrFvsSourceTables = new string[] { "FVS_CutList"};
            string[] arrTreeAccdbs = new string[] { "PREPOST_FVS_CUTLIST.ACCDB" };

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//RunAppend_UpdateFVSTreeTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            ado_data_access oAdo = new ado_data_access();
            DataMgr oDataMgr = new DataMgr();


            oAdo.m_intError = 0;
            oAdo.m_strError = "";
            //
            //make sure all the tables and columns exist
            //
            oAdo.m_strSQL = "";

            string seqNumWorkTable = "cutlist_seq_num_work_table";
            for (y = 0; y <= m_oPrePostDbFileItem_Collection.Count - 1; y++)
            {
                strDbFile = m_oPrePostDbFileItem_Collection.Item(y).DbFileName.Trim().ToUpper();
                if (arrTreeAccdbs.Contains(strDbFile))
                {
                    int idxTable = Array.IndexOf(arrTreeAccdbs, strDbFile);
                    string strFvsTreeTable = Tables.FVS.DefaultFVSCutTreeTableName; ;

                    for (x = 0; x <= m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Count - 1; x++)
                    {
                        TableLinkItem oTableLinkItem = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(x);
                        if (m_oPrePostDbFileItem_Collection.Item(y).TableType.Trim().ToUpper() ==
                            oTableLinkItem.FVSOutputTableName.Trim().ToUpper() && oTableLinkItem.FVSOutputTable)
                        {
                            strFVSOutTable = oTableLinkItem.FVSOutputTableName.Trim().ToUpper();
                            //strFVSOutTableLink = oTableLinkItem.LinkedTableName.Trim().ToUpper();
                            strFVSOutTableLink = arrFvsSourceTables[idxTable];

                            oConn = m_oPrePostDbFileItem_Collection.Item(y).Connection;
                            string strDirectory = m_oPrePostDbFileItem_Collection.Item(y).Directory;

                            m_intProgressStepCurrentCount = 0;
                            if (m_strRxCycleArray == null)
                                m_strRxCycleArray = new string[1];

                            m_intProgressStepTotalCount = 8 + (m_strRxCycleArray.Length * 5);
                            /**************************************************************
                             **delete records in the fvs_tree table that have the current
                             **package
                             **************************************************************/
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
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + m_ado.m_strSQL + "\r\n");

                                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE: " + System.DateTime.Now.ToString() + "\r\n");
                            }

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            if (p_intError == 0)
                            {

                                for (z = 0; z <= m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Count - 1; z++)
                                {
                                    if (m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).TableName ==
                                             m_oQueries.m_oDataSource.getValidDataSourceTableName(Datasource.TableTypes.FiaTreeSpeciesReference))
                                    {
                                        strFiaTreeSpeciesRefTableLink = m_oPrePostDbFileItem_Collection.Item(y).m_oTableLinkItemCollection1.Item(z).LinkedTableName;
                                    }
                                }

                                // Keep track of whether there are trees to calculate metrics
                                bool bRunFcs = false;
                                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTreeTempDbFile)))
                                {
                                    conn.Open();

                                    //
                                    //COPY THE SEQNUM MATRIX TABLE
                                    //

                                    if (oDataMgr.TableExist(conn, seqNumWorkTable))
                                        oDataMgr.SqlNonQuery(conn, "DROP TABLE " + seqNumWorkTable);
                                    frmMain.g_oTables.m_oFvs.CreateSQLiteFVSOutputPrePostSeqNumAuditGenericTable(oDataMgr, conn, seqNumWorkTable);
                                    var seqMdb = $@"/FVSOUT_{p_strVariant}_P{p_strPackage}-{p_strRx1}-{p_strRx2}-{p_strRx3}-{p_strRx4}_BIOSUM.ACCDB";
                                    string strSeqNumConn = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + @"/fvs/data/" +
                                        p_strVariant + seqMdb;
                                    using (System.Data.OleDb.OleDbConnection seqConn = new System.Data.OleDb.OleDbConnection(m_ado.getMDBConnString(strSeqNumConn, "", "")))
                                    {
                                        seqConn.Open();
                                        // Using FVS_SUMMARY_PREPOST_SEQNUM_MATRIX instead of FVS_CUTLIST_PREPOST_SEQNUM_MATRIX
                                        // because not all rxs generate a cutlist. Sequence numbers should be the same for CUTLIST and SUMMARY
                                        oAdo.m_strSQL = "select * from FVS_SUMMARY_PREPOST_SEQNUM_MATRIX";
                                        oAdo.CreateDataTable(seqConn, oAdo.m_strSQL, seqNumWorkTable, false);
                                        oAdo.m_DataTable.Columns.Add("RXPACKAGE", typeof(String));  // The SQLite matrix had rxpackage column
                                        using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter("select * from " + seqNumWorkTable, conn))
                                        {
                                            using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                            {
                                                using (var transaction = conn.BeginTransaction())
                                                {
                                                    da.InsertCommand = cb.GetInsertCommand();
                                                    int rows = da.Update(oAdo.m_DataTable);
                                                    transaction.Commit();
                                                }
                                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated results table " + seqNumWorkTable + " \r\n");
                                                }
                                            }
                                        }
                                    }
                                    if (oAdo.m_DataSet != null)
                                    {
                                        oAdo.m_DataSet.Clear();
                                        oAdo.m_DataSet.Dispose();
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
                                                              "FROM FVSOUT." + strCasesTable + " c, FVSOUT." + strFVSOutTableLink + " t " +
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
                                                    "FROM FVSOUT." + strCasesTable + " c, FVSOUT." + strFVSOutTableLink + " t " +
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
                                                                            "(SELECT standid,year " +
                                                                             "FROM " + seqNumWorkTable + " " +
                                                                             "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  b " +
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
                                                "FROM FVSOUT." + strCasesTable + " c, FVSOUT." + strFVSOutTableLink + " t " +
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
                                                       "FROM " + strCasesTable + " c," + strFVSOutTableLink + " t " +
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
                                                       "CASE WHEN t.dbh < 1.0 AND t.TPA > 0 THEN 1 ELSE null END AS STATUSCD, " +     // @ToDo: fix for sqlite! sets statuscd for seedlings
                                                       "'Y' AS FvsCreatedTree_YN," +
                                                       "'" + m_strDateTimeCreated + "' AS DateTimeCreated " +
                                                       "FROM FVSOUT." + strCasesTable + " c," + strFVSOutTableLink + " t " +
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
                                                                            "(SELECT standid,year " +
                                                                             "FROM " + seqNumWorkTable + " " +
                                                                             "WHERE CYCLE" + strCycle + "_PRE_YN='Y')  b " +
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
                                    ODBCMgr odbcMgr = new ODBCMgr();
                                    // Create ODBC entry for the temporary cut list db file
                                    if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName))
                                    {
                                        odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName);
                                    }
                                    odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName, strTreeTempDbFile);
                                    m_dao.CreateSQLiteTableLink(oConn.DataSource, Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                        Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                        ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName, strTreeTempDbFile, true);
                                    m_dao.CreateSQLiteTableLink(oConn.DataSource, strFvsTreeTable,
                                        strFvsTreeTable, ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName, strTreeTempDbFile, true);
                                    // Sleep to ensure table link is complete
                                    Thread.Sleep(5000);

                                    //join plot, cond, and tree table to oracle input tree volumes table.
                                    //NOTE: this query handles existing FIADB trees that have been grown forward.
                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2(
                                                       Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                       m_oQueries.m_oFIAPlot.m_strTreeTable,
                                                       m_oQueries.m_oFIAPlot.m_strPlotTable,
                                                       m_oQueries.m_oFIAPlot.m_strCondTable);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                               m_intProgressStepCurrentCount,
                                               m_intProgressStepTotalCount);


                                    //Set DIAHTCD for FIADB Cycle<>1 trees to their Cycle=1 DIAHTCD values
                                    //Set standing_dead_code from tree table on FIADB trees
                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFiadb(m_oQueries.m_oFIAPlot.m_strTreeTable);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    //Set DIAHTCD for FVS-Created trees using FIA_TREE_SPECIES_REF.WOODLAND_YN
                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFvs(strFiaTreeSpeciesRefTableLink);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    //Set STATUSCD for seedlings; It is populated from FCS for all other trees
                                    oAdo.m_strSQL = $@"UPDATE {strFvsTreeTable} b 
                                                   INNER JOIN {m_oQueries.m_oFIAPlot.m_strTreeTable} t 
                                                   ON t.biosum_cond_id=b.biosum_cond_id AND t.fvs_tree_id=b.fvs_tree_id
                                                   SET b.statuscd=t.statuscd
                                                   WHERE rxpackage='{p_strPackage.Trim()}' AND fvs_variant='{p_strVariant.Trim()}' and dbh < 1.0";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    //Drop the table link when statuscd complete
                                    oAdo.m_strSQL = $@"DROP TABLE {strFvsTreeTable}";
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    //Update FVSCreatedTrees precipitation=plot.precipitation
                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step2a(
                                                       Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                       m_oQueries.m_oFIAPlot.m_strPlotTable,
                                                       m_oQueries.m_oFIAPlot.m_strCondTable);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    //join cond table to oracle input tree volumes table.
                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step3(
                                                      Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                      m_oQueries.m_oFIAPlot.m_strCondTable);



                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");

                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");



                                    //populate treeclcd column
                                    if (oAdo.TableExist(oConn, "CULL_TOTAL_WORK_TABLE"))
                                        oAdo.SqlNonQuery(oConn, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step4(
                                                      "cull_total_work_table",
                                                      Tables.VolumeAndBiomass.BiosumVolumesInputTable);


                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                               m_intProgressStepCurrentCount,
                                               m_intProgressStepTotalCount);


                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                                                        "cull_total_work_table",
                                                        Tables.VolumeAndBiomass.BiosumVolumesInputTable);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                               m_intProgressStepCurrentCount,
                                               m_intProgressStepTotalCount);

                                    oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                                                      "cull_total_work_table",
                                                      Tables.VolumeAndBiomass.BiosumVolumesInputTable);

                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                               m_intProgressStepCurrentCount,
                                               m_intProgressStepTotalCount);

                                    // Drop link to BiosumVolumesInputTable to avoid conflicts
                                    oAdo.m_strSQL = "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable;
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                    // DELETE ODBC ENTRY
                                    if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName))
                                    {
                                        odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName);
                                    }

                                    // CLOSE MS ACCESS CONNECTION
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

                                        System.Threading.Thread.Sleep(2000);

                                        //insert records 
                                        //from 
                                        //table biosum_volumes_input (BiosumVolumesInputTable)
                                        //into 
                                        //table fcs_biosum_volumes_input (FcsBiosumVolumesInputTable)
                                        oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteBiosumCalcTable_Step7(
                                                    Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                    $"FCS.{Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");

                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                        UpdateTherm(m_frmTherm.progressBar1, 1, 6);

                                        //
                                        //RUN JAVA APP TO SEND TO ORACLE AND CALCULATE VOLUME/BIOMASS
                                        //
                                        if (m_intError == 0)
                                        {
                                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.lblMsg, "Text", "Package:" + p_strPackage.Trim() + " Running volume calculations for " + strFvsTreeTable);
                                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                                            //oDataMgr.CloseAndDisposeConnection(oDataMgr.m_Connection, true);
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
                                        UpdateTherm(m_frmTherm.progressBar1, 2, 6);
                                        //
                                        //UPDATE MSACCESS WITH CALCULATED VALUES
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
                                            //update calculated fields 
                                            //from 
                                            //biosum_calc table
                                            //into 
                                            //table fvs_tree
                                            oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step9(
                                                strFvsTreeTable, Tables.VolumeAndBiomass.BiosumVolumeCalcTable, p_strPackage.Trim());
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oDataMgr.m_strSQL + "\r\n");
                                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                                            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                            oDataMgr.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step10(
                                                strFvsTreeTable, Tables.VolumeAndBiomass.BiosumVolumeCalcTable, p_strPackage.Trim());
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
                                            oDataMgr.CloseAndDisposeConnection(oDataMgr.m_Connection, true);
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

                                        // DELETE SEEDLINGS FROM CUTTREE TABLE; RESIDTREE TABLE TBD
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
                                            seqNumWorkTable, "cutlist_rowid_work_table"};
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
                                p_intError = oDataMgr.m_intError;
                            }
                            p_intError = oDataMgr.m_intError;
                        }

                        if (oDataMgr.m_intError != 0 || m_intError != 0) break;

                    }
                    if (oDataMgr.m_intError != 0 || m_intError != 0) break;
                }
            }
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

        private int GetTreeClassCode(System.Data.OleDb.OleDbDataReader p_oDR)
        {
            int intSpCd = Convert.ToInt32(p_oDR["spcd"]);
            int intStatusCd = Convert.ToInt32(p_oDR["statuscd"]);
            double dblDia = Convert.ToDouble(p_oDR["dbh"]);
            double dblCull = Convert.ToDouble(p_oDR["cull"]);
            double dblRoughCull = Convert.ToDouble(p_oDR["roughcull"]);
            byte bytDecayCd = Convert.ToByte(p_oDR["decaycd"]);
            double dblCullTotal = dblCull + dblRoughCull;
            int intTreeClCd = 4;

            if (intSpCd == 62 || intSpCd == 65 || intSpCd == 66 || intSpCd == 106 ||
                intSpCd == 133 || intSpCd == 138 || intSpCd == 304 || intSpCd == 321 ||
                intSpCd == 322 || intSpCd == 475 || intSpCd == 756 || intSpCd == 758 ||
                intSpCd == 990)
            {
                intTreeClCd = 3;
            }
            else if (intStatusCd == 2)
            {
                intTreeClCd = 3;
                if (bytDecayCd > 1) intTreeClCd = 4;
                else if (dblDia < 9 && intSpCd < 300) intTreeClCd = 4;
            }
            else if (dblCullTotal < 75) intTreeClCd = 2;
            else if (dblRoughCull > 37.5) intTreeClCd = 3;
            else intTreeClCd = 4;
            return intTreeClCd;

            //IIF(SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3,IIF(StatusCd=2,3,IIF(cull + roughcull < 75,2,IIF(roughcull > 37.5,3,4))))
                



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
			string strRx1="";
			string strRx2="";
			string strRx3="";
			string strRx4="";
			string strPackage="";
			string strVariant="";
			string strRx="";
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(this.lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;

			Tables oTables = new Tables();
			

			
			

			string strOutDirAndFile;
			string strDbFile;
			string strCasesTable;
            string strAuditDbFile;
            string strPotFireBaseYearFile;
            
			

			
			string[] strSourceTableArray=null;
			ado_data_access oAdo = new ado_data_access();
			
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
            if (m_ado.m_OleDbConnection.State == ConnectionState.Closed)
            {
                m_ado.OpenConnection(m_strTempMDBFileConnectionString);
               
            }
            m_oRxTools.LoadFVSOutputPrePostRxCycleSeqNum(m_ado, m_ado.m_OleDbConnection, m_oFVSPrePostSeqNumItemCollection);

           

			
			int x,y;
			
			try
			{
				bSkip=false;

               

				if (this.m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
					this.m_ado.m_OleDbConnection.Close();

				while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
				{
					System.Threading.Thread.Sleep(1000);
				}


				intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv,"Count",false);
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
                            m_bPotFireBaseYearTableExist = true;
                            m_strPotFireBaseYearLinkedTableName = "FVS_POTFIRE_BASEYEAR";
                            m_strPotFireStandardLinkedTableName = "FVS_POTFIRE_STANDARD";

                            m_intProgressStepTotalCount = 20;
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

                            strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE1, "Text", false);
                            strRx1 = strRx1.Trim();

                            strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE2, "Text", false);
                            strRx2 = strRx2.Trim();

                            strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE3, "Text", false);
                            strRx3 = strRx3.Trim();

                            strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE4, "Text", false);
                            strRx4 = strRx4.Trim();

                            //find the package item in the package collection
                            for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                            {
                                if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim() == strRx1.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim() == strRx2.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim() == strRx3.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim() == strRx4.Trim() &&
                                    this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
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
                            if (strRx1.Trim().Length > 0 && strRx1.Trim() != "000") this.m_strRxCycleList = "1,";
                            if (strRx2.Trim().Length > 0 && strRx2.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                            if (strRx3.Trim().Length > 0 && strRx3.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                            if (strRx4.Trim().Length > 0 && strRx4.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                            if (this.m_strRxCycleList.Trim().Length > 0)
                                this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

                            this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");

                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                            strDbFile = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_MDBOUT, "Text", false);
                            strDbFile = strDbFile.Trim();

                            strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                            strOutDirAndFile = strOutDirAndFile.Trim();
                            strOutDirAndFile = strOutDirAndFile + "\\" + strVariant + "\\" + strDbFile;

                            strAuditDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                            strAuditDbFile = strAuditDbFile.Trim();
                            strAuditDbFile = strAuditDbFile + "\\" + strVariant + "\\" + strDbFile.Replace(".MDB", "_BIOSUM.ACCDB");

                            
                            m_oRxTools.CreateFVSOutputTableLinks(strAuditDbFile, strOutDirAndFile);

                            CreatePotFireTables(strAuditDbFile,strOutDirAndFile,strVariant,strPackage);

                            uc_filesize_monitor1.BeginMonitoringFile(
                            strOutDirAndFile,
                            2000000000, "2gb");
                            uc_filesize_monitor1.Information = "FVS output file";

                            if (uc_filesize_monitor2.File.Trim().Length > 0) uc_filesize_monitor2.EndMonitoringFile();
                            uc_filesize_monitor2.BeginMonitoringFile(strAuditDbFile, 2000000000, "2GB");
                            uc_filesize_monitor2.Information = "BIOSUM DB file containing PREPOST SEQNUM MATRIX and AUDIT tables for processing FVS Output DB file " + strDbFile; 




                            oAdo.OpenConnection(oAdo.getMDBConnString(strAuditDbFile, "", ""));

                            m_strLogFile = strAuditDbFile.Trim() + "_Audit_" + m_strLogDate.Replace(" ", "_") + ".txt";


                            frmMain.g_oUtils.WriteText(m_strLogFile, "AUDIT LOG \r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "--------- \r\n\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Date/Time:" + System.DateTime.Now.ToString().Trim() + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Database File:" + strDbFile + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Variant:" + strVariant + " \r\n");
                            frmMain.g_oUtils.WriteText(m_strLogFile, "Treatment:" + strRx + " \r\n\r\n");



                            strSourceTableArray = oAdo.getTableNames(oAdo.m_OleDbConnection);
                            

                            frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_CASES-----\r\n");

                            //check to ensure the variant in the fvs cases table
                            //matches the current variant
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " FVS_CASES");

                            strCasesTable = "fvs_cases";
                            bResult = oAdo.ValuesExistNotEqualToTargetValue(oAdo.m_OleDbConnection,
                                strCasesTable,
                                "variant",
                                strVariant.Trim(),
                                false);

                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            if (oAdo.m_intError == oAdo.ErrorCodeNoErrors && bResult == true)
                            {
                                intItemError = -1;
                                strItemError = strItemError + "ERROR:Incorrect variant found in FVS_Cases.variant column";
                                strItemDialogMsg = strItemDialogMsg + strDbFile + ": Incorrect variant found in FVS_Cases.variant column\r\n";
                                frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR:Incorrect variant found in variant column\r\n\r\n");

                            }
                            else if (oAdo.m_intError == oAdo.ErrorCodeTableNotFound)
                            {
                                intItemError = -1;
                                strItemError = strItemError + "FVS_Cases table missing\r\n";
                                strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Cases table missing\r\n";
                                frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: table missing\r\n\r\n");
                            }
                            else if (oAdo.m_intError == oAdo.ErrorCodeColumnNotFound)
                            {
                                intItemError = -1;
                                strItemError = strItemError + "FVS_Cases.variant column not found\r\n";
                                strItemDialogMsg = strItemDialogMsg + strDbFile + ":FVS_Cases.variant column not found\r\n";
                                frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: variant column not found\r\n\r\n");
                            }
                            else
                            {
                                Validate_FVSCaseId(strCasesTable, oAdo, oAdo.m_OleDbConnection, ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);
                                if (intItemError == 0)
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n\r\n");
                                else
                                {
                                    strItemDialogMsg = strItemDialogMsg + strDbFile + "\r\n_________________________________\r\n\r\n" + strItemError;
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: " + strItemError + "\r\n\r\n");
                                }
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            //
                            //check for duplicate standid+year records for tables that are 
                            //only supposed to have one stand+year combination represented
                            //
                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Check For duplicate standid+year");
                                CheckForDuplicateStandIdandYear(oAdo, oAdo.m_OleDbConnection, strSourceTableArray, strVariant, strPackage, true, ref strItemError, ref intItemError);
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);
                           


                            if (intItemError==0) 
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " Create SeqNum Matrix tables");
                                CreatePrePostSeqNumMatrixTables(oAdo, oAdo.m_OleDbConnection,strPackage,true);
                            }
                            m_intProgressStepCurrentCount++;
                            UpdateTherm(m_frmTherm.progressBar1,
                                        m_intProgressStepCurrentCount,
                                        m_intProgressStepTotalCount);

                            //check if summary table exists
                            if (intItemError == 0)
                            {
                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " FVS_SUMMARY");
                                frmMain.g_oUtils.WriteText(m_strLogFile, "-----FVS_SUMMARY-----\r\n");
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "FVS_Summary"))
                                {
                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit...FVS_SUMMARY");

                                    if (intItemError == 0)
                                    {
                                        //
                                        //get fvs_summary configuration
                                        //
                                        GetPrePostSeqNumConfiguration("FVS_SUMMARY",strPackage);
                                        //check pre and post-treatment seqnum assignments
                                        this.Validate_FvsSummaryPrePostSeqNum(oAdo, oAdo.m_OleDbConnection, ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);
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


                                    strSourceTableArray = oAdo.getTableNames(oAdo.m_OleDbConnection);
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

                                               

                                                this.Validate_TreeListTables(oAdo, oAdo.m_OleDbConnection, "fvs_treelist", "fvs_summary", ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);

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

                                                this.Validate_TreeListTables(oAdo, oAdo.m_OleDbConnection, "fvs_cutlist", "fvs_summary", ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);





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
                                                    this.Validate_FVSTreeId(oAdo, oAdo.m_OleDbConnection, strAuditDbFile, "FVS_Cases", "FVS_CutList", strVariant, strPackage, strRx1, strRx2, strRx3, strRx4, true, ref intItemWarning, ref strItemWarning, ref intItemError, ref strItemError);
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

                                                

                                                this.Validate_TreeListTables(oAdo, oAdo.m_OleDbConnection, "fvs_atrtlist", "fvs_summary", ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);

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

                                                
                                                this.Validate_PotFire(oAdo, oAdo.m_OleDbConnection, "FVS_POTFIRE", strVariant, ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);

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

                                                this.Validate_FVSGenericTable(oAdo, oAdo.m_OleDbConnection, "FVS_STRCLASS", "FVS_STRCLASS", ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);

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
                                                if (strSourceTableArray[y].Trim().ToUpper() !=
                                                    m_oQueries.m_oFvs.m_strFVSWesternTreeSpeciesTable.Trim().ToUpper() &&
                                                    strSourceTableArray[y].Trim().ToUpper() !=
                                                    m_oQueries.m_oFvs.m_strFVSEasternTreeSpeciesTable.Trim().ToUpper())
                                                {
                                                    frmMain.g_oUtils.WriteText(m_strLogFile, "-----" + strSourceTableArray[y].Trim().ToUpper() + "-----\r\n");
                                                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing Audit..." + strSourceTableArray[y].Trim());


                                                    intItemWarning = 0;
                                                    strItemWarning = "";

                                                    this.Validate_FVSGenericTable(oAdo, oAdo.m_OleDbConnection, strSourceTableArray[y].Trim(), strSourceTableArray[y], ref intItemError, ref strItemError, ref intItemWarning, ref strItemWarning, true);

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


                                            }

                                            if (strSourceTableArray[y].Trim().ToUpper() != "FVS_TREELIST" &&
                                                strSourceTableArray[y].Trim().ToUpper() != "FVS_CUTLIST" &&
                                                strSourceTableArray[y].Trim().ToUpper() != "FVS_ATRTLIST")
                                            {

                                            }
                                            if (intItemError != 0) break;
                                        }

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

                            if (oAdo.TableExist(oAdo.m_OleDbConnection, "ATRTList_work"))
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE ATRTlist_work");

                            if (oAdo.TableExist(oAdo.m_OleDbConnection, "cutlist_work"))
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE cutlist_work");

                            if (oAdo.TableExist(oAdo.m_OleDbConnection, "treelist_work"))
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE treelist_work");


                            oAdo.CloseConnection(oAdo.m_OleDbConnection);
                            frmMain.g_oDelegate.ExecuteListViewItemsMethod(oLv, "Refresh");

                            //compact and repair
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Compact and Repair");
                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                            System.Threading.Thread.Sleep(5000);
                            m_dao.CompactMDB(strAuditDbFile);
                            

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
				if (oAdo.m_OleDbConnection != null)
				{
					if (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
					{
						oAdo.CloseConnection(oAdo.m_OleDbConnection);
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
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn">Connection to the FVSOUT_VARIANT_P000-000-000-000-000.MDB file</param>
        /// <param name="p_strFVSOutTables">tables in the fvsoutput file</param>
        /// <param name="p_strRxPackage">treatment package</param>
        /// <param name="p_bAudit">audit routine</param>
        /// <param name="p_strError"></param>
        /// <param name="p_intError"></param>
        private void CheckForDuplicateStandIdandYear(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn,string[] p_strFVSOutTables,string p_strVariant, string p_strRxPackage,bool p_bAudit,ref string p_strError, ref int p_intError)
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
                        p_oAdo.m_strSQL = "SELECT DISTINCT b.standid,b.year,b.rowcount " +
                                          "FROM " + p_strFVSOutTables[y] + " a," +
                                            "(SELECT count(*) as rowcount,standid,year " +
                                             "FROM " + p_strFVSOutTables[y] + " GROUP BY STANDID,YEAR) b " +
                                          "WHERE a.standid=b.standid AND a.year=b.year AND b.rowcount > 1";
                        p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                        if (p_oAdo.m_OleDbDataReader.HasRows)
                        {
                            p_intError=-1;
                            while (p_oAdo.m_OleDbDataReader.Read())
                            {
                                if (p_oAdo.m_OleDbDataReader["standid"] != System.DBNull.Value && 
                                    p_oAdo.m_OleDbDataReader["year"] != System.DBNull.Value && 
                                    p_oAdo.m_OleDbDataReader["rowcount"] != System.DBNull.Value)
                                        p_strError = p_strError + "ERROR: [Duplicate StandId+Year] Variant: " + p_strVariant + " Package: " + p_strRxPackage + " Table:" + p_strFVSOutTables[y] + " StandId: " + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " Year: " + p_oAdo.m_OleDbDataReader["year"].ToString().Trim() + " Row Count:" + p_oAdo.m_OleDbDataReader["rowcount"].ToString().Trim() + "\r\n"; 
                            }
                        }
                        p_oAdo.m_OleDbDataReader.Close();
                    }
                }
            }
            if (p_intError != 0 && p_bAudit)
                frmMain.g_oUtils.WriteText(m_strLogFile, p_strError + "\r\n\r\n");
        }
        /// <summary>
        /// Create the FVS_POTFIRE table. If the configuration includes baseyear than combine the baseyear POTFIRE table to the 
        /// standard POTFIRE table.  If the baseyear POTFIRE table is not used then just use the standard POTFIRE table. It is 
        /// assumed that the standard FVS_POTFIRE table link already exists as a link in the p_strDbFile when this routine is called.
        /// </summary>
        /// <param name="p_strDbFile">The FVSOUT_P000-000-000-000_BIOSUM.ACCDB or any accdb file</param>
        /// <param name="p_strVariant">FVS variant</param>
        private void CreatePotFireTables(string p_strAuditDbFile, string p_strFVSOutDbFile,string p_strVariant,string p_strRxPackageId)
        {

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// CreatePotFireTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            int x;
            string strPotFireBaseYearFile = "";
            string strPotFireTable = "";
            ado_data_access oAdo;


            dao_data_access oDao = new dao_data_access();
            //
            //CHECK TO SEE IF THE BASEYEAR FILE AND TABLE EXIST
            //
            strPotFireBaseYearFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
            strPotFireBaseYearFile = strPotFireBaseYearFile.Trim();
            strPotFireBaseYearFile = strPotFireBaseYearFile + "\\" + p_strVariant + "\\FVSOUT_" + p_strVariant + "_POTFIRE_BaseYr.MDB";

            //see which potfire table
            if (oDao.TableExists(p_strFVSOutDbFile, "FVS_POTFIRE")) strPotFireTable = "FVS_POTFIRE";
            else if (oDao.TableExists(p_strFVSOutDbFile, "FVS_POTFIRE_EAST")) strPotFireTable = "FVS_POTFIRE_EAST";

            if (!System.IO.File.Exists(strPotFireBaseYearFile) || !oDao.TableExists(strPotFireBaseYearFile, strPotFireTable))
            {
                m_bPotFireBaseYearTableExist = false;
                oDao.m_DaoWorkspace.Close();
                oDao = null;
                return;
            }

            //get the PREPOST SeqNum configuration for the POTFIRE table
            GetPrePostSeqNumConfiguration(strPotFireTable,p_strRxPackageId);
            //
            //CHECK TO SEE IF THE CONFIGURATION INCLUDES BASEYEAR
            //
            if (m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN == "N" &&
                m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN == "N" &&
                m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN == "N" &&
                m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN == "N")
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
                oAdo = new ado_data_access();
                oAdo.OpenConnection(oAdo.getMDBConnString(p_strAuditDbFile,"",""));
                if (oAdo.m_intError == 0)
                {
                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "FVS_POTFIRE_TEMP"))
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE FVS_POTFIRE_TEMP");

                    oAdo.m_strSQL = "SELECT * INTO FVS_POTFIRE_TEMP FROM " + strPotFireTable;
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strPotFireTable);
                    oAdo.m_strSQL = "SELECT * INTO " + strPotFireTable + " FROM FVS_POTFIRE_TEMP";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE FVS_POTFIRE_TEMP");
                    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                    oAdo = null;

                }
                return;
            }
            //
            //PROCESS THE BASEYEAR AND STANDARD POTFIRE TABLES INTO ONE FVS_POTFIRE TABLE
            //
            oDao.OpenDb(p_strAuditDbFile);
            oDao.RenameTable(oDao.m_DaoDatabase, strPotFireTable,m_strPotFireStandardLinkedTableName,true,true);
            oDao.CreateTableLink(oDao.m_DaoDatabase,m_strPotFireBaseYearLinkedTableName, strPotFireBaseYearFile,strPotFireTable,true);
            oDao.m_DaoDatabase.Close();
            oDao.m_DaoWorkspace.Close();
            oDao = null;
            //create the new FVS_POTFIRE table by inserting the baseyear POTFIRE records
            oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strAuditDbFile,"",""));
            if (oAdo.m_intError == 0)
            {
                if (oAdo.TableExist(oAdo.m_OleDbConnection, "tempBASEYEAR")) oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE TEMPBASEYEAR");
                if (oAdo.TableExist(oAdo.m_OleDbConnection, "BASEYEAR")) oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE BASEYEAR");
                if (oAdo.TableExist(oAdo.m_OleDbConnection, "NONBASEYEAR")) oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE NONBASEYEAR");
                string[] strSQL = Queries.FVS.FVSOutputTable_PrePostPotFireBaseYearSQL(
                    m_strPotFireBaseYearLinkedTableName,
                    m_strPotFireStandardLinkedTableName,
                    "FVS_POTFIRE");

                for (x = 0; x <= strSQL.Length - 1; x++)
                {
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL[x] + "\r\n");
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL[x]);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    if (oAdo.m_intError != 0) break;
                }
                if (oAdo.m_intError == 0)
                {
                    //drop the baseyear_yn column
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_POTFIRE DROP COLUMN baseyear_yn\r\n");
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_POTFIRE DROP COLUMN baseyear_yn");
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_strPotFireStandardLinkedTableName)) oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_strPotFireStandardLinkedTableName);
                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_strPotFireBaseYearLinkedTableName)) oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_strPotFireBaseYearLinkedTableName);
                }
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
            }
            oAdo = null;
        }
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
			m_intProgressOverallTotalCount=0;
			m_intProgressStepCurrentCount=0;
			m_strError="";
			m_strWarning="";
			m_intWarning=0;
			m_intProgressOverallCurrentCount=0;
			string strRx1="";
			string strRx2="";
			string strRx3="";
			string strRx4="";
			string strPackage="";
			string strVariant="";
			string strRx="";
            string strSQL = "";
            System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(this.lstFvsOutput, false);
            System.Windows.Forms.ListViewItem oLvItem = null;

			Tables oTables = new Tables();

            string strVariantList = "";
           
			

			string strOutDirAndFile;
			string strDbFile;
			
            string strAuditDbFile;
            
			

			
			string[] strSourceTableArray=null;
			ado_data_access oAdo = new ado_data_access();
			
			bool bSkip;
			bool bResult=false;
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
                bSkip = false;



                if (this.m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
                    this.m_ado.m_OleDbConnection.Close();

                while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
                {
                    System.Threading.Thread.Sleep(1000);
                }


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
                            bSkip = true;




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

                            strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE1, "Text", false);
                            strRx1 = strRx1.Trim();

                            strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE2, "Text", false);
                            strRx2 = strRx2.Trim();

                            strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE3, "Text", false);
                            strRx3 = strRx3.Trim();

                            strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE4, "Text", false);
                            strRx4 = strRx4.Trim();

                            
                            //check to see if this variant already processed
                           
                                //find the package item in the package collection
                                for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
                                {
                                    if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim() == strRx1.Trim() &&
                                        this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim() == strRx2.Trim() &&
                                        this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim() == strRx3.Trim() &&
                                        this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim() == strRx4.Trim() &&
                                        this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
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
                                if (strRx1.Trim().Length > 0 && strRx1.Trim() != "000") this.m_strRxCycleList = "1,";
                                if (strRx2.Trim().Length > 0 && strRx2.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
                                if (strRx3.Trim().Length > 0 && strRx3.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
                                if (strRx4.Trim().Length > 0 && strRx4.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

                                if (this.m_strRxCycleList.Trim().Length > 0)
                                    this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

                                this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");

                                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
                                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

                                strDbFile = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_MDBOUT, "Text", false);
                                strDbFile = strDbFile.Trim();

                                strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                                strOutDirAndFile = strOutDirAndFile.Trim();
                                strOutDirAndFile = strOutDirAndFile + "\\" + strVariant + "\\" + strDbFile;

                                strAuditDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                                strAuditDbFile = strAuditDbFile.Trim();
                                strAuditDbFile = strAuditDbFile + "\\" + strVariant + "\\PostAudit.accdb";
                                string strTempCutListTable = "tmpCutTree";
                                //string strTableLinkName = Tables.FVS.DefaultFVSCutTreeTableName;

                                int intTreeTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREE");
                                int intCondTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("CONDITION");
                                int intPlotTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("PLOT");
                                int intRxTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREATMENT PRESCRIPTIONS");
                                int intRxPackageTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREATMENT PACKAGES");

                                if (System.IO.File.Exists(strAuditDbFile))
                                {
                                    oAdo.OpenConnection(oAdo.getMDBConnString(strAuditDbFile, "", ""));
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE]))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE]);
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE]))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE]);
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE]))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE]);
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.TABLE]))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.TABLE]);
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE]))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE]);
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, strTempCutListTable))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE " + strTempCutListTable);
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "rxpackage_work_table"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE rxpackage_work_table");
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "fvs_tree_unique_biosum_plot_id_work_table"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE fvs_tree_unique_biosum_plot_id_work_table");
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "fvs_tree_biosum_plot_id_work_table"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE fvs_tree_biosum_plot_id_work_table");
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "cond_biosum_cond_id_work_table"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE cond_biosum_cond_id_work_table");
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "plot_biosum_plot_id_work_table"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE plot_biosum_plot_id_work_table");
                                    }
                                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "tree_fvs_tree_id_work_table"))
                                    {
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "DROP TABLE tree_fvs_tree_id_work_table");
                                    }
                                    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                                }

                                dao_data_access oDao = new dao_data_access();

                                if (System.IO.File.Exists(strAuditDbFile) == false)
                                    oDao.CreateMDB(strAuditDbFile);

                            ODBCMgr odbcmgr = new ODBCMgr();
                            // Set up an ODBC DSN for the FVSOUT_TREE_LIST.db
                            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsOutTreeListDsnName))
                            {
                                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsOutTreeListDsnName);
                            }
                            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.FvsOutTreeListDsnName,
                                m_strFvsTreeDb);
                            //oDao.CreateTableLink(strAuditDbFile, strTableLinkName, strFvsTreeFile, "fvs_tree", false);
                            // Prepare tmpCutTree for next run
                            m_dbConn = SQLite.GetConnectionString(m_strFvsTreeDb);
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_dbConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, strTempCutListTable))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE " + strTempCutListTable);
                                }
                                frmMain.g_oTables.m_oFvs.CreateFVSOutTreeTable(SQLite, conn, strTempCutListTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Created table " + strTempCutListTable + "\r\n");
                                SQLite.SqlNonQuery(conn, $@"INSERT INTO {strTempCutListTable} SELECT * FROM {Tables.FVS.DefaultFVSCutTreeTableName} 
                                                         WHERE FVS_VARIANT ='{strVariant}' AND RXPACKAGE = '{strPackage}'");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "Populated table " + strTempCutListTable + "\r\n");
                            }
                            oDao.CreateSQLiteTableLink(strAuditDbFile, strTempCutListTable, strTempCutListTable,
                                ODBCMgr.DSN_KEYS.FvsOutTreeListDsnName, m_strFvsTreeDb);

                            //tree table link
                            oDao.CreateTableLink(strAuditDbFile, m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE],
                                                      m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.PATH].Trim() + "\\" +
                                                       m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.MDBFILE].Trim(),
                                                      m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE], true);
                                //condition table link

                                oDao.CreateTableLink(strAuditDbFile, m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE],
                                                      m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.PATH].Trim() + "\\" +
                                                       m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.MDBFILE].Trim(),
                                                      m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE], true);
                                //plot table link

                                oDao.CreateTableLink(strAuditDbFile, m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE],
                                                      m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.PATH].Trim() + "\\" +
                                                       m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.MDBFILE].Trim(),
                                                      m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE], true);

                                //rx table link
                                oDao.CreateTableLink(strAuditDbFile, m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.TABLE],
                                                     m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.PATH].Trim() + "\\" +
                                                      m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.MDBFILE].Trim(),
                                                     m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.TABLE], true);

                                //rx package table link
                                oDao.CreateTableLink(strAuditDbFile, m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE],
                                                    m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.PATH].Trim() + "\\" +
                                                     m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.MDBFILE].Trim(),
                                                    m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE], true);

                                oDao.m_DaoWorkspace.Close();
                                oDao = null;
                                System.Threading.Thread.Sleep(2000);

                                oAdo.OpenConnection(oAdo.getMDBConnString(strAuditDbFile, "", ""));




                                uc_filesize_monitor1.BeginMonitoringFile(strAuditDbFile, 2000000000, "2GB");
                                uc_filesize_monitor1.Information = "BIOSUM DB file containing FVS OUTPUT PRE/POST AUDIT tables";


                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);

                               

                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_SUMMARY") == false)
                                {
                                    oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistSUMMARYtableSQL("audit_Post_SUMMARY");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                }
                                

                                string[] sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryFVS(
                                    m_oQueries.m_oDataSource.m_strDataSource[intRxTable,Datasource.TABLE],
                                    m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE],
                                    m_oQueries.m_oDataSource.m_strDataSource[intTreeTable,Datasource.TABLE],
                                    m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE],
                                    m_oQueries.m_oDataSource.m_strDataSource[intCondTable,Datasource.TABLE],
                                    "audit_Post_SUMMARY", strTempCutListTable,
                                    strPackage);

                                for (y = 0; y <= sqlArray.Length - 1; y++)
                                {
                                    strSQL = sqlArray[y];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                    m_intProgressStepCurrentCount++;
                                    UpdateTherm(m_frmTherm.progressBar1,
                                                m_intProgressStepCurrentCount,
                                                m_intProgressStepTotalCount);

                                }
                                //
                                //process NOVALUE_ERROR column
                                //
                                sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_NOVALUE_ERROR(
                                    "audit_Post_NOVALUE_ERROR",
                                    "audit_Post_SUMMARY",
                                    strTempCutListTable,
                                    strPackage);
                                
                                //check if NOVALUE_ERROR table exists
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_NOVALUE_ERROR") == false)
                                {
                                    oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL("audit_Post_NOVALUE_ERROR");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                }
                            else if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, "audit_Post_NOVALUE_ERROR", "treeval"))
                            {
                                // Accommodate new fields for v5.8.10
                                oAdo.m_strSQL = "DROP TABLE audit_Post_NOVALUE_ERROR";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL("audit_Post_NOVALUE_ERROR");
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            }
                            else
                                {
                                    //delete any records from a previous run for the current variant and rxpackage
                                    strSQL = sqlArray[0];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);    
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                                                    
                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any no value errors
                                int intRowCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM audit_Post_SUMMARY WHERE RXPACKAGE='" + strPackage + "' AND NOVALUE_ERROR IS NOT NULL AND LEN(TRIM(NOVALUE_ERROR)) > 0 AND TRIM(NOVALUE_ERROR) <> 'NA' AND VAL(NOVALUE_ERROR) > 0", "AUDIT_POST_SUMMARY");
                                if (intRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = sqlArray[1];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
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
                                    strPackage);

                                //check if VALUE_ERROR table exists

                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_VALUE_ERROR") == false)
                                {
                                    oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL("audit_Post_VALUE_ERROR");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                }
                            else if (!oAdo.ColumnExist(oAdo.m_OleDbConnection, "audit_Post_VALUE_ERROR", "treeval"))
                            {
                                // Accommodate new fields for v5.8.10
                                oAdo.m_strSQL = "DROP TABLE audit_Post_VALUE_ERROR";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL("audit_Post_VALUE_ERROR");
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            }
                            else
                                {
                                    //delete any records from a previous run for the current variant and rxpackage
                                    strSQL = sqlArray[0];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any no value errors
                                intRowCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM audit_Post_SUMMARY WHERE RXPACKAGE='" + strPackage + "' AND VALUE_ERROR IS NOT NULL AND LEN(TRIM(VALUE_ERROR)) > 0 AND TRIM(VALUE_ERROR) <> 'NA' AND VAL(VALUE_ERROR) > 0 AND TRIM(COLUMN_NAME) <> 'DBH'", "AUDIT_POST_SUMMARY");
                                if (intRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = sqlArray[1];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
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
                                sqlArray = Queries.FVS.FVSOutputTable_AuditPostSummaryDetailFVS_NOTFOUND_ERROR(
                                    "audit_Post_NOTFOUND_ERROR",
                                    "audit_Post_SUMMARY",
                                    strTempCutListTable,
                                    strPackage,
                                     m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE].Trim(),
                                     m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE].Trim(),
                                     m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE].Trim(),
                                     m_oQueries.m_oDataSource.m_strDataSource[intRxTable, Datasource.TABLE].Trim(),
                                     m_oQueries.m_oDataSource.m_strDataSource[intRxPackageTable, Datasource.TABLE].Trim(),
                                     "rxpackage_work_table");
                                //check if NOVALUE_ERROR table exists
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_NOTFOUND_ERROR") == false)
                                {
                                    oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL("audit_Post_NOTFOUND_ERROR");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                }
                                else
                                {
                                    //delete any records from a previous run for the current variant and rxpackage
                                    strSQL = sqlArray[0];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any not found in table errors
                                strSQL = "SELECT COUNT(*) FROM audit_Post_SUMMARY " +
                                         "WHERE rxpackage='" + strPackage + "' AND " + 
                                               "(NF_IN_COND_TABLE_ERROR IS NOT NULL AND LEN(TRIM(NF_IN_COND_TABLE_ERROR)) > 0 AND TRIM(NF_IN_COND_TABLE_ERROR) <> 'NA' AND VAL(NF_IN_COND_TABLE_ERROR) > 0) OR " +
                                               "(NF_IN_PLOT_TABLE_ERROR IS NOT NULL AND LEN(TRIM(NF_IN_PLOT_TABLE_ERROR)) > 0 AND TRIM(NF_IN_PLOT_TABLE_ERROR) <> 'NA' AND VAL(NF_IN_PLOT_TABLE_ERROR) > 0) OR " +
                                               "(NF_IN_RX_TABLE_ERROR IS NOT NULL AND LEN(TRIM(NF_IN_RX_TABLE_ERROR)) > 0 AND TRIM(NF_IN_RX_TABLE_ERROR) <> 'NA' AND VAL(NF_IN_RX_TABLE_ERROR) > 0) OR " +
                                               "(NF_IN_RXPACKAGE_TABLE_ERROR IS NOT NULL AND LEN(TRIM(NF_IN_RXPACKAGE_TABLE_ERROR)) > 0 AND TRIM(NF_IN_RXPACKAGE_TABLE_ERROR) <> 'NA' AND VAL(NF_IN_RXPACKAGE_TABLE_ERROR) > 0) OR " +
                                               "(NF_RXPACKAGE_RXCYCLE_RX_ERROR IS NOT NULL AND LEN(TRIM(NF_RXPACKAGE_RXCYCLE_RX_ERROR)) > 0 AND TRIM(NF_RXPACKAGE_RXCYCLE_RX_ERROR) <> 'NA' AND VAL(NF_RXPACKAGE_RXCYCLE_RX_ERROR) > 0) OR " +
                                               "(NF_IN_TREE_TABLE_ERROR IS NOT NULL AND LEN(TRIM(NF_IN_TREE_TABLE_ERROR)) > 0 AND TRIM(NF_IN_TREE_TABLE_ERROR) <> 'NA' AND VAL(NF_IN_TREE_TABLE_ERROR) > 0)";
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                intRowCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, strSQL, "audit_Post_SUMMARY");
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                if (intRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = sqlArray[1];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
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

                                //check if SPCDCHANGE_WARNING table exists
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_SPCDCHANGE_WARNING") == false)
                                {
                                    oAdo.m_strSQL = Tables.FVS.Audit.Post.CreateFVSPostAuditCutlistFVSFIA_TREEMATCHINGtableSQL("audit_Post_SPCDCHANGE_WARNING","WARNING_DESC");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                                }
                                else
                                {
                                    //delete any records from a previous run for the current variant and rxpackage
                                    strSQL = sqlArray[0];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
                                //check if any no value errors
                                intRowCount = oAdo.getRecordCount(oAdo.m_OleDbConnection, "SELECT COUNT(*) FROM audit_Post_SUMMARY WHERE RXPACKAGE='" + strPackage + "' AND TREE_SPECIES_CHANGE_WARNING IS NOT NULL AND LEN(TRIM(TREE_SPECIES_CHANGE_WARNING)) > 0 AND TRIM(TREE_SPECIES_CHANGE_WARNING) <> 'NA' AND VAL(TREE_SPECIES_CHANGE_WARNING) > 0", "AUDIT_POST_SUMMARY");
                                if (intRowCount > 0)
                                {
                                    //insert the new audit records
                                    strSQL = sqlArray[1];
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                                }
                                m_intProgressStepCurrentCount++;
                                UpdateTherm(m_frmTherm.progressBar1,
                                            m_intProgressStepCurrentCount,
                                            m_intProgressStepTotalCount);
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
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_NOVALUE_ERROR"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_NOVALUE_ERROR\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_NOVALUE_ERROR\r\n---------------------------\r\n");

                                   //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_NOVALUE_ERROR WHERE RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                                    if (oAdo.m_OleDbDataReader.HasRows)
                                    {
                                        intItemError=-1;
                                        strItemError = strItemError + "\r\n\r\naudit_Post_NOVALUE_ERROR\r\n---------------------------\r\n";
                                        while (oAdo.m_OleDbDataReader.Read())
                                        {
                                            
                                            strItemError = strItemError + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg  + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    oAdo.m_OleDbDataReader.Close();
                                    oAdo.m_OleDbDataReader.Dispose();
                                }
                                //NOTFOUND errors
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_NOTFOUND_ERROR"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_NOTFOUND_ERROR\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_NOTFOUND_ERROR\r\n---------------------------\r\n");
                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_NOTFOUND_ERROR WHERE RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                                    if (oAdo.m_OleDbDataReader.HasRows)
                                    {
                                        intItemError = -1;
                                        strItemError = strItemError + "\r\n\r\naudit_Post_NOTFOUND_ERROR\r\n---------------------------\r\n";

                                        while (oAdo.m_OleDbDataReader.Read())
                                        {
                                           
                                            strItemError = strItemError + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    oAdo.m_OleDbDataReader.Close();
                                    oAdo.m_OleDbDataReader.Dispose();
                                }
                                //VALUE errors
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_VALUE_ERROR"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_VALUE_ERROR\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_VALUE_ERROR\r\n---------------------------\r\n");
                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_VALUE_ERROR WHERE RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                                    if (oAdo.m_OleDbDataReader.HasRows)
                                    {
                                        intItemError = -1;
                                        strItemError = strItemError + "\r\n\r\naudit_Post_VALUE_ERROR\r\n---------------------------\r\n";

                                        while (oAdo.m_OleDbDataReader.Read())
                                        {
                                            
                                            strItemError = strItemError + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    oAdo.m_OleDbDataReader.Close();
                                    oAdo.m_OleDbDataReader.Dispose();
                                }
                                //SPCD CHANGE WARNINGS
                                if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_SPCDCHANGE_WARNING"))
                                {
                                    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_SPCDCHANGE_WARNING\r\n---------------------------\r\n";
                                    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\naudit_Post_SPCDCHANGE_WARNING\r\n---------------------------\r\n");

                                    //see if any records
                                    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,WARNING_DESC FROM audit_Post_SPCDCHANGE_WARNING WHERE RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,WARNING_DESC";
                                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                                    if (oAdo.m_OleDbDataReader.HasRows)
                                    {
                                        intItemWarning = -1;
                                        strItemWarning = strItemWarning + "\r\n\r\naudit_Post_SPCDCHANGE_WARNING\r\n---------------------------\r\n";
                                        while (oAdo.m_OleDbDataReader.Read())
                                        {
                                            
                                            strItemWarning = strItemWarning + "WARNING: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["WARNING_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            strItemDialogMsg = strItemDialogMsg + "WARNING: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["WARNING_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                                            frmMain.g_oUtils.WriteText(m_strLogFile, "WARNING: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["WARNING_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                                        }
                                    }
                                    else
                                    {
                                        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                                        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                                    }
                                    oAdo.m_OleDbDataReader.Close();
                                    oAdo.m_OleDbDataReader.Dispose();
                                }
                            // TREEMATCH errors
                            //if (oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_TREEMATCH_ERROR"))
                            //{
                            //    strItemDialogMsg = strItemDialogMsg + "\r\n\r\naudit_Post_TREEMATCH_ERROR\r\n---------------------------\r\n";
                            //    frmMain.g_oUtils.WriteText(m_strLogFile, "\r\n\r\audit_Post_TREEMATCH_ERROR\r\n---------------------------\r\n");
                            //    //see if any records
                            //    strSQL = "SELECT COUNT(*) AS ROWCOUNT,COLUMN_NAME,ERROR_DESC FROM audit_Post_TREEMATCH_ERROR WHERE RXPACKAGE='" + strPackage + "' GROUP BY COLUMN_NAME,ERROR_DESC";
                            //    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, strSQL);
                            //    if (oAdo.m_OleDbDataReader.HasRows)
                            //    {
                            //        intItemError = -1;
                            //        strItemError = strItemError + "\r\n\r\naudit_Post_TREEMATCH_ERROR\r\n---------------------------\r\n";

                            //        while (oAdo.m_OleDbDataReader.Read())
                            //        {

                            //            strItemError = strItemError + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                            //            strItemDialogMsg = strItemDialogMsg + "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n";
                            //            frmMain.g_oUtils.WriteText(m_strLogFile, "ERROR: COLUMN:" + oAdo.m_OleDbDataReader["COLUMN_NAME"].ToString() + " MSG:" + oAdo.m_OleDbDataReader["ERROR_DESC"] + " Records:" + oAdo.m_OleDbDataReader["ROWCOUNT"].ToString().Trim() + "\r\n");
                            //        }
                            //    }
                            //    else
                            //    {
                            //        strItemDialogMsg = strItemDialogMsg + "OK\r\n";
                            //        frmMain.g_oUtils.WriteText(m_strLogFile, "OK\r\n");
                            //    }
                            //    oAdo.m_OleDbDataReader.Close();
                            //    oAdo.m_OleDbDataReader.Dispose();
                            //}

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

                            // Delete tmpCutTree table from SQLite database
                            m_dbConn = SQLite.GetConnectionString(m_strFvsTreeDb);
                            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_dbConn))
                            {
                                conn.Open();
                                if (SQLite.TableExist(conn, strTempCutListTable))
                                {
                                    SQLite.SqlNonQuery(conn, "DROP TABLE " + strTempCutListTable);
                                }
                            }

                            oAdo.CloseConnection(oAdo.m_OleDbConnection);
                            oAdo.m_OleDbConnection.Dispose();
                                
                                //compact and repair when file size is 70 percent of 2GB
                                if (uc_filesize_monitor1.CurrentPercent(strAuditDbFile,2000000000) > 70)
                                {
                                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Compact and Repair");
                                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");
                                    System.Threading.Thread.Sleep(5000);
                                    m_dao.CompactMDB(strAuditDbFile);
                                }
                            
                                
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

                System.Threading.Thread.Sleep(2000);
                this.FVSRecordsFinished();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (oAdo.m_OleDbConnection != null)
                {
                    if (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
                    {
                        oAdo.CloseConnection(oAdo.m_OleDbConnection);
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

                m_oRxTools.CreateFVSPrePostSeqNumTables(p_oAdo, p_oConn, m_oFVSPrePostSeqNumItem, p_strSourceTableName, p_strSourceTableName, p_bAudit, m_bDebug, m_strDebugFile);
            }
            else
            {
                if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, p_strSourceTableName + " table does not exist.\r\n\r\n");
                }
            }

            


        }
        private void CreateSummaryTableFVSPrePostYearWorkTables(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strSourceTableName, string p_strSourceLinkedTableName)
        {

            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateSummaryTableFVSPrePostYearWorkTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            
           
            




            if (p_strSourceTableName.Trim().ToUpper() == "FVS_SUMMARY")
            {
                if (p_oAdo.TableExist(p_oConn, m_strFVSSummaryAuditPrePostSeqNumTable))
                    p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE " + m_strFVSSummaryAuditPrePostSeqNumTable);

                if (p_oAdo.TableExist(p_oConn, m_strFVSSummaryAuditPrePostSeqNumCountsTable))
                    p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE " + m_strFVSSummaryAuditPrePostSeqNumCountsTable);

                if (p_oAdo.TableExist(p_oConn,  m_strFVSSummaryAuditYearCountsTable))
                    p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE " + m_strFVSSummaryAuditYearCountsTable);


                frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSeqNumAuditGenericTable(p_oAdo, p_oConn, m_strFVSSummaryAuditPrePostSeqNumTable);

                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditPrePostGenericSQL("", "FVS_SUMMARY", false);

                p_oAdo.m_strSQL = "INSERT INTO " + m_strFVSSummaryAuditPrePostSeqNumTable + " " +
                                  p_oAdo.m_strSQL;

                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePrePostGenericSQL(
                    m_oFVSPrePostSeqNumItem, m_strFVSSummaryAuditPrePostSeqNumTable);
                
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
                
                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPrePostSeqNumCount
                    (m_oFVSPrePostSeqNumItem, m_strFVSSummaryAuditPrePostSeqNumCountsTable, m_strFVSSummaryAuditPrePostSeqNumTable);
               
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_PrePostGenericSQL(
                    m_strFVSSummaryAuditYearCountsTable, "FVS_SUMMARY", false);

                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

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
		private void CreateGenericTablePrePostYearWorkTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn, string p_strSourceTableName, string p_strSourceLinkedTableName)
		{


            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateGenericTablePrePostYearWorkTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

			//drop tables if they exist
			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"audit_pre_post_rx_year_" + p_strSourceTableName.Trim()))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE audit_pre_post_rx_year_" + p_strSourceTableName.Trim());

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4");

				

			//pre and post treatment year summary audit table
			p_oAdo.m_strSQL = Tables.FVS.CreateFVSOutputRxPrePostCycleYearTableSQL("audit_pre_post_rx_year_" + p_strSourceTableName.Trim());
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);

			//populate the audit_pre_post_rx_year table with the FVSOUT table standid
			p_oAdo.m_strSQL = "INSERT INTO " + 
				                "audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " (standid) " + 
				               "SELECT DISTINCT standid FROM " + p_strSourceLinkedTableName.Trim();

            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1",
																				 p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","1");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
            p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2",
																		         p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","2");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3",
																			     p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","3");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4",
																		     	p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","4");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				              "INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1 b " + 
				              "ON a.standid=b.standid " + 
				              "SET a.pre_year1=b.pre_year1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year2=b.pre_year2";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year3=b.pre_year3";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year4=b.pre_year4";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPostYears("temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work","audit_pre_post_rx_year_" + p_strSourceTableName.Trim(),p_strSourceLinkedTableName.Trim());
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
			
			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditUpdatePostYears("audit_pre_post_rx_year_" + p_strSourceTableName.Trim(),"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " " + 
				              "SET pre_year1 = IIF(pre_year1 IS NULL,-1,pre_year1)," + 
								  "pre_year2 = IIF(pre_year2 IS NULL,-1,pre_year2)," + 
				                  "pre_year3 = IIF(pre_year3 IS NULL,-1,pre_year3)," + 
				                  "pre_year4 = IIF(pre_year4 IS NULL,-1,pre_year4)," + 
								  "post_year1 = IIF(post_year1 IS NULL,-1,post_year1)," + 
								  "post_year2 = IIF(post_year2 IS NULL,-1,post_year2)," + 
								  "post_year3 = IIF(post_year3 IS NULL,-1,post_year3)," + 
				                  "post_year4 = IIF(post_year4 IS NULL,-1,post_year4)";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

					

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4");


			
			

		}
		private void CreateStrClassTablePrePostYearWorkTables(ado_data_access p_oAdo,System.Data.OleDb.OleDbConnection p_oConn, string p_strSourceTableName, string p_strSourceLinkedTableName)
		{
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateStrClassTablePrePostYearWorkTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

			//drop tables if they exist
			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"audit_pre_post_rx_year_" + p_strSourceTableName.Trim()))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE audit_pre_post_rx_year_" + p_strSourceTableName.Trim());

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work1"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work1");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work2"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work2");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work3"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work3");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work4"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work4");

				

			//pre and post treatment year summary audit table
			p_oAdo.m_strSQL = Tables.FVS.CreateFVSOutputRxPrePostCycleYearTableSQL("audit_pre_post_rx_year_" + p_strSourceTableName.Trim());
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			//populate the audit_pre_post_rx_year table with the FVSOUT table standid
			p_oAdo.m_strSQL = "INSERT INTO " + 
				"audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " (standid) " + 
				"SELECT DISTINCT standid FROM " + p_strSourceLinkedTableName.Trim();

            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1",
				p_strSourceLinkedTableName.Trim(),"audit_pre_post_rx_year_fvs_summary","1");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2",
				p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","2");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3",
				p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","3");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditSelectIntoPreYears("temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4",
				p_strSourceLinkedTableName,"audit_pre_post_rx_year_fvs_summary","4");
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year1=b.pre_year1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year2=b.pre_year2";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year3=b.pre_year3";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4 b " + 
				"ON a.standid=b.standid " + 
				"SET a.pre_year4=b.pre_year4";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


			p_oAdo.m_strSQL = "SELECT b.pre_year1 AS post_year1,b.standid " + 
				              "INTO temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work1 " + 
			                  "FROM audit_pre_post_rx_year_" + p_strSourceTableName + " b," + 
							       p_strSourceLinkedTableName.Trim() + " a " + 
				              "WHERE b.standid=a.standid AND b.pre_year1=a.year AND a.removal_code=1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "SELECT b.pre_year2 AS post_year2,b.standid " + 
				"INTO temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work2 " + 
				"FROM audit_pre_post_rx_year_" + p_strSourceTableName + " b," + 
				p_strSourceLinkedTableName.Trim() + " a " + 
				"WHERE b.standid=a.standid AND b.pre_year2=a.year AND a.removal_code=1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "SELECT b.pre_year3 AS post_year3,b.standid " + 
				"INTO temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work3 " + 
				"FROM audit_pre_post_rx_year_" + p_strSourceTableName + " b," + 
				p_strSourceLinkedTableName.Trim() + " a " + 
				"WHERE b.standid=a.standid AND b.pre_year3=a.year AND a.removal_code=1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "SELECT b.pre_year4 AS post_year4,b.standid " + 
				"INTO temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work4 " + 
				"FROM audit_pre_post_rx_year_" + p_strSourceTableName + " b," + 
				p_strSourceLinkedTableName.Trim() + " a " + 
				"WHERE b.standid=a.standid AND b.pre_year4=a.year AND a.removal_code=1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work1 b " + 
				"ON a.standid=b.standid " + 
				"SET a.post_year1=b.post_year1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work2 b " + 
				"ON a.standid=b.standid " + 
				"SET a.post_year2=b.post_year2";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work3 b " + 
				"ON a.standid=b.standid " + 
				"SET a.post_year3=b.post_year3";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " a " + 
				"INNER JOIN temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work4 b " + 
				"ON a.standid=b.standid " + 
				"SET a.post_year4=b.post_year4";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


			p_oAdo.m_strSQL = "UPDATE audit_pre_post_rx_year_" + p_strSourceTableName.Trim() + " " + 
				"SET pre_year1 = IIF(pre_year1 IS NULL,-1,pre_year1)," + 
				"pre_year2 = IIF(pre_year2 IS NULL,-1,pre_year2)," + 
				"pre_year3 = IIF(pre_year3 IS NULL,-1,pre_year3)," + 
				"pre_year4 = IIF(pre_year4 IS NULL,-1,pre_year4)," + 
				"post_year1 = IIF(post_year1 IS NULL,-1,post_year1)," + 
				"post_year2 = IIF(post_year2 IS NULL,-1,post_year2)," + 
				"post_year3 = IIF(post_year3 IS NULL,-1,post_year3)," + 
				"post_year4 = IIF(post_year4 IS NULL,-1,post_year4)";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
			p_oAdo.SqlNonQuery(p_oConn,p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

					

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_POST_RX_YEAR_work"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work1");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work2");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work3");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_pre_rx_year_work4");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work1"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work1");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work2"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work2");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work3"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work3");

			if (p_oAdo.TableExist(p_oConn,"temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work4"))
				p_oAdo.SqlNonQuery(p_oConn,"DROP TABLE temp_" + p_strSourceTableName.Trim() + "_post_rx_year_work4");


			
			

		}


		
		
		private int Validate_FvsCasesVariant(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,string p_strTableName,string p_strVariant)
		{
			int intError=0;

			

			//check to ensure the variant in the fvs cases table
			//matches the current variant
			if (p_oAdo.getRecordCount(p_oConn,
				"SELECT COUNT(*) FROM (SELECT TOP 1 * FROM " + p_strTableName + " WHERE TRIM(variant) <> '" + p_strVariant + "')",
				p_strTableName) > 0)
			{
				intError=-1;
			}

			return intError;
			


		}
        private void Validate_FVSCaseId(string p_strTableName,ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, ref int p_intItemError, ref string p_strItemError, ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings)
        {
            //ERRORS
           
             

            Validate_MultipleCaseId(p_strTableName, p_oAdo, p_oConn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning, p_bDoWarnings);

            if (p_intItemError != 0) return;
            if (p_bDoWarnings == false) return;

            
        }
        /// <summary>
        /// Validate that every SeqNum value the user specified for PRE-POST values is found in the FVS_SUMMARY table
        /// </summary>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings"></param>
        private void Validate_FvsSummaryPrePostSeqNum(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, ref int p_intItemError, ref string p_strItemError, ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings)
        {
           
            
           
            int intCycleLength = 10;
           

            //get the cycle length for this package
            if (this.m_oRxPackageItem != null) intCycleLength = this.m_oRxPackageItem.RxCycleLength;

           
            Validate_SeqNumExistance("FVS_SUMMARY",p_oAdo, p_oConn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning, p_bDoWarnings);


           
			
        }
        /// <summary>
        /// Biosum cannot process FVS Output tables that have multiple caseid for a single standid in the FVS_CASES table.
        /// </summary>
        /// <param name="p_strFVSOutputTable"></param>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings"></param>
        private void Validate_MultipleCaseId(string p_strFVSOutputTable, ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, ref int p_intItemError, ref string p_strItemError, ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// Validate_MultipleCaseId\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (p_oAdo.TableExist(p_oConn, "temp_caseidcount"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_caseidcount");
            //each standid,year should be represented 1 time.
            p_oAdo.m_strSQL = "SELECT COUNT(*) AS ROWCOUNT,STANDID INTO temp_caseidcount FROM " + p_strFVSOutputTable + " GROUP BY STANDID";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

            p_oAdo.m_strSQL = "SELECT COUNT(*) AS RECORDCOUNT FROM temp_caseidcount WHERE ROWCOUNT > 1";
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
            if ((int)p_oAdo.getRecordCount(p_oConn, p_oAdo.m_strSQL, "temp_caseidcount") > 0)
            {
                p_intItemError = -1;
                p_strItemError = "FVS_CASEID table cannot contain more than one standid instance.\r\nTo resolve the problem, delete all the FVS Output tables and rerun the FVS treatments.";

            }
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

            if (p_oAdo.TableExist(p_oConn, "temp_caseidcount"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_caseidcount");

        }
        /// <summary>
        /// Validate that the user specified SEQNUM exists in the fvs output table
        /// </summary>
        /// <param name="p_strFVSOutputTable"></param>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings">AUDIT:true APPEND:false</param>
        private void Validate_SeqNumExistance(string p_strFVSOutputTable, ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, ref int p_intItemError, ref string p_strItemError, ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings)
        {
            string strCol = "";
            int x,z;

            string strAuditPrePostSeqNumCountsTable = "audit_" + p_strFVSOutputTable + "_prepost_seqnum_counts_table";

            p_oAdo.m_strSQL = "SELECT * FROM " + strAuditPrePostSeqNumCountsTable + " WHERE standid IS NOT NULL";

           

                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                if (p_oAdo.m_intError == 0)
                {
                    if (p_oAdo.m_OleDbDataReader.HasRows)
                    {
                        if (p_bDoWarnings)
                        {
                            z = 0;
                            while (p_oAdo.m_OleDbDataReader.Read())
                            {
                                //PRE TREATMENT
                                for (x = 1; x <= 4; x++)
                                {

                                    try
                                    {
                                        strCol = "pre_cycle" + x.ToString().Trim() + "rows";
                                        if (p_oAdo.m_OleDbDataReader[strCol] != DBNull.Value &&
                                            Convert.ToInt32(p_oAdo.m_OleDbDataReader[strCol]) == 0)
                                        {
                                            switch (x)
                                            {
                                                case 1:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum + " for cycle 1 PRE treatment values.\r\n";
                                                    break;
                                                case 2:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum + " for cycle 2 PRE treatment values.\r\n";
                                                    break;
                                                case 3:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum + " for cycle 3 PRE treatment values.\r\n";
                                                    break;
                                                case 4:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum + " for cycle 4 PRE treatment values.\r\n";
                                                    break;
                                            }
                                            p_intItemWarning = -3;
                                            break;
                                        }
                                        else if (p_oAdo.m_OleDbDataReader[strCol] != DBNull.Value &&
                                           Convert.ToInt32(p_oAdo.m_OleDbDataReader[strCol]) > 1)
                                        {
                                            switch (x)
                                            {
                                                case 1:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNum + " for cycle 1 PRE treatment values.\r\n";
                                                    break;
                                                case 2:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNum + " for cycle 2 PRE treatment values.\r\n";
                                                    break;
                                                case 3:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNum + " for cycle 3 PRE treatment values.\r\n";
                                                    break;
                                                case 4:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNum + " for cycle 4 PRE treatment values.\r\n";
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
                                        if (p_oAdo.m_OleDbDataReader[strCol] != DBNull.Value &&
                                            Convert.ToInt32(p_oAdo.m_OleDbDataReader[strCol]) == 0)
                                        {
                                            switch (x)
                                            {
                                                case 1:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum + " for cycle 1 POST treatment values.\r\n";
                                                    break;
                                                case 2:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum + " for cycle 2 POST treatment values.\r\n";
                                                    break;
                                                case 3:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum + " for cycle 3 POST treatment values.\r\n";
                                                    break;
                                                case 4:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " does not have Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum + " for cycle 4 POST treatment values.\r\n";
                                                    break;
                                            }
                                            p_intItemWarning = -5;
                                            break;
                                        }
                                        else if (p_oAdo.m_OleDbDataReader[strCol] != DBNull.Value &&
                                           Convert.ToInt32(p_oAdo.m_OleDbDataReader[strCol]) > 1)
                                        {
                                            switch (x)
                                            {
                                                case 1:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle1PostSeqNum + " for cycle 1 POST treatment values.\r\n";
                                                    break;
                                                case 2:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle2PostSeqNum + " for cycle 2 POST treatment values.\r\n";
                                                    break;
                                                case 3:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle3PostSeqNum + " for cycle 3 POST treatment values.\r\n";
                                                    break;
                                                case 4:
                                                    p_strItemWarning = p_strItemWarning + "WARNING: Stand" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " has more than one of the same Sequence Number " + m_oFVSPrePostSeqNumItem.RxCycle4PostSeqNum + " for cycle 4 POST treatment values.\r\n";
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
                    p_oAdo.m_OleDbDataReader.Close();
                }
                else
                {
                    p_intItemError = p_oAdo.m_intError;
                    p_strItemError = p_strItemError + "ERROR:" + p_oAdo.m_strError + "\r\n";
                }
            
        }
            
        /// <summary>
        /// Validation routine to check the FVS_TreeList,FVS_CutList,and FVS_ATRTLIST tables. Validate Checks:
        /// 1. Check for tree list records
        /// 2. Check if there are any standid,year records in the treelist table that 
        /// are not found in the fvs_summary table.  Every treelist standid,year combination should 
        /// be found in the fvs_summary table.
        /// </summary>
        /// <param name="p_oAdo"></param>
        /// <param name="p_oConn"></param>
        /// <param name="p_strTreeListTableName"></param>
        /// <param name="p_strSummaryTableName"></param>
        /// <param name="p_intItemError"></param>
        /// <param name="p_strItemError"></param>
        /// <param name="p_intItemWarning"></param>
        /// <param name="p_strItemWarning"></param>
        /// <param name="p_bDoWarnings"></param>
        private void Validate_TreeListTables(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,
            string p_strTreeListTableName, string p_strSummaryTableName,
            ref int p_intItemError, ref string p_strItemError,
            ref int p_intItemWarning, ref string p_strItemWarning, bool p_bDoWarnings)
        {
            string strWorkListTable = "";
            int x;
            string strSQL = "";
            
            //ERRORS
            //
            //see if any records 
            //
            p_oAdo.m_strSQL = "SELECT TOP 1 COUNT(*) FROM " + m_strFVSSummaryAuditYearCountsTable + " a WHERE a." + p_strTreeListTableName + " > 0";
            if ((int)p_oAdo.getRecordCount(p_oConn, p_oAdo.m_strSQL, strWorkListTable) == 0)
            {
                p_intItemError = -2;
                p_strItemError = p_strItemError + "ERROR: No trees in " + p_strTreeListTableName.Trim() + " table\r\n";
                return;
            }
            //
            //ensure the cut list standid,treatment year exists in the fvs summary table
            //
            if (p_oAdo.TableExist(p_oConn, "temp_treelist"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_treelist");

            if (p_oAdo.TableExist(p_oConn, "temp_summary"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_summary");

            if (p_oAdo.TableExist(p_oConn, "temp_missingrows"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_missingrows");

            string[] strSQLArray = Queries.FVS.FVSOutputTable_AuditSelectTreeListCyleYearExistInFVSSummaryTableSQL(
                                         "temp_treelist", "temp_summary","temp_missingrows",p_strTreeListTableName,m_strFVSSummaryAuditYearCountsTable);

            for (x = 0; x <= strSQLArray.Length - 1; x++)
            {
                strSQL = strSQLArray[x];
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + strSQL + "\r\n");
                p_oAdo.SqlNonQuery(p_oConn, strSQL);
                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

            }


            p_oAdo.m_strSQL = "SELECT * FROM temp_missingrows";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);

            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                p_intItemError = -4;
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    p_strItemError = p_strItemError + "ERROR: STANDID:" + p_oAdo.m_OleDbDataReader["standid"].ToString().Trim() + " " +
                        "YEAR:" + p_oAdo.m_OleDbDataReader["year"].ToString().Trim() + " " +
                        "TREECOUNT:" + p_oAdo.m_OleDbDataReader["treecount"].ToString().Trim() + " Standid and year not found in the fvs_summary table\r\n";
                }

            }
            p_oAdo.m_OleDbDataReader.Close();

            if (p_oAdo.TableExist(p_oConn, "temp_missingrows"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE temp_missingrows");


            if (p_intItemError != 0) return;

        }
		/// <summary>
		/// The validation is two fold:
        /// 1. Validate that every tree, except FVS created trees, in the FVS_CUTLIST can be found in the FIA tree table
        /// 2. Check if a tree is cut multiple times
        /// An error is returned if cutlist trees can't be found in the tree table.
        /// A warning is returned if a tree is cut more than once.
        /// </summary>
		/// <param name="p_oAdo"></param>
		/// <param name="p_oConn"></param>
		/// <param name="p_strFVSOutDBFile"></param>
		/// <param name="p_strFVSCasesTableName"></param>
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
        private void Validate_FVSTreeId(ado_data_access p_oAdo, 
                                        System.Data.OleDb.OleDbConnection p_oConn,
                                        string p_strFVSOutDBFile,
                                        string p_strFVSCasesTableName,
                                        string p_strFVSTreeTableNameToAudit,
                                        string p_strVariant,
                                        string p_strRxPackage,
                                        string p_strRx1,
                                        string p_strRx2,
                                        string p_strRx3,
                                        string p_strRx4,
                                        bool p_bAudit,
                                        ref int p_intItemWarning, 
                                        ref string p_strItemWarning,
                                        ref int p_intItemError, 
                                        ref string p_strItemError)
                                        
        {
            int x,y;
            string strRxCycle;
            string strRx="";
            string strRxYear;
            int intTreeTable;
            int intCondTable;
            int intPlotTable;
            string strConn = "";
            if (p_bAudit)
            {
                strConn = p_oConn.ConnectionString;
                intTreeTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("TREE");
                intCondTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("CONDITION");
                intPlotTable = m_oQueries.m_oDataSource.getDataSourceTableNameRow("PLOT");
                if (p_oAdo.TableExist(p_oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE]))
                {
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE]);
                }
                if (p_oAdo.TableExist(p_oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE]))
                {
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE]);
                }
                if (p_oAdo.TableExist(p_oAdo.m_OleDbConnection, m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE]))
                {
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, "DROP TABLE " + m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE]);
                }
                p_oAdo.CloseConnection(p_oAdo.m_OleDbConnection);


                dao_data_access oDao = new dao_data_access();


                //tree table link
                oDao.CreateTableLink(p_strFVSOutDBFile, m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE],
                                      m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.PATH].Trim() + "\\" +
                                       m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.MDBFILE].Trim(),
                                      m_oQueries.m_oDataSource.m_strDataSource[intTreeTable, Datasource.TABLE], true);
                //condition table link

                m_dao.CreateTableLink(p_strFVSOutDBFile, m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE],
                                      m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.PATH].Trim() + "\\" +
                                       m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.MDBFILE].Trim(),
                                      m_oQueries.m_oDataSource.m_strDataSource[intCondTable, Datasource.TABLE], true);
                //plot table link

                m_dao.CreateTableLink(p_strFVSOutDBFile, m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE],
                                      m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.PATH].Trim() + "\\" +
                                       m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.MDBFILE].Trim(),
                                      m_oQueries.m_oDataSource.m_strDataSource[intPlotTable, Datasource.TABLE], true);

                oDao.m_DaoWorkspace.Close();
                oDao = null;
                System.Threading.Thread.Sleep(2000);

                p_oAdo.OpenConnection(strConn);
            }

            //
            //validate fvs_tree_id
            //
            //create temp fvs_tree table
            if (p_oAdo.TableExist(p_oAdo.m_OleDbConnection,"audit_fvs_tree_id"))
                p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection,"DROP TABLE audit_fvs_tree_id");

            frmMain.g_oTables.m_oFvs.CreateFVSTreeIdAudit(p_oAdo, p_oAdo.m_OleDbConnection, "audit_fvs_tree_id");

            //
			//loop through all the rx cycles for this package and append them to the fvs tree
			//
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
                    strRxYear = "(t.standid=p.standid AND " +
                        "t.year=p.pre_year" + strRxCycle + ")";


                    
                    p_oAdo.m_strSQL = Queries.FVS.FVSOutputTable_AuditFVSTreeId(
                                                "audit_fvs_tree_id",
                                                    p_strFVSCasesTableName,
                                                    p_strFVSTreeTableNameToAudit,
                                                    "FVS_CUTLIST_PREPOST_SEQNUM_MATRIX",
                                                    p_strRxPackage,
                                                    strRx,
                                                    strRxCycle,
                                                    strRxYear);

                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
                    p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");




                }
            }
            p_oAdo.m_strSQL = "UPDATE audit_fvs_tree_id  i " +
                                      "INNER JOIN ((" + m_oQueries.m_oFIAPlot.m_strTreeTable + " t " +
                                                      "INNER JOIN " + m_oQueries.m_oFIAPlot.m_strCondTable + " c " +
                                                      "ON t.biosum_cond_id=c.biosum_cond_id) " +
                                                      "INNER JOIN " + m_oQueries.m_oFIAPlot.m_strPlotTable + " p " +
                                                      "ON p.biosum_plot_id=c.biosum_plot_id) " +
                                      "ON i.fvs_tree_id=t.fvs_tree_id AND i.biosum_cond_id=t.biosum_cond_id " +
                                      "SET i.FOUND_FvsTreeId_YN='Y'";

            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + p_oAdo.m_strSQL + "\r\n");
            p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (m_bDebug && frmMain.g_intDebugLevel > 2)
                this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

            int intCount = (int)p_oAdo.getSingleDoubleValueFromSQLQuery(p_oAdo.m_OleDbConnection,"SELECT COUNT(*) AS TTLCOUNT FROM audit_fvs_tree_id WHERE FOUND_FvsTreeId_YN='N'","TEMP");

            if (intCount > 0)
            {
                if (p_bAudit) p_strItemError = p_strItemError + "ERROR: " + intCount.ToString() + " fvs_cutlist trees for variant " + p_strVariant + " could not be found in the FIADB tree table by matching FVS_Tree_Id (See audit table audit_fvs_tree_id)\r\n";
                else
                    p_strItemError = p_strItemError +  " " + intCount.ToString() + " fvs_cutlist trees for variant " + p_strVariant + " could not be found in the FIADB tree table by matching FVS_Tree_Id (See audit table audit_fvs_tree_id)\r\n";
                p_intItemError = -5;

            }
        }
		
		private void Validate_PotFire(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,
			string p_strPotFireTableName,string p_strVariant,
			ref int p_intItemError, ref string p_strItemError, 
			ref int p_intItemWarning, ref string p_strItemWarning,bool p_bDoWarnings)
		{
			
			
            
           

			

			//ERRORS
            if ((m_oFVSPrePostSeqNumItem.RxCycle1PreSeqNumBaseYearYN == "Y" ||
                m_oFVSPrePostSeqNumItem.RxCycle2PreSeqNumBaseYearYN == "Y" ||
                m_oFVSPrePostSeqNumItem.RxCycle3PreSeqNumBaseYearYN == "Y" ||
                m_oFVSPrePostSeqNumItem.RxCycle4PreSeqNumBaseYearYN == "Y") &&
                m_bPotFireBaseYearTableExist==false)
            {
                p_intItemError = -2;
                p_strItemError = "ERROR: POTFIRE Base year file and/or table missing";
                return;
            }
			//
			//see if any records 
			//
			p_oAdo.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM " + p_strPotFireTableName + ")";
			if ((int)p_oAdo.getRecordCount(p_oConn,p_oAdo.m_strSQL,"FVS_POTFIRE") == 0)
			{ 
				p_intItemError=-2;
				p_strItemError=p_strItemError + "ERROR: No fvs potfire records\r\n";
				return;
			}

            if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "N")
                Validate_SeqNumExistance(p_strPotFireTableName, p_oAdo, p_oConn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning, p_bDoWarnings);

		}
		
		private void Validate_FVSGenericTable(ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, 
			string p_strFvsTableName, string p_strFvsLinkedTableName, ref int p_intItemError, ref string p_strItemError, 
			ref int p_intItemWarning, ref string p_strItemWarning,bool p_bDoWarnings)
		{
			
			string strField=p_strFvsTableName + "_count";
			string strAuditPrePostTable = "audit_pre_post_rx_year_" + p_strFvsTableName;
						

			//ERRORS
			//
			//see if any records 
			//
			p_oAdo.m_strSQL = "SELECT COUNT(*) FROM (SELECT TOP 1 standid FROM " + p_strFvsLinkedTableName.Trim() + ")";
			if ((int)p_oAdo.getRecordCount(p_oConn,p_oAdo.m_strSQL,p_strFvsLinkedTableName) == 0)
			{ 
				p_intItemError=-2;
				p_strItemError=p_strItemError + "ERROR: No " + p_strFvsLinkedTableName + " records\r\n";
				return;
			}

            if (m_oFVSPrePostSeqNumItem.UseSummaryTableSeqNumYN == "N")
                Validate_SeqNumExistance(p_strFvsLinkedTableName, p_oAdo, p_oConn, ref p_intItemError, ref p_strItemError, ref p_intItemWarning, ref p_strItemWarning, p_bDoWarnings);
			

			if (p_intItemError != 0) return;
			if (p_bDoWarnings==false) return;


			
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
            RunAppend_CloseDbConnections(m_oPrePostDbFileItem_Collection);
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

			string strSearch = this.lstFvsOutput.SelectedItems[0].SubItems[COL_MDBOUT].Text.Trim().ToUpper().Replace(".MDB","_BIOSUM.ACCDB") + "_AUDIT_*.txt";
			
			string strDirectory = this.txtOutDir.Text.Trim() + "\\" + lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
			
			string[]  strFiles= System.IO.Directory.GetFiles(strDirectory,strSearch);

			FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();

			oDlg.uc_select_list_item1.lblTitle.Text = "Open Audit Log File";
			oDlg.uc_select_list_item1.listBox1.Sorted = true;
			for (int x=0;x<=strFiles.Length - 1;x++)
			{
				oDlg.uc_select_list_item1.listBox1.Items.Add(strFiles[x].Substring(strDirectory.Length+1,strFiles[x].Length - strDirectory.Length - 1));
			}
			if (oDlg.uc_select_list_item1.listBox1.Items.Count > 0) oDlg.uc_select_list_item1.listBox1.SelectedIndex = oDlg.uc_select_list_item1.listBox1.Items.Count-1;
			oDlg.uc_select_list_item1.lblMsg.Text = "Log File Contents of " + strDirectory;
			oDlg.uc_select_list_item1.lblMsg.Show();

			oDlg.uc_select_list_item1.Show();

			DialogResult result = oDlg.ShowDialog();
			if (result==DialogResult.OK)
			{
				string strDirAndFile = strDirectory + "\\" + oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim();
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


		    string strDbFile = this.lstFvsOutput.SelectedItems[0].SubItems[COL_MDBOUT].Text.Trim();
            strDbFile = strDbFile.Replace(".MDB", "_BIOSUM.ACCDB");
			string strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir,"Text",false);
			strOutDirAndFile=strOutDirAndFile.Trim()  + "\\" + lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();	
			strOutDirAndFile = strOutDirAndFile  + "\\" + strDbFile;
            if (System.IO.File.Exists(strOutDirAndFile))
            {
                ado_data_access oAdo = new ado_data_access();
                strConn = oAdo.getMDBConnString(strOutDirAndFile, "", "");
                oAdo.OpenConnection(strConn);

                if (!oAdo.TableExist(oAdo.m_OleDbConnection, this.m_strFVSSummaryAuditYearCountsTable))
                {
                    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                    oAdo = null;
                    string strWarnMessage = "No PRE-APPEND audit tables exist in the file " + strOutDirAndFile + ". The PRE-APPEND Audit tables cannot be displayed.";
                    MessageBox.Show(strWarnMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                FIA_Biosum_Manager.frmGridView oFrm = new frmGridView();



             



                oFrm.Text = "Database: Browse (PRE-APPEND Audit Tables)";
                if (m_strFVSPreAppendAuditTables != null)
                {
                    for (int x = 0; x <= m_strFVSPreAppendAuditTables.Count - 1; x++)
                    {
                        if (oAdo.TableExist(oAdo.m_OleDbConnection, m_strFVSPreAppendAuditTables[x].Trim()))
                        {
                            oFrm.LoadDataSet(oAdo.m_OleDbConnection, strConn, "SELECT * FROM " + m_strFVSPreAppendAuditTables[x].Trim(), m_strFVSPreAppendAuditTables[x].Trim());
                        }
                    }
                }

                oAdo.CloseConnection(oAdo.m_OleDbConnection);
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
            if (this.lstFvsOutput.SelectedItems.Count > 0 && frmMain.g_oDelegate.CurrentThreadProcessIdle)
            {
                m_oLvAlternateColors.DelegateListViewItem(lstFvsOutput.SelectedItems[0]);

                //Enable/Disable PRE-APPEND Audit Tables
                string strDbFile = this.lstFvsOutput.SelectedItems[0].SubItems[COL_MDBOUT].Text.Trim();
                strDbFile = strDbFile.Replace(".MDB", "_BIOSUM.ACCDB");
                string strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                strOutDirAndFile = strOutDirAndFile.Trim() + "\\" + lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
                strOutDirAndFile = strOutDirAndFile + "\\" + strDbFile;
                if (System.IO.File.Exists(strOutDirAndFile))
                   btnAuditDb.Enabled = true;
                else
                   btnAuditDb.Enabled = false;

                //Enable/Disable POST-APPEND Audit Tables
                string strAuditDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
                strAuditDbFile = strAuditDbFile.Trim();
                string strVariant = lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
                strAuditDbFile = strAuditDbFile + "\\" + strVariant + "\\PostAudit.accdb";

                if (System.IO.File.Exists(strAuditDbFile))
                    btnPostAppendAuditDb.Enabled = true;
                else
                    btnPostAppendAuditDb.Enabled = false;

                //Enable/Disable Open Pre Audit Log button
                btnViewLogFile.Enabled = false;
                string strDirectory = this.txtOutDir.Text.Trim() + "\\" + lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
                if (System.IO.Directory.Exists(strDirectory) == true)
                {
                    string strSearch = this.lstFvsOutput.SelectedItems[0].SubItems[COL_MDBOUT].Text.Trim().ToUpper().Replace(".MDB","_BIOSUM.ACCDB") + "_AUDIT_*.txt";
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
                for (int i = 0; i < this.lstFvsOutput.CheckedItems.Count; i++)
                {
                    var lvItem = this.lstFvsOutput.CheckedItems[i];
                    string strVariant = lvItem.SubItems[COL_VARIANT].Text.Trim();
                    string strRxPackage = lvItem.SubItems[COL_PACKAGE].Text.Trim();
                    long lngTreeRecords = -1;
                    string strSQL = $@"SELECT COUNT(*) FROM {Tables.FVS.DefaultFVSCutTreeTableName} 
                                           WHERE FVS_VARIANT = '{strVariant}' and RXPACKAGE = '{strRxPackage}'";
                    lngTreeRecords = SQLite.getRecordCount(conn, strSQL, Tables.FVS.DefaultFVSCutTreeTableName);

                    if (lngTreeRecords < 1)
                    {
                        DialogResult result = MessageBox.Show("!!Warning!!\r\n-----------\r\nCut Tree Table for  " + strVariant + " " + strRxPackage + " " +
                            " Does Not Exist. Continue Auditing ?(Y/N)",
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
            // Check checked item count
            if (this.lstFvsOutput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
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
        
        //private void ConvertAlphaSpCd()
        //{
        //    if (this.lstFvsOutput.CheckedItems.Count == 0)
        //    {
        //        MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
        //        return;
        //    }

        //    if (this.m_intError == 0)
        //    {


        //        this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
        //            "FVS Output", "2");
        //        this.m_frmTherm.lblMsg.Text = "";
        //        this.cmbStep.Enabled = false;
        //        this.btnExecute.Enabled = false;
        //        this.btnChkAll.Enabled = false;
        //        this.btnClearAll.Enabled = false;
        //        this.btnRefresh.Enabled = false;
        //        this.btnClose.Enabled = false;
        //        this.btnHelp.Enabled = false;
        //        this.btnCancel.Visible = false;
        //        this.btnViewLogFile.Enabled = false;
        //        this.btnViewPostLogFile.Enabled = false;
        //        this.btnAuditDb.Enabled = false;
        //        this.btnPostAppendAuditDb.Enabled = false;
        //       // this.grpboxSpCdConvert.Enabled = false;


        //        frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
        //        frmMain.g_oDelegate.CurrentThreadProcessDone = false;
        //        frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
        //        frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(RunConvertAlphaSpCd_Start));
        //        frmMain.g_oDelegate.InitializeThreadEvents();
        //        frmMain.g_oDelegate.m_oThread.IsBackground = true;
        //        frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
        //        frmMain.g_oDelegate.m_oThread.Start();


        //    }
        //}
        //private void RunConvertAlphaSpCd_Start()
        //{

        //    frmMain.g_oDelegate.InitializeThreadEvents();
        //    frmMain.g_oDelegate.m_oEventStopThread.Reset();
        //    frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
        //    frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.RunConvertAlphaSpCd_Main));
        //    frmMain.g_oDelegate.m_oThread.IsBackground = true;
        //    frmMain.g_oDelegate.m_oThread.Start();


        //}
        //private void RunConvertAlphaSpCd_Main()
        //{


        //    frmMain.g_oDelegate.CurrentThreadProcessName = "main";
        //    frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
        //    System.Threading.Thread.Sleep(2000);

        //    this.m_intError = 0;
        //    int intCount = 0;
           
        //    m_strError = "";
        //    m_strWarning = "";
        //    m_intWarning = 0;

        //    m_intProgressOverallCurrentCount = 0;
        //    m_intProgressOverallTotalCount = 0;
        //    m_intProgressStepCurrentCount = 0;
        //    m_intProgressStepTotalCount = 0;

        //    string strRx1 = "";
        //    string strRx2 = "";
        //    string strRx3 = "";
        //    string strRx4 = "";
        //    string strPackage = "";
        //    string strVariant = "";
        //    System.Windows.Forms.ListView oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(this.lstFvsOutput, false);
        //    System.Windows.Forms.ListViewItem oLvItem = null;
        //    bool bSkip = false;


        //    Tables oTables = new Tables();



        //    string strOutDirAndFile;
        //    string strDbFile;
            



        //    string[] strSourceTableArray = null;
        //    ado_data_access oAdo = new ado_data_access();

        //    if (m_bDebug)
        //        frmMain.g_oUtils.WriteText(m_strDebugFile, "*****START*****" + System.DateTime.Now.ToString() + "\r\n");


        //    int x, y,z;
        //    int intTranslatorTable = 0;
        //    try
        //    {


        //        if (this.m_ado.m_OleDbConnection.State == System.Data.ConnectionState.Open)
        //            this.m_ado.m_OleDbConnection.Close();

        //        while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
        //        {
        //            System.Threading.Thread.Sleep(1000);
        //        }
                

        //        intCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);
        //        for (x = 0; x <= intCount - 1; x++)
        //        {
        //            oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
        //            m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = true;
        //            m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
        //            //see if checked
        //            if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false))
        //                m_intProgressOverallTotalCount++;
        //        }
        //        //
        //        //INITIALIZE OVERALL PROGRESS BAR
        //        //
        //        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
        //        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
        //        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
        //        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
        //        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
        //        frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);

        //        for (x = 0; x <= intCount - 1; x++)
        //        {
        //            oLvItem = (System.Windows.Forms.ListViewItem)frmMain.g_oDelegate.GetListViewItem(oLv, x, false);
        //            this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
        //            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");
        //            this.m_oLvAlternateColors.DelegateListViewSubItem(oLvItem, x, COL_RUNSTATUS);
        //            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "");

        //            if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false) == true)
        //            {

        //                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 100);
        //                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
        //                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 0);



                       


        //                m_intProgressStepCurrentCount = 0;
        //                m_intProgressStepTotalCount = 5;



        //                this.m_oLvAlternateColors.m_oRowCollection.Item(x).m_oColumnCollection.Item(COL_RUNSTATUS).UpdateColumn = false;
        //                frmMain.g_oDelegate.ExecuteControlMethodWithParam((System.Windows.Forms.Control)oLv, "EnsureVisible", new object[] { x });
        //                frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Selected", true);
        //                frmMain.g_oDelegate.SetListViewItemPropertyValue(oLv, x, "Focused", true);



        //                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.DarkGoldenrod);
        //                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
        //                frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Convert FVS Alpha Codes To FIA Codes");

        //                //get the variant
        //                strVariant = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_VARIANT, "Text", false);
        //                strVariant = strVariant.Trim();

        //                //get the package and treatments
        //                strPackage = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_PACKAGE, "Text", false);
        //                strPackage = strPackage.Trim();

        //                strRx1 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE1, "Text", false);
        //                strRx1 = strRx1.Trim();

        //                strRx2 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE2, "Text", false);
        //                strRx2 = strRx2.Trim();

        //                strRx3 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE3, "Text", false);
        //                strRx3 = strRx3.Trim();

        //                strRx4 = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_RXCYCLE4, "Text", false);
        //                strRx4 = strRx4.Trim();

        //                //find the package item in the package collection
        //                for (y = 0; y <= this.m_oRxPackageItem_Collection.Count - 1; y++)
        //                {
        //                    if (this.m_oRxPackageItem_Collection.Item(y).SimulationYear1Rx.Trim() == strRx1.Trim() &&
        //                        this.m_oRxPackageItem_Collection.Item(y).SimulationYear2Rx.Trim() == strRx2.Trim() &&
        //                        this.m_oRxPackageItem_Collection.Item(y).SimulationYear3Rx.Trim() == strRx3.Trim() &&
        //                        this.m_oRxPackageItem_Collection.Item(y).SimulationYear4Rx.Trim() == strRx4.Trim() &&
        //                        this.m_oRxPackageItem_Collection.Item(y).RxPackageId.Trim() == strPackage.Trim())
        //                        break;


        //                }
        //                if (y <= m_oRxPackageItem_Collection.Count - 1)
        //                {
        //                    this.m_oRxPackageItem = new RxPackageItem();
        //                    m_oRxPackageItem.CopyProperties(m_oRxPackageItem_Collection.Item(y), m_oRxPackageItem);
        //                }
        //                else
        //                {
        //                    this.m_oRxPackageItem = null;
        //                }

        //                //get the list of treatment cycle year fields to reference for this package
        //                this.m_strRxCycleList = "";
        //                if (strRx1.Trim().Length > 0 && strRx1.Trim() != "000") this.m_strRxCycleList = "1,";
        //                if (strRx2.Trim().Length > 0 && strRx2.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "2,";
        //                if (strRx3.Trim().Length > 0 && strRx3.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "3,";
        //                if (strRx4.Trim().Length > 0 && strRx4.Trim() != "000") this.m_strRxCycleList = this.m_strRxCycleList + "4,";

        //                if (this.m_strRxCycleList.Trim().Length > 0)
        //                    this.m_strRxCycleList = this.m_strRxCycleList.Substring(0, this.m_strRxCycleList.Length - 1);

        //                this.m_strRxCycleArray = frmMain.g_oUtils.ConvertListToArray(this.m_strRxCycleList, ",");

        //                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim());
        //                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

        //                strDbFile = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, COL_MDBOUT, "Text", false);
        //                strDbFile = strDbFile.Trim();

        //                strOutDirAndFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
        //                strOutDirAndFile = strOutDirAndFile.Trim();

                        



        //                strOutDirAndFile = strOutDirAndFile + "\\" + strVariant + "\\" + strDbFile;

        //                uc_filesize_monitor1.BeginMonitoringFile(
        //                    strOutDirAndFile,
        //                    2000000000, "2gb");

        //                m_intProgressStepCurrentCount++;
        //                UpdateTherm(this.m_frmTherm.progressBar1,
        //                            m_intProgressStepCurrentCount,
        //                            m_intProgressStepTotalCount);

                        
        //                dao_data_access oDao = new dao_data_access();
        //                oDao.OpenDb(strOutDirAndFile);
        //                intTranslatorTable=m_oQueries.m_oDataSource.getValidTableNameRow("FVS WESTERN TREE SPECIES TRANSLATOR");

        //                if (oDao.TableExists(oDao.m_DaoDatabase, m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, 4].Trim()))
        //                    oDao.DeleteTableFromMDB(oDao.m_DaoDatabase, m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.TABLE].Trim());
        //                oDao.CreateTableLink(oDao.m_DaoDatabase,
        //                    m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.TABLE].Trim(),
        //                    m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.PATH].Trim() + "\\" +
        //                    m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.MDBFILE].Trim(),
        //                    m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.TABLE].Trim());
        //                oDao.m_DaoDatabase.Close();
        //                oDao.m_DaoWorkspace.Close();
        //                oDao = null;

        //                m_intProgressStepCurrentCount++;
        //                UpdateTherm(this.m_frmTherm.progressBar1,
        //                            m_intProgressStepCurrentCount,
        //                            m_intProgressStepTotalCount);

                        
        //                System.Threading.Thread.Sleep(3000);


        //                oAdo.OpenConnection(oAdo.getMDBConnString(strOutDirAndFile, "", ""));



        //                oAdo.DisplayErrors = false;

        //                strSourceTableArray = oAdo.getTableNames(oAdo.m_OleDbConnection);

        //                m_intProgressStepCurrentCount++;
        //                UpdateTherm(this.m_frmTherm.progressBar1,
        //                            m_intProgressStepCurrentCount,
        //                            m_intProgressStepTotalCount);
        //                bSkip = true;
        //                for (y = 0; y <= strSourceTableArray.Length - 1; y++)
        //                {
        //                    if (strSourceTableArray[y] == null) break;
        //                    bSkip = true;
        //                    for (z = 0; z <= Tables.FVS.g_strFVSOutTablesArray.Length - 1; z++)
        //                    {
        //                        if (strSourceTableArray[y].Trim().ToUpper() ==
        //                            Tables.FVS.g_strFVSOutTablesArray[z].Trim().ToUpper())
        //                        {
        //                            bSkip = false; break;
        //                        }

        //                    }
        //                    if (bSkip == false)
        //                    {
        //                        //
        //                        //FVS_TREELIST
        //                        //
        //                        if (strSourceTableArray[y].Trim().ToUpper() == "FVS_TREELIST")
        //                        {
        //                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " FVS_TREELIST");
        //                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

        //                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing...FVS_Treelist");
        //                            if ((double)oAdo.getSingleDoubleValueFromSQLQuery(oAdo.m_OleDbConnection, "SELECT COUNT(*) AS ROWCOUNT FROM FVS_TREELIST", "FVS_TREELIST") > 0)
        //                            {
        //                                if (oAdo.ColumnExist(oAdo.m_OleDbConnection, "FVS_TREELIST", "species_temp"))
        //                                {
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_TREELIST DROP COLUMN species_temp\r\n");
        //                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_TREELIST DROP COLUMN species_temp");
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                                }

        //                                oAdo.m_strSQL = "ALTER TABLE FVS_TREELIST ADD COLUMN species_temp TEXT(10)";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

        //                                oAdo.m_strSQL = "UPDATE FVS_TREELIST SET species_temp=species";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


        //                                oAdo.m_strSQL = "UPDATE FVS_TREELIST a " +
        //                                                "INNER JOIN " + m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.TABLE].Trim() + " b " +
        //                                                "ON a.species_temp=b.USDA_PLANTS_SYMBOL " +
        //                                                "SET a.species=b.fia_spcd";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

        //                                if (oAdo.ColumnExist(oAdo.m_OleDbConnection, "FVS_TREELIST", "species_temp"))
        //                                {
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_TREELIST DROP COLUMN species_temp\r\n");
        //                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_TREELIST DROP COLUMN species_temp");
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                                }
        //                            }

        //                        }
        //                        //
        //                        //FVS_CUTLIST
        //                        //
        //                        else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_CUTLIST")
        //                        {
        //                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " FVS_CUTLIST");
        //                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

        //                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing...FVS_Cutlist");
        //                            if ((double)oAdo.getSingleDoubleValueFromSQLQuery(oAdo.m_OleDbConnection, "SELECT COUNT(*) AS ROWCOUNT FROM FVS_CUTLIST", "FVS_CUTLIST") > 0)
        //                            {
        //                                if (oAdo.ColumnExist(oAdo.m_OleDbConnection, "FVS_CUTLIST", "species_temp"))
        //                                {
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_CUTLIST DROP COLUMN species_temp\r\n");
        //                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_CUTLIST DROP COLUMN species_temp");
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                                }

        //                                oAdo.m_strSQL = "ALTER TABLE FVS_CUTLIST ADD COLUMN species_temp TEXT(10)";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

        //                                oAdo.m_strSQL = "UPDATE FVS_CUTLIST SET species_temp=species";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


        //                                oAdo.m_strSQL = "UPDATE FVS_CUTLIST a " +
        //                                                "INNER JOIN " + m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.TABLE].Trim() + " b " +
        //                                                "ON a.species_temp=b.USDA_PLANTS_SYMBOL " +
        //                                                "SET a.species=b.fia_spcd";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

        //                                if (oAdo.ColumnExist(oAdo.m_OleDbConnection, "FVS_CUTLIST", "species_temp"))
        //                                {
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_CUTLIST DROP COLUMN species_temp\r\n");
        //                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_CUTLIST DROP COLUMN species_temp");
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                                }
        //                            }


        //                        }
        //                        //
        //                        //FVS_ATRTLIST
        //                        //
        //                        else if (strSourceTableArray[y].Trim().ToUpper() == "FVS_ATRTLIST")
        //                        {
        //                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "Processing Variant:" + strVariant.Trim() + " Package:" + strPackage.Trim() + " FVS_ATRTLIST");
        //                            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Refresh");

        //                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "Processing...FVS_ATRTList");
        //                            if ((double)oAdo.getSingleDoubleValueFromSQLQuery(oAdo.m_OleDbConnection, "SELECT COUNT(*) AS ROWCOUNT FROM FVS_ATRTLIST", "FVS_ATRTLIST") > 0)
        //                            {
        //                                if (oAdo.ColumnExist(oAdo.m_OleDbConnection, "FVS_ATRTLIST", "species_temp"))
        //                                {
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_ATRTLIST DROP COLUMN species_temp\r\n\r\n");
        //                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_ATRTLIST DROP COLUMN species_temp");
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                                }

        //                                oAdo.m_strSQL = "ALTER TABLE FVS_ATRTLIST ADD COLUMN species_temp TEXT(10)";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

        //                                oAdo.m_strSQL = "UPDATE FVS_ATRTLIST SET species_temp=species";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");


        //                                oAdo.m_strSQL = "UPDATE FVS_ATRTLIST a " +
        //                                                "INNER JOIN " + m_oQueries.m_oDataSource.m_strDataSource[intTranslatorTable, Datasource.TABLE].Trim() + " b " +
        //                                                "ON a.species_temp=b.USDA_PLANTS_SYMBOL " +
        //                                                "SET a.species=b.fia_spcd";
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n\r\n");
        //                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                    this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

        //                                if (oAdo.ColumnExist(oAdo.m_OleDbConnection, "FVS_ATRTLIST", "species_temp"))
        //                                {
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\nALTER TABLE FVS_ATRTLIST DROP COLUMN species_temp\r\n");
        //                                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, "ALTER TABLE FVS_ATRTLIST DROP COLUMN species_temp");
        //                                    if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                                        this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                                }
        //                            }




        //                        }
        //                    }

        //                }


        //                m_intProgressStepCurrentCount++;
        //                UpdateTherm(this.m_frmTherm.progressBar1,
        //                            m_intProgressStepCurrentCount,
        //                            m_intProgressStepTotalCount);

        //                if (oAdo.m_intError==0)
        //                {
        //                    if (oAdo.TableExist(oAdo.m_OleDbConnection, "FVS_CASES"))
        //                    {
        //                        oAdo.m_strSQL = "UPDATE FVS_CASES SET BIOSUM_FVSAlphaToFIANumeric_YN='Y'";
        //                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                            this.WriteText(m_strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
        //                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
        //                        if (m_bDebug && frmMain.g_intDebugLevel > 2)
        //                            this.WriteText(m_strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");
        //                        frmMain.g_oDelegate.SetListViewTextValue(
        //                            oLv,x,COL_CHECKBOX,Convert.ToString(frmMain.g_oDelegate.GetListViewTextValue(oLv,x,COL_CHECKBOX,false).Replace("c","")));

        //                    }
        //                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Green);
        //                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "ForeColor", Color.White);
        //                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "DONE");

        //                }
        //                else 
        //                {
                            
        //                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "BackColor", Color.Red);
        //                    frmMain.g_oDelegate.SetListViewSubItemPropertyValue(oLv, x, COL_RUNSTATUS, "Text", "ERROR");
        //                }
                        
        //                oAdo.CloseConnection(oAdo.m_OleDbConnection);
        //                frmMain.g_oDelegate.ExecuteListViewItemsMethod(oLv, "Refresh");

        //                UpdateTherm(this.m_frmTherm.progressBar1,
        //                            m_intProgressStepTotalCount,
        //                            m_intProgressStepTotalCount);
        //                System.Threading.Thread.Sleep(2000);

        //                m_intProgressOverallCurrentCount++;
        //                UpdateTherm(m_frmTherm.progressBar2,
        //                            m_intProgressOverallCurrentCount,
        //                            m_intProgressOverallTotalCount);

        //            }
        //        }
        //        UpdateTherm(m_frmTherm.progressBar2,
        //                    m_intProgressOverallTotalCount,
        //                    m_intProgressOverallTotalCount);

                
        //        System.Threading.Thread.Sleep(2000);
        //        this.FVSRecordsFinished();
        //    }
        //    catch (System.Threading.ThreadInterruptedException err)
        //    {
        //        MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
        //    }
        //    catch (System.Threading.ThreadAbortException err)
        //    {
        //        if (oAdo.m_OleDbConnection != null)
        //        {
        //            if (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
        //            {
        //                oAdo.CloseConnection(oAdo.m_OleDbConnection);
        //            }
        //        }
        //        this.ThreadCleanUp();
        //        this.CleanupThread();

        //    }
        //    catch (Exception err)
        //    {
        //        MessageBox.Show("!!Error!! \n" +
        //            "Module - uc_fvs_output:ConvertAlphaSpCd_Main  \n" +
        //            "Err Msg - " + err.Message.ToString().Trim(),
        //            "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
        //            System.Windows.Forms.MessageBoxIcon.Exclamation);
        //        this.m_intError = -1;
        //    }

        //    if (m_bDebug)
        //        frmMain.g_oUtils.WriteText(m_strDebugFile, "*****END*****" + System.DateTime.Now.ToString() + "\r\n");
           
        //    CleanupThread();


        //    frmMain.g_oDelegate.CurrentThreadProcessDone = true;
        //    frmMain.g_oDelegate.m_oEventThreadStopped.Set();
        //    this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        //}

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
            string strAuditDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)this.txtOutDir, "Text", false);
            strAuditDbFile = strAuditDbFile.Trim();
            string strVariant = lstFvsOutput.SelectedItems[0].SubItems[COL_VARIANT].Text.Trim();
            strAuditDbFile = strAuditDbFile + "\\" + strVariant + "\\PostAudit.accdb";

            if (System.IO.File.Exists(strAuditDbFile))
            {
                ado_data_access oAdo = new ado_data_access();
                strConn = oAdo.getMDBConnString(strAuditDbFile, "", "");
                oAdo.OpenConnection(strConn);

                if (!oAdo.TableExist(oAdo.m_OleDbConnection, "audit_Post_SUMMARY"))
                {
                    oAdo.CloseConnection(oAdo.m_OleDbConnection);
                    oAdo = null;
                    string strWarnMessage = "No POST-APPEND audit tables exist in the file " + strAuditDbFile + ". The POST-APPEND Audit tables cannot be displayed.";
                    MessageBox.Show(strWarnMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                FIA_Biosum_Manager.frmGridView oFrm = new frmGridView();



                oFrm.Text = "Database: Browse (POST-APPEND Audit Tables)";
                if (m_strFVSPostAppendAuditTables != null)
                {
                    for (int x = 0; x <= m_strFVSPostAppendAuditTables.Count - 1; x++)
                    {
                        if (oAdo.TableExist(oAdo.m_OleDbConnection, m_strFVSPostAppendAuditTables[x].Trim()))
                        {
                            oFrm.LoadDataSet(oAdo.m_OleDbConnection, strConn, "SELECT * FROM " + m_strFVSPostAppendAuditTables[x].Trim(), m_strFVSPostAppendAuditTables[x].Trim());
                        }
                    }
                }

                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oFrm.TileGridViews();
                oFrm.Show();
                oFrm.Focus();
            }
            else
            {
                MessageBox.Show("The file " + strAuditDbFile + " does not exist");
            }
        }

        private bool FvsOutWithCutList()
        {
            if (System.IO.File.Exists(m_strFvsOutDb))
            {
                string strConn = SQLite.GetConnectionString(m_strFvsOutDb);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    if (SQLite.TableExist(conn, "FVS_CUTLIST"))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void RunCreateFVSOut_BioSum_Start()
        {
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

            ado_data_access oAdo = new ado_data_access();

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


                        CreatePrePostSeqNumMatrixSqliteTables(strTreeTempDbFile, strRxPackage, false);

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
                                SQLite.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultFVSInForestTreeTableName} WHERE RXPACKAGE ='{strRxPackage}'";
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
                            Datasource oDatasource = new Datasource(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim());
                            string strTempMDBFile = oDatasource.CreateMDBAndTableDataSourceLinks();
                            strConn = oAdo.getMDBConnString(strTempMDBFile, "", "");
                            using (System.Data.OleDb.OleDbConnection accessConn = new System.Data.OleDb.OleDbConnection(strConn))
                            {
                                accessConn.Open();
                                ODBCMgr odbcMgr = new ODBCMgr();
                                // Create ODBC entry for the temporary cut list db file
                                if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName))
                                {
                                    odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName);
                                }
                                odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName, strTreeTempDbFile);
                                m_dao.CreateSQLiteTableLink(strTempMDBFile, Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                    Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                    ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName, strTreeTempDbFile, true);
                                // Sleep to ensure table link is complete
                                Thread.Sleep(5000);

                                //NOTE: this query handles existing FIADB trees that have been grown forward.
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step2(
                                                   Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                   m_oQueries.m_oFIAPlot.m_strTreeTable,
                                                   m_oQueries.m_oFIAPlot.m_strPlotTable,
                                                   m_oQueries.m_oFIAPlot.m_strCondTable);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //Set DIAHTCD for FIADB Cycle<>1 trees to their Cycle=1 DIAHTCD values
                                //Set standing_dead_code from tree table on FIADB trees
                                oAdo.m_strSQL = oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFiadb(m_oQueries.m_oFIAPlot.m_strTreeTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //Set DIAHTCD for FVS-Created trees using FIA_TREE_SPECIES_REF.WOODLAND_YN
                                string strFiaTreeSpeciesRefTableLink = Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName;
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculationDiaHtCdFvs(strFiaTreeSpeciesRefTableLink);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //Update FVSCreatedTrees precipitation=plot.precipitation
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputSQLiteTableForVolumeCalculation_Step2a(
                                                   Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                   m_oQueries.m_oFIAPlot.m_strPlotTable,
                                                   m_oQueries.m_oFIAPlot.m_strCondTable);

                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //Start cull calculations
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step3(
                                                  Tables.VolumeAndBiomass.BiosumVolumesInputTable,
                                                  m_oQueries.m_oFIAPlot.m_strCondTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                //populate treeclcd column
                                if (oAdo.TableExist(accessConn, "CULL_TOTAL_WORK_TABLE"))
                                    oAdo.SqlNonQuery(accessConn, "DROP TABLE CULL_TOTAL_WORK_TABLE");
                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.BuildInputTableForVolumeCalculation_Step4(
                                                  "cull_total_work_table",
                                                  Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                                                    "cull_total_work_table",
                                                    Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                oAdo.m_strSQL = Queries.VolumeAndBiomass.FVSOut.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                                                  "cull_total_work_table",
                                                  Tables.VolumeAndBiomass.BiosumVolumesInputTable);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                // Drop link to BiosumVolumesInputTable to avoid conflicts
                                oAdo.m_strSQL = "DROP TABLE " + Tables.VolumeAndBiomass.BiosumVolumesInputTable;
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "START: " + System.DateTime.Now.ToString() + "\r\n" + oAdo.m_strSQL + "\r\n");
                                oAdo.SqlNonQuery(accessConn, oAdo.m_strSQL);
                                if (m_bDebug && frmMain.g_intDebugLevel > 2)
                                    this.WriteText(strDebugFile, "DONE:" + System.DateTime.Now.ToString() + "\r\n\r\n");

                                // DELETE ODBC ENTRY
                                if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName))
                                {
                                    odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsOutTemporaryDsnName);
                                }
                            }  // Closing MS Access connection

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
                                        strRxPackage);
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
                                                        WHERE rxpackage = '{strRxPackage}'";
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
            oAdo.m_intError = 0;
            oAdo.m_strError = "";
            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }

        private void CreatePrePostSeqNumMatrixSqliteTables(string p_strDbFile, string p_strRxPackageId, bool p_bAudit)
        {
            int z;
            string[] strSourceTableArray = new string[0];
            string strFvsOutDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(strFvsOutDb)))
            {
                conn.Open();
                strSourceTableArray = SQLite.getTableNames(conn);
                CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, "FVS_SUMMARY", p_strRxPackageId, p_bAudit);
                CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, "FVS_CUTLIST", p_strRxPackageId, p_bAudit);
                CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, "FVS_POTFIRE", p_strRxPackageId, p_bAudit);

                for (z = 0; z <= strSourceTableArray.Length - 1; z++)
                {
                    if (strSourceTableArray[z] == null) break;

                    if (RxTools.ValidFVSTable(strSourceTableArray[z]))
                    {
                        if (strSourceTableArray[z].Trim().ToUpper() != "FVS_SUMMARY" &&
                            strSourceTableArray[z].Trim().ToUpper() != "FVS_CUTLIST" &&
                            strSourceTableArray[z].Trim().ToUpper() != "FVS_POTFIRE")
                        {

                            CreateFVSPrePostSeqNumWorkTables(conn, p_strDbFile, strSourceTableArray[z], p_strRxPackageId, p_bAudit);

                        }
                    }
                }
            }
        }

        private void CreateFVSPrePostSeqNumWorkTables(System.Data.SQLite.SQLiteConnection conn, string p_strTreeTempDbFile, string p_strSourceTableName, string p_strRxPackageId, bool p_bAudit)
        {
            if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//CreateFVSPrePostSeqNumWorkTables\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Process Table:" + p_strSourceTableName + "\r\n\r\n");
            }

            if (p_strSourceTableName.Trim().ToUpper() == "FVS_CASES") return;


            if (SQLite.TableExist(conn, p_strSourceTableName))
            {
                if (m_oFVSPrePostSeqNumItemCollection == null) m_oFVSPrePostSeqNumItemCollection = new FVSPrePostSeqNumItem_Collection();
                string strParamConn = m_ado.getMDBConnString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile, "", "");
                using (System.Data.OleDb.OleDbConnection accessConn = new System.Data.OleDb.OleDbConnection(strParamConn))
                {
                    accessConn.Open();
                    m_oRxTools.LoadFVSOutputPrePostRxCycleSeqNum(m_ado, accessConn, m_oFVSPrePostSeqNumItemCollection);
                }

                GetPrePostSeqNumConfiguration(p_strSourceTableName, p_strRxPackageId);
                m_oRxTools.CreateFVSPrePostSeqNumTables(p_strTreeTempDbFile, m_oFVSPrePostSeqNumItem, p_strSourceTableName, p_strSourceTableName, 
                    p_bAudit, m_bDebug, m_strDebugFile, p_strRxPackageId);

            }
            else
            {
                if (m_bDebug && frmMain.g_intDebugLevel > 1 && !p_bAudit)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, p_strSourceTableName + " table does not exist.\r\n\r\n");
                }
            }
        }
    }
}

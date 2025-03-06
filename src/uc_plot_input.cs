using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Threading;
using SQLite.ADO;


namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_plot_input.
	/// </summary>
	public class uc_plot_input : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.RadioButton rdoFilterNone;
		private System.Windows.Forms.RadioButton rdoFilterByMenu;
		private System.Windows.Forms.RadioButton rdoFilterByFile;
		private System.Windows.Forms.TextBox txtFilterByFile;
		private System.Windows.Forms.Button btnFilterByFileBrowse;
		private System.Windows.Forms.GroupBox grpboxFilter;
		private System.Windows.Forms.Button btnFilterHelp;
		private System.Windows.Forms.Button btnFilterPrevious;
		private System.Windows.Forms.Button btnFilterNext;
		private System.Windows.Forms.Button btnFilterCancel;
		private System.Windows.Forms.ListView lstFilterByState;
		private System.Windows.Forms.Button btnFilterByStateUnselect;
		private System.Windows.Forms.Button btnFilterByStateSelect;
		private System.Windows.Forms.GroupBox grpboxFilterByPlot;
		private System.Windows.Forms.Button btnFilterByPlotUnselect;
		private System.Windows.Forms.Button btnFilterByPlotSelect;
		private System.Windows.Forms.ListView lstFilterByPlot;
		private System.Windows.Forms.Button btnFilterByPlotHelp;
		private System.Windows.Forms.Button btnFilterByPlotPrevious;
		private System.Windows.Forms.Button btnFilterByPlotNext;
		private System.Windows.Forms.Button btnFilterByPlotCancel;
		private System.Windows.Forms.Button btnFilterByPlotFinish;
		public int m_DialogHt;
		public int m_DialogWd;
		private System.Windows.Forms.Button btnFilterByStateFinish;
		private System.Windows.Forms.Button btnFilterByStateHelp;
		private System.Windows.Forms.Button btnFilterByStatePrevious;
		private System.Windows.Forms.Button btnFilterByStateNext;
		private System.Windows.Forms.Button btnFilterByStateCancel;
        private System.Windows.Forms.Button btnFilterFinish;
		private System.Windows.Forms.GroupBox grpboxFilterByState;
		private string m_strPlotTxtInputFile;
		private string m_strCondTxtInputFile;
		private string m_strTreeTxtInputFile;
		private string m_strSiteTreeTxtInputFile;
		private string m_strPopEvalTxtInputFile;
		private string m_strLoadedPopEvalTxtInputFile="";
		private string m_strPopEstUnitTxtInputFile;
		private string m_strLoadedPopEstUnitTxtInputFile="";
		private string m_strPpsaTxtInputFile;
		private string m_strLoadedPpsaTxtInputFile="";
		private string m_strPopStratumTxtInputFile;
		private string m_strLoadedPopStratumTxtInputFile="";
		private string m_strCurrentTxtInputFile="";



		private string m_strLoadedPopEvalInputTable="";
		private string m_strLoadedPopEstUnitInputTable="";
		private string m_strLoadedPpsaInputTable="";
		private string m_strLoadedPopStratumInputTable="";
		private string m_strLoadedFiadbInputFile="";
		private string m_strCurrentFiadbInputFile="";
		private string m_strCurrentFiadbTable="";
		private string m_strCurrentBiosumTable="";
		private string m_strCurrentProcess="";
        private string m_strMasterDbFile = "";


		private string m_strTempMDBFile;
        private string m_strTempDbFile;
        private string m_strTempMDBFileConn;
        private System.Data.OleDb.OleDbConnection m_connTempMDBFile;
		private FIA_Biosum_Manager.ado_data_access m_ado;
		private System.Data.DataTable m_dtStateCounty;
		private System.Data.DataTable m_dtPlot;
		private int m_intError;
        private string m_strError;
		//private string m_strError;
		private string m_strPlotTable;
		private string m_strCondTable;
		private string m_strTreeTable;
		private string m_strSiteTreeTable;
        private string m_strSeedlingTable = frmMain.g_oTables.m_oFIAPlot.DefaultSeedlingTableName;
        private string m_strPopEvalTable = frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName;
		private string m_strPopEstUnitTable = frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName;
		private string m_strPpsaTable = frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName;
		private string m_strPopStratumTable = frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName;
        private string m_strBiosumPopStratumAdjustmentFactorsTable;
        private string m_strTreeMacroPlotBreakPointDiaTable;
        private string m_strDwmCwdTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName;
        private string m_strDwmFwdTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName;
        private string m_strDwmDuffLitterTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName;
        private string m_strDwmTransectSegmentTable = frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName;
		private string m_strSQL;

		private FIA_Biosum_Manager.frmTherm m_frmTherm;
		private string m_strPlotIdList="";
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.CheckBox chkForested;
        private System.Windows.Forms.CheckBox chkNonForested;
		private bool m_bAllCountiesSelected=true;
		private bool m_bAllPlotsSelected=true;
		private string m_strStateCountySQL;
		private string m_strStateCountyPlotSQL;
		private string m_strIDBInv;
		private string m_strLoadedFIADBEvalId="";
		private string m_strLoadedFIADBRsCd="";
		private string m_strCurrFIADBEvalId="";
		private string m_strCurrFIADBRsCd="";
		private string m_strTableType;
		private System.Windows.Forms.GroupBox groupBox8;
		private System.Windows.Forms.Button btnMDBTreeBrowse;
		private System.Windows.Forms.TextBox txtMDBTree;
		private System.Windows.Forms.GroupBox groupBox9;
		private System.Windows.Forms.Button btnMDBCondBrowse;
		private System.Windows.Forms.TextBox txtMDBCond;
		private System.Windows.Forms.GroupBox groupBox10;
		private System.Windows.Forms.Button btnMDBPlotBrowse;
		private System.Windows.Forms.TextBox txtMDBPlotTable;
		private System.Windows.Forms.TextBox txtMDBCondTable;
		private System.Windows.Forms.TextBox txtMDBTreeTable;
		private System.Windows.Forms.TextBox txtMDBPlot;
		private System.Windows.Forms.Button btnIDBInvAppend;
		private System.Windows.Forms.ListView lstIDBInv;
		private System.Windows.Forms.Button btnIDBInvHelp;
		private System.Windows.Forms.Button btnIDBInvPrevious;
		private System.Windows.Forms.Button btnIDBInvNext;
		private System.Windows.Forms.Button btnIDBInvCancel;
		private System.Windows.Forms.GroupBox grpboxIDBInv;
		private System.Threading.Thread thdProcessRecords;
		private int m_intAddedPlotRows=0;
		private int m_intAddedCondRows=0;
		private int m_intAddedTreeRows=0;
		private int m_intAddedSiteTreeRows=0;
	    private int m_intAddedDwmCwdRows=0;
	    private int m_intAddedDwmFwdRows=0;
	    private int m_intAddedDwmDuffLitterRows=0;
	    private int m_intAddedDwmTransectSegmentRows=0;
        private int m_intAddedSeedlingRows = 0;
        private System.Windows.Forms.GroupBox grpboxFIADBInv;
		private System.Windows.Forms.Button btnFIADBInvAppend;
		private System.Windows.Forms.ListView lstFIADBInv;
		private System.Windows.Forms.Button btnFIADBInvHelp;
		private System.Windows.Forms.Button btnFIADBInvPrevious;
		private System.Windows.Forms.Button btnFIADBInvNext;
		private System.Windows.Forms.Button btnFIADBInvCancel;
		private Datasource m_oDatasource=null;

		private bool m_bLoadStateCountyList=true;
		private bool m_bLoadStateCountyPlotList=true;
		private System.Windows.Forms.GroupBox grpboxMDBFiadbInput;
		private System.Windows.Forms.GroupBox groupBox15;
		private System.Windows.Forms.GroupBox groupBox16;
		private System.Windows.Forms.GroupBox groupBox17;
		private System.Windows.Forms.GroupBox groupBox18;
		private System.Windows.Forms.GroupBox groupBox19;
		private System.Windows.Forms.GroupBox groupBox20;
		private System.Windows.Forms.GroupBox groupBox21;
		private System.Windows.Forms.GroupBox groupBox22;
		private System.Windows.Forms.Button btnboxMDBFiadbInputFile;
		private System.Windows.Forms.ComboBox cmbFiadbPlotTable;
		private System.Windows.Forms.ComboBox cmbFiadbCondTable;
		private System.Windows.Forms.GroupBox groupBox23;
		private System.Windows.Forms.ComboBox cmbFiadbTreeTable;
		private System.Windows.Forms.ComboBox cmbFiadbSeedlingTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopEvalTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopEstUnitTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopStratumTable;
		private System.Windows.Forms.ComboBox cmbFiadbPpsaTable;
		private System.Windows.Forms.Button btnMDBFiadbInputFinish;
		private System.Windows.Forms.Button btnMDBFiadbInputHelp;
		private System.Windows.Forms.Button btnMDBFiadbInputPrev;
		private System.Windows.Forms.Button btnMDBFiadbInputNext;
		private System.Windows.Forms.Button btnMDBFiadbInputCancel;
		private System.Windows.Forms.TextBox txtFiadbInputFile;
		private System.Windows.Forms.GroupBox groupBox24;
		private System.Windows.Forms.ComboBox cmbFiadbSiteTreeTable;
		private System.Windows.Forms.GroupBox groupBox25;
		private System.Windows.Forms.TextBox txtMDBSiteTreeTable;
		private System.Windows.Forms.Button btnMDBSiteTreeBrowse;
        private System.Windows.Forms.TextBox txtMDBSiteTree;
        private Label label2;
        private ComboBox cmbCondPropPercent;
        private Label label1;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
        private string m_pdfFile = @"Help\DATABASE_Help.pdf";
        private bool m_bLoadSeedlings = true;
        private CheckBox chkDwmImport;


        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public SQLite.ADO.DataMgr SQLite { get; set; } = new SQLite.ADO.DataMgr();

        private Label lblFSNetwork;
        private CheckBox ckImportSeedlings;
        private GroupBox groupBox2;
        private ComboBox cmbFiadbPlotGeomTable;
        private Label label16;
    
        public FIA_Biosum_Manager.ado_data_access MSAccess { get; set; }

        public frmDialog ReferenceFormDialog { set; get; } = null;

        public uc_plot_input()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			Initialize();

			// TODO: Add any initialization after the InitializeComponent call

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_plot_input));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxMDBFiadbInput = new System.Windows.Forms.GroupBox();
            this.groupBox24 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbSiteTreeTable = new System.Windows.Forms.ComboBox();
            this.groupBox23 = new System.Windows.Forms.GroupBox();
            this.txtFiadbInputFile = new System.Windows.Forms.TextBox();
            this.btnboxMDBFiadbInputFile = new System.Windows.Forms.Button();
            this.groupBox15 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPopEstUnitTable = new System.Windows.Forms.ComboBox();
            this.groupBox16 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPpsaTable = new System.Windows.Forms.ComboBox();
            this.groupBox17 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPopStratumTable = new System.Windows.Forms.ComboBox();
            this.groupBox18 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPopEvalTable = new System.Windows.Forms.ComboBox();
            this.groupBox19 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbSeedlingTable = new System.Windows.Forms.ComboBox();
            this.groupBox20 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbTreeTable = new System.Windows.Forms.ComboBox();
            this.groupBox21 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbCondTable = new System.Windows.Forms.ComboBox();
            this.groupBox22 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPlotTable = new System.Windows.Forms.ComboBox();
            this.btnMDBFiadbInputFinish = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputHelp = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputPrev = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputNext = new System.Windows.Forms.Button();
            this.btnMDBFiadbInputCancel = new System.Windows.Forms.Button();
            this.grpboxFIADBInv = new System.Windows.Forms.GroupBox();
            this.btnFIADBInvAppend = new System.Windows.Forms.Button();
            this.lstFIADBInv = new System.Windows.Forms.ListView();
            this.btnFIADBInvHelp = new System.Windows.Forms.Button();
            this.btnFIADBInvPrevious = new System.Windows.Forms.Button();
            this.btnFIADBInvNext = new System.Windows.Forms.Button();
            this.btnFIADBInvCancel = new System.Windows.Forms.Button();
            this.grpboxIDBInv = new System.Windows.Forms.GroupBox();
            this.btnIDBInvAppend = new System.Windows.Forms.Button();
            this.lstIDBInv = new System.Windows.Forms.ListView();
            this.btnIDBInvHelp = new System.Windows.Forms.Button();
            this.btnIDBInvPrevious = new System.Windows.Forms.Button();
            this.btnIDBInvNext = new System.Windows.Forms.Button();
            this.btnIDBInvCancel = new System.Windows.Forms.Button();
            this.grpboxFilterByState = new System.Windows.Forms.GroupBox();
            this.btnFilterByStateFinish = new System.Windows.Forms.Button();
            this.btnFilterByStateUnselect = new System.Windows.Forms.Button();
            this.btnFilterByStateSelect = new System.Windows.Forms.Button();
            this.lstFilterByState = new System.Windows.Forms.ListView();
            this.btnFilterByStateHelp = new System.Windows.Forms.Button();
            this.btnFilterByStatePrevious = new System.Windows.Forms.Button();
            this.btnFilterByStateNext = new System.Windows.Forms.Button();
            this.btnFilterByStateCancel = new System.Windows.Forms.Button();
            this.grpboxFilter = new System.Windows.Forms.GroupBox();
            this.btnFilterFinish = new System.Windows.Forms.Button();
            this.btnFilterHelp = new System.Windows.Forms.Button();
            this.btnFilterPrevious = new System.Windows.Forms.Button();
            this.btnFilterNext = new System.Windows.Forms.Button();
            this.btnFilterCancel = new System.Windows.Forms.Button();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.ckImportSeedlings = new System.Windows.Forms.CheckBox();
            this.chkDwmImport = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbCondPropPercent = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.rdoFilterNone = new System.Windows.Forms.RadioButton();
            this.rdoFilterByMenu = new System.Windows.Forms.RadioButton();
            this.chkNonForested = new System.Windows.Forms.CheckBox();
            this.chkForested = new System.Windows.Forms.CheckBox();
            this.btnFilterByFileBrowse = new System.Windows.Forms.Button();
            this.txtFilterByFile = new System.Windows.Forms.TextBox();
            this.rdoFilterByFile = new System.Windows.Forms.RadioButton();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpboxFilterByPlot = new System.Windows.Forms.GroupBox();
            this.btnFilterByPlotFinish = new System.Windows.Forms.Button();
            this.btnFilterByPlotUnselect = new System.Windows.Forms.Button();
            this.btnFilterByPlotSelect = new System.Windows.Forms.Button();
            this.lstFilterByPlot = new System.Windows.Forms.ListView();
            this.btnFilterByPlotHelp = new System.Windows.Forms.Button();
            this.btnFilterByPlotPrevious = new System.Windows.Forms.Button();
            this.btnFilterByPlotNext = new System.Windows.Forms.Button();
            this.btnFilterByPlotCancel = new System.Windows.Forms.Button();
            this.groupBox25 = new System.Windows.Forms.GroupBox();
            this.txtMDBSiteTreeTable = new System.Windows.Forms.TextBox();
            this.btnMDBSiteTreeBrowse = new System.Windows.Forms.Button();
            this.txtMDBSiteTree = new System.Windows.Forms.TextBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.txtMDBTreeTable = new System.Windows.Forms.TextBox();
            this.btnMDBTreeBrowse = new System.Windows.Forms.Button();
            this.txtMDBTree = new System.Windows.Forms.TextBox();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.txtMDBCondTable = new System.Windows.Forms.TextBox();
            this.btnMDBCondBrowse = new System.Windows.Forms.Button();
            this.txtMDBCond = new System.Windows.Forms.TextBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.txtMDBPlotTable = new System.Windows.Forms.TextBox();
            this.btnMDBPlotBrowse = new System.Windows.Forms.Button();
            this.txtMDBPlot = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPlotGeomTable = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.grpboxMDBFiadbInput.SuspendLayout();
            this.groupBox24.SuspendLayout();
            this.groupBox23.SuspendLayout();
            this.groupBox15.SuspendLayout();
            this.groupBox16.SuspendLayout();
            this.groupBox17.SuspendLayout();
            this.groupBox18.SuspendLayout();
            this.groupBox19.SuspendLayout();
            this.groupBox20.SuspendLayout();
            this.groupBox21.SuspendLayout();
            this.groupBox22.SuspendLayout();
            this.grpboxFIADBInv.SuspendLayout();
            this.grpboxIDBInv.SuspendLayout();
            this.grpboxFilterByState.SuspendLayout();
            this.grpboxFilter.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.grpboxFilterByPlot.SuspendLayout();
            this.groupBox25.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpboxMDBFiadbInput);
            this.groupBox1.Controls.Add(this.grpboxFIADBInv);
            this.groupBox1.Controls.Add(this.grpboxIDBInv);
            this.groupBox1.Controls.Add(this.grpboxFilterByState);
            this.groupBox1.Controls.Add(this.grpboxFilter);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.grpboxFilterByPlot);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(880, 3375);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // grpboxMDBFiadbInput
            // 
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox2);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox24);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox23);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox15);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox16);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox17);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox18);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox19);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox20);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox21);
            this.grpboxMDBFiadbInput.Controls.Add(this.groupBox22);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputFinish);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputHelp);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputPrev);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputNext);
            this.grpboxMDBFiadbInput.Controls.Add(this.btnMDBFiadbInputCancel);
            this.grpboxMDBFiadbInput.Location = new System.Drawing.Point(19, 70);
            this.grpboxMDBFiadbInput.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxMDBFiadbInput.Name = "grpboxMDBFiadbInput";
            this.grpboxMDBFiadbInput.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxMDBFiadbInput.Size = new System.Drawing.Size(1344, 450);
            this.grpboxMDBFiadbInput.TabIndex = 35;
            this.grpboxMDBFiadbInput.TabStop = false;
            this.grpboxMDBFiadbInput.Text = "FIADB SQLite Database File Input";
            // 
            // groupBox24
            // 
            this.groupBox24.Controls.Add(this.cmbFiadbSiteTreeTable);
            this.groupBox24.Location = new System.Drawing.Point(425, 340);
            this.groupBox24.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox24.Name = "groupBox24";
            this.groupBox24.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox24.Size = new System.Drawing.Size(390, 56);
            this.groupBox24.TabIndex = 51;
            this.groupBox24.TabStop = false;
            this.groupBox24.Text = "Site Tree ";
            // 
            // cmbFiadbSiteTreeTable
            // 
            this.cmbFiadbSiteTreeTable.Location = new System.Drawing.Point(10, 20);
            this.cmbFiadbSiteTreeTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbSiteTreeTable.Name = "cmbFiadbSiteTreeTable";
            this.cmbFiadbSiteTreeTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbSiteTreeTable.TabIndex = 3;
            // 
            // groupBox23
            // 
            this.groupBox23.Controls.Add(this.txtFiadbInputFile);
            this.groupBox23.Controls.Add(this.btnboxMDBFiadbInputFile);
            this.groupBox23.Location = new System.Drawing.Point(25, 20);
            this.groupBox23.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox23.Name = "groupBox23";
            this.groupBox23.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox23.Size = new System.Drawing.Size(790, 60);
            this.groupBox23.TabIndex = 50;
            this.groupBox23.TabStop = false;
            this.groupBox23.Text = "SQLite File";
            // 
            // txtFiadbInputFile
            // 
            this.txtFiadbInputFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFiadbInputFile.Location = new System.Drawing.Point(10, 16);
            this.txtFiadbInputFile.Margin = new System.Windows.Forms.Padding(4);
            this.txtFiadbInputFile.Name = "txtFiadbInputFile";
            this.txtFiadbInputFile.Size = new System.Drawing.Size(719, 30);
            this.txtFiadbInputFile.TabIndex = 2;
            // 
            // btnboxMDBFiadbInputFile
            // 
            this.btnboxMDBFiadbInputFile.Image = ((System.Drawing.Image)(resources.GetObject("btnboxMDBFiadbInputFile.Image")));
            this.btnboxMDBFiadbInputFile.Location = new System.Drawing.Point(740, 12);
            this.btnboxMDBFiadbInputFile.Margin = new System.Windows.Forms.Padding(4);
            this.btnboxMDBFiadbInputFile.Name = "btnboxMDBFiadbInputFile";
            this.btnboxMDBFiadbInputFile.Size = new System.Drawing.Size(40, 40);
            this.btnboxMDBFiadbInputFile.TabIndex = 1;
            this.btnboxMDBFiadbInputFile.Click += new System.EventHandler(this.btnboxMDBFiadbInputFile_Click);
            // 
            // groupBox15
            // 
            this.groupBox15.Controls.Add(this.cmbFiadbPopEstUnitTable);
            this.groupBox15.Location = new System.Drawing.Point(425, 148);
            this.groupBox15.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox15.Name = "groupBox15";
            this.groupBox15.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox15.Size = new System.Drawing.Size(390, 56);
            this.groupBox15.TabIndex = 49;
            this.groupBox15.TabStop = false;
            this.groupBox15.Text = "Population Estimation Unit";
            // 
            // cmbFiadbPopEstUnitTable
            // 
            this.cmbFiadbPopEstUnitTable.Location = new System.Drawing.Point(10, 21);
            this.cmbFiadbPopEstUnitTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbPopEstUnitTable.Name = "cmbFiadbPopEstUnitTable";
            this.cmbFiadbPopEstUnitTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbPopEstUnitTable.TabIndex = 4;
            // 
            // groupBox16
            // 
            this.groupBox16.Controls.Add(this.cmbFiadbPpsaTable);
            this.groupBox16.Location = new System.Drawing.Point(425, 275);
            this.groupBox16.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox16.Name = "groupBox16";
            this.groupBox16.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox16.Size = new System.Drawing.Size(390, 56);
            this.groupBox16.TabIndex = 48;
            this.groupBox16.TabStop = false;
            this.groupBox16.Text = "Population Plot Stratum Assignment";
            // 
            // cmbFiadbPpsaTable
            // 
            this.cmbFiadbPpsaTable.Location = new System.Drawing.Point(10, 20);
            this.cmbFiadbPpsaTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbPpsaTable.Name = "cmbFiadbPpsaTable";
            this.cmbFiadbPpsaTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbPpsaTable.TabIndex = 4;
            // 
            // groupBox17
            // 
            this.groupBox17.Controls.Add(this.cmbFiadbPopStratumTable);
            this.groupBox17.Location = new System.Drawing.Point(425, 211);
            this.groupBox17.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox17.Name = "groupBox17";
            this.groupBox17.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox17.Size = new System.Drawing.Size(390, 56);
            this.groupBox17.TabIndex = 47;
            this.groupBox17.TabStop = false;
            this.groupBox17.Text = "Population Stratum";
            // 
            // cmbFiadbPopStratumTable
            // 
            this.cmbFiadbPopStratumTable.Location = new System.Drawing.Point(10, 21);
            this.cmbFiadbPopStratumTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbPopStratumTable.Name = "cmbFiadbPopStratumTable";
            this.cmbFiadbPopStratumTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbPopStratumTable.TabIndex = 4;
            // 
            // groupBox18
            // 
            this.groupBox18.Controls.Add(this.cmbFiadbPopEvalTable);
            this.groupBox18.Location = new System.Drawing.Point(425, 89);
            this.groupBox18.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox18.Name = "groupBox18";
            this.groupBox18.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox18.Size = new System.Drawing.Size(390, 56);
            this.groupBox18.TabIndex = 46;
            this.groupBox18.TabStop = false;
            this.groupBox18.Text = "Population Evaluation";
            // 
            // cmbFiadbPopEvalTable
            // 
            this.cmbFiadbPopEvalTable.Location = new System.Drawing.Point(10, 20);
            this.cmbFiadbPopEvalTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbPopEvalTable.Name = "cmbFiadbPopEvalTable";
            this.cmbFiadbPopEvalTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbPopEvalTable.TabIndex = 4;
            // 
            // groupBox19
            // 
            this.groupBox19.Controls.Add(this.cmbFiadbSeedlingTable);
            this.groupBox19.Location = new System.Drawing.Point(24, 340);
            this.groupBox19.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox19.Name = "groupBox19";
            this.groupBox19.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox19.Size = new System.Drawing.Size(390, 56);
            this.groupBox19.TabIndex = 45;
            this.groupBox19.TabStop = false;
            this.groupBox19.Text = "Seedling Data";
            // 
            // cmbFiadbSeedlingTable
            // 
            this.cmbFiadbSeedlingTable.Location = new System.Drawing.Point(10, 20);
            this.cmbFiadbSeedlingTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbSeedlingTable.Name = "cmbFiadbSeedlingTable";
            this.cmbFiadbSeedlingTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbSeedlingTable.TabIndex = 3;
            // 
            // groupBox20
            // 
            this.groupBox20.Controls.Add(this.cmbFiadbTreeTable);
            this.groupBox20.Location = new System.Drawing.Point(25, 275);
            this.groupBox20.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox20.Name = "groupBox20";
            this.groupBox20.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox20.Size = new System.Drawing.Size(390, 56);
            this.groupBox20.TabIndex = 44;
            this.groupBox20.TabStop = false;
            this.groupBox20.Text = "Tree Data";
            // 
            // cmbFiadbTreeTable
            // 
            this.cmbFiadbTreeTable.Location = new System.Drawing.Point(10, 18);
            this.cmbFiadbTreeTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbTreeTable.Name = "cmbFiadbTreeTable";
            this.cmbFiadbTreeTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbTreeTable.TabIndex = 2;
            // 
            // groupBox21
            // 
            this.groupBox21.Controls.Add(this.cmbFiadbCondTable);
            this.groupBox21.Location = new System.Drawing.Point(25, 211);
            this.groupBox21.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox21.Name = "groupBox21";
            this.groupBox21.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox21.Size = new System.Drawing.Size(390, 56);
            this.groupBox21.TabIndex = 43;
            this.groupBox21.TabStop = false;
            this.groupBox21.Text = "Condition Data";
            // 
            // cmbFiadbCondTable
            // 
            this.cmbFiadbCondTable.Location = new System.Drawing.Point(10, 20);
            this.cmbFiadbCondTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbCondTable.Name = "cmbFiadbCondTable";
            this.cmbFiadbCondTable.Size = new System.Drawing.Size(369, 24);
            this.cmbFiadbCondTable.TabIndex = 1;
            // 
            // groupBox22
            // 
            this.groupBox22.Controls.Add(this.cmbFiadbPlotTable);
            this.groupBox22.Location = new System.Drawing.Point(25, 89);
            this.groupBox22.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox22.Name = "groupBox22";
            this.groupBox22.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox22.Size = new System.Drawing.Size(390, 56);
            this.groupBox22.TabIndex = 42;
            this.groupBox22.TabStop = false;
            this.groupBox22.Text = "Plot Data";
            // 
            // cmbFiadbPlotTable
            // 
            this.cmbFiadbPlotTable.Location = new System.Drawing.Point(9, 20);
            this.cmbFiadbPlotTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbPlotTable.Name = "cmbFiadbPlotTable";
            this.cmbFiadbPlotTable.Size = new System.Drawing.Size(370, 24);
            this.cmbFiadbPlotTable.TabIndex = 0;
            // 
            // btnMDBFiadbInputFinish
            // 
            this.btnMDBFiadbInputFinish.Enabled = false;
            this.btnMDBFiadbInputFinish.Location = new System.Drawing.Point(730, 408);
            this.btnMDBFiadbInputFinish.Margin = new System.Windows.Forms.Padding(4);
            this.btnMDBFiadbInputFinish.Name = "btnMDBFiadbInputFinish";
            this.btnMDBFiadbInputFinish.Size = new System.Drawing.Size(90, 30);
            this.btnMDBFiadbInputFinish.TabIndex = 7;
            this.btnMDBFiadbInputFinish.Text = "Append";
            // 
            // btnMDBFiadbInputHelp
            // 
            this.btnMDBFiadbInputHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnMDBFiadbInputHelp.Location = new System.Drawing.Point(30, 408);
            this.btnMDBFiadbInputHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnMDBFiadbInputHelp.Name = "btnMDBFiadbInputHelp";
            this.btnMDBFiadbInputHelp.Size = new System.Drawing.Size(80, 30);
            this.btnMDBFiadbInputHelp.TabIndex = 3;
            this.btnMDBFiadbInputHelp.Text = "Help";
            this.btnMDBFiadbInputHelp.Click += new System.EventHandler(this.btnMDBFiadbInputHelp_Click);
            // 
            // btnMDBFiadbInputPrev
            // 
            this.btnMDBFiadbInputPrev.Enabled = false;
            this.btnMDBFiadbInputPrev.Location = new System.Drawing.Point(530, 408);
            this.btnMDBFiadbInputPrev.Margin = new System.Windows.Forms.Padding(4);
            this.btnMDBFiadbInputPrev.Name = "btnMDBFiadbInputPrev";
            this.btnMDBFiadbInputPrev.Size = new System.Drawing.Size(90, 30);
            this.btnMDBFiadbInputPrev.TabIndex = 5;
            this.btnMDBFiadbInputPrev.TabStop = false;
            this.btnMDBFiadbInputPrev.Text = "< Previous";
            // 
            // btnMDBFiadbInputNext
            // 
            this.btnMDBFiadbInputNext.Location = new System.Drawing.Point(620, 408);
            this.btnMDBFiadbInputNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnMDBFiadbInputNext.Name = "btnMDBFiadbInputNext";
            this.btnMDBFiadbInputNext.Size = new System.Drawing.Size(90, 30);
            this.btnMDBFiadbInputNext.TabIndex = 6;
            this.btnMDBFiadbInputNext.Text = "Next >";
            this.btnMDBFiadbInputNext.Click += new System.EventHandler(this.btnMDBFiadbInputNext_Click);
            // 
            // btnMDBFiadbInputCancel
            // 
            this.btnMDBFiadbInputCancel.Location = new System.Drawing.Point(420, 408);
            this.btnMDBFiadbInputCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnMDBFiadbInputCancel.Name = "btnMDBFiadbInputCancel";
            this.btnMDBFiadbInputCancel.Size = new System.Drawing.Size(80, 30);
            this.btnMDBFiadbInputCancel.TabIndex = 4;
            this.btnMDBFiadbInputCancel.Text = "Cancel";
            this.btnMDBFiadbInputCancel.Click += new System.EventHandler(this.btnMDBFiadbInputCancel_Click);
            // 
            // grpboxFIADBInv
            // 
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvAppend);
            this.grpboxFIADBInv.Controls.Add(this.lstFIADBInv);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvHelp);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvPrevious);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvNext);
            this.grpboxFIADBInv.Controls.Add(this.btnFIADBInvCancel);
            this.grpboxFIADBInv.Location = new System.Drawing.Point(20, 2855);
            this.grpboxFIADBInv.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxFIADBInv.Name = "grpboxFIADBInv";
            this.grpboxFIADBInv.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxFIADBInv.Size = new System.Drawing.Size(840, 450);
            this.grpboxFIADBInv.TabIndex = 34;
            this.grpboxFIADBInv.TabStop = false;
            this.grpboxFIADBInv.Text = "Select FIADB Inventory Evalulation";
            // 
            // btnFIADBInvAppend
            // 
            this.btnFIADBInvAppend.Enabled = false;
            this.btnFIADBInvAppend.Location = new System.Drawing.Point(730, 406);
            this.btnFIADBInvAppend.Margin = new System.Windows.Forms.Padding(4);
            this.btnFIADBInvAppend.Name = "btnFIADBInvAppend";
            this.btnFIADBInvAppend.Size = new System.Drawing.Size(90, 30);
            this.btnFIADBInvAppend.TabIndex = 34;
            this.btnFIADBInvAppend.Text = "Append";
            this.btnFIADBInvAppend.Click += new System.EventHandler(this.btnFIADBInvAppend_Click);
            // 
            // lstFIADBInv
            // 
            this.lstFIADBInv.FullRowSelect = true;
            this.lstFIADBInv.GridLines = true;
            this.lstFIADBInv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFIADBInv.HideSelection = false;
            this.lstFIADBInv.Location = new System.Drawing.Point(20, 40);
            this.lstFIADBInv.Margin = new System.Windows.Forms.Padding(4);
            this.lstFIADBInv.MultiSelect = false;
            this.lstFIADBInv.Name = "lstFIADBInv";
            this.lstFIADBInv.Size = new System.Drawing.Size(1312, 349);
            this.lstFIADBInv.TabIndex = 30;
            this.lstFIADBInv.UseCompatibleStateImageBehavior = false;
            this.lstFIADBInv.View = System.Windows.Forms.View.Details;
            this.lstFIADBInv.SelectedIndexChanged += new System.EventHandler(this.lstFIADBInv_SelectedIndexChanged);
            // 
            // btnFIADBInvHelp
            // 
            this.btnFIADBInvHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFIADBInvHelp.Location = new System.Drawing.Point(20, 406);
            this.btnFIADBInvHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnFIADBInvHelp.Name = "btnFIADBInvHelp";
            this.btnFIADBInvHelp.Size = new System.Drawing.Size(80, 30);
            this.btnFIADBInvHelp.TabIndex = 23;
            this.btnFIADBInvHelp.Text = "Help";
            // 
            // btnFIADBInvPrevious
            // 
            this.btnFIADBInvPrevious.Location = new System.Drawing.Point(530, 406);
            this.btnFIADBInvPrevious.Margin = new System.Windows.Forms.Padding(4);
            this.btnFIADBInvPrevious.Name = "btnFIADBInvPrevious";
            this.btnFIADBInvPrevious.Size = new System.Drawing.Size(90, 30);
            this.btnFIADBInvPrevious.TabIndex = 22;
            this.btnFIADBInvPrevious.Text = "< Previous";
            this.btnFIADBInvPrevious.Click += new System.EventHandler(this.btnFIADBInvPrevious_Click);
            // 
            // btnFIADBInvNext
            // 
            this.btnFIADBInvNext.Location = new System.Drawing.Point(620, 406);
            this.btnFIADBInvNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnFIADBInvNext.Name = "btnFIADBInvNext";
            this.btnFIADBInvNext.Size = new System.Drawing.Size(90, 30);
            this.btnFIADBInvNext.TabIndex = 21;
            this.btnFIADBInvNext.Text = "Next >";
            this.btnFIADBInvNext.Click += new System.EventHandler(this.btnFIADBInvNext_Click);
            // 
            // btnFIADBInvCancel
            // 
            this.btnFIADBInvCancel.Location = new System.Drawing.Point(420, 406);
            this.btnFIADBInvCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnFIADBInvCancel.Name = "btnFIADBInvCancel";
            this.btnFIADBInvCancel.Size = new System.Drawing.Size(80, 30);
            this.btnFIADBInvCancel.TabIndex = 20;
            this.btnFIADBInvCancel.Text = "Cancel";
            this.btnFIADBInvCancel.Click += new System.EventHandler(this.btnFIADBInvCancel_Click);
            // 
            // grpboxIDBInv
            // 
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvAppend);
            this.grpboxIDBInv.Controls.Add(this.lstIDBInv);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvHelp);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvPrevious);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvNext);
            this.grpboxIDBInv.Controls.Add(this.btnIDBInvCancel);
            this.grpboxIDBInv.Location = new System.Drawing.Point(20, 2395);
            this.grpboxIDBInv.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxIDBInv.Name = "grpboxIDBInv";
            this.grpboxIDBInv.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxIDBInv.Size = new System.Drawing.Size(840, 450);
            this.grpboxIDBInv.TabIndex = 33;
            this.grpboxIDBInv.TabStop = false;
            this.grpboxIDBInv.Text = "Select IDB Inventory";
            // 
            // btnIDBInvAppend
            // 
            this.btnIDBInvAppend.Enabled = false;
            this.btnIDBInvAppend.Location = new System.Drawing.Point(730, 408);
            this.btnIDBInvAppend.Margin = new System.Windows.Forms.Padding(4);
            this.btnIDBInvAppend.Name = "btnIDBInvAppend";
            this.btnIDBInvAppend.Size = new System.Drawing.Size(90, 30);
            this.btnIDBInvAppend.TabIndex = 34;
            this.btnIDBInvAppend.Text = "Append";
            this.btnIDBInvAppend.Click += new System.EventHandler(this.btnIDBInvAppend_Click);
            // 
            // lstIDBInv
            // 
            this.lstIDBInv.FullRowSelect = true;
            this.lstIDBInv.GridLines = true;
            this.lstIDBInv.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstIDBInv.HideSelection = false;
            this.lstIDBInv.Location = new System.Drawing.Point(20, 40);
            this.lstIDBInv.Margin = new System.Windows.Forms.Padding(4);
            this.lstIDBInv.MultiSelect = false;
            this.lstIDBInv.Name = "lstIDBInv";
            this.lstIDBInv.Size = new System.Drawing.Size(799, 349);
            this.lstIDBInv.TabIndex = 30;
            this.lstIDBInv.UseCompatibleStateImageBehavior = false;
            this.lstIDBInv.View = System.Windows.Forms.View.Details;
            this.lstIDBInv.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstIDBInv_ItemCheck);
            // 
            // btnIDBInvHelp
            // 
            this.btnIDBInvHelp.Location = new System.Drawing.Point(20, 408);
            this.btnIDBInvHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnIDBInvHelp.Name = "btnIDBInvHelp";
            this.btnIDBInvHelp.Size = new System.Drawing.Size(80, 30);
            this.btnIDBInvHelp.TabIndex = 23;
            this.btnIDBInvHelp.Text = "Help";
            // 
            // btnIDBInvPrevious
            // 
            this.btnIDBInvPrevious.Location = new System.Drawing.Point(530, 408);
            this.btnIDBInvPrevious.Margin = new System.Windows.Forms.Padding(4);
            this.btnIDBInvPrevious.Name = "btnIDBInvPrevious";
            this.btnIDBInvPrevious.Size = new System.Drawing.Size(90, 30);
            this.btnIDBInvPrevious.TabIndex = 22;
            this.btnIDBInvPrevious.Text = "< Previous";
            this.btnIDBInvPrevious.Click += new System.EventHandler(this.btnIDBInvPrevious_Click);
            // 
            // btnIDBInvNext
            // 
            this.btnIDBInvNext.Location = new System.Drawing.Point(620, 408);
            this.btnIDBInvNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnIDBInvNext.Name = "btnIDBInvNext";
            this.btnIDBInvNext.Size = new System.Drawing.Size(90, 30);
            this.btnIDBInvNext.TabIndex = 21;
            this.btnIDBInvNext.Text = "Next >";
            this.btnIDBInvNext.Click += new System.EventHandler(this.btnIDBInvNext_Click);
            // 
            // btnIDBInvCancel
            // 
            this.btnIDBInvCancel.Location = new System.Drawing.Point(420, 408);
            this.btnIDBInvCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnIDBInvCancel.Name = "btnIDBInvCancel";
            this.btnIDBInvCancel.Size = new System.Drawing.Size(80, 30);
            this.btnIDBInvCancel.TabIndex = 20;
            this.btnIDBInvCancel.Text = "Cancel";
            this.btnIDBInvCancel.Click += new System.EventHandler(this.btnIDBInvCancel_Click);
            // 
            // grpboxFilterByState
            // 
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateFinish);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateUnselect);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateSelect);
            this.grpboxFilterByState.Controls.Add(this.lstFilterByState);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateHelp);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStatePrevious);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateNext);
            this.grpboxFilterByState.Controls.Add(this.btnFilterByStateCancel);
            this.grpboxFilterByState.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.grpboxFilterByState.Location = new System.Drawing.Point(20, 995);
            this.grpboxFilterByState.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxFilterByState.Name = "grpboxFilterByState";
            this.grpboxFilterByState.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxFilterByState.Size = new System.Drawing.Size(840, 450);
            this.grpboxFilterByState.TabIndex = 31;
            this.grpboxFilterByState.TabStop = false;
            this.grpboxFilterByState.Text = "Filter By State And County";
            // 
            // btnFilterByStateFinish
            // 
            this.btnFilterByStateFinish.Location = new System.Drawing.Point(730, 406);
            this.btnFilterByStateFinish.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStateFinish.Name = "btnFilterByStateFinish";
            this.btnFilterByStateFinish.Size = new System.Drawing.Size(90, 30);
            this.btnFilterByStateFinish.TabIndex = 34;
            this.btnFilterByStateFinish.Text = "Append";
            this.btnFilterByStateFinish.Click += new System.EventHandler(this.btnFilterByStateFinish_Click);
            // 
            // btnFilterByStateUnselect
            // 
            this.btnFilterByStateUnselect.Location = new System.Drawing.Point(700, 228);
            this.btnFilterByStateUnselect.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStateUnselect.Name = "btnFilterByStateUnselect";
            this.btnFilterByStateUnselect.Size = new System.Drawing.Size(110, 60);
            this.btnFilterByStateUnselect.TabIndex = 32;
            this.btnFilterByStateUnselect.Text = "Clear All";
            this.btnFilterByStateUnselect.Click += new System.EventHandler(this.btnFilterByStateUnselect_Click);
            // 
            // btnFilterByStateSelect
            // 
            this.btnFilterByStateSelect.Location = new System.Drawing.Point(700, 148);
            this.btnFilterByStateSelect.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStateSelect.Name = "btnFilterByStateSelect";
            this.btnFilterByStateSelect.Size = new System.Drawing.Size(110, 60);
            this.btnFilterByStateSelect.TabIndex = 31;
            this.btnFilterByStateSelect.Text = "Select All";
            this.btnFilterByStateSelect.Click += new System.EventHandler(this.btnFilterByStateSelect_Click);
            // 
            // lstFilterByState
            // 
            this.lstFilterByState.CheckBoxes = true;
            this.lstFilterByState.FullRowSelect = true;
            this.lstFilterByState.GridLines = true;
            this.lstFilterByState.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFilterByState.HideSelection = false;
            this.lstFilterByState.Location = new System.Drawing.Point(170, 40);
            this.lstFilterByState.Margin = new System.Windows.Forms.Padding(4);
            this.lstFilterByState.Name = "lstFilterByState";
            this.lstFilterByState.Size = new System.Drawing.Size(499, 349);
            this.lstFilterByState.TabIndex = 30;
            this.lstFilterByState.UseCompatibleStateImageBehavior = false;
            this.lstFilterByState.View = System.Windows.Forms.View.Details;
            this.lstFilterByState.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstFilterByState_ItemCheck);
            // 
            // btnFilterByStateHelp
            // 
            this.btnFilterByStateHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterByStateHelp.Location = new System.Drawing.Point(20, 406);
            this.btnFilterByStateHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStateHelp.Name = "btnFilterByStateHelp";
            this.btnFilterByStateHelp.Size = new System.Drawing.Size(80, 30);
            this.btnFilterByStateHelp.TabIndex = 23;
            this.btnFilterByStateHelp.Text = "Help";
            this.btnFilterByStateHelp.Click += new System.EventHandler(this.btnFilterByStateHelp_Click);
            // 
            // btnFilterByStatePrevious
            // 
            this.btnFilterByStatePrevious.Location = new System.Drawing.Point(530, 406);
            this.btnFilterByStatePrevious.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStatePrevious.Name = "btnFilterByStatePrevious";
            this.btnFilterByStatePrevious.Size = new System.Drawing.Size(90, 30);
            this.btnFilterByStatePrevious.TabIndex = 22;
            this.btnFilterByStatePrevious.Text = "< Previous";
            this.btnFilterByStatePrevious.Click += new System.EventHandler(this.btnFilterByStatePrevious_Click);
            // 
            // btnFilterByStateNext
            // 
            this.btnFilterByStateNext.Location = new System.Drawing.Point(620, 406);
            this.btnFilterByStateNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStateNext.Name = "btnFilterByStateNext";
            this.btnFilterByStateNext.Size = new System.Drawing.Size(90, 30);
            this.btnFilterByStateNext.TabIndex = 21;
            this.btnFilterByStateNext.Text = "Next >";
            this.btnFilterByStateNext.Click += new System.EventHandler(this.btnFilterByStateNext_Click);
            // 
            // btnFilterByStateCancel
            // 
            this.btnFilterByStateCancel.Location = new System.Drawing.Point(420, 406);
            this.btnFilterByStateCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByStateCancel.Name = "btnFilterByStateCancel";
            this.btnFilterByStateCancel.Size = new System.Drawing.Size(80, 30);
            this.btnFilterByStateCancel.TabIndex = 20;
            this.btnFilterByStateCancel.Text = "Cancel";
            this.btnFilterByStateCancel.Click += new System.EventHandler(this.btnFilterByStateCancel_Click);
            // 
            // grpboxFilter
            // 
            this.grpboxFilter.Controls.Add(this.btnFilterFinish);
            this.grpboxFilter.Controls.Add(this.btnFilterHelp);
            this.grpboxFilter.Controls.Add(this.btnFilterPrevious);
            this.grpboxFilter.Controls.Add(this.btnFilterNext);
            this.grpboxFilter.Controls.Add(this.btnFilterCancel);
            this.grpboxFilter.Controls.Add(this.groupBox7);
            this.grpboxFilter.Location = new System.Drawing.Point(20, 530);
            this.grpboxFilter.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxFilter.Name = "grpboxFilter";
            this.grpboxFilter.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxFilter.Size = new System.Drawing.Size(840, 450);
            this.grpboxFilter.TabIndex = 30;
            this.grpboxFilter.TabStop = false;
            this.grpboxFilter.Text = "Filter Options";
            // 
            // btnFilterFinish
            // 
            this.btnFilterFinish.Enabled = false;
            this.btnFilterFinish.Location = new System.Drawing.Point(730, 408);
            this.btnFilterFinish.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterFinish.Name = "btnFilterFinish";
            this.btnFilterFinish.Size = new System.Drawing.Size(90, 30);
            this.btnFilterFinish.TabIndex = 5;
            this.btnFilterFinish.Text = "Append";
            this.btnFilterFinish.Click += new System.EventHandler(this.btnFilterFinish_Click);
            // 
            // btnFilterHelp
            // 
            this.btnFilterHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterHelp.Location = new System.Drawing.Point(20, 408);
            this.btnFilterHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterHelp.Name = "btnFilterHelp";
            this.btnFilterHelp.Size = new System.Drawing.Size(80, 30);
            this.btnFilterHelp.TabIndex = 2;
            this.btnFilterHelp.Text = "Help";
            this.btnFilterHelp.Click += new System.EventHandler(this.btnFilterHelp_Click);
            // 
            // btnFilterPrevious
            // 
            this.btnFilterPrevious.Location = new System.Drawing.Point(530, 408);
            this.btnFilterPrevious.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterPrevious.Name = "btnFilterPrevious";
            this.btnFilterPrevious.Size = new System.Drawing.Size(90, 30);
            this.btnFilterPrevious.TabIndex = 4;
            this.btnFilterPrevious.Text = "< Previous";
            this.btnFilterPrevious.Click += new System.EventHandler(this.btnFilterPrevious_Click);
            // 
            // btnFilterNext
            // 
            this.btnFilterNext.Enabled = false;
            this.btnFilterNext.Location = new System.Drawing.Point(620, 408);
            this.btnFilterNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterNext.Name = "btnFilterNext";
            this.btnFilterNext.Size = new System.Drawing.Size(90, 30);
            this.btnFilterNext.TabIndex = 5;
            this.btnFilterNext.Text = "Next >";
            this.btnFilterNext.Click += new System.EventHandler(this.btnFilterNext_Click);
            // 
            // btnFilterCancel
            // 
            this.btnFilterCancel.Location = new System.Drawing.Point(420, 408);
            this.btnFilterCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterCancel.Name = "btnFilterCancel";
            this.btnFilterCancel.Size = new System.Drawing.Size(80, 30);
            this.btnFilterCancel.TabIndex = 3;
            this.btnFilterCancel.Text = "Cancel";
            this.btnFilterCancel.Click += new System.EventHandler(this.btnFilterCancel_Click);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.ckImportSeedlings);
            this.groupBox7.Controls.Add(this.chkDwmImport);
            this.groupBox7.Controls.Add(this.label2);
            this.groupBox7.Controls.Add(this.cmbCondPropPercent);
            this.groupBox7.Controls.Add(this.label1);
            this.groupBox7.Controls.Add(this.rdoFilterNone);
            this.groupBox7.Controls.Add(this.rdoFilterByMenu);
            this.groupBox7.Controls.Add(this.chkNonForested);
            this.groupBox7.Controls.Add(this.chkForested);
            this.groupBox7.Controls.Add(this.btnFilterByFileBrowse);
            this.groupBox7.Controls.Add(this.txtFilterByFile);
            this.groupBox7.Controls.Add(this.rdoFilterByFile);
            this.groupBox7.Location = new System.Drawing.Point(106, 24);
            this.groupBox7.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox7.Size = new System.Drawing.Size(649, 361);
            this.groupBox7.TabIndex = 1;
            this.groupBox7.TabStop = false;
            // 
            // ckImportSeedlings
            // 
            this.ckImportSeedlings.AutoSize = true;
            this.ckImportSeedlings.Checked = true;
            this.ckImportSeedlings.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ckImportSeedlings.Location = new System.Drawing.Point(76, 181);
            this.ckImportSeedlings.Margin = new System.Windows.Forms.Padding(4);
            this.ckImportSeedlings.Name = "ckImportSeedlings";
            this.ckImportSeedlings.Size = new System.Drawing.Size(219, 21);
            this.ckImportSeedlings.TabIndex = 13;
            this.ckImportSeedlings.Text = "Load seedling data to BioSum";
            this.ckImportSeedlings.UseVisualStyleBackColor = true;
            // 
            // chkDwmImport
            // 
            this.chkDwmImport.AutoSize = true;
            this.chkDwmImport.Checked = true;
            this.chkDwmImport.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwmImport.Location = new System.Drawing.Point(76, 208);
            this.chkDwmImport.Margin = new System.Windows.Forms.Padding(4);
            this.chkDwmImport.Name = "chkDwmImport";
            this.chkDwmImport.Size = new System.Drawing.Size(290, 21);
            this.chkDwmImport.TabIndex = 10;
            this.chkDwmImport.Text = "Load down wood material data to BioSum";
            this.chkDwmImport.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(72, 328);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(453, 17);
            this.label2.TabIndex = 9;
            this.label2.Text = "percent, condition status will be changed from forested to nonsampled.";
            // 
            // cmbCondPropPercent
            // 
            this.cmbCondPropPercent.FormattingEnabled = true;
            this.cmbCondPropPercent.Location = new System.Drawing.Point(479, 294);
            this.cmbCondPropPercent.Margin = new System.Windows.Forms.Padding(4);
            this.cmbCondPropPercent.Name = "cmbCondPropPercent";
            this.cmbCondPropPercent.Size = new System.Drawing.Size(64, 24);
            this.cmbCondPropPercent.TabIndex = 8;
            this.cmbCondPropPercent.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCondPropPercent_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(72, 298);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(404, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "When a forested condition has a condition proportion less than";
            // 
            // rdoFilterNone
            // 
            this.rdoFilterNone.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterNone.Location = new System.Drawing.Point(50, 112);
            this.rdoFilterNone.Margin = new System.Windows.Forms.Padding(4);
            this.rdoFilterNone.Name = "rdoFilterNone";
            this.rdoFilterNone.Size = new System.Drawing.Size(500, 40);
            this.rdoFilterNone.TabIndex = 0;
            this.rdoFilterNone.Text = "Input All Plots (Not Recommended)";
            this.rdoFilterNone.Click += new System.EventHandler(this.rdoFilterNone_Click);
            // 
            // rdoFilterByMenu
            // 
            this.rdoFilterByMenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByMenu.Location = new System.Drawing.Point(50, 79);
            this.rdoFilterByMenu.Margin = new System.Windows.Forms.Padding(4);
            this.rdoFilterByMenu.Name = "rdoFilterByMenu";
            this.rdoFilterByMenu.Size = new System.Drawing.Size(535, 40);
            this.rdoFilterByMenu.TabIndex = 1;
            this.rdoFilterByMenu.Text = "Filter Plots By Menu Selection (State, County, And Plot)";
            this.rdoFilterByMenu.Click += new System.EventHandler(this.rdoFilterByMenu_Click);
            // 
            // chkNonForested
            // 
            this.chkNonForested.Location = new System.Drawing.Point(166, 153);
            this.chkNonForested.Margin = new System.Windows.Forms.Padding(4);
            this.chkNonForested.Name = "chkNonForested";
            this.chkNonForested.Size = new System.Drawing.Size(140, 20);
            this.chkNonForested.TabIndex = 6;
            this.chkNonForested.Text = "Non Forested";
            // 
            // chkForested
            // 
            this.chkForested.Checked = true;
            this.chkForested.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkForested.Location = new System.Drawing.Point(76, 153);
            this.chkForested.Margin = new System.Windows.Forms.Padding(4);
            this.chkForested.Name = "chkForested";
            this.chkForested.Size = new System.Drawing.Size(90, 20);
            this.chkForested.TabIndex = 5;
            this.chkForested.Text = "Forested";
            // 
            // btnFilterByFileBrowse
            // 
            this.btnFilterByFileBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnFilterByFileBrowse.Image")));
            this.btnFilterByFileBrowse.Location = new System.Drawing.Point(510, 42);
            this.btnFilterByFileBrowse.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByFileBrowse.Name = "btnFilterByFileBrowse";
            this.btnFilterByFileBrowse.Size = new System.Drawing.Size(40, 40);
            this.btnFilterByFileBrowse.TabIndex = 4;
            this.btnFilterByFileBrowse.Click += new System.EventHandler(this.btnFilterByFileBrowse_Click);
            // 
            // txtFilterByFile
            // 
            this.txtFilterByFile.Enabled = false;
            this.txtFilterByFile.Location = new System.Drawing.Point(76, 51);
            this.txtFilterByFile.Margin = new System.Windows.Forms.Padding(4);
            this.txtFilterByFile.Name = "txtFilterByFile";
            this.txtFilterByFile.Size = new System.Drawing.Size(409, 22);
            this.txtFilterByFile.TabIndex = 3;
            // 
            // rdoFilterByFile
            // 
            this.rdoFilterByFile.Checked = true;
            this.rdoFilterByFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoFilterByFile.Location = new System.Drawing.Point(50, 12);
            this.rdoFilterByFile.Margin = new System.Windows.Forms.Padding(4);
            this.rdoFilterByFile.Name = "rdoFilterByFile";
            this.rdoFilterByFile.Size = new System.Drawing.Size(500, 40);
            this.rdoFilterByFile.TabIndex = 2;
            this.rdoFilterByFile.TabStop = true;
            this.rdoFilterByFile.Text = "Filter By File (Text File Containing Plot_CN numbers)";
            this.rdoFilterByFile.CheckedChanged += new System.EventHandler(this.rdoFilterByFile_CheckedChanged);
            this.rdoFilterByFile.Click += new System.EventHandler(this.rdoFilterByFile_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(4, 19);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(872, 30);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Plot Data Input";
            // 
            // grpboxFilterByPlot
            // 
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotFinish);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotUnselect);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotSelect);
            this.grpboxFilterByPlot.Controls.Add(this.lstFilterByPlot);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotHelp);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotPrevious);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotNext);
            this.grpboxFilterByPlot.Controls.Add(this.btnFilterByPlotCancel);
            this.grpboxFilterByPlot.Location = new System.Drawing.Point(20, 1455);
            this.grpboxFilterByPlot.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxFilterByPlot.Name = "grpboxFilterByPlot";
            this.grpboxFilterByPlot.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxFilterByPlot.Size = new System.Drawing.Size(840, 450);
            this.grpboxFilterByPlot.TabIndex = 32;
            this.grpboxFilterByPlot.TabStop = false;
            this.grpboxFilterByPlot.Text = "Filter By Plot";
            this.grpboxFilterByPlot.Enter += new System.EventHandler(this.groupBox6_Enter);
            // 
            // btnFilterByPlotFinish
            // 
            this.btnFilterByPlotFinish.Location = new System.Drawing.Point(730, 406);
            this.btnFilterByPlotFinish.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotFinish.Name = "btnFilterByPlotFinish";
            this.btnFilterByPlotFinish.Size = new System.Drawing.Size(90, 30);
            this.btnFilterByPlotFinish.TabIndex = 33;
            this.btnFilterByPlotFinish.Text = "Append";
            this.btnFilterByPlotFinish.Click += new System.EventHandler(this.btnFilterByPlotFinish_Click);
            // 
            // btnFilterByPlotUnselect
            // 
            this.btnFilterByPlotUnselect.Location = new System.Drawing.Point(700, 225);
            this.btnFilterByPlotUnselect.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotUnselect.Name = "btnFilterByPlotUnselect";
            this.btnFilterByPlotUnselect.Size = new System.Drawing.Size(110, 60);
            this.btnFilterByPlotUnselect.TabIndex = 32;
            this.btnFilterByPlotUnselect.Text = "Clear All";
            this.btnFilterByPlotUnselect.Click += new System.EventHandler(this.btnFilterByPlotUnselect_Click);
            // 
            // btnFilterByPlotSelect
            // 
            this.btnFilterByPlotSelect.Location = new System.Drawing.Point(700, 145);
            this.btnFilterByPlotSelect.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotSelect.Name = "btnFilterByPlotSelect";
            this.btnFilterByPlotSelect.Size = new System.Drawing.Size(110, 60);
            this.btnFilterByPlotSelect.TabIndex = 31;
            this.btnFilterByPlotSelect.Text = "Select All";
            this.btnFilterByPlotSelect.Click += new System.EventHandler(this.btnFilterByPlotSelect_Click);
            // 
            // lstFilterByPlot
            // 
            this.lstFilterByPlot.CheckBoxes = true;
            this.lstFilterByPlot.FullRowSelect = true;
            this.lstFilterByPlot.GridLines = true;
            this.lstFilterByPlot.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFilterByPlot.HideSelection = false;
            this.lstFilterByPlot.Location = new System.Drawing.Point(170, 40);
            this.lstFilterByPlot.Margin = new System.Windows.Forms.Padding(4);
            this.lstFilterByPlot.MultiSelect = false;
            this.lstFilterByPlot.Name = "lstFilterByPlot";
            this.lstFilterByPlot.Size = new System.Drawing.Size(499, 349);
            this.lstFilterByPlot.TabIndex = 30;
            this.lstFilterByPlot.UseCompatibleStateImageBehavior = false;
            this.lstFilterByPlot.View = System.Windows.Forms.View.Details;
            // 
            // btnFilterByPlotHelp
            // 
            this.btnFilterByPlotHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnFilterByPlotHelp.Location = new System.Drawing.Point(20, 406);
            this.btnFilterByPlotHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotHelp.Name = "btnFilterByPlotHelp";
            this.btnFilterByPlotHelp.Size = new System.Drawing.Size(80, 30);
            this.btnFilterByPlotHelp.TabIndex = 23;
            this.btnFilterByPlotHelp.Text = "Help";
            this.btnFilterByPlotHelp.Click += new System.EventHandler(this.btnFilterByPlotHelp_Click);
            // 
            // btnFilterByPlotPrevious
            // 
            this.btnFilterByPlotPrevious.Location = new System.Drawing.Point(530, 406);
            this.btnFilterByPlotPrevious.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotPrevious.Name = "btnFilterByPlotPrevious";
            this.btnFilterByPlotPrevious.Size = new System.Drawing.Size(90, 30);
            this.btnFilterByPlotPrevious.TabIndex = 22;
            this.btnFilterByPlotPrevious.Text = "< Previous";
            this.btnFilterByPlotPrevious.Click += new System.EventHandler(this.btnFilterByPlotPrevious_Click);
            // 
            // btnFilterByPlotNext
            // 
            this.btnFilterByPlotNext.Enabled = false;
            this.btnFilterByPlotNext.Location = new System.Drawing.Point(620, 406);
            this.btnFilterByPlotNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotNext.Name = "btnFilterByPlotNext";
            this.btnFilterByPlotNext.Size = new System.Drawing.Size(90, 30);
            this.btnFilterByPlotNext.TabIndex = 21;
            this.btnFilterByPlotNext.Text = "Next >";
            // 
            // btnFilterByPlotCancel
            // 
            this.btnFilterByPlotCancel.Location = new System.Drawing.Point(420, 406);
            this.btnFilterByPlotCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnFilterByPlotCancel.Name = "btnFilterByPlotCancel";
            this.btnFilterByPlotCancel.Size = new System.Drawing.Size(80, 30);
            this.btnFilterByPlotCancel.TabIndex = 20;
            this.btnFilterByPlotCancel.Text = "Cancel";
            this.btnFilterByPlotCancel.Click += new System.EventHandler(this.btnFilterByPlotCancel_Click);
            // 
            // groupBox25
            // 
            this.groupBox25.Controls.Add(this.txtMDBSiteTreeTable);
            this.groupBox25.Controls.Add(this.btnMDBSiteTreeBrowse);
            this.groupBox25.Controls.Add(this.txtMDBSiteTree);
            this.groupBox25.Location = new System.Drawing.Point(24, 248);
            this.groupBox25.Name = "groupBox25";
            this.groupBox25.Size = new System.Drawing.Size(624, 73);
            this.groupBox25.TabIndex = 31;
            this.groupBox25.TabStop = false;
            this.groupBox25.Text = "Site Tree Data";
            // 
            // txtMDBSiteTreeTable
            // 
            this.txtMDBSiteTreeTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBSiteTreeTable.Location = new System.Drawing.Point(408, 27);
            this.txtMDBSiteTreeTable.Name = "txtMDBSiteTreeTable";
            this.txtMDBSiteTreeTable.Size = new System.Drawing.Size(152, 30);
            this.txtMDBSiteTreeTable.TabIndex = 1;
            // 
            // btnMDBSiteTreeBrowse
            // 
            this.btnMDBSiteTreeBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBSiteTreeBrowse.Image")));
            this.btnMDBSiteTreeBrowse.Location = new System.Drawing.Point(573, 23);
            this.btnMDBSiteTreeBrowse.Name = "btnMDBSiteTreeBrowse";
            this.btnMDBSiteTreeBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBSiteTreeBrowse.TabIndex = 2;
            this.btnMDBSiteTreeBrowse.Click += new System.EventHandler(this.btnMDBSiteTreeBrowse_Click);
            // 
            // txtMDBSiteTree
            // 
            this.txtMDBSiteTree.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBSiteTree.Location = new System.Drawing.Point(17, 27);
            this.txtMDBSiteTree.Name = "txtMDBSiteTree";
            this.txtMDBSiteTree.Size = new System.Drawing.Size(383, 30);
            this.txtMDBSiteTree.TabIndex = 0;
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.txtMDBTreeTable);
            this.groupBox8.Controls.Add(this.btnMDBTreeBrowse);
            this.groupBox8.Controls.Add(this.txtMDBTree);
            this.groupBox8.Location = new System.Drawing.Point(24, 167);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(624, 73);
            this.groupBox8.TabIndex = 30;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Tree Data";
            // 
            // txtMDBTreeTable
            // 
            this.txtMDBTreeTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBTreeTable.Location = new System.Drawing.Point(408, 27);
            this.txtMDBTreeTable.Name = "txtMDBTreeTable";
            this.txtMDBTreeTable.Size = new System.Drawing.Size(152, 30);
            this.txtMDBTreeTable.TabIndex = 1;
            // 
            // btnMDBTreeBrowse
            // 
            this.btnMDBTreeBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBTreeBrowse.Image")));
            this.btnMDBTreeBrowse.Location = new System.Drawing.Point(573, 23);
            this.btnMDBTreeBrowse.Name = "btnMDBTreeBrowse";
            this.btnMDBTreeBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBTreeBrowse.TabIndex = 2;
            this.btnMDBTreeBrowse.Click += new System.EventHandler(this.btnMDBTreeBrowse_Click);
            // 
            // txtMDBTree
            // 
            this.txtMDBTree.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBTree.Location = new System.Drawing.Point(17, 27);
            this.txtMDBTree.Name = "txtMDBTree";
            this.txtMDBTree.Size = new System.Drawing.Size(383, 30);
            this.txtMDBTree.TabIndex = 0;
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.txtMDBCondTable);
            this.groupBox9.Controls.Add(this.btnMDBCondBrowse);
            this.groupBox9.Controls.Add(this.txtMDBCond);
            this.groupBox9.Location = new System.Drawing.Point(24, 95);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(624, 65);
            this.groupBox9.TabIndex = 29;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Condition Data";
            // 
            // txtMDBCondTable
            // 
            this.txtMDBCondTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBCondTable.Location = new System.Drawing.Point(408, 24);
            this.txtMDBCondTable.Name = "txtMDBCondTable";
            this.txtMDBCondTable.Size = new System.Drawing.Size(152, 30);
            this.txtMDBCondTable.TabIndex = 1;
            // 
            // btnMDBCondBrowse
            // 
            this.btnMDBCondBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBCondBrowse.Image")));
            this.btnMDBCondBrowse.Location = new System.Drawing.Point(573, 18);
            this.btnMDBCondBrowse.Name = "btnMDBCondBrowse";
            this.btnMDBCondBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBCondBrowse.TabIndex = 2;
            this.btnMDBCondBrowse.Click += new System.EventHandler(this.btnMDBCondBrowse_Click);
            // 
            // txtMDBCond
            // 
            this.txtMDBCond.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBCond.Location = new System.Drawing.Point(17, 24);
            this.txtMDBCond.Name = "txtMDBCond";
            this.txtMDBCond.Size = new System.Drawing.Size(383, 30);
            this.txtMDBCond.TabIndex = 0;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.txtMDBPlotTable);
            this.groupBox10.Controls.Add(this.btnMDBPlotBrowse);
            this.groupBox10.Controls.Add(this.txtMDBPlot);
            this.groupBox10.Location = new System.Drawing.Point(24, 22);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(624, 66);
            this.groupBox10.TabIndex = 28;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "Plot Data";
            // 
            // txtMDBPlotTable
            // 
            this.txtMDBPlotTable.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBPlotTable.Location = new System.Drawing.Point(408, 26);
            this.txtMDBPlotTable.Name = "txtMDBPlotTable";
            this.txtMDBPlotTable.Size = new System.Drawing.Size(152, 30);
            this.txtMDBPlotTable.TabIndex = 1;
            // 
            // btnMDBPlotBrowse
            // 
            this.btnMDBPlotBrowse.Image = ((System.Drawing.Image)(resources.GetObject("btnMDBPlotBrowse.Image")));
            this.btnMDBPlotBrowse.Location = new System.Drawing.Point(573, 22);
            this.btnMDBPlotBrowse.Name = "btnMDBPlotBrowse";
            this.btnMDBPlotBrowse.Size = new System.Drawing.Size(32, 32);
            this.btnMDBPlotBrowse.TabIndex = 2;
            this.btnMDBPlotBrowse.Click += new System.EventHandler(this.btnMDBPlotBrowse_Click);
            // 
            // txtMDBPlot
            // 
            this.txtMDBPlot.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMDBPlot.Location = new System.Drawing.Point(17, 26);
            this.txtMDBPlot.Name = "txtMDBPlot";
            this.txtMDBPlot.Size = new System.Drawing.Size(383, 30);
            this.txtMDBPlot.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbFiadbPlotGeomTable);
            this.groupBox2.Location = new System.Drawing.Point(24, 148);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(390, 56);
            this.groupBox2.TabIndex = 52;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Plot Geometry Data";
            // 
            // cmbFiadbPlotGeomTable
            // 
            this.cmbFiadbPlotGeomTable.Location = new System.Drawing.Point(9, 20);
            this.cmbFiadbPlotGeomTable.Margin = new System.Windows.Forms.Padding(4);
            this.cmbFiadbPlotGeomTable.Name = "cmbFiadbPlotGeomTable";
            this.cmbFiadbPlotGeomTable.Size = new System.Drawing.Size(370, 24);
            this.cmbFiadbPlotGeomTable.TabIndex = 0;
            // 
            // uc_plot_input
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "uc_plot_input";
            this.Size = new System.Drawing.Size(880, 3375);
            this.groupBox1.ResumeLayout(false);
            this.grpboxMDBFiadbInput.ResumeLayout(false);
            this.groupBox24.ResumeLayout(false);
            this.groupBox23.ResumeLayout(false);
            this.groupBox23.PerformLayout();
            this.groupBox15.ResumeLayout(false);
            this.groupBox16.ResumeLayout(false);
            this.groupBox17.ResumeLayout(false);
            this.groupBox18.ResumeLayout(false);
            this.groupBox19.ResumeLayout(false);
            this.groupBox20.ResumeLayout(false);
            this.groupBox21.ResumeLayout(false);
            this.groupBox22.ResumeLayout(false);
            this.grpboxFIADBInv.ResumeLayout(false);
            this.grpboxIDBInv.ResumeLayout(false);
            this.grpboxFilterByState.ResumeLayout(false);
            this.grpboxFilter.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.grpboxFilterByPlot.ResumeLayout(false);
            this.groupBox25.ResumeLayout(false);
            this.groupBox25.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void groupBox6_Enter(object sender, System.EventArgs e)
		{
		
		}
		private void Initialize()
		{
            this.Width = 1100;
            this.m_DialogWd = this.Width + 10;
			this.m_DialogHt = this.groupBox1.Top + this.grpboxMDBFiadbInput.Top + this.grpboxMDBFiadbInput.Height + 100 ;

		
					
			this.grpboxFilterByState.Left = this.grpboxMDBFiadbInput.Left;
			this.grpboxFilterByState.Width = this.grpboxMDBFiadbInput.Width;
			this.grpboxFilterByState.Height = this.grpboxMDBFiadbInput.Height;
			this.grpboxFilterByState.Top = this.grpboxMDBFiadbInput.Top;
            this.btnFilterByStateHelp.Location = this.btnMDBFiadbInputHelp.Location;
            this.btnFilterByStateCancel.Location = this.btnMDBFiadbInputCancel.Location;
            this.btnFilterByStatePrevious.Location = this.btnMDBFiadbInputPrev.Location;
			this.btnFilterByStateNext.Location = this.btnMDBFiadbInputNext.Location;
			this.btnFilterByStateFinish.Location = this.btnMDBFiadbInputFinish.Location;
			this.grpboxFilterByState.Visible=false;	

			this.grpboxFilter.Left = this.grpboxMDBFiadbInput.Left;
			this.grpboxFilter.Width = this.grpboxMDBFiadbInput.Width;
			this.grpboxFilter.Height = this.grpboxMDBFiadbInput.Height;
			this.grpboxFilter.Top = this.grpboxMDBFiadbInput.Top;
			this.btnFilterHelp.Location = this.btnMDBFiadbInputHelp.Location;
			this.btnFilterCancel.Location = this.btnMDBFiadbInputCancel.Location;
			this.btnFilterPrevious.Location = this.btnMDBFiadbInputPrev.Location;
			this.btnFilterNext.Location = this.btnMDBFiadbInputNext.Location;
			this.btnFilterFinish.Location = this.btnMDBFiadbInputFinish.Location;
			this.grpboxFilter.Visible=false;	

			this.grpboxFilterByPlot.Left = this.grpboxMDBFiadbInput.Left;
			this.grpboxFilterByPlot.Width = this.grpboxMDBFiadbInput.Width;
			this.grpboxFilterByPlot.Height = this.grpboxMDBFiadbInput.Height;
			this.grpboxFilterByPlot.Top = this.grpboxMDBFiadbInput.Top;
			this.btnFilterByPlotHelp.Location = this.btnMDBFiadbInputHelp.Location;
			this.btnFilterByPlotCancel.Location = this.btnMDBFiadbInputCancel.Location;
			this.btnFilterByPlotPrevious.Location = this.btnMDBFiadbInputPrev.Location;
			this.btnFilterByPlotNext.Location = this.btnMDBFiadbInputNext.Location;
			this.btnFilterByPlotFinish.Location = this.btnMDBFiadbInputFinish.Location;
			this.grpboxFilterByPlot.Visible=false;

			this.grpboxIDBInv.Left = this.grpboxMDBFiadbInput.Left;
			this.grpboxIDBInv.Width = this.grpboxMDBFiadbInput.Width;
			this.grpboxIDBInv.Height = this.grpboxMDBFiadbInput.Height;
			this.grpboxIDBInv.Top = this.grpboxMDBFiadbInput.Top;
			this.btnIDBInvHelp.Location = this.btnMDBFiadbInputHelp.Location;
			this.btnIDBInvCancel.Location = this.btnMDBFiadbInputCancel.Location;
			this.btnIDBInvPrevious.Location = this.btnMDBFiadbInputPrev.Location;
			this.btnIDBInvNext.Location = this.btnMDBFiadbInputNext.Location;
			this.btnIDBInvAppend.Location = this.btnMDBFiadbInputFinish.Location;
			this.grpboxIDBInv.Visible=false;	

			this.grpboxFIADBInv.Left = this.grpboxMDBFiadbInput.Left;
			this.grpboxFIADBInv.Width = this.grpboxMDBFiadbInput.Width;
			this.grpboxFIADBInv.Height = this.grpboxMDBFiadbInput.Height;
			this.grpboxFIADBInv.Top = this.grpboxMDBFiadbInput.Top;
			this.btnFIADBInvHelp.Location = this.btnMDBFiadbInputHelp.Location;
            this.btnFIADBInvHelp.Click += new System.EventHandler(this.btnFIADBInvHelp_Click);
			this.btnFIADBInvCancel.Location = this.btnMDBFiadbInputCancel.Location;
			this.btnFIADBInvPrevious.Location = this.btnMDBFiadbInputPrev.Location;
			this.btnFIADBInvNext.Location = this.btnMDBFiadbInputNext.Location;
			this.btnFIADBInvAppend.Location = this.btnMDBFiadbInputFinish.Location;
			this.grpboxFIADBInv.Visible=false;	

			this.lstFilterByState.Clear();
			this.lstFilterByState.Columns.Add(" ", 100, HorizontalAlignment.Center); 
			this.lstFilterByState.Columns.Add("State", 100, HorizontalAlignment.Left);
			this.lstFilterByState.Columns.Add("County", 100, HorizontalAlignment.Left);

			//create state,count table
			
			this.m_dtStateCounty = new DataTable("statecounty");
			this.m_dtStateCounty.Columns.Add("statecd",typeof(string));
			this.m_dtStateCounty.Columns.Add("countycd",typeof(string));

			// two columns in the Primary Key.
			DataColumn[] colPk = new DataColumn[2];
			colPk[0] = this.m_dtStateCounty.Columns["statecd"];
			colPk[1] = this.m_dtStateCounty.Columns["countycd"];
			this.m_dtStateCounty.PrimaryKey = colPk;


			//create state,county,plot table
			this.m_dtPlot = new DataTable("statecountyplot");
			this.m_dtPlot.Columns.Add("statecd",typeof(string));
			this.m_dtPlot.Columns.Add("countycd",typeof(string));
			this.m_dtPlot.Columns.Add("plot",typeof(string));

            for (int x = 1; x <= 99; x++)
            {
                cmbCondPropPercent.Items.Add(x.ToString().Trim());
            }
            cmbCondPropPercent.Text = "25";

            this.m_oEnv = new env();

		}

        private void InitializeDatasource()
		{
			string strProjDir=frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
			
			m_oDatasource = new Datasource();
			m_oDatasource.LoadTableColumnNamesAndDataTypes=false;
			m_oDatasource.LoadTableRecordCount=false;
			m_oDatasource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
			m_oDatasource.m_strDataSourceTableName = "datasource";
			m_oDatasource.m_strScenarioId="";
			m_oDatasource.populate_datasource_array();

			//get table names
			this.m_strPlotTable = m_oDatasource.getValidDataSourceTableName("PLOT");
			this.m_strCondTable = m_oDatasource.getValidDataSourceTableName("CONDITION");
			this.m_strTreeTable = m_oDatasource.getValidDataSourceTableName("TREE");
			this.m_strSiteTreeTable = m_oDatasource.getValidDataSourceTableName("SITE TREE");
            this.m_strBiosumPopStratumAdjustmentFactorsTable = m_oDatasource.getValidDataSourceTableName("BIOSUM POP STRATUM ADJUSTMENT FACTORS");
            this.m_strTreeMacroPlotBreakPointDiaTable = m_oDatasource.getValidDataSourceTableName("FIA TREE MACRO PLOT BREAKPOINT DIAMETER");
		}

		private void btnInvTypeCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFIADBTxtInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFilterPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFilter.Visible=false;
            this.grpboxMDBFiadbInput.Visible = true;
		}

		private void rdoFilterByFile_Click(object sender, System.EventArgs e)
		{
				this.btnFilterFinish.Enabled=false;
				this.chkForested.Enabled=false;
				this.chkNonForested.Enabled=false;
				this.btnFilterNext.Enabled=true;
				this.txtFilterByFile.Enabled=true;
				this.btnFilterByFileBrowse.Enabled=true;
		}

		private void rdoFilterByMenu_Click(object sender, System.EventArgs e)
		{
			//if (rdoFilterByFile.Checked==false && this.txtFilterByFile.Enabled==true) 
			//{
			this.btnFilterFinish.Enabled=false;
			this.chkForested.Enabled=true;
			this.chkNonForested.Enabled=true;
			this.btnFilterNext.Enabled=true;
			this.txtFilterByFile.Enabled=false;
			this.btnFilterByFileBrowse.Enabled=false;
			//}
			
		}

		private void rdoFilterNone_Click(object sender, System.EventArgs e)
		{

			this.btnFilterFinish.Enabled=false;
			this.chkForested.Enabled=true;
			this.chkNonForested.Enabled=true;
			this.btnFilterNext.Enabled=true;
			this.txtFilterByFile.Enabled=false;
			this.btnFilterByFileBrowse.Enabled=false;
		}

		private void btnFilterCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFilterFinish_Click(object sender, System.EventArgs e)
		{
			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
			this.m_intError=0;

            CalculateAdjustments_Start();
            if (m_intError == 0)
            {
                if (this.rdoFilterNone.Checked == true)
                {
                    LoadMDBPlotCondTreeData_Start();
                }
                else if (this.rdoFilterByFile.Checked == true)
                {
                    if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
                    {
                        this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                        {
                            this.LoadMDBPlotCondTreeData_Start();
                        }
                    }
                    else
                    {
                        MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!", "Add Plot Data", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                }
			}
		}

		private string CreateBiosumPlotId(System.Data.DataRow p_dr)
		{
			string strBiosumPlotId="";
			string strInvId ="";
			string strStateCd = "";
			string strCycle="";
			string strSubCycle="";
			string strCountyCd="";
			string strPlot="";
			string strForestBlm="";

			//inventory id
            strBiosumPlotId = "1";
			if (p_dr["measyear"] != System.DBNull.Value &&
				p_dr["measyear"].ToString().Trim().Length > 0)
			{
				strInvId = p_dr["measyear"].ToString().Trim();
			}
			else
			{
				strInvId = "9999";
			}

			strBiosumPlotId = strBiosumPlotId + strInvId;


			
			//state
			if (p_dr["statecd"] != System.DBNull.Value &&
				p_dr["statecd"].ToString().Trim().Length > 0)
			{
				strStateCd= p_dr["statecd"].ToString().Trim();
			}

			switch (strStateCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "99";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strStateCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strStateCd.Trim();
					break;
			}

			//cycle
			if (p_dr["cycle"] != System.DBNull.Value &&
				p_dr["cycle"].ToString().Trim().Length > 0)
			{
				strCycle = p_dr["cycle"].ToString().Trim();
			}
				
			switch (strCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCycle.Trim();
					break;
			}

			//subcycle
			if (p_dr["subcycle"] != System.DBNull.Value &&
				p_dr["subcycle"].ToString().Trim().Length > 0)
			{
				strSubCycle = p_dr["subcycle"].ToString().Trim();
			}
				
			switch (strSubCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strSubCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strSubCycle.Trim();
					break;
			}


			//countycode

			if (p_dr["countycd"] != System.DBNull.Value &&
				p_dr["countycd"].ToString().Trim().Length > 0)
			{
				strCountyCd = p_dr["countycd"].ToString().Trim();
			}

			switch (strCountyCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strCountyCd.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strCountyCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCountyCd.Trim();
					break;
			}

			//plot
			if (p_dr["plot"] != System.DBNull.Value &&
				p_dr["plot"].ToString().Trim().Length > 0)
			{
				strPlot = p_dr["plot"].ToString().Trim();
			}


			switch (strPlot.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "9999999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "000000" + strPlot.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "00000" + strPlot.Trim();
					break;
				case 3:
					strBiosumPlotId = strBiosumPlotId + "0000" + strPlot.Trim();
					break;
				case 4:
					strBiosumPlotId = strBiosumPlotId + "000" + strPlot.Trim();
					break;
				case 5:
					strBiosumPlotId = strBiosumPlotId + "00" + strPlot.Trim();
					break;
				case 6:
					strBiosumPlotId = strBiosumPlotId + "0" + strPlot.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strPlot.Trim();
					break;
			}


			
			//forest or blm district - need value for pnw idb unique key value
		    strForestBlm="000";

			switch (strForestBlm.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strForestBlm.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strForestBlm.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strForestBlm.Trim();
					break;
			}
			return strBiosumPlotId;
		}
		
        private string CreateBiosumPlotId(System.Data.OleDb.OleDbDataReader p_dr)
		{
			string strBiosumPlotId="";
			string strInvId ="";
			string strStateCd = "";
			string strCycle="";
			string strSubCycle="";
			string strCountyCd="";
			string strPlot="";
			string strForestBlm="";

			
			//inventory id
			strBiosumPlotId = "1";
			if (p_dr["measyear"] != System.DBNull.Value &&
				p_dr["measyear"].ToString().Trim().Length > 0)
			{
				strInvId = p_dr["measyear"].ToString().Trim();
			}
			else
			{
				strInvId = "9999";
			}

			strBiosumPlotId = strBiosumPlotId + strInvId;


			
			//state
			if (p_dr["statecd"] != System.DBNull.Value &&
				p_dr["statecd"].ToString().Trim().Length > 0)
			{
				strStateCd= p_dr["statecd"].ToString().Trim();
			}

			switch (strStateCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "99";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strStateCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strStateCd.Trim();
					break;
			}

			//cycle
			if (p_dr["cycle"] != System.DBNull.Value &&
				p_dr["cycle"].ToString().Trim().Length > 0)
			{
				strCycle = p_dr["cycle"].ToString().Trim();
			}
				
			switch (strCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCycle.Trim();
					break;
			}

			//subcycle
			if (p_dr["subcycle"] != System.DBNull.Value &&
				p_dr["subcycle"].ToString().Trim().Length > 0)
			{
				strSubCycle = p_dr["subcycle"].ToString().Trim();
			}
				
			switch (strSubCycle.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "00";
					break;
				case 1:
					strBiosumPlotId =  strBiosumPlotId + "0" + strSubCycle.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strSubCycle.Trim();
					break;
			}


			//countycode

			if (p_dr["countycd"] != System.DBNull.Value &&
				p_dr["countycd"].ToString().Trim().Length > 0)
			{
				strCountyCd = p_dr["countycd"].ToString().Trim();
			}

			switch (strCountyCd.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strCountyCd.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strCountyCd.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strCountyCd.Trim();
					break;
			}

			//plot
			if (p_dr["plot"] != System.DBNull.Value &&
				p_dr["plot"].ToString().Trim().Length > 0)
			{
				strPlot = p_dr["plot"].ToString().Trim();
			}


			switch (strPlot.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "9999999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "000000" + strPlot.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "00000" + strPlot.Trim();
					break;
				case 3:
					strBiosumPlotId = strBiosumPlotId + "0000" + strPlot.Trim();
					break;
				case 4:
					strBiosumPlotId = strBiosumPlotId + "000" + strPlot.Trim();
					break;
				case 5:
					strBiosumPlotId = strBiosumPlotId + "00" + strPlot.Trim();
					break;
				case 6:
					strBiosumPlotId = strBiosumPlotId + "0" + strPlot.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strPlot.Trim();
					break;
			}


			
			//forest or blm district - need value for pnw idb unique key value
		    strForestBlm="000";

			switch (strForestBlm.Trim().Length)
			{
				case 0:
					strBiosumPlotId = strBiosumPlotId + "999";
					break;
				case 1:
					strBiosumPlotId = strBiosumPlotId + "00" + strForestBlm.Trim();
					break;
				case 2:
					strBiosumPlotId = strBiosumPlotId + "0" + strForestBlm.Trim();
					break;
				default:
					strBiosumPlotId = strBiosumPlotId + strForestBlm.Trim();
					break;
			}
			return strBiosumPlotId;
		}
        private void CleanupThread()
        {
           // ((frmDialog)this.ParentForm).m_frmMain.Visible = true;
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog,"Visible",true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            
        }
        private void ThreadCleanUp()
        {
            try
            {
               // ((frmDialog)this.ParentForm).m_frmMain.Visible = true;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);

                if (this.m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Dispose");

                    this.m_frmTherm = null;
                }
                
            }
            catch
            {
            }

        }
        private void CancelThread()
        {
            bool bAbort = frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel processing (Y/N)?");
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
        private void CalculateAdjustments_Process()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.CalculateAdjustments_Process\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
          
            int x = 0;
            string strFIADBDbFile = "";

            m_intAddedPlotRows = 0;
            m_intAddedCondRows = 0;
            m_intAddedTreeRows = 0;
            m_intAddedSiteTreeRows = 0;

            this.m_intError = 0;
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();

            //-----------PREPARATION FOR CALCULATING ADJUSTMENTS---------//

            try
            {
                //instatiate the oledb data access class
                this.m_ado = new ado_data_access();

                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);



                //open the temp db file
                string strConnection = SQLite.GetConnectionString(m_strTempDbFile);
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
                {
                    con.Open();
                    if (!GetBooleanValue((Control)m_frmTherm, "AbortProcess"))
                    {
                        this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Drop Work Tables");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                        if (SQLite.TableExist(con, "BIOSUM_PLOT"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_PLOT");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 20);
                        if (SQLite.TableExist(con, "BIOSUM_COND"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_COND");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                        if (SQLite.TableExist(con, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName))
                            SQLite.SqlNonQuery(con, "DROP TABLE " + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
                        SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                        if (SQLite.TableExist(con, "BIOSUM_PPSA"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_PPSA");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                        if (SQLite.TableExist(con, "BIOSUM_EUS_TEMP"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_EUS_TEMP");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 60);
                        if (SQLite.TableExist(con, "BIOSUM_PPSA_DENIED_ACCESS"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_PPSA_DENIED_ACCESS");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 70);
                        if (SQLite.TableExist(con, "BIOSUM_PPSA_TEMP"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_PPSA_TEMP");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                        if (SQLite.TableExist(con, "BIOSUM_EUS_TEMP"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_EUS_TEMP");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 90);
                        if (SQLite.TableExist(con, "BIOSUM_EUS_ACCESS"))
                            SQLite.SqlNonQuery(con, "DROP TABLE BIOSUM_EUS_ACCESS");

                        SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                        System.Threading.Thread.Sleep(2000);
                        SetThermValue(m_frmTherm.progressBar2, "Value", 30);
                        strFIADBDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtFiadbInputFile, "Text", false);
                        strFIADBDbFile = strFIADBDbFile.Trim();
                        string[] strSql = Queries.FIAPlot.FIADBPlotInput_CalculateAdjustmentFactorsSQL(
                            "POP_PLOT_STRATUM_ASSGN",
                            "POP_ESTN_UNIT",
                            "POP_STRATUM",
                            "POP_EVAL",
                            "PLOT",
                            "COND",
                             m_strCurrFIADBRsCd,
                             m_strCurrFIADBEvalId,
                             frmMain.g_oDelegate.GetControlPropertyValue(cmbCondPropPercent, "Text", false).ToString().Trim(),
                             strFIADBDbFile);
                        SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                        this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Calculate Adjustment Factors For RsCd=" + m_strCurrFIADBRsCd + " and EvalId=" + m_strCurrFIADBEvalId);

                        for (x = 0; x <= strSql.Length - 1; x++)
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, strSql[x] + "\r\n");
                            SQLite.SqlNonQuery(con, strSql[x]);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 20 + x + 5);
                            if (SQLite.m_intError != 0) break;
                        }

                        // Add indexes to BIOSUM_PLOT and BIOSUM_COND tables
                        SQLite.AddIndex(con, "BIOSUM_PLOT", "BIOSUM_PLOT_idx1", "CN");
                        SQLite.AddIndex(con, "BIOSUM_COND", "BIOSUM_COND_idx1", "CN");
                    }
                }

                m_intError = SQLite.m_intError;
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                
                //create tablelinks to the projects main folder
                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    //SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                    //this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Create table links");
                    //create a table link from the newly created BIOSUM_ADJFACTORS into the master.mdb

                    //instatiate dao for creating links in the temp table
                    //to the fiadb plot, cond, and tree input tables
                    //dao_data_access oDao = new dao_data_access();
                    //SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                    //create links to the fiadb input tables in the temp mdb file
                    //plot table
                    //strSourceTableName = "BIOSUM_PLOT";
                    //strDestTableLinkName = "fiadb_plot_input";
                    //oDao.CreateTableLink(this.m_strTempMDBFile,strDestTableLinkName,strFIADBDbFile,strSourceTableName,true);
                    //@ToDo: This is where it breaks because I have stopped creating the MS Access m_strTempMDBFile
                    //Should no longer be needed with conversion to SQLite. Use SQLite attach instead
                    //oDao.CreateSQLiteTableLink(this.m_strTempMDBFile, strSourceTableName, strDestTableLinkName,
                    //    ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
                    //SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                    //cond table
                    //strSourceTableName = "BIOSUM_COND";
                    //strDestTableLinkName = "fiadb_cond_input";
                    //if (oDao.m_intError==0) oDao.CreateSQLiteTableLink(this.m_strTempMDBFile, strSourceTableName, strDestTableLinkName,
                    //                            ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
                    //SetThermValue(m_frmTherm.progressBar1, "Value", 70);
                    //biosum adjustment factors table
                    //strSourceTableName = frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName;
                    //strDestTableLinkName = "fiadb_biosum_adjustment_factors_input";
                    //if (oDao.m_intError==0) oDao.CreateSQLiteTableLink(this.m_strTempMDBFile, strSourceTableName, strDestTableLinkName,
                    //                            ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
                    //m_intError = oDao.m_intError;
                    
                    //destroy the object and release it from memory
                    //oDao.m_DaoWorkspace.Close();
                    //oDao.m_DaoWorkspace = null;
                    //oDao = null;

                    //m_intError = SQLite.m_intError;
                    //SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    //System.Threading.Thread.Sleep(2000);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 70);                  
                    
                }
                //delete any records from the production biosum adjustment factor table that did not previously complete processing (error or user cancelled)
                //or any previous rscd and evalid that equal the current ones
                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                    this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Deleting Old Data");
                    //open the connection to the temp mdb file
                    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
                    {
                        con.Open();
                        SQLite.m_strSQL = $@"ATTACH DATABASE '{m_strMasterDbFile}' AS MASTER";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(con, SQLite.m_strSQL);
                        //delete any previous rscd and evalid that equal the current ones
                        SQLite.m_strSQL = $@"DELETE FROM MASTER.{this.m_strBiosumPopStratumAdjustmentFactorsTable} WHERE biosum_status_cd=9 or (rscd={m_strCurrFIADBRsCd} and evalid={m_strCurrFIADBEvalId})";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(con, SQLite.m_strSQL);
                        SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                        m_intError = SQLite.m_intError;

                        //append work table to production table
                        if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                        {
                            this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Appending New Data");
                            //SQLite.m_strSQL = $@"INSERT INTO MASTER.{this.m_strBiosumPopStratumAdjustmentFactorsTable} SELECT a.*,9 AS biosum_status_cd FROM biosum_pop_stratum_adjustment_factors a";
                            SQLite.m_strSQL = $@"INSERT INTO MASTER.biosum_pop_stratum_adjustment_factors 
                                SELECT stratum_cn,rscd,evalid,eval_descr,estn_unit,estn_unit_descr,stratumcd,
                                p2pointcnt_man,double_sampling,stratum_area,expns,pmh_macr,pmh_sub,pmh_micr,
                                pmh_cond,adj_factor_macr,adj_factor_subp,adj_factor_micr,9 AS biosum_status_cd 
                                FROM biosum_pop_stratum_adjustment_factors";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(con, SQLite.m_strSQL);
                            SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                            SetThermValue(m_frmTherm.progressBar2, "Value", 100);
                        }
                    }
                }



                if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {

                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
                }

                CalculateAdjustments_Finish();

            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (oConn != null)
                {
                    if (oConn.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(oConn);
                    }
                    oConn = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.CancelThreadCleanup();
                this.ThreadCleanUp();
                this.CleanupThread();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_plot_input.CalculateAdjustments_Process  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {

            }

            if (oConn != null)
            {
                if (oConn.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(oConn);
                }
                oConn = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);




            CalculateAdjustments_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }
		private void LoadMDBPlotCondTreeData_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
			string strBiosumPlotId="";
			string strFields="";
			
			int x=0;

			string strSourceTableLink="";
            m_intAddedPlotRows=0;
		    m_intAddedCondRows=0;
		    m_intAddedTreeRows=0;
			m_intAddedSiteTreeRows=0;

		    this.m_intError=0;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.LoadMDBPlotCondTreeData_Process\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

            //-----------PREPARATION FOR ADDING PLOT RECORDS---------//
            try
            {
                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1,"Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1,"Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1,"Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //tree table
                string strTreeSource = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbTreeTable, "Text", false);
                //seedling table
                string strSeedlingSource = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbSeedlingTable, "Text", false);
                if (strSeedlingSource.Trim().Length > 0 && strSeedlingSource.Trim() != "<Optional Table>" && Checked(ckImportSeedlings))
                {
                    // Do nothing
                }
                else
                {
                    m_bLoadSeedlings = false;
                }

                //site tree
                string strSiteTreeSource = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbSiteTreeTable, "Text", false);

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 10);

                System.Data.DataTable dtPlotSchema = new DataTable();
                System.Data.DataTable dtCondSchema = new DataTable();
                System.Data.DataTable dtTreeSchema = new DataTable();
                System.Data.DataTable dtSiteTreeSchema = new DataTable();
                System.Data.DataTable dtFIADBPlotSchema = new DataTable();
                System.Data.DataTable dtFIADBCondSchema = new DataTable();
                System.Data.DataTable dtFIADBTreeSchema = new DataTable();
                System.Data.DataTable dtFIADBSiteTreeSchema = new DataTable();
                DataTable dtFIADBSeedlingSchema = new DataTable();

                string strConn = SQLite.GetConnectionString(m_strMasterDbFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    //Before processing new plot information, delete any records that were not completely processed
                    DeleteFromTablesWhereFilterSqlite(conn, new string[] { this.m_strPlotTable },
                        " WHERE biosum_status_cd=9 OR LENGTH(biosum_plot_id)=0;");
                    DeleteFromTablesWhereFilterSqlite(conn, new string[]
                    {m_strCondTable, m_strTreeTable, m_strSiteTreeTable}, " WHERE biosum_status_cd=9;");
                    if (m_intError == 0)
                        m_intError = SQLite.m_intError;

                    if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 20);

                        /****************************************************************
                         **get the table structure that results from executing the sql
                         ****************************************************************/
                        //get the fiabiosum table structures
                        dtPlotSchema = SQLite.getTableSchema(conn, "select * from " + this.m_strPlotTable);
                        dtCondSchema = SQLite.getTableSchema(conn, "select * from " + this.m_strCondTable);
                        dtTreeSchema = SQLite.getTableSchema(conn, "select * from " + this.m_strTreeTable);
                        dtSiteTreeSchema = SQLite.getTableSchema(conn, "select * from " + this.m_strSiteTreeTable);

                        SQLite.m_strSQL = $@"ATTACH DATABASE '{m_strTempDbFile}' AS TEMPDB";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        SQLite.m_strSQL = $@"ATTACH DATABASE '{m_strCurrentFiadbInputFile}' AS FIADB";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        //get the fiadb table structures
                        dtFIADBPlotSchema = SQLite.getTableSchema(conn, "select * from TEMPDB.BIOSUM_PLOT");
                        dtFIADBCondSchema = SQLite.getTableSchema(conn, "select * from TEMPDB.BIOSUM_COND");
                        dtFIADBTreeSchema = SQLite.getTableSchema(conn, $@"select * from FIADB.{strTreeSource}");
                        if (m_bLoadSeedlings)
                        {
                            dtFIADBSeedlingSchema = SQLite.getTableSchema(conn, "select * from FIADB." + strSeedlingSource);
                        }
                        dtFIADBSiteTreeSchema = SQLite.getTableSchema(conn, $@"select * from FIADB.{strSiteTreeSource}");
                        m_intError = SQLite.m_intError;
                    }

                    if (this.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        //-------------PLOT TABLE----------------//
                        //build field list string to insert sql by matching columns in thebiosum and fiadb plot tables
                        strFields = CreateStrFieldsFromDataTables(dtPlotSchema, dtFIADBPlotSchema);

                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Plot Table: Insert New Plot Records");
                        if (Checked(rdoFilterByFile) == true && m_strPlotIdList.Trim().Length > 0 &&
                            !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                        {
                            string strDelimiter = ",";
                            string[] strPlotIdArray = m_strPlotIdList.Split(strDelimiter.ToCharArray());
                            SQLite.m_strSQL = "CREATE TABLE TEMPDB.input_cn (CN CHAR(34))";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                                    SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            for (x = 0; x <= strPlotIdArray.Length - 1; x++)
                            {
                                if (strPlotIdArray[x] != null && strPlotIdArray[x].Trim().Length > 0)
                                {
                                    SQLite.m_strSQL = "INSERT INTO input_cn (CN) VALUES (" + strPlotIdArray[x].Trim() + ")";
                                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                                        SQLite.m_strSQL + "\r\n");
                                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                                }
                            }
                            SQLite.m_strSQL = "CREATE TABLE TEMPDB.tempplot as SELECT '999999999999999999999999' AS biosum_plot_id,9 AS biosum_status_cd, p.* " +
                                "FROM BIOSUM_PLOT p ,TEMPDB." + this.m_strPpsaTable + " ppsa, input_cn" +
                                " WHERE p.cn=input_cn.cn AND " +
                                "p.cn=ppsa.plt_cn AND " +
                                "ppsa.rscd = " + this.m_strCurrFIADBRsCd + " AND " +
                                "ppsa.evalid = " + this.m_strCurrFIADBEvalId;
                        }
                        else
                        {

                            /********************************************************
                             **create plot input insert command
                             ********************************************************/
                            //check the user defined filters
                            SQLite.m_strSQL = SQLite.m_strSQL = "CREATE TABLE TEMPDB.tempplot as SELECT '999999999999999999999999' AS biosum_plot_id,9 AS biosum_status_cd, p.* " +
                                "FROM BIOSUM_PLOT p ,TEMPDB." + this.m_strPpsaTable + " ppsa" +
                                " WHERE p.cn=ppsa.plt_cn AND " +
                                "ppsa.rscd = " + this.m_strCurrFIADBRsCd + " AND " +
                                "ppsa.evalid = " + this.m_strCurrFIADBEvalId;
                        }

                        if (Checked(rdoFilterNone))
                        {
                            //forested/nonforested filters
                            if (Checked(chkNonForested) &&
                                Checked(chkForested))
                            {
                                //all plots

                            }
                            else if (Checked(chkForested))
                            {
                                //forested plots
                                SQLite.m_strSQL = SQLite.m_strSQL + " AND p.plot_status_cd = 1";
                            }
                            else
                            {
                                //nonforested plots
                                SQLite.m_strSQL = SQLite.m_strSQL + " AND p.plot_status_cd IS NULL OR p.plot_status_cd <> 1";
                            }
                        }

                        else if (Checked(rdoFilterByMenu))
                        {
                            if (Checked(chkNonForested) &&
                                Checked(chkForested))
                            {
                                //all plots
                            }
                            else if (Checked(chkForested))
                            {
                                //forested plots
                                SQLite.m_strSQL = SQLite.m_strSQL + " AND (p.plot_status_cd = 1) ";
                            }
                            else
                            {
                                //nonforested plots
                                SQLite.m_strSQL = SQLite.m_strSQL + " AND (p.plot_status_cd IS NULL OR p.plot_status_cd <> 1) ";
                            }
                            if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
                            {
                                //state,county,plot filter
                                this.BuildFilterByPlotString("ppsa.statecd", "ppsa.countycd", "ppsa.plot", false);
                                SQLite.m_strSQL += " AND " + this.m_strStateCountyPlotSQL.Trim() + ";";
                            }
                            else
                            {
                                //state,county filter
                                this.BuildFilterByStateCountyString("ppsa.statecd", "ppsa.countycd", false);
                                SQLite.m_strSQL += " AND " + this.m_strStateCountySQL.Trim() + ";";
                            }

                        }

                        //insert new plot records
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Plot Table: Update Biosum_Plot_Id Column");
                        System.Data.SQLite.SQLiteTransaction p_transTempPlot = conn.BeginTransaction();

                        //update the biosum_plot_id column
                        SQLite.m_DataSet = new DataSet("FIADB");
                        SQLite.m_DataAdapter = new System.Data.SQLite.SQLiteDataAdapter();
                        SQLite.AddSQLQueryToDataSet(conn,
                            ref SQLite.m_DataAdapter,
                            ref SQLite.m_DataSet,
                            ref p_transTempPlot,
                            "SELECT * FROM tempplot", "tempplot");

                        SQLite.ConfigureDataAdapterUpdateCommand(conn,
                            SQLite.m_DataAdapter,
                            p_transTempPlot,
                            "SELECT biosum_plot_id FROM tempplot",
                            "SELECT CN FROM tempplot",
                            "tempplot");

                        for (x = 0; x <= SQLite.m_DataSet.Tables["tempplot"].Rows.Count - 1; x++)
                        {
                            strBiosumPlotId = this.CreateBiosumPlotId(SQLite.m_DataSet.Tables["tempplot"].Rows[x]);
                            SQLite.m_DataSet.Tables["tempplot"].Rows[x].BeginEdit();
                            SQLite.m_DataSet.Tables["tempplot"].Rows[x]["biosum_plot_id"] = strBiosumPlotId;
                            SQLite.m_DataSet.Tables["tempplot"].Rows[x].EndEdit();
                        }
                        SQLite.m_DataAdapter.Update(SQLite.m_DataSet.Tables["tempplot"]);
                        p_transTempPlot.Commit();
                        SQLite.m_DataSet.Tables["tempplot"].AcceptChanges();
                        p_transTempPlot = null;
                        SQLite.m_DataAdapter.Dispose();
                        SQLite.m_DataAdapter = null;
                        m_intError = SQLite.m_intError;
                    }
                }

 



                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                    //insert the new plot records into the plot table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (biosum_plot_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_plot_id),biosum_status_cd," + strFields + " FROM tempplot";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    //create plot column update work table
                    this.m_strSQL = "SELECT biosum_plot_id, statecd as cond_ttl " +
                        "INTO plot_column_updates_work_table FROM " + this.m_strPlotTable.Trim() + " WHERE 1=2;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    System.Threading.Thread.Sleep(10000);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1,"Value",40);
                    SetThermValue(m_frmTherm.progressBar2,"Value",20);
                    //-------------CONDITION TABLE----------------//
                    strSourceTableLink = "fiadb_cond_input";
                    //build field list string to insert sql by matching FIADB and BioSum Cond columns
                    strFields = CreateStrFieldsFromDataTables(dtFIADBCondSchema, dtCondSchema);
                    /********************************************************
                     **create condition input insert command
                     ********************************************************/
                    //check the user defined filters
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Condition Table: Insert New  Records");
                    this.m_ado.m_strSQL = "SELECT p.biosum_plot_id, TRIM(p.biosum_plot_id) + TRIM(CSTR(c.condid)) AS biosum_cond_id,9 AS biosum_status_cd,c.*" +
                        " INTO tempcond FROM " + strSourceTableLink + " c, " +
                        this.m_strPlotTable +  " p" +
                        " WHERE c.plt_cn=p.cn AND" +
                        " p.biosum_status_cd=9";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                    //insert the new condition records into the condition table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strCondTable + " (biosum_plot_id,biosum_cond_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_plot_id),TRIM(biosum_cond_id),biosum_status_cd," + strFields + " FROM tempcond";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    // Trying this since I couldn't update from SQLite strSourceTableLink
                    // Get ODBC error 'Operation Must Use an updateable query'
                    // Seems to work vs creating a new temp table with this field
                    // Added new column to tempcond with varchar datatype to link to cond table
                    m_ado.AddColumn(this.m_connTempMDBFile, "tempcond", "CN_JOIN", "VARCHAR", "34");
                    // Populate new column
                    if (m_ado.m_intError == 0)
                    {
                        m_ado.m_strSQL = "UPDATE tempcond SET CN_JOIN = TRIM(CN)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    }

                    m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " d " +
                        "INNER JOIN tempcond s " +
                        "ON d.cn = s.CN_JOIN " +
                        "SET d.cond_status_cd = s.cond_status_cd";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    //create cond column work table
                    this.m_strSQL = "SELECT biosum_cond_id, qmd_all_inch,qmd_hwd_inch," +
                        "qmd_swd_inch,tpacurr,hwd_tpacurr,swd_tpacurr,ba_ft2_ac," +
                        "hwd_ba_ft2_ac,swd_ba_ft2_ac,vol_ac_grs_stem_ttl_ft3," +
                        "hwd_vol_ac_grs_stem_ttl_ft3,swd_vol_ac_grs_stem_ttl_ft3," +
                        "vol_ac_grs_ft3, hwd_vol_ac_grs_ft3," +
                        "swd_vol_ac_grs_ft3,volcsgrs," +
                        "hwd_volcsgrs, swd_volcsgrs INTO cond_column_updates_work_table " +
                        "FROM " + this.m_strCondTable.Trim() + " WHERE 1=2;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_strSQL + "\r\n");
                    if (m_ado.m_intError == 0)
                        m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_strSQL);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 55);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 40);
                    //-------------TREE TABLE----------------//
                    strSourceTableLink = "fiadb_tree_input";
                    //build field list string to insert sql by matching FIADB and BioSum Tree columns
                    strFields = CreateStrFieldsFromDataTables(dtFIADBTreeSchema, dtTreeSchema);
                    /********************************************************
                     **create tree input insert command
                     ********************************************************/
                    //check the user defined filters
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Tree Table: Insert New  Records");
                    this.m_ado.m_strSQL = "SELECT TRIM(p.biosum_plot_id) + TRIM(CSTR(t.condid)) AS biosum_cond_id,9 AS biosum_status_cd,t.* INTO temptree FROM " + strSourceTableLink + " t " +
                        " INNER JOIN " + this.m_strPlotTable + " p ON t.plt_cn=TRIM(p.cn) " +
                        " WHERE p.biosum_status_cd=9 AND t.statuscd<>0;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 60);
                    //insert the new tree records into the tree table; Note that temptree is used later for GRM processing
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strTreeTable + " (biosum_cond_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_cond_id),biosum_status_cd," + strFields + " FROM temptree";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                }

                // SEEDLINGS
                if (m_intError == 0 && m_bLoadSeedlings == true && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 65);
                    //-------------SEEDLING TABLE----------------//
                    strSourceTableLink = frmMain.g_oTables.m_oFIAPlot.DefaultSeedlingTableName;
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Seedling Table: Insert New  Records");
                    if (this.m_ado.TableExist(this.m_connTempMDBFile, "tempseedling"))
                    {
                        this.m_ado.m_strSQL = "DROP TABLE tempseedling";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                        if (m_ado.m_intError == 0)
                            this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    }
                    this.m_ado.m_strSQL = "SELECT TRIM(p.biosum_plot_id) + TRIM(CSTR(s.condid)) AS biosum_cond_id,9 AS biosum_status_cd, 0.1 as dia, 1 as diahtcd, " +
                        "'1' + Format(SPCD,'000') + '00' + SUBP AS fvs_tree_id, 1 as statuscd, s.* INTO tempseedling FROM " + 
                        strSourceTableLink + " s " + " INNER JOIN " + this.m_strPlotTable + " p ON s.plt_cn=TRIM(p.cn) " +
                        " WHERE p.biosum_status_cd=9;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    //Set DIAHTCD for Seedlings using FIA_TREE_SPECIES_REF.WOODLAND_YN
                    this.m_ado.m_strSQL = $@"UPDATE tempseedling t 
                                             INNER JOIN FIA_TREE_SPECIES_REF ref ON cint(t.spcd)=ref.spcd
                                             SET t.diahtcd=IIF(ref.woodland_yn='N', 1, 2)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    //Prepend CN with "S" to indicate seedlings
                    this.m_ado.m_strSQL = $@"UPDATE tempseedling t                                             
                                             SET CN = 'S' + CN";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    if (m_ado.m_intError == 0)
                        this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    //build field list string to insert sql by matching FIADB and BioSum Tree columns
                    strFields = CreateStrFieldsFromDataTables(dtFIADBSeedlingSchema, dtTreeSchema);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 70);
                    //insert the new seedling records into the tree table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strTreeTable + " (biosum_cond_id,biosum_status_cd,dia,diahtcd,fvs_tree_id,statuscd," + strFields + ") " +
                        "SELECT TRIM(biosum_cond_id),biosum_status_cd,dia,diahtcd,fvs_tree_id,statuscd, " + strFields + " FROM tempseedling";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    m_intError = m_ado.m_intError;
                }



                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                    //update the cullbf column
                    this.m_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " " +
                        "SET cullbf=IIF(cullbf IS NULL," +
                        "IIF(cull IS NOT NULL AND roughcull IS NOT NULL," +
                        "cull + roughcull," +
                        "IIF(cull IS NOT NULL,cull," +
                        "IIF(roughcull IS NOT NULL,roughcull,0))),cullbf)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    //-------------SITE TREE TABLE----------------//
                    strSourceTableLink = "fiadb_site_tree_input";
                    //build field list string to insert sql by matching biosum fiadb SiteTree table
                    strFields = CreateStrFieldsFromDataTables(dtFIADBSiteTreeSchema, dtSiteTreeSchema);
                    /********************************************************
                     **create site tree input insert command
                     ********************************************************/
                    //check the user defined filters
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Site Tree Table: Insert New  Records");
                    this.m_ado.m_strSQL = "SELECT TRIM(p.biosum_plot_id) AS biosum_plot_id,9 AS biosum_status_cd,t.* INTO tempsitetree FROM " + strSourceTableLink + " t " +
                        " INNER JOIN " + this.m_strPlotTable + " p ON t.plt_cn=TRIM(p.cn) " +
                        " WHERE p.biosum_status_cd=9";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);
                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                    //insert the new condition records into the condition table
                    m_ado.m_strSQL = "INSERT INTO " + this.m_strSiteTreeTable + " (biosum_plot_id,biosum_status_cd," + strFields + ") " +
                        "SELECT TRIM(biosum_plot_id),biosum_status_cd," + strFields + " FROM tempsitetree";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, this.m_ado.m_strSQL);

                    m_intError = m_ado.m_intError;
                }

                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                    SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                    m_intError = UpdateColumns(m_ado);
                }

				//Down Woody Materials Section
                if (m_intError == 0 && Checked(chkDwmImport) &&
                    !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
                {
                    SetThermValue(m_frmTherm.progressBar1, "Value",
                        GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                    SetThermValue(m_frmTherm.progressBar2, "Value", 80);
                    //m_intError = ImportDownWoodyMaterials(m_ado);
                    m_intError = ImportDownWoodyMaterialsSqlite();
                }

				//Growth Removal Mortality Section
                //if (m_intError == 0 && Checked(chkGrmImport) &&
                //    !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
                //{
                //    SetThermValue(m_frmTherm.progressBar1, "Value",
                //        GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                //    SetThermValue(m_frmTherm.progressBar2, "Value", 90);
                //    m_intError = ImportGrowthRemovalMortality(m_ado);
                //}

                //The connection should be closed if necessary and reopened
                if (m_connTempMDBFile != null && m_connTempMDBFile.State == System.Data.ConnectionState.Open)
                {
                    m_connTempMDBFile.Dispose();
                }

                using (m_connTempMDBFile = new OleDbConnection(m_strTempMDBFileConn))
                {
                    m_connTempMDBFile.Open();

                    //If no ADO errors have occured,
                    //count the number of records added to each table and
                    //set their biosum_status_cd values to 1
                    if (this.m_intError == 0 &&
                        !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
                    {

                        //Record counts associated with imported plots for each table
                        m_intAddedPlotRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strPlotTable);
                        m_intAddedCondRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strCondTable);
                        m_intAddedTreeRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strTreeTable);
                        m_intAddedSeedlingRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strSeedlingTable);
                        m_intAddedSiteTreeRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strSiteTreeTable);

                        if (Checked(chkDwmImport))
                        {
                            //m_intAddedDwmCwdRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strDwmCwdTable);
                            //m_intAddedDwmFwdRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strDwmFwdTable);
                            //m_intAddedDwmDuffLitterRows =
                            //    GetNewRecordCount(m_ado, m_connTempMDBFile, m_strDwmDuffLitterTable);
                            //m_intAddedDwmTransectSegmentRows = GetNewRecordCount(m_ado, m_connTempMDBFile,
                            //    m_strDwmTransectSegmentTable);
                            m_intAddedDwmCwdRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strDwmCwdTable + "_master");
                            m_intAddedDwmFwdRows = GetNewRecordCount(m_ado, m_connTempMDBFile, m_strDwmFwdTable + "_master");
                            m_intAddedDwmDuffLitterRows =
                                GetNewRecordCount(m_ado, m_connTempMDBFile, m_strDwmDuffLitterTable + "_master");
                            m_intAddedDwmTransectSegmentRows = GetNewRecordCount(m_ado, m_connTempMDBFile,
                                m_strDwmTransectSegmentTable + "_master");
                        }

                        //Successfully imported and updated plot data. Set biosum_status_cd to 1
                        //string[] arrTables = new string[]{m_strPlotTable, m_strCondTable, m_strTreeTable, m_strPopEvalTable, m_strPopStratumTable,
                        //    m_strPpsaTable, m_strPopEstUnitTable, m_strSiteTreeTable,
                        //    m_strBiosumPopStratumAdjustmentFactorsTable,
                        //    m_strDwmCwdTable, m_strDwmFwdTable, m_strDwmDuffLitterTable, m_strDwmTransectSegmentTable};
                        string[] arrTables = new string[]{m_strPlotTable, m_strCondTable, m_strTreeTable, m_strPopEvalTable, m_strPopStratumTable,
                            m_strPpsaTable, m_strPopEstUnitTable, m_strSiteTreeTable,
                            m_strBiosumPopStratumAdjustmentFactorsTable,
                            m_strDwmCwdTable + "_master", m_strDwmFwdTable + "_master", m_strDwmDuffLitterTable + "_master", m_strDwmTransectSegmentTable + "_master"};
                        foreach (string table in arrTables)
                        {
                            if (m_ado.TableExist(m_connTempMDBFile, table))
                            {
                                m_ado.m_strSQL = "UPDATE " + table +
                                                 " SET biosum_status_cd = 1 WHERE biosum_status_cd = 9;";
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, this.m_ado.m_strSQL + "\r\n");
                                m_ado.SqlNonQuery(m_connTempMDBFile, m_ado.m_strSQL);
                            }
                        }


                        // Save POP tables to SQLite
                        SaveSqlitePopTables(this.m_strCurrFIADBEvalId, this.m_strCurrFIADBRsCd);
                        SetThermValue(m_frmTherm.progressBar1, "Value",
                        GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                        SetThermValue(m_frmTherm.progressBar2, "Value",
                        GetThermValue(m_frmTherm.progressBar2, "Maximum"));
                        frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Button) m_frmTherm.btnCancel,
                            "Visible", false);

                        MessageBox.Show(CreateRecordsAddedMessage(), "Add Plot Data");

                        this.m_strLoadedPopEstUnitInputTable =
                            (string) frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.ComboBox) cmbFiadbPopEstUnitTable, "Text", false);
                        this.m_strLoadedPopStratumInputTable =
                            (string) frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.ComboBox) cmbFiadbPopStratumTable, "Text", false);
                        this.m_strLoadedPpsaInputTable =
                            (string) frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.ComboBox) cmbFiadbPpsaTable, "Text", false);
                        this.m_strLoadedFIADBEvalId = this.m_strCurrFIADBEvalId;
                        this.m_strLoadedFIADBRsCd = this.m_strCurrFIADBRsCd;
                        this.m_strLoadedFiadbInputFile =
                            (string) frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.TextBox) txtFiadbInputFile, "Text", false);
                        System.Threading.Thread.Sleep(1000);
                    }

                    //An ADO error occurred, so delete all the incompletely processed new records
                    else
                    {
                        DeleteFromTablesWhereFilter(m_ado, m_connTempMDBFile, new string[]
                            {
                                m_strPlotTable, m_strCondTable, m_strTreeTable, m_strPopEvalTable, m_strPopStratumTable,
                                m_strPpsaTable, m_strPopEstUnitTable, m_strSiteTreeTable,
                                m_strBiosumPopStratumAdjustmentFactorsTable,
                                //m_strDwmCwdTable, m_strDwmFwdTable, m_strDwmDuffLitterTable,
                                //m_strDwmTransectSegmentTable
                                m_strDwmCwdTable + "_master", m_strDwmFwdTable + "_master", m_strDwmDuffLitterTable + "_master",
                                m_strDwmTransectSegmentTable + "_master"
                            },
                            " WHERE biosum_status_cd=9;");
                        MessageBox.Show("!!Error Occured Adding Plot Records: 0 Records Added!!", "FIA Biosum",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                }

                this.m_ado.m_DataSet.Clear();
                this.m_ado.m_DataSet.Dispose();
                this.m_ado = null;

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);

                LoadMDBPlotCondTreeData_Finish();
            }
            catch (System.Threading.ThreadInterruptedException err)
			{
				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
                if (this.m_connTempMDBFile != null)
                {
                    if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(m_connTempMDBFile);
                    }
                    m_connTempMDBFile = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.CancelThreadCleanup();
			    this.ThreadCleanUp();
				this.CleanupThread();
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_plot_input:mdbFIADBFileInput  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			finally
			{
               
			}

            if (this.m_connTempMDBFile != null)
            {
                if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(m_connTempMDBFile);
                }
                m_connTempMDBFile = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
			if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"Visible",false);
            LoadMDBPlotCondTreeData_Finish();
            CleanupThread();
            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
		}

	    private string CreateRecordsAddedMessage()
	    {
	        string strMessage = "Successfully Appended\n" +
	                            m_intAddedPlotRows.ToString().Trim() + " Plots\n" +
	                            m_intAddedCondRows.ToString().Trim() + " Conditions \n" +
	                            m_intAddedTreeRows.ToString().Trim() + " Trees\n" +
                                m_intAddedSeedlingRows.ToString().Trim() + " Seedlings\n" +
                                m_intAddedSiteTreeRows.ToString().Trim() + " Site Trees";

	        if (Checked(chkDwmImport))
	        {
	            strMessage += "\r\n" +
	                          m_intAddedDwmTransectSegmentRows.ToString().Trim() + " Transect Segments\r\n" +
	                          m_intAddedDwmCwdRows.ToString().Trim() + " CWD Pieces\r\n" +
	                          m_intAddedDwmFwdRows.ToString().Trim() + " FWD Transects\r\n" +
	                          m_intAddedDwmDuffLitterRows.ToString().Trim() + " Duff and Litter Samples";
	        }

	        return strMessage;
	    }

	    /// <summary>
        /// Return the record count of the table if it exists 
        /// </summary>
        /// <param name="p_ado"></param>
        /// <param name="p_conn"></param>
        /// <param name="table"></param>
        /// <returns></returns>
	    private int GetNewRecordCount(ado_data_access p_ado, OleDbConnection p_conn, string table)
        {
            if (p_ado.TableExist(p_conn, table) || table.Equals(m_strSeedlingTable))
            {
                int intReturn = -1;
                switch (table)
                {
                    case "fiadb_seedling_input":
                        intReturn = (int)p_ado.getRecordCount(p_conn,
                            String.Concat("select count(*) from ", m_strTreeTable, " where biosum_status_cd=9 and dia=0.1;"), m_strTreeTable);
                        break;
                    case "tree":
                        intReturn = (int)p_ado.getRecordCount(p_conn,
                            String.Concat("select count(*) from ", table, " where biosum_status_cd=9 and dia<>0.1;"), table);
                        break;
                    default:
                        intReturn = (int)p_ado.getRecordCount(p_conn,
                    String.Concat("select count(*) from ", table, " where biosum_status_cd=9;"), table);
                        break;
                }
                return intReturn;
            }
            else
            {
                return 0;
            }
        }

        private int GetNewRecordCountSqlite(ado_data_access p_ado, OleDbConnection p_conn, string table)
        {
            if (p_ado.TableExist(p_conn, table) || table.Equals(m_strSeedlingTable))
            {
                int intReturn = -1;
                switch (table)
                {
                    case "fiadb_seedling_input":
                        intReturn = (int)p_ado.getRecordCount(p_conn,
                            String.Concat("select count(*) from ", m_strTreeTable, " where biosum_status_cd=9 and dia=0.1;"), m_strTreeTable);
                        break;
                    case "tree":
                        intReturn = (int)p_ado.getRecordCount(p_conn,
                            String.Concat("select count(*) from ", table, " where biosum_status_cd=9 and dia<>0.1;"), table);
                        break;
                    default:
                        intReturn = (int)p_ado.getRecordCount(p_conn,
                    String.Concat("select count(*) from ", table, " where biosum_status_cd='9';"), table);
                        break;
                }
                return intReturn;
            }
            else
            {
                return 0;
            }
        }

	    /// <summary>
	    /// Delete records from strTableNames in the database p_ado is currently connected to using the strDeleteFilter.
	    /// </summary>
	    /// <param name="strTableName">An array of table names to delete</param>
	    /// <param name="strDeleteFilter">A WHERE clause for the delete query</param>
	    private void DeleteFromTablesWhereFilter(ado_data_access p_ado, OleDbConnection conn, string[] strTableNames, string strDeleteFilter)
	    {
	        foreach (string table in strTableNames)
	        {
	            if (p_ado.TableExist(conn, table))
	            {
	                p_ado.m_strSQL = "DELETE FROM " + table + strDeleteFilter;
	                p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
	                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
	                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
	                        p_ado.m_strSQL + "\r\n");
	            }
	        }
	    }

        private void DeleteFromTablesWhereFilterSqlite(System.Data.SQLite.SQLiteConnection conn, string[] strTableNames, string strDeleteFilter)
        {
            foreach (string table in strTableNames)
            {
                if (SQLite.TableExist(conn, table))
                {
                    SQLite.m_strSQL = "DELETE FROM " + table + strDeleteFilter;
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            SQLite.m_strSQL + "\r\n");
                }
            }
        }

	    /// <summary>
	    /// After importing DWM table data from FIADB to Master.mdb, join the DWM table with the plot/cond tables to get the BSCID
	    /// </summary>
	    /// <param name="p_ado">Reference to the temporary database used to link FIADB and Master during plot input phase</param>
	    /// <param name="strDestTable">The DWM table to update</param>
	    private void UpdateDwmBiosumCondIds(ado_data_access p_ado, string strDestTable = null)
	    {
	        p_ado.m_strSQL = String.Format(
	            "UPDATE {0} t INNER JOIN ({1} p INNER JOIN {2} c ON c.biosum_plot_id = p.biosum_plot_id) ON (p.cn = t.plt_cn) AND (t.CONDID = c.condid) " +
	            "SET t.biosum_cond_id=c.biosum_cond_id WHERE t.biosum_cond_id='9999999999999999999999999' AND p.biosum_status_cd=9;",
	            strDestTable, m_strPlotTable, m_strCondTable);
	        p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
	        m_intError = p_ado.m_intError;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
	    }

        private void UpdateDwmBiosumCondIdsSqlite(ado_data_access p_ado, string strDestTable = null)
        {
            //p_ado.m_strSQL = String.Format(
            //    "SELECT c.biosum_cond_id, p.biosum_status_cd INTO temptable FROM {0} t INNER JOIN ({1} p INNER JOIN {2} c ON c.biosum_plot_id = p.biosum_plot_id) ON (TRIM(p.cn) = TRIM(t.PLT_CN)) AND (TRIM(t.CONDID) = TRIM(c.condid))",
            //    strDestTable + "_master", m_strPlotTable, m_strCondTable);
            //p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
            //p_ado.m_strSQL = String.Format(
            //    "UPDATE  {0} a, temptable b SET a.biosum_cond_id = b.biosum_cond_id WHERE a.biosum_cond_id = '9999999999999999999999999' AND b.biosum_status_cd = 9",
            //    strDestTable + "_master");
            //p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
            //p_ado.m_strSQL = "DROP TABLE temptable";
            //p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
            p_ado.m_strSQL = String.Format(
            "UPDATE {0} t INNER JOIN ({1} p INNER JOIN {2} c ON c.biosum_plot_id = p.biosum_plot_id) ON (TRIM(p.cn) = TRIM(t.plt_cn)) AND (TRIM(t.CONDID) = TRIM(c.condid)) " +
            "SET t.biosum_cond_id=c.biosum_cond_id WHERE t.biosum_cond_id='9999999999999999999999999' AND p.biosum_status_cd=9;",
               strDestTable + "_master", m_strPlotTable, m_strCondTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
            p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
            m_intError = p_ado.m_intError;
        }
        
	    /// <summary>
	    /// Creates temporary DWM table from FIADB that associates dwm data with a biosum_cond_id. 
	    /// Assumes that the FIADB DWM table has a plt_cn and condid.
	    /// </summary>
	    /// <param name="strSourceTable"></param>
	    /// <param name="strTempTable"></param>
	    private void SelectIntoTempDWMTableFromSourceDWMTable(ado_data_access p_ado, string strSourceTable = null, string strTempTable = null)
	    {
	        p_ado.m_strSQL = String.Format(
	            "SELECT TRIM(c.biosum_cond_id) AS biosum_cond_id, 9 as biosum_status_cd, t.* INTO {0} " +
	            "FROM {1} t INNER JOIN ({2} p INNER JOIN {3} c ON c.biosum_plot_id = p.biosum_plot_id) ON (p.cn = t.plt_cn) AND (t.CONDID = c.condid) " +
	            "WHERE p.biosum_status_cd=9;", strTempTable, strSourceTable, m_strPlotTable, m_strCondTable);
	        p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
	        m_intError = p_ado.m_intError;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
	    }

        /// <summary>
        /// Generates and executes an SQL "INSERT INTO dest_tablename (columns) SELECT columns FROM source_tablename" statement
        /// strInsertFields contains a subset of the source table columns that are common with the dest table.
        /// </summary>
        /// <param name="strSourceTable"></param>
        /// <param name="strDestTable"></param>
        /// <param name="strInsertFields"></param>
        private void InsertIntoDestTableFromSourceTable(ado_data_access p_ado, 
            string strSourceTable = null, string strDestTable = null,
            string strSourceFields = null, string strDestFields = null,
            bool InsertBiosumCondIdAndStatusCode = false)
        {
            string strInsertIntoValues = "INSERT INTO {0} ({1}) ";
            string strSelectColumns = "SELECT {2} FROM {3} dwm INNER JOIN {4} p ON dwm.plt_cn=p.cn;";
            if (InsertBiosumCondIdAndStatusCode)
            {
                strInsertIntoValues = "INSERT INTO {0} (biosum_cond_id, biosum_status_cd, {1}) ";
                strSelectColumns =
                    "SELECT '9999999999999999999999999' AS biosum_cond_id, 9 AS biosum_status_cd, {2} FROM {3} dwm INNER JOIN {4} p ON dwm.plt_cn=trim(p.cn);";
            }
            p_ado.m_strSQL = String.Format(strInsertIntoValues + strSelectColumns, 
                strDestTable, strDestFields, strSourceFields, strSourceTable, "tempplot");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
            p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
            m_intError = p_ado.m_intError;
        }

        private void InsertIntoDestTableFromSourceTableSqlite (ado_data_access p_ado,
            string strSourceTable = null, string strDestTable = null,
            string strSourceFields = null, string strDestFields = null,
            bool InsertBiosumCondIdAndStatusCode = false)
        {
            string strInsertIntoValues = "INSERT INTO {0} ({1}) ";
            string strSelectColumns = "SELECT {2} FROM {3} dwm INNER JOIN {4} p ON dwm.plt_cn=p.cn;";
            if (InsertBiosumCondIdAndStatusCode)
            {
                strInsertIntoValues = "INSERT INTO {0} (biosum_cond_id, biosum_status_cd, {1}) ";
                strSelectColumns =
                    "SELECT '9999999999999999999999999' AS biosum_cond_id, 9 AS biosum_status_cd, {2} FROM {3} dwm INNER JOIN {4} p ON dwm.plt_cn=trim(p.cn);";
            }
            p_ado.m_strSQL = String.Format(strInsertIntoValues + strSelectColumns,
                strDestTable + "_master", strDestFields, strSourceFields, strSourceTable + "_fiadb", "tempplot");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
            p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);

            m_intError = p_ado.m_intError;
        }

	    private string CreateStrFieldsFromDataTables(DataTable dtSourceSchema=null, DataTable dtDestSchema=null, string strTablePrefix="")
	    {
	        string strFields;
	        int x;
	        string strCol;
	        int y;
	        strFields = "";
	        for (x = 0; x <= dtDestSchema.Rows.Count - 1; x++)
	        {
	            strCol = dtDestSchema.Rows[x]["columnname"].ToString().Trim();
	            //see if there is an equivalent FIADB column
	            for (y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
	            {
	                if (strCol.Trim().ToUpper() == dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper())
	                {
                        if (strCol=="VALUE")
                        {
                            strCol = strTablePrefix + "`" + strCol + "`";
                        }
	                    if (strFields.Trim().Length == 0)
	                    {
	                        strFields = strTablePrefix + strCol;
	                    }
	                    else
	                    {
	                        strFields += "," + strTablePrefix + strCol;
	                    }
	                    break;
	                }
	            }
	        }
	        return strFields;
	    }

   
	    private int UpdateColumns(FIA_Biosum_Manager.ado_data_access p_ado)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.UpdateColumns\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

			//create work tables


            string strColumns = "";
            string strValues = "";
			string strTime = System.DateTime.Now.ToString();
            string strConn = "";
				
			//----------------------COND COLUMN UPDATES-----------------------//
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//----------------------COND AND TREE COLUMN UPDATES-----------------------//\r\n");
               
                SetThermValue(m_frmTherm.progressBar1,"Maximum",42);
                SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                
                SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Condition Proportion Column...Stand By");
			    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
					//update the condition proportion column
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
									 "INNER JOIN ((" + this.m_strPpsaTable + " ppsa " + 
												 "INNER JOIN " + this.m_strPlotTable + " p " + 
												 "ON ppsa.plt_cn= p.cn) " + 
												 "INNER JOIN " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " ps " + 
												 "ON ppsa.stratum_cn=ps.stratum_cn) " + 
									 "ON c.biosum_plot_id = p.biosum_plot_id " + 
									 "SET condprop = IIf(ps.pmh_macr Is Not Null And ps.pmh_macr>0," + 
														 "c.condprop_unadj/ps.pmh_macr," + 
													"IIf(ps.pmh_sub Is Not Null And ps.pmh_sub>0," + 
														 "c.condprop_unadj/ps.pmh_sub," + 
													"IIf(ps.pmh_micr Is Not Null And ps.pmh_micr>0," + 
														 "c.condprop_unadj/ps.pmh_micr,0)))";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Condition Acres Column...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

					//update acres column
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
									 "INNER JOIN ((" + this.m_strPpsaTable + " ppsa " + 
												 "INNER JOIN " + this.m_strPlotTable + " p " + 
												 "ON ppsa.plt_cn= p.cn) " +
                                                 "INNER JOIN " + this.m_strBiosumPopStratumAdjustmentFactorsTable + " ps " + 
												 "ON ppsa.stratum_cn=ps.stratum_cn) " + 
									 "ON c.biosum_plot_id = p.biosum_plot_id " + 
									 "SET acres = IIF( c.condprop IS NOT NULL and ps.expns IS NOT NULL," + 
													"c.condprop * ps.expns,0)";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                    //update condprop_specific column for when plot.macro_breakpoint_dia has a value
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Updating Tree Condprop Specific Column...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
                    
                    p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                        "INNER JOIN (" + m_strCondTable + " c INNER JOIN " + m_strPlotTable + " p ON c.biosum_plot_id=p.biosum_plot_id) " +
                                        "ON t.biosum_cond_id = c.biosum_cond_id " +
                                        "SET t.condprop_specific = " +
                                        "IIF(c.micrprop_unadj IS NOT NULL AND t.dia < 5," +
                                            "c.micrprop_unadj," +
                                        "IIF(c.subpprop_unadj IS NOT NULL AND " +
                                            "p.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                            "t.dia >= 5 AND " +
                                            "t.dia < p.MACRO_BREAKPOINT_DIA," +
                                            "c.subpprop_unadj," +
                                        "IIF(c.macrprop_unadj IS NOT NULL AND " +
                                            "p.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                            "t.dia >= p.MACRO_BREAKPOINT_DIA," +
                                            "c.macrprop_unadj))) " + 
                                        "WHERE t.biosum_status_cd=9";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
                    p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);
                    //check if null values 
                    p_ado.m_strSQL = "SELECT COUNT(*) AS ROWCOUNT " +
                                     "FROM " + m_strTreeTable + " t," +
                                               m_strTreeMacroPlotBreakPointDiaTable + " bp " +
                                     "WHERE t.biosum_status_cd=9 AND " +
                                           "t.condprop_specific IS NULL AND " +
                                           "t.statecd = bp.statecd AND " +
                                           "t.unitcd = bp.unitcd";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
                    //check if condprop_specific null and exists in the tree macro plot breakpoint diameter table
                    if ((double)p_ado.getSingleDoubleValueFromSQLQuery(m_connTempMDBFile, p_ado.m_strSQL, "temp") > 0)
                    {
                        //got some nulls
                        p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                         "INNER JOIN ((" + m_strCondTable + " c " +
                                            "INNER JOIN " + m_strPlotTable + " p ON c.biosum_plot_id=p.biosum_plot_id) " +
                                            "INNER JOIN " + m_strTreeMacroPlotBreakPointDiaTable + " bp ON p.statecd = bp.statecd AND p.unitcd=bp.unitcd) " +
                                         "ON t.biosum_cond_id = c.biosum_cond_id " +
                                         "SET t.condprop_specific = " +
                                         "IIF(c.micrprop_unadj IS NOT NULL AND t.dia < 5," +
                                                "c.micrprop_unadj," +
                                         "IIF(c.subpprop_unadj IS NOT NULL AND " +
                                             "bp.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                             "t.dia >= 5 AND t.dia < bp.MACRO_BREAKPOINT_DIA," +
                                                "c.subpprop_unadj," +
                                         "IIF(c.macrprop_unadj IS NOT NULL AND " +
                                             "bp.MACRO_BREAKPOINT_DIA IS NOT NULL AND " +
                                             "t.dia >= bp.MACRO_BREAKPOINT_DIA," +
                                                "c.macrprop_unadj))) " +
                                         "WHERE t.biosum_status_cd=9";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);
                    }
                    else
                    {
                        //see if we have nulls and no MACRO PLOT for the unit code
                         p_ado.m_strSQL = "SELECT COUNT(*) AS ROWCOUNT FROM tree a," + 
                                              "(SELECT t.* FROM tree t " + 
                                               "WHERE NOT EXISTS " + 
                                                   "(SELECT  * FROM TreeMacroPlotBreakPointDia bp " + 
                                                    "WHERE t.statecd=bp.statecd AND t.unitcd=bp.unitcd)) b " + 
                                         "WHERE a.CN=b.CN AND a.condprop_specific IS NULL AND a.biosum_status_cd=9";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
                        //handle for those states and units that do not have macro plot
                        if ((double)p_ado.getSingleDoubleValueFromSQLQuery(m_connTempMDBFile, p_ado.m_strSQL, "temp") > 0)
                        {
                            p_ado.m_strSQL = "UPDATE " + m_strTreeTable + " t " +
                                             "INNER JOIN (" + m_strCondTable + " c " +
                                             "INNER JOIN " + m_strPlotTable + " p ON c.biosum_plot_id=p.biosum_plot_id) " +
                                             "ON t.biosum_cond_id = c.biosum_cond_id AND t.condid = c.condid " +
                                             "SET t.condprop_specific = " +
                                             "IIF(c.micrprop_unadj IS NOT NULL AND t.dia < 5," +
                                                "c.micrprop_unadj," +
                                             "IIF(c.subpprop_unadj IS NOT NULL AND t.dia >= 5," +
                                                "c.subpprop_unadj)) " +
                                             "WHERE t.biosum_status_cd=9";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
                            p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);

                        }

                    }

            		//Update fvs_tree_id column for tracking a tree between BioSum and FVS for lifetime of project
            		SetLabelValue(m_frmTherm.lblMsg, "Text", "Updating Tree fvs_tree_id Column...Stand By");
            		frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control) this.m_frmTherm, "Refresh");
            		p_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " SET fvs_tree_id = CStr(subp*1000+tree) where dia > 0.1;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
                    this.m_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);

                    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Tree tpacurr Column...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
					//update tree tpacurr column
                    p_ado.m_strSQL = "UPDATE " + this.m_strTreeTable + " t " +
                        "SET tpacurr = IIF(t.tpa_unadj IS NOT NULL AND t.condprop_specific IS NOT NULL," +
                                       "t.tpa_unadj / t.condprop_specific,0)";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, m_ado.m_strSQL + "\r\n");

					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                    //
                    //drybiot,drybiom,voltsgrs processing
                    //                   
                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Start Volume and Biomass Calculations...Stand By");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

                        //step 5 - delete and create work table
                        if (p_ado.TableExist(this.m_connTempMDBFile, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable))
                            p_ado.SqlNonQuery(this.m_connTempMDBFile, "DROP TABLE " + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                        frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(p_ado, this.m_connTempMDBFile, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);


                        var treeToFcsBiosumVolumesInputTable = new List<Tuple<string, string>>
                        {
                            Tuple.Create("ACTUALHT", "ACTUALHT"),
                            // ("BALIVE", "BALIVE"), //cond table
                            Tuple.Create("BFSND", "BFSND"),
                            Tuple.Create("BOLEHT", "BOLEHT"),
                            Tuple.Create("CENTROID_DIA", "CENTROID_DIA"),
                            Tuple.Create("CENTROID_DIA_HT_ACTUAL", "CENTROID_DIA_HT_ACTUAL"),
                            Tuple.Create("CFSND", "CFSND"),
                            Tuple.Create("CND_CN", "BIOSUM_COND_ID AS CND_CN"),
                            Tuple.Create("COUNTYCD", "COUNTYCD"),
                            Tuple.Create("CR", "CR"),
                            Tuple.Create("CULL", "CULL"),
                            Tuple.Create("CULLBF", "CULLBF"),
                            Tuple.Create("CULLCF", "CULLCF"),
                            Tuple.Create("CULLDEAD", "CULLDEAD"),
                            Tuple.Create("CULLFORM", "CULLFORM"),
                            Tuple.Create("CULLMSTOP", "CULLMSTOP"),
                            Tuple.Create("CULL_FLD", "CULL_FLD"),
                            // ("DECAYCD", "null as DECAYCD"), // TODO: Update queries following this will set it.
                            Tuple.Create("DIA", "IIF(DIA IS NOT NULL, ROUND(DIA, 2), DIA)"),
                            Tuple.Create("DIAHTCD", "DIAHTCD"),
                            Tuple.Create("FORMCL", "FORMCL"),
                            Tuple.Create("HT", "HT"),
                            Tuple.Create("HTDMP", "HTDMP"),
                            Tuple.Create("INVYR", "INVYR"),
                            Tuple.Create("PLOT", "CLNG(MID(BIOSUM_COND_ID, 16, 6)) AS PLOT"),
                            Tuple.Create("PLT_CN", "MID(BIOSUM_COND_ID, 1, LEN(BIOSUM_COND_ID)-1) AS PLT_CN"),
                            // ("PRECIPITATION", "PRECIPITATION"), //plot table
                            Tuple.Create("ROUGHCULL", "ROUGHCULL"),
                            Tuple.Create("SAWHT", "SAWHT"),
                            Tuple.Create("SITREE", "SITREE"),
                            Tuple.Create("SPCD", "SPCD"),
                            Tuple.Create("STANDING_DEAD_CD", "STANDING_DEAD_CD"),
                            Tuple.Create("STATECD", "STATECD"),
                            Tuple.Create("STATUSCD", "STATUSCD"),
                            Tuple.Create("SUBP", "SUBP"),
                            Tuple.Create("TOTAGE", "TOTAGE"),
                            Tuple.Create("TREE", "TREE"),
                            Tuple.Create("TREECLCD", "TREECLCD"),
                            Tuple.Create("TRE_CN", "CN AS TRE_CN"),
                            Tuple.Create("UPPER_DIA", "UPPER_DIA"),
                            Tuple.Create("UPPER_DIA_HT", "UPPER_DIA_HT"),
                            Tuple.Create("VOL_LOC_GRP", "'' AS VOL_LOC_GRP"),
                            Tuple.Create("WDLDSTEM", "WDLDSTEM"),
                        };

                        strColumns = string.Join(",", treeToFcsBiosumVolumesInputTable.Select(e => e.Item1));
                        strValues = string.Join(",", treeToFcsBiosumVolumesInputTable.Select(e => e.Item2));

                        //insert records
                        p_ado.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step1(Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable, m_strTreeTable, strColumns, strValues);
                        
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);

                        p_ado.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step2(Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable, m_strTreeTable,m_strPlotTable,m_strCondTable);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);


                        p_ado.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step3(Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable, m_strCondTable);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);

                        //populate treeclcd column
                        if (p_ado.TableExist(m_connTempMDBFile, "CULL_TOTAL_WORK_TABLE"))
                            p_ado.SqlNonQuery(m_connTempMDBFile, "DROP TABLE CULL_TOTAL_WORK_TABLE");

                        p_ado.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step4(
                                          "cull_total_work_table",
                                          Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                       
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);


                        p_ado.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                            "cull_total_work_table", Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);


                        p_ado.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                                       "cull_total_work_table", Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                        p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);

                            frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesWorkTable(p_ado, m_connTempMDBFile, Tables.VolumeAndBiomass.SqliteWorkTable);
                            p_ado.m_strSQL = $@"INSERT INTO {Tables.VolumeAndBiomass.SqliteWorkTable} 
                                                SELECT * FROM {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable}";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n\r\n");
                            p_ado.SqlNonQuery(m_connTempMDBFile, p_ado.m_strSQL);
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
                                int COUNT = 0;
                                SQLite.ADO.DataMgr oSQLite = new SQLite.ADO.DataMgr();

                                SQLite = oSQLite;
                                MSAccess = p_ado;
                                var columnsAndDataTypes = Tables.VolumeAndBiomass.ColumnsAndDataTypes;
                                strColumns = string.Join(",", columnsAndDataTypes.Select(item => item.Item1));

                                //
                                //CONNECT TO SQLITE AND REMOVE DATA FROM SQLITE DB
                                //
                                oSQLite.OpenConnection(false, 1,
                                    frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" +
                                    Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase, "BIOSUM");
                                oSQLite.SqlNonQuery(oSQLite.m_Connection, $"DELETE FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable}");
                                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                                System.Threading.Thread.Sleep(2000);

                                //
                                //INSERT ALL THE MS ACCESS DATA INTO SQLITE
                                //
                                int intTotalRecs = (int) p_ado.getRecordCount(m_connTempMDBFile, $"SELECT COUNT(*) AS ROWCOUNT FROM {Tables.VolumeAndBiomass.SqliteWorkTable}", Tables.VolumeAndBiomass.SqliteWorkTable);

                                p_ado.SqlQueryReader(m_connTempMDBFile, $"SELECT * FROM {Tables.VolumeAndBiomass.SqliteWorkTable}");
                                if (p_ado.m_OleDbDataReader.HasRows)
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
                                        SetThermValue(m_frmTherm.progressBar1, "Maximum", intTotalRecs);

                                        while (MSAccess.m_OleDbDataReader.Read())
                                        {
                                            COUNT++;
                                            strValues = utils.GetParsedInsertValues(MSAccess.m_OleDbDataReader, columnsAndDataTypes);
                                            command.CommandText = $"INSERT INTO {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} ({strColumns}) VALUES ({strValues})";
                                            command.ExecuteNonQuery();
                                            //frmMain.g_oDelegate.SetControlPropertyValue((Control)lblSQLite2Msg, "Text", "INSERT DATA: " + COUNT.ToString() + " of " + intTotalCount.ToString());
                                            frmMain.g_oDelegate.SetControlPropertyValue(
                                                (System.Windows.Forms.Control) m_frmTherm.lblMsg, "Text",
                                                "Prepare Tree Data for Volume and Biomass Calculations...Stand By [" +
                                                COUNT.ToString() + "/" + intTotalRecs.ToString() + "]");
                                            frmMain.g_oDelegate.ExecuteControlMethod(
                                                (System.Windows.Forms.Control) this.m_frmTherm.lblMsg, "Refresh");
                                            if (COUNT <= intTotalRecs)
                                                SetThermValue(m_frmTherm.progressBar1, "Value", COUNT);
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

                                strConn = m_connTempMDBFile.ConnectionString;
                                MSAccess.m_OleDbDataReader.Close();
                                MSAccess.m_OleDbDataReader.Dispose();
                                MSAccess.CloseConnection(m_connTempMDBFile);
                                m_connTempMDBFile.Dispose();
                                //oAdo = null;
                                SetThermValue(m_frmTherm.progressBar1, "Value", 1);
                                //
                                //RUN JAVA APP TO CALCULATE VOLUME/BIOMASS
                                //
                                if (m_intError == 0)
                                {
                                    oSQLite.CloseAndDisposeConnection(oSQLite.m_Connection, true);
                                    frmMain.g_oUtils.RunProcess(
                                        frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum", "fcs_tree_calc.bat",
                                        "BAT");
                                    if (System.IO.File.Exists(frmMain.g_oEnv.strApplicationDataDirectory +
                                                              "\\FIABiosum\\fcs_error_msg.txt"))
                                    {
                                        // Read entire text file content in one string  
                                        m_strError = System.IO.File.ReadAllText(
                                            frmMain.g_oEnv.strApplicationDataDirectory +
                                            "\\FIABiosum\\fcs_error_msg.txt");
                                        if (m_strError.IndexOf("JAVA.EXE", 0) > 0)
                                            m_strError = "Problem detected running JAVA.EXE";
                                        m_intError = -2;
                                    }
                                }

                                SetThermValue(m_frmTherm.progressBar1, "Value", 2);
                                //
                                //UPDATE MSACCESS WITH CALCULATED VALUES
                                //
                                if (m_intError == 0)
                                {
                                    //oAdo = new ado_data_access();

                                    //MSAccess = p_ado;

                                    SQLite.OpenConnection(false, 1,
                                        frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" +
                                        Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase, "BIOSUM");

                                    //COUNT = Convert.ToInt32(_oSQLite.getSingleDoubleValueFromSQLQuery(_oSQLite.m_Connection,"SELECT COUNT(*) AS ROWCOUNT FROM BIOSUM_CALC","biosum_calc"));

                                    MSAccess.OpenConnection(strConn, ref m_connTempMDBFile);

                                    if (MSAccess.TableExist(m_connTempMDBFile,
                                        Tables.VolumeAndBiomass.BiosumCalcOutputTable))
                                        MSAccess.SqlNonQuery(m_connTempMDBFile,
                                            $"DROP TABLE {Tables.VolumeAndBiomass.BiosumCalcOutputTable}");

                                    System.Threading.Thread.Sleep(3000);

                                    MSAccess.SqlNonQuery(m_connTempMDBFile,
                                        $"SELECT * INTO {Tables.VolumeAndBiomass.BiosumCalcOutputTable} FROM {Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable} WHERE 1=2");


                                    //MSAccessBeginTransaction("BIOSUM_VOLUME_INPUT", "TRE_CN,VOLCSGRS_CALC,VOLCFGRS_CALC,VOLCFNET_CALC,DRYBIOT_CALC,DRYBIOM_CALC,VOLTSGRS_CALC", "TRE_CN",COUNT , "");

                                    intTotalRecs = Convert.ToInt32(SQLite.getSingleDoubleValueFromSQLQuery(SQLite.m_Connection, 
                                        $"SELECT COUNT(*) AS ROWCOUNT FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} WHERE VOLTSGRS_CALC IS NOT NULL", 
                                        Tables.VolumeAndBiomass.BiosumVolumeCalcTable));

                                    SetThermValue(m_frmTherm.progressBar1, "Maximum", intTotalRecs);
                                    if (intTotalRecs < 50)
                                    {
                                        SetThermValue(m_frmTherm.progressBar1, "Maximum", 50);
                                    }

                                    oSQLite.SqlQueryReader(oSQLite.m_Connection,
                                        $"SELECT * FROM {Tables.VolumeAndBiomass.BiosumVolumeCalcTable} WHERE VOLTSGRS_CALC IS NOT NULL");

                                    if (oSQLite.m_DataReader.HasRows)
                                    {
                                        System.Data.OleDb.OleDbTransaction transaction;
                                        System.Data.OleDb.OleDbCommand command = m_connTempMDBFile.CreateCommand();
                                        // Start a local transaction
                                        transaction = m_connTempMDBFile.BeginTransaction(IsolationLevel.ReadCommitted);
                                        // Assign transaction object for a pending local transaction
                                        command.Transaction = transaction;
                                        try
                                        {
                                            COUNT = 0;
                                            SetThermValue(m_frmTherm.progressBar1, "Value", COUNT);

                                            while (oSQLite.m_DataReader.Read())
                                            {
                                                COUNT++;
                                                if (oSQLite.m_DataReader["TRE_CN"] != DBNull.Value &&
                                                    Convert.ToString(oSQLite.m_DataReader["TRE_CN"]).Trim().Length > 0)
                                                {
                                                    strValues = utils.GetParsedInsertValues(oSQLite.m_DataReader, columnsAndDataTypes);
                                                    command.CommandText = $"INSERT INTO {Tables.VolumeAndBiomass.BiosumCalcOutputTable} ({strColumns}) VALUES ({strValues})";
                                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 3)
                                                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, command.CommandText + "\r\n\r\n");
                                                    command.ExecuteNonQuery();
                                                    SetThermValue(m_frmTherm.progressBar1, "Value", COUNT);
                                                }

                                                frmMain.g_oDelegate.SetControlPropertyValue(
                                                    (System.Windows.Forms.Control) m_frmTherm.lblMsg, "Text",
                                                    "Update MSAccess with SQLite Results..Stand By [" +
                                                    COUNT.ToString() + "/" + intTotalRecs.ToString() + "]");
                                                frmMain.g_oDelegate.ExecuteControlMethod(
                                                    (System.Windows.Forms.Control) this.m_frmTherm.lblMsg, "Refresh");
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
                                            //update VOLTSGRS
                                            if (m_intError == 0)
                                            {
                                                MSAccess.m_strSQL =
                                                    Queries.VolumeAndBiomass.FIAPlotInput.WriteCalculatedVolumeAndBiomassColumnsToTreeTable(Tables.VolumeAndBiomass.BiosumCalcOutputTable);
                                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, MSAccess.m_strSQL + "\r\n\r\n");
                                                    MSAccess.SqlNonQuery(m_connTempMDBFile, MSAccess.m_strSQL);
                                            }
                                        }

                                        transaction.Dispose();
                                        transaction = null;
                                    }

                                    //MSAccess.CloseConnection(MSAccess.m_OleDbConnection);
                                    //MSAccess.m_OleDbConnection.Dispose();
                                    //oAdo = null;
                                    oSQLite.CloseAndDisposeConnection(oSQLite.m_Connection, true);
                                    oSQLite = null;
                                }

                                SetThermValue(m_frmTherm.progressBar1, "Value", m_frmTherm.progressBar1.Maximum);
                            }
  

                    
                

                SetLabelValue(m_frmTherm.lblMsg,"Text", "Updating Condition Table Columns...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

				
                SetThermValue(m_frmTherm.progressBar1, "Value", 6);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//tpa column
					//sum trees per acre on a condition 
					//for live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,tpacurr) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tottpa as tpa  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(tpacurr) as tottpa " + 
						"FROM " + this.m_strTreeTable + " " +
                        //issue_253: include saplings in tpacurr
                        //"WHERE  dia >= 5 AND statuscd=1 " + 
                        "WHERE statuscd=1 " +
                        "GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 7);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.TPACURR = u.TPACURR;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 8);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//swd_tpa
					//sum trees per acre on a condition 
					//for softwood live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_tpacurr) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.totswdtpa as swd_tpa  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(tpacurr) as totswdtpa " + 
						"FROM " + this.m_strTreeTable + " " +
                        //issue_253: include saplings in tpacurr  
                        //"WHERE spcd < 300 AND dia >= 5 AND statuscd=1 " +
                        "WHERE spcd < 300 AND statuscd = 1 " +
                        "GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 9);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_TPACURR = u.SWD_TPACURR;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//hwd tpa
					//sum trees per acre on a condition 
					//for hardwood live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_tpacurr) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tothwdtpa as hwd_tpacurr  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(tpacurr) as tothwdtpa " + 
						"FROM " + this.m_strTreeTable + " " +
                        //issue_253: include saplings in tpacurr  
                        //"WHERE spcd > 299 AND dia >= 5 AND statuscd=1 " + 
                        "WHERE spcd > 299 AND statuscd=1 " +
                        "GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 11);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_TPACURR = u.HWD_TPACURR;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 12);
                

                //vol_ac_grs_ft3
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//total
					//for all live trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,vol_ac_grs_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tot_volgrsft3 as vol_ac_grs_ft3  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) as tot_volgrsft3 " + 
						"FROM " + this.m_strTreeTable + " WHERE volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd=1 AND dia >=5 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 13);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.VOL_AC_GRS_FT3 = u.VOL_AC_GRS_FT3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 14);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//hwd
					//for all live hardwood trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_vol_ac_grs_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tot_volgrsft3 as hwd_vol_ac_grs_ft3  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) as tot_volgrsft3 " + 
					 	 "FROM " + this.m_strTreeTable + " " + 
                         "WHERE spcd > 299 AND volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd=1 AND dia >=5 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 15);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_VOL_AC_GRS_FT3 = u.HWD_VOL_AC_GRS_FT3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 16);



                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//SWD
					//for all live softwood trees >= 5 inches in diameter 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_vol_ac_grs_ft3) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.tot_volgrsft3 as swd_vol_ac_grs_ft3  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) as tot_volgrsft3 " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE spcd < 300 AND volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd=1 AND dia >=5 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 17);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_VOL_AC_GRS_FT3 = u.SWD_VOL_AC_GRS_FT3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value",18);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					//ba_ft2_ac basal area column
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,ba_ft2_ac) " + 
						"SELECT a.biosum_cond_id, b.tottemp AS ba_ft2_ac " + 
						"FROM " + this.m_strCondTable + " a, " +  
						"(SELECT biosum_cond_id, SUM((.005454154 * dia^2)  * tpacurr)  AS tottemp " + 
						 "FROM " + this.m_strTreeTable + " " + 
						 "WHERE biosum_status_cd=9  AND statuscd=1 " + 
						 "GROUP BY biosum_cond_id)  b " + 
						"WHERE a.biosum_cond_id = b.biosum_cond_id ";


					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 19);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.BA_FT2_AC = u.BA_FT2_AC;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 20);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					//swd_ba_ft2_ac softwood basal area 
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_ba_ft2_ac) " + 
						"SELECT a.biosum_cond_id, b.tottemp AS swd_ba_ft2_ac " + 
						"FROM " + this.m_strCondTable + " a, " +  
						"(SELECT biosum_cond_id, SUM((.005454154 * dia^2)  * tpacurr)  AS tottemp " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE biosum_status_cd=9 AND " + 
						      "spcd < 300 AND statuscd=1 " + 
						"GROUP BY biosum_cond_id)  b " + 
						"WHERE a.biosum_cond_id = b.biosum_cond_id ";


					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 21);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_BA_FT2_AC = u.SWD_BA_FT2_AC";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 22);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					//hardwood ba_ft2_ac
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_ba_ft2_ac) " + 
						"SELECT a.biosum_cond_id, b.tottemp AS hwd_ba_ft2_ac " + 
						"FROM " + this.m_strCondTable + " a, " +  
						"(SELECT biosum_cond_id, SUM((.005454154 * dia^2)  * tpacurr)  AS tottemp " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE biosum_status_cd=9 AND " + 
						"spcd > 299 AND statuscd=1 " + 
						"GROUP BY biosum_cond_id)  b " + 
						"WHERE a.biosum_cond_id = b.biosum_cond_id ";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 23);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_BA_FT2_AC = u.HWD_BA_FT2_AC";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 24);

				//volcsgrs  
				//gross sawlog
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,volcsgrs) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as  volcsgrs " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcsgrs) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"GROUP BY biosum_cond_id) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 25);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.volcsgrs = u.volcsgrs;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);

				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 26);


				//swd_volcsgrs      
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,swd_volcsgrs) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as swd_volcsgrs  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcsgrs) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE SPCD < 300 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 27);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.SWD_volcsgrs = u.SWD_volcsgrs";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 28);

				//hwd_volcsgrs      
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,hwd_volcsgrs) " + 
						"SELECT DISTINCT(a.biosum_cond_id),a.ttl as hwd_volcsgrs  " + 
						"FROM " + this.m_strTreeTable + " t, " + 
						"(SELECT biosum_cond_id, SUM(volcsgrs) as ttl " + 
						"FROM " + this.m_strTreeTable + " " + 
						"WHERE SPCD > 299 " + 
						"GROUP BY biosum_cond_id ) a " + 
						"WHERE t.biosum_status_cd=9 AND " + 
						"a.biosum_cond_id=t.biosum_cond_id;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 29);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.HWD_volcsgrs = u.HWD_volcsgrs";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 30);








            //qmd_all_inch 
            // quadratic mean diameter for all the live trees on the condition
            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,qmd_all_inch) " +
                        "SELECT c.biosum_cond_id, SQR(c.ba_ft2_ac/(.005454154 * c.tpacurr)) as qmd_all_inch " + 
						"FROM " + this.m_strCondTable + " c " + 
						"WHERE c.biosum_status_cd=9 AND " + 
						      "c.ba_ft2_ac IS NOT NULL AND " + 
							  "c.ba_ft2_ac <> 0 AND " + 
						      "c.tpacurr IS NOT NULL AND " + 
						      "c.tpacurr <> 0;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 31);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " +
                        "SET c.qmd_all_inch = u.qmd_all_inch;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);

				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 32);


            //qmd_swd_inch      
            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id,qmd_swd_inch) " +
                        "SELECT c.biosum_cond_id, SQR(c.swd_ba_ft2_ac/(.005454154 * c.swd_tpacurr)) as qmd_swd_inch " + 
						"FROM " + this.m_strCondTable + " c " + 
						"WHERE c.biosum_status_cd=9 AND " + 
						"c.swd_ba_ft2_ac IS NOT NULL AND " + 
						"c.swd_ba_ft2_ac <> 0 AND " + 
						"c.swd_tpacurr IS NOT NULL AND " + 
						"c.swd_tpacurr <> 0;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 33);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " +
                        "SET c.qmd_swd_inch = u.qmd_swd_inch";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 34);

            // qmd_hwd_inch    
            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

					p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table (biosum_cond_id, qmd_hwd_inch) " +
                        "SELECT c.biosum_cond_id, SQR(c.hwd_ba_ft2_ac/(.005454154 * c.hwd_tpacurr)) as qmd_hwd_inch " + 
						"FROM " + this.m_strCondTable + " c " + 
						"WHERE c.biosum_status_cd=9 AND " + 
						"c.hwd_ba_ft2_ac IS NOT NULL AND " + 
						"c.hwd_ba_ft2_ac <> 0 AND " + 
						"c.hwd_tpacurr IS NOT NULL AND " + 
						"c.hwd_tpacurr <> 0;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 35);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " +
                        "SET c.qmd_hwd_inch = u.qmd_hwd_inch";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 36);



				//VOL_AC_GRS_STEM_TTL_FT
                //gross wood volume of the total stem from ground to tip
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{
					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table(biosum_cond_id, vol_ac_grs_stem_ttl_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id),a.ttl as vol_ac_grs_stem_ttl_ft3 " +
                    "FROM " + this.m_strTreeTable + " t, " +
                    "(SELECT biosum_cond_id, SUM(voltsgrs * tpacurr) as ttl " +
                    "FROM " + this.m_strTreeTable + " " +
                    "WHERE voltsgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd = 1 AND dia >= 1 " +
                    "GROUP BY biosum_cond_id ) a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";

                    strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 37);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.vol_ac_grs_stem_ttl_ft3 = u.vol_ac_grs_stem_ttl_ft3;";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);

				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 38);


				//hwd_vol_ac_grs_stem_ttl_ft
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table(biosum_cond_id, hwd_vol_ac_grs_stem_ttl_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id),a.ttl as  hwd_vol_ac_grs_stem_ttl_ft3 " +
                    "FROM " + this.m_strTreeTable + " t, " +
                    "(SELECT biosum_cond_id, SUM(voltsgrs * tpacurr) as ttl " +
                    "FROM " + this.m_strTreeTable + " " +
                    "WHERE spcd > 299 AND voltsgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd = 1 AND dia >= 1 " +
                    "GROUP BY biosum_cond_id ) a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";

                    strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 39);

                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.hwd_vol_ac_grs_stem_ttl_ft3 = u.hwd_vol_ac_grs_stem_ttl_ft3";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 40);

				//swd_vol_ac_grs_stem_ttl_ft     
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "DELETE FROM cond_column_updates_work_table;";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

                p_ado.m_strSQL = "INSERT INTO cond_column_updates_work_table(biosum_cond_id, swd_vol_ac_grs_stem_ttl_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id),a.ttl as swd_vol_ac_grs_stem_ttl_ft3 " +
                    "FROM " + this.m_strTreeTable + " t, " +
                    "(SELECT biosum_cond_id, SUM(voltsgrs * tpacurr) as ttl " +
                    "FROM " + this.m_strTreeTable + " " +
                    "WHERE spcd< 300 AND voltsgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd = 1 AND dia >= 1 " +
                    "GROUP BY biosum_cond_id ) a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";

                    strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 41);
                if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
				{

					p_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " c " + 
						"INNER JOIN cond_column_updates_work_table u " + 
						"ON c.biosum_cond_id = u.biosum_cond_id " + 
						"SET c.swd_vol_ac_grs_stem_ttl_ft3 = u.swd_vol_ac_grs_stem_ttl_ft3";

					strTime = System.DateTime.Now.ToString();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
					p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 42);
			//----------------------PLOT COLUMN UPDATES-----------------------//
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//----------------------PLOT COLUMN UPDATES-----------------------//\r\n");
            SetThermValue(m_frmTherm.progressBar1, "Maximum", 4);
            SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
            SetThermValue(m_frmTherm.progressBar1, "Value", 0);
            SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Plot Table Columns...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{

				/********************************************
				 **update the plot half state field
				 ********************************************/
				p_ado.m_strSQL = "UPDATE " + this.m_strPlotTable + " p " + 
					" INNER JOIN " + this.m_strCondTable + " c " + 
					"ON p.biosum_plot_id = c.biosum_plot_id  " + 
					" SET p.half_state = MID(c.vol_loc_grp,5,LEN(TRIM(c.vol_loc_grp))) " + 
					" WHERE c.condid=1;";

				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
				//MessageBox.Show(strTime);

				
				 
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 1);

            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				/***************************************************
				 **update the number of conditions on each plot
				 ***************************************************/
				//use the biosum_plot_input as our work table so delete all records
				p_ado.m_strSQL = "DELETE FROM plot_column_updates_work_table ";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);

				//insert the condition counts into the work table
				p_ado.m_strSQL = "INSERT INTO plot_column_updates_work_table (biosum_plot_id, cond_ttl) " + 
					" SELECT biosum_plot_id , COUNT(biosum_plot_id) " + 
					" FROM " + this.m_strCondTable + 
					" GROUP BY biosum_plot_id;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 2);

            if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{

				p_ado.m_strSQL = "UPDATE " + this.m_strPlotTable  + " p " + 
					"INNER JOIN plot_column_updates_work_table i " + 
					"ON  p.biosum_plot_id = i.biosum_plot_id " + 
					"SET p.num_cond = i.cond_ttl, p.one_cond_yn = IIF(i.cond_ttl > 1,'N','Y');";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
				p_ado.SqlNonQuery(this.m_connTempMDBFile,p_ado.m_strSQL);
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 3);

		    //Set plot.gis_yard_dist_ft to fiadb_fvs_variant.MoveDistance_ft if available
		    if (p_ado.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
		    {
		        if (p_ado.ColumnExist(this.m_connTempMDBFile, Tables.Reference.DefaultFiadbFVSVariantTableName, "MoveDistance_ft"))
		        {
		            p_ado.m_strSQL = String.Format(
		                @"UPDATE {0} p INNER JOIN {1} ffv 
                        ON p.statecd=ffv.statecd and p.countycd=ffv.countycd and p.plot=ffv.plot 
                        SET p.gis_yard_dist_ft=ffv.MoveDistance_ft 
                        WHERE ffv.MoveDistance_ft IS NOT NULL and p.biosum_status_cd=9;",
		                m_strPlotTable, Tables.Reference.DefaultFiadbFVSVariantTableName);
		            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
		                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
		            p_ado.SqlNonQuery(this.m_connTempMDBFile, p_ado.m_strSQL);
		        }
		    }
		    SetThermValue(m_frmTherm.progressBar1, "Value", 4);

		    return p_ado.m_intError;
		}

	    private int ImportDownWoodyMaterials(ado_data_access p_ado)
	    {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.ImportDownWoodyMaterials\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            

            SetLabelValue(m_frmTherm.lblMsg,"Text","Importing DWM data...Stand By");
            SetThermValue(m_frmTherm.progressBar1,"Maximum", 14);
            SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
            SetThermValue(m_frmTherm.progressBar1, "Value", 0);

		    string strFIADBDbFile = "";
		    strFIADBDbFile = (string) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox) txtFiadbInputFile, "Text", false);
		    strFIADBDbFile = strFIADBDbFile.Trim();
		    String strFiaCWD = "fiadb_dwm_cwd_input";
		    String strFiaFWD = "fiadb_dwm_fwd_input";
		    String strFiaDL = "fiadb_dwm_dufflitter_input";
		    String strFiaTS = "fiadb_dwm_transect_segment_input";

		    p_ado.CloseConnection(m_connTempMDBFile);
	        m_connTempMDBFile.Dispose();
	        m_connTempMDBFile = null;


            //If any of the FIADB source DWM tables do not exist,
            //show message, uncheck the DWM checkbox, return early
            DataMgr oDataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFIADBDbFile)))
	        {
	            conn.Open();
                if (!oDataMgr.TableExist(conn, m_strDwmCwdTable) ||
                    !oDataMgr.TableExist(conn, m_strDwmFwdTable) ||
                    !oDataMgr.TableExist(conn, m_strDwmDuffLitterTable) ||
                    !oDataMgr.TableExist(conn, m_strDwmTransectSegmentTable))
                {
	                Func<bool, string, string> result = (boolTableExists, tableName) =>
	                {
	                    if (!boolTableExists) return "\r\n - " + tableName;
	                    else return "";
	                };
	                DialogResult dlgResult = MessageBox.Show(String.Format(
	                        "!!Error!!\nModule - uc_plot_input:ImportDownWoodyMaterials\n" + "Err Msg - " +
	                        "At least one FIADB Source DWM table was not found:{0}{1}{2}{3}\r\nDo you wish to continue plot data input without DWM?",
	                        result(oDataMgr.TableExist(conn, m_strDwmCwdTable), m_strDwmCwdTable),
	                        result(oDataMgr.TableExist(conn, m_strDwmFwdTable), m_strDwmFwdTable),
	                        result(oDataMgr.TableExist(conn, m_strDwmDuffLitterTable), m_strDwmDuffLitterTable),
	                        result(oDataMgr.TableExist(conn, m_strDwmTransectSegmentTable), m_strDwmTransectSegmentTable)),
	                    "FIA Biosum",
	                    MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation);
	                //Disable functionality related to DWM option down the pipeline
	                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.CheckBox) chkDwmImport,
	                    "Checked", false);

	                if (dlgResult == DialogResult.No)
	                {
	                    return -1; //terminates plot input processing.
	                }
	                else if (dlgResult == DialogResult.Yes)
	                {
                        //m_intError == 0 keeps performing plot input.
	                    return 0;
	                }
	            }
	        }

		    dao_data_access p_dao = new dao_data_access();
            //Link to FIADB source tables in temporary database
            p_dao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmCwdTable, strFiaCWD,  ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
            p_dao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmFwdTable, strFiaFWD,  ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
            p_dao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmDuffLitterTable, strFiaDL,  ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
            p_dao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmTransectSegmentTable, strFiaTS,  ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);

            //Link to Master_AUX.accdb DWM source tables in temporary database
            string strMasterAuxDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.accdb"; 
		    p_dao.CreateTableLink(m_strTempMDBFile, m_strDwmCwdTable, strMasterAuxDb, m_strDwmCwdTable);
		    p_dao.CreateTableLink(m_strTempMDBFile, m_strDwmFwdTable, strMasterAuxDb, m_strDwmFwdTable);
		    p_dao.CreateTableLink(m_strTempMDBFile, m_strDwmDuffLitterTable, strMasterAuxDb, m_strDwmDuffLitterTable);
		    p_dao.CreateTableLink(m_strTempMDBFile, m_strDwmTransectSegmentTable, strMasterAuxDb, m_strDwmTransectSegmentTable);

		    m_intError = p_dao.m_intError;
		    p_dao.m_DaoWorkspace.Close();
	        p_dao.m_DaoWorkspace = null;
	        p_dao.m_DaoDbEngine = null;
		    p_dao = null;

            m_connTempMDBFile = new OleDbConnection(m_strTempMDBFileConn);
	        m_connTempMDBFile.Open();

            SetThermValue(m_frmTherm.progressBar1, "Value", 1);

		    String strSourceFields = "";
		    String strDestFields = "";
		    System.Data.DataTable dtDwmCwd = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + m_strDwmCwdTable);
		    System.Data.DataTable dtDwmFwd = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + m_strDwmFwdTable);
		    System.Data.DataTable dtDwmDuffLitter = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + m_strDwmDuffLitterTable);
		    System.Data.DataTable dtDwmTransectSegment = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + m_strDwmTransectSegmentTable);
		    System.Data.DataTable dtFIADBDwmCwd = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + strFiaCWD);
		    System.Data.DataTable dtFIADBDwmFwd = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + strFiaFWD);
		    System.Data.DataTable dtFIADBDwmDuffLitter = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + strFiaDL);
		    System.Data.DataTable dtFIADBDwmTransectSegment = p_ado.getTableSchema(m_connTempMDBFile, "select * from " + strFiaTS);


            //Preemptively remove any records that were not imported successfully 
	        DeleteFromTablesWhereFilter(p_ado, m_connTempMDBFile,
	            new string[]
	                {m_strDwmCwdTable, m_strDwmFwdTable, m_strDwmDuffLitterTable, m_strDwmTransectSegmentTable},
	            " WHERE biosum_status_cd=9");

	        //DWM Coarse Woody Debris FIADB into Master.MDB Table
	        if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
	        {
	            SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM CWD...Stand By");
	            SetThermValue(m_frmTherm.progressBar1, "Value", 2);
	            strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmCwd, dtDestSchema: dtDwmCwd);
	            strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmCwd, dtDestSchema: dtDwmCwd,
	                strTablePrefix: "dwm.");
	            InsertIntoDestTableFromSourceTable(p_ado, strSourceTable: strFiaCWD, strDestTable: m_strDwmCwdTable,
	                strDestFields: strDestFields, strSourceFields: strSourceFields, InsertBiosumCondIdAndStatusCode: true);
	            UpdateDwmBiosumCondIds(p_ado, m_strDwmCwdTable);
	            m_intError = p_ado.m_intError;
	        }

	        //DWM Fine Woody Debris FIADB into Master.MDB Table
	        if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
	        {
	            SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM FWD...Stand By");
	            SetThermValue(m_frmTherm.progressBar1, "Value", 4);
	            strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmFwd, dtDestSchema: dtDwmFwd);
	            strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmFwd, dtDestSchema: dtDwmFwd,
	                strTablePrefix: "dwm.");
	            InsertIntoDestTableFromSourceTable(p_ado, strSourceTable: strFiaFWD, strDestTable: m_strDwmFwdTable,
	                strDestFields: strDestFields, strSourceFields: strSourceFields, InsertBiosumCondIdAndStatusCode: true);
	            UpdateDwmBiosumCondIds(p_ado, m_strDwmFwdTable);
	            m_intError = p_ado.m_intError;
	        }

	        //DWM Duff Litter Fuel FIADB into Master.MDB Table
	        if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
	        {
	            SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM DuffLitter...Stand By");
	            SetThermValue(m_frmTherm.progressBar1, "Value", 6);
	            strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmDuffLitter,
	                dtDestSchema: dtDwmDuffLitter);
	            strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmDuffLitter,
	                dtDestSchema: dtDwmDuffLitter, strTablePrefix: "dwm.");
	            InsertIntoDestTableFromSourceTable(p_ado, strSourceTable: strFiaDL,
	                strDestTable: m_strDwmDuffLitterTable,
	                strDestFields: strDestFields, strSourceFields: strSourceFields,
	                InsertBiosumCondIdAndStatusCode: true);
                p_ado.SqlNonQuery(m_connTempMDBFile,
                    String.Format(
                        "UPDATE {0} SET duffdep=0 WHERE duffdep IS NULL AND duff_nonsample_reasn_cd IS NULL;", //or other metacodes are null? dlf_sample_method, dl_status_cd, etc.
                        m_strDwmDuffLitterTable));
                p_ado.SqlNonQuery(m_connTempMDBFile,
                    String.Format(
                        "UPDATE {0} SET littdep=0 WHERE littdep IS NULL AND litter_nonsample_reasn_cd IS NULL;", //or other metacodes are null? dlf_sample_method, dl_status_cd, etc.
                        m_strDwmDuffLitterTable));
	            UpdateDwmBiosumCondIds(p_ado, m_strDwmDuffLitterTable);
	            m_intError = p_ado.m_intError;
	        }

	        //DWM Transect Segment FIADB into Master.MDB Table
	        if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control) m_frmTherm, "AbortProcess"))
	        {
	            SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM Transect Segment...Stand By");
	            SetThermValue(m_frmTherm.progressBar1, "Value", 8);
	            strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmTransectSegment,
	                dtDestSchema: dtDwmTransectSegment);
	            strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmTransectSegment,
	                dtDestSchema: dtDwmTransectSegment, strTablePrefix: "dwm.");
	            InsertIntoDestTableFromSourceTable(p_ado, strSourceTable: strFiaTS,
	                strDestTable: m_strDwmTransectSegmentTable,
	                strDestFields: strDestFields, strSourceFields: strSourceFields, InsertBiosumCondIdAndStatusCode: true);
	            UpdateDwmBiosumCondIds(p_ado, m_strDwmTransectSegmentTable);
	            m_intError = p_ado.m_intError;
	        }

            SetLabelValue(m_frmTherm.lblMsg,"Text","");
            SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
	        return p_ado.m_intError;
	    }

        private int ImportDownWoodyMaterialsSqlite()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.ImportDownWoodyMaterials\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }


            SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM data...Stand By");
            SetThermValue(m_frmTherm.progressBar1, "Maximum", 14);
            SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
            SetThermValue(m_frmTherm.progressBar1, "Value", 0);

            string strFIADBDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtFiadbInputFile, "Text", false);
            strFIADBDbFile = strFIADBDbFile.Trim();
            String strFiaCWD = "fiadb." + m_strDwmCwdTable;
            String strFiaFWD = "fiadb." + m_strDwmFwdTable;
            String strFiaDL = "fiadb." + m_strDwmDuffLitterTable;
            String strFiaTS = "fiadb." + m_strDwmTransectSegmentTable;

            string strMasterAuxDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";

            ODBCMgr oODBC = new ODBCMgr();
            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();

            if (oODBC.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterAuxDsnName))
            {
                oODBC.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName);
            }
            oODBC.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName, strMasterAuxDb);

            //If any of the FIADB source DWM tables do not exist,
            //show message, uncheck the DWM checkbox, return early
            DataMgr oDataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strMasterAuxDb)))
            {
                conn.Open();

                oDataMgr.m_strSQL = "ATTACH DATABASE '" + strFIADBDbFile + "' AS fiadb";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                if (!oDataMgr.AttachedTableExist(conn, m_strDwmCwdTable) ||
                    !oDataMgr.AttachedTableExist(conn, m_strDwmFwdTable) ||
                    !oDataMgr.AttachedTableExist(conn, m_strDwmDuffLitterTable) ||
                    !oDataMgr.AttachedTableExist(conn, m_strDwmTransectSegmentTable))
                {
                    Func<bool, string, string> result = (boolTableExists, tableName) =>
                    {
                        if (!boolTableExists) return "\r\n - " + tableName;
                        else return "";
                    };
                    DialogResult dlgResult = MessageBox.Show(String.Format(
                            "!!Error!!\nModule - uc_plot_input:ImportDownWoodyMaterials\n" + "Err Msg - " +
                            "At least one FIADB Source DWM table was not found:{0}{1}{2}{3}\r\nDo you wish to continue plot data input without DWM?",
                            result(oDataMgr.AttachedTableExist(conn, m_strDwmCwdTable), m_strDwmCwdTable),
                            result(oDataMgr.AttachedTableExist(conn, m_strDwmFwdTable), m_strDwmFwdTable),
                            result(oDataMgr.AttachedTableExist(conn, m_strDwmDuffLitterTable), m_strDwmDuffLitterTable),
                            result(oDataMgr.AttachedTableExist(conn, m_strDwmTransectSegmentTable), m_strDwmTransectSegmentTable)),
                        "FIA Biosum",
                        MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    //Disable functionality related to DWM option down the pipeline
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.CheckBox)chkDwmImport,
                        "Checked", false);

                    if (dlgResult == DialogResult.No)
                    {
                        return -1; //terminates plot input processing.
                    }
                    else if (dlgResult == DialogResult.Yes)
                    {
                        //m_intError == 0 keeps performing plot input.
                        return 0;
                    }
                }

                String strSourceFields = "";
                String strDestFields = "";
                System.Data.DataTable dtDwmCwd = oDataMgr.getTableSchema(conn, "select * from " + m_strDwmCwdTable);
                System.Data.DataTable dtDwmFwd = oDataMgr.getTableSchema(conn, "select * from " + m_strDwmFwdTable);
                System.Data.DataTable dtDwmDuffLitter = oDataMgr.getTableSchema(conn, "select * from " + m_strDwmDuffLitterTable);
                System.Data.DataTable dtDwmTransectSegment = oDataMgr.getTableSchema(conn, "select * from " + m_strDwmTransectSegmentTable);
                System.Data.DataTable dtFIADBDwmCwd = oDataMgr.getTableSchema(conn, "select * from " + strFiaCWD);
                System.Data.DataTable dtFIADBDwmFwd = oDataMgr.getTableSchema(conn, "select * from " + strFiaFWD);
                System.Data.DataTable dtFIADBDwmDuffLitter = oDataMgr.getTableSchema(conn, "select * from " + strFiaDL);
                System.Data.DataTable dtFIADBDwmTransectSegment = oDataMgr.getTableSchema(conn, "select * from " + strFiaTS);

                //Preemptively remove any records that were not imported successfully 
                DeleteFromTablesWhereFilterSqlite(conn,
                    new string[]
                        {m_strDwmCwdTable, m_strDwmFwdTable, m_strDwmDuffLitterTable, m_strDwmTransectSegmentTable},
                    " WHERE biosum_status_cd=9");

                //Link FIADB tables to temp accdb
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmCwdTable, m_strDwmCwdTable + "_fiadb", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile);
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmFwdTable, m_strDwmFwdTable + "_fiadb", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile);
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmDuffLitterTable, m_strDwmDuffLitterTable + "_fiadb", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile);
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmTransectSegmentTable, m_strDwmTransectSegmentTable + "_fiadb", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile);

                //Link master_aux.db tables to temp accdb
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmCwdTable, m_strDwmCwdTable + "_master", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strMasterAuxDb);
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmFwdTable, m_strDwmFwdTable + "_master", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strMasterAuxDb);
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmDuffLitterTable, m_strDwmDuffLitterTable + "_master", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strMasterAuxDb);
                oDao.CreateSQLiteTableLink(m_strTempMDBFile, m_strDwmTransectSegmentTable, m_strDwmTransectSegmentTable + "_master", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strMasterAuxDb);

                //Sleep until tables are linked
                int i = 0;
                do
                {
                    // break out of loop if it runs too long
                    if (i > 20)
                    {
                        System.Windows.Forms.MessageBox.Show("An error occurred while trying to attach FIADB and master_aux tables! " +
                        "Validate the contents of this database before trying to run Treatment Optimizer.", "FIA Biosum");
                        break;
                    }
                    System.Threading.Thread.Sleep(1000);
                    i++;
                }
                while (!this.m_ado.TableExist(m_connTempMDBFile, m_strDwmCwdTable + "_fiadb") || !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmCwdTable + "_master") ||
                !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmFwdTable + "_fiadb") || !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmFwdTable + "_master") ||
                !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmDuffLitterTable + "_fiadb") || !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmDuffLitterTable + "_master") ||
                !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmTransectSegmentTable + "_fiadb") || !this.m_ado.TableExist(m_connTempMDBFile, m_strDwmTransectSegmentTable + "_master"));

                Thread.Sleep(5000);

                //DWM Coarse Woody Debris FIADB into master_aux.db Table
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM CWD...Stand By");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 2);
                    strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmCwd, dtDestSchema: dtDwmCwd);
                    strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmCwd, dtDestSchema: dtDwmCwd,
                        strTablePrefix: "dwm.");
                    InsertIntoDestTableFromSourceTableSqlite(oAdo, strSourceTable: m_strDwmCwdTable, strDestTable: m_strDwmCwdTable,
                        strDestFields: strDestFields, strSourceFields: strSourceFields, InsertBiosumCondIdAndStatusCode: true);
                    //oAdo.AddPrimaryKey(m_connTempMDBFile, m_strDwmCwdTable + "_master", "pk1",  "CN");
                    
                    UpdateDwmBiosumCondIdsSqlite(oAdo, m_strDwmCwdTable);
                    m_intError = oDataMgr.m_intError;
                }

                //DWM Fine Woody Debris FIADB into master_aux.db Table
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM FWD...Stand By");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 4);
                    strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmFwd, dtDestSchema: dtDwmFwd);
                    strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmFwd, dtDestSchema: dtDwmFwd,
                        strTablePrefix: "dwm.");
                    InsertIntoDestTableFromSourceTableSqlite(oAdo, strSourceTable: m_strDwmFwdTable, strDestTable: m_strDwmFwdTable,
                        strDestFields: strDestFields, strSourceFields: strSourceFields, InsertBiosumCondIdAndStatusCode: true);
                    UpdateDwmBiosumCondIdsSqlite(oAdo, m_strDwmFwdTable);
                    m_intError = oDataMgr.m_intError;
                }

                //DWM Duff Litter Fuel FIADB into Master.MDB Table
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM DuffLitter...Stand By");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 6);
                    strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmDuffLitter,
                        dtDestSchema: dtDwmDuffLitter);
                    strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmDuffLitter,
                        dtDestSchema: dtDwmDuffLitter, strTablePrefix: "dwm.");
                    InsertIntoDestTableFromSourceTableSqlite(oAdo, strSourceTable: m_strDwmDuffLitterTable,
                        strDestTable: m_strDwmDuffLitterTable,
                        strDestFields: strDestFields, strSourceFields: strSourceFields,
                        InsertBiosumCondIdAndStatusCode: true);
                    oDataMgr.m_strSQL = "UPDATE " + m_strDwmDuffLitterTable + " SET duffdep=0 WHERE duffdep IS NULL AND duff_nonsample_reasn_cd IS NULL";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = "UPDATE " + m_strDwmDuffLitterTable + " SET littdep=0 WHERE littdep IS NULL AND litter_nonsample_reasn_cd IS NULL";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    UpdateDwmBiosumCondIdsSqlite(oAdo, m_strDwmDuffLitterTable);
                    m_intError = oDataMgr.m_intError;
                }

                //DWM Transect Segment FIADB into Master.MDB Table
                if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                {
                    SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM Transect Segment...Stand By");
                    SetThermValue(m_frmTherm.progressBar1, "Value", 8);
                    strDestFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmTransectSegment,
                        dtDestSchema: dtDwmTransectSegment);
                    strSourceFields = CreateStrFieldsFromDataTables(dtSourceSchema: dtFIADBDwmTransectSegment,
                        dtDestSchema: dtDwmTransectSegment, strTablePrefix: "dwm.");
                    InsertIntoDestTableFromSourceTableSqlite(oAdo, strSourceTable: m_strDwmTransectSegmentTable,
                        strDestTable: m_strDwmTransectSegmentTable,
                        strDestFields: strDestFields, strSourceFields: strSourceFields, InsertBiosumCondIdAndStatusCode: true);
                    UpdateDwmBiosumCondIdsSqlite(oAdo, m_strDwmTransectSegmentTable);
                    m_intError = oDataMgr.m_intError;
                }
            }

            SetLabelValue(m_frmTherm.lblMsg, "Text", "");
            SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
            return oDataMgr.m_intError;
        }

      //  private int ImportGrowthRemovalMortality(ado_data_access p_ado)
      //  {
      //   string strFIADBDbFile = "";
      //   strFIADBDbFile = (string) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox) txtFiadbInputFile, "Text", false);
      //   strFIADBDbFile = strFIADBDbFile.Trim();
      //      string strMasterAuxDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.accdb";
      //         m_strGRMComponentTable = (string) frmMain.g_oDelegate.GetControlPropertyValue(cmbGrmComponentTable, "Text", false);

        //         if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
        //         {
        //             frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
        //             frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.ImportGrowthRemovalMortality\r\n");
        //             frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
        //         }

        //         DataMgr oDataMgr = new DataMgr();
        //         using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFIADBDbFile)))
        //         {
        //             conn.Open();
        //             if (!oDataMgr.TableExist(conn, m_strGRMComponentTable))
        //          {
        //              DialogResult dlgResult = MessageBox.Show(
        //                  "!!Error!!\nModule - uc_plot_input:ImportGrowthRemovalMortality\n" + "Err Msg - " +
        //                  m_strGRMComponentTable +
        //                  " was not found in FIADB Source!\r\nDo you wish to continue plot data input without GRM Calibration Data?",
        //                  "FIA Biosum", MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation);

        //              //Disable functionality related to GRM option down the pipeline
        //                 frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.CheckBox) chkGrmImport, "Checked", false); //unchecked value?
        //              if (dlgResult == DialogResult.No)
        //              {
        //                  return -1; //terminates plot input processing.
        //              }
        //              else if (dlgResult == DialogResult.Yes)
        //              {
        //                  //m_intError == 0 keeps performing plot input.
        //                  return 0;
        //              }
        //          }
        //      }

        //      using (var conn = new OleDbConnection(p_ado.getMDBConnString(strMasterAuxDb, "","")))
        //      {
        //          conn.Open();
        //          string message = "";
        //          bool bGrmStandTableNotFound = !p_ado.TableExist(conn, m_strGrmStandTable);
        //          bool bGrmTreeTableNotFound = !p_ado.TableExist(conn, m_strGrmTreeTable);
        //          if (bGrmStandTableNotFound)
        //          {
        //              message += m_strGrmStandTable + " could not be found in Master_Aux!\r\n";
        //          }
        //          if (bGrmTreeTableNotFound)
        //          {
        //              message += m_strGrmTreeTable + " could not be found in Master_Aux!\r\n";
        //          }
        //          if (bGrmStandTableNotFound || bGrmTreeTableNotFound)
        //          {
        //              DialogResult dlgResult = MessageBox.Show(
        //                  "!!Error!!\nModule - uc_plot_input:ImportGrowthRemovalMortality\n" + "Err Msg - " +
        //                  message + "\r\nDo you wish to continue plot data input without GRM Calibration Data?",
        //                  "FIA Biosum", MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation);

        //              //Disable functionality related to GRM option down the pipeline
        //              frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.CheckBox)chkGrmImport, "Checked", false);
        //              if (dlgResult == DialogResult.No)
        //              {
        //                  return -1; //terminates plot input processing.
        //              }
        //              else if (dlgResult == DialogResult.Yes)
        //              {
        //                  //m_intError == 0 keeps performing plot input.
        //                  return 0;
        //              }
        //          }
        //      }

        //      dao_data_access p_dao = new dao_data_access();
        //         //Link to FIADB source tables in temporary database
        //         p_dao.CreateSQLiteTableLink(m_strTempMDBFile, m_strGRMComponentTable, m_strGRMComponentTable, ODBCMgr.DSN_KEYS.PlotInputDsnName,
        //             strFIADBDbFile);
        //         p_dao.CreateSQLiteTableLink(m_strTempMDBFile, (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbPlotTable,
        //                 "Text", false), "FIADB_PLOT_TABLE", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
        //         p_dao.CreateSQLiteTableLink(m_strTempMDBFile, (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbCondTable,
        //                 "Text", false), "FIADB_COND_TABLE", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
        //         p_dao.CreateSQLiteTableLink(m_strTempMDBFile, (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.ComboBox)cmbFiadbTreeTable,
        //                 "Text", false), "FIADB_TREE_TABLE", ODBCMgr.DSN_KEYS.PlotInputDsnName, strFIADBDbFile, true);
        //      //Link to Master_AUX.accdb GRM dest tables in temporary database
        //      p_dao.CreateTableLink(m_strTempMDBFile, m_strGrmStandTable, strMasterAuxDb, m_strGrmStandTable);
        //      p_dao.CreateTableLink(m_strTempMDBFile, m_strGrmTreeTable, strMasterAuxDb, m_strGrmTreeTable);

        //      m_intError = p_dao.m_intError;
        //if (p_dao.m_intError != 0) {
        //	return p_dao.m_intError;
        //}
        //      p_dao.m_DaoWorkspace.Close();
        //      p_dao.m_DaoWorkspace = null;
        //      p_dao.m_DaoDbEngine = null;
        //      p_dao = null;

        //      using (var conn = new OleDbConnection(p_ado.getMDBConnString(m_strTempMDBFile, "", "")))
        //      {
        //          conn.Open();

        //             //Check if any records with biosum_status_cd=9 exist in the GRM stand and tree tables and remove them
        //          p_ado.m_strSQL = "DELETE FROM " + m_strGrmStandTable + " WHERE biosum_status_cd=9;";
        //          p_ado.SqlNonQuery(conn, p_ado.m_strSQL);
        //          p_ado.m_strSQL = "DELETE FROM " + m_strGrmTreeTable + " WHERE biosum_status_cd=9;";
        //          p_ado.SqlNonQuery(conn, p_ado.m_strSQL);

        //             //Insert into GRM_Stand
        //             p_ado.m_strSQL = "INSERT INTO " + m_strGrmStandTable +
        //                              " (biosum_cond_id, biosum_plot_id, biosum_status_cd, plt_cn, prev_plt_cn, measurement_period) " +
        //                              "SELECT c.biosum_cond_id, tp.biosum_plot_id, 9 as biosum_status_cd, tp.cn, fp.cn, " +
        //                                  "IIF(tp.remper IS NOT NULL, CInt(tp.remper), CInt(DateDiff('m', " +
        //                                  "DateSerial(fp.measyear,fp.measmon,fp.measday), " +
        //                                  "DateSerial(tp.measyear,tp.measmon,tp.measday))/12.0)) " +
        //                              "FROM (tempplot tp INNER JOIN fiadb_plot_table fp ON trim(tp.prev_plt_cn) = fp.cn) " +
        //                                  "INNER JOIN tempcond c ON trim(c.PLT_CN) = trim(tp.CN);";
        //             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
        //                 frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
        //             p_ado.SqlNonQuery(conn, p_ado.m_strSQL);

        //          //Insert into GRM_Tree
        //          p_ado.m_strSQL = "INSERT INTO " + m_strGrmTreeTable +
        //                              " (biosum_cond_id, biosum_status_cd, fvs_tree_id, tre_cn, prev_tre_cn, " +
        //                                  "dia_begin, dia_end, " +
        //                                  "micr_component_al_forest, tre_statuscd, prev_tre_statuscd) " +
        //                           "SELECT tt.biosum_cond_id, 9 as biosum_status_cd, " +
        //                                  "tt.SUBP*1000+tt.TREE, tt.CN, grm.PREV_TRE_CN, " +
        //                                  "grm.DIA_BEGIN, grm.DIA_END, " +
        //                                  "grm.MICR_COMPONENT_AL_FOREST, tt.STATUSCD, ft.STATUSCD " +
        //                           "FROM (" + m_strGRMComponentTable + " grm " +
        //                                  "INNER JOIN temptree tt ON grm.TRE_CN = tt.CN) " +
        //                                  "LEFT JOIN FIADB_TREE_TABLE ft ON grm.PREV_TRE_CN = ft.CN; ";
        //             if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
        //                 frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, p_ado.m_strSQL + "\r\n");
        //             p_ado.SqlNonQuery(conn, p_ado.m_strSQL);
        //      }

        //      return p_ado.m_intError;
        //  }

        private void ThermCancel(object sender, System.EventArgs e)
		{
			string strMsg = "Do you wish to cancel appending plot data (Y/N)?";
			DialogResult result = MessageBox.Show(strMsg,"Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			switch (result) 
			{
				case DialogResult.Yes:
					this.m_frmTherm.AbortProcess = true;
					this.m_frmTherm.Hide();
					return;
				case DialogResult.No:
					return;
			}                
		}
		/// <summary>
		/// create a delimited string list from a text file
		/// that has a single column of data with multiple rows
		/// </summary>
		/// <param name="p_strTxtFile">text file containing the column of data</param>
		/// <param name="p_strTxtFileDelimiter">specified character between list items</param>
		/// <param name="p_strListDelimiter">specified character between list items</param>
		/// <param name="p_bNumericDataType">specifies if the column data to retrieve in the text file is numeric</param>
		/// <returns></returns>
		private string CreateDelimitedStringList(string p_strTxtFile,string p_strTxtFileDelimiter, string p_strListDelimiter,bool p_bNumericDataType)
		{
			//The DataSet to Return
			//DataSet result = new DataSet();
			this.m_intError=0;
			string strList="";
			string str="";
			try
			{
				//Open the file in a stream reader.
				System.IO.StreamReader s = new System.IO.StreamReader(p_strTxtFile);
				//Read the rest of the data in the file.        
				string AllData = s.ReadToEnd();
    
				//Split off each row at the Carriage Return/Line Feed
				//Default line ending in most <A class=iAs style="FONT-WEIGHT: normal; FONT-SIZE: 100%; PADDING-BOTTOM: 1px; COLOR: darkgreen; BORDER-BOTTOM: darkgreen 0.07em solid; BACKGROUND-COLOR: transparent; TEXT-DECORATION: underline" href="#" target=_blank itxtdid="2592535">windows</A> exports.  
				string[] rows = AllData.Split("\r\n".ToCharArray());
 
				//Now add each row to the DataSet        
				foreach(string r in rows)
				{
					//Split the row at the delimiter.
					string[] items = r.Split(p_strTxtFileDelimiter.ToCharArray());
					str = items[0].Trim();  //plot_cn in first column
					str = str.Replace("\"",""); //remove any quotations
					if (str.Trim().Length > 0)
					{
						if (strList.Trim().Length == 0)
						{
							if (p_bNumericDataType == true)
							{
								strList = str.Trim();
							}
							else
							{
								strList = "'" + str.Trim() + "'";
							}
						}
						else
						{
							if (p_bNumericDataType == true)
							{
								strList = strList + p_strListDelimiter.Trim() + str.Trim();
							}
							else
							{
								strList = strList + p_strListDelimiter.Trim() + "'" + str.Trim() + "'";
							}

						}
					}
				}
			}
			catch (Exception caught)
			{
				this.m_intError=-1;
				MessageBox.Show("!!Error: CreateDelimitedStringList() Routine Error Msg:" + caught.Message);
			}
			return strList;
		}

		private void btnFilterByFileBrowse_Click(object sender, System.EventArgs e)
		{
			
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = "Text File With PLOT_CN data";
			OpenFileDialog1.Filter = "Text File (*.TXT;*.DAT) |*.txt;*.dat";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					this.txtFilterByFile.Text = OpenFileDialog1.FileName.Trim();
				}
				else 
				{
				}
				OpenFileDialog1 = null;

			}
		}

		private void btnFilterNext_Click(object sender, System.EventArgs e)
		{
				
		    if (this.LoadMDBFiadbPopEvalTable() && m_intError==0)
			{	
			    this.m_strLoadedPopEvalInputTable=this.cmbFiadbPopEvalTable.Text;
				this.FIADBLoadInv();

                List<string> lstFiadbStates = QueryFiadbStates();
                IDictionary<string, List<string>> dictStateEval = QueryStateEvalids();
                if (dictStateEval.Keys.Count > 0 && lstFiadbStates.Count > 0)
                {
                    bool bExists = false;
                    // Check to see if any states in list are in the dict keys
                    foreach (var strState in lstFiadbStates)
                    {
                        if (dictStateEval.Keys.Contains(strState))
                        {
                            bExists = true;
                            break;
                        }
                    }
                    if (bExists)
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("The POP_PLT_STRATUM_ASSIGN table in this project contains records with the following ");
                        sb.Append("combinations of STATE and EVALID for the FIADB tables you are about to load. ");
                        sb.Append("It is STRONGLY advised to include only one EVALID per STATE in a BioSum project. ");
                        sb.Append("Departing from this advisory presents risks to integrity and interpretability of project ");
                        sb.Append("data that have not been fully explored.\n\r");
                        foreach (var strState in lstFiadbStates)
                        {
                            if (dictStateEval.Keys.Contains(strState))
                            {
                                List<string> lstEvalid = dictStateEval[strState];
                                string csv = String.Join(", ", lstEvalid.Select(x => x.ToString()).ToArray());
                                sb.Append(strState + " -> " + csv);
                            }
                            sb.Append("\n");
                        }
                        MessageBox.Show(sb.ToString(), "FIA Biosum");
                    }
                }						
			}

			if (this.m_intError==0)
			{
			    if (this.rdoFilterByMenu.Checked==true)
					{
						this.btnFIADBInvAppend.Enabled=false;
						this.btnFIADBInvNext.Enabled=true;
					}
					else
					{
						this.btnFIADBInvAppend.Enabled=true;
						this.btnFIADBInvNext.Enabled=false;
					}
					this.grpboxFIADBInv.Visible=true;
					this.grpboxFilter.Visible=false;
			}
			
		}

		private void mdbInputStateCounty()
		{
	
			string strState="";
			string strCounty="";
			int intAddedPlotRows=0;
    

			this.m_dtStateCounty.Clear();        
			this.lstFilterByState.Clear();
			this.lstFilterByState.Columns.Add(" ", 100, HorizontalAlignment.Center); 
			this.lstFilterByState.Columns.Add("State", 100, HorizontalAlignment.Left);
			this.lstFilterByState.Columns.Add("County", 100, HorizontalAlignment.Left);

			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
			this.m_intError=0;

            string strSqliteFile = this.txtFiadbInputFile.Text.Trim();
            DataMgr oDataMgr = new DataMgr();
            string strTable = this.cmbFiadbPpsaTable.Text.Trim();
            string strConn = oDataMgr.GetConnectionString(strSqliteFile);
            string strSQL = "";

			if (this.chkNonForested.Checked == true && this.chkForested.Checked==true)
			{

					strSQL = "SELECT statecd,countycd " + 
						"FROM " + strTable +  " " + 
						"WHERE RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"EVALID = " + this.m_strCurrFIADBEvalId + " " +  
						"GROUP BY statecd,countycd;";
			}
			else if (this.chkForested.Checked==true)
			{

                strSQL = "SELECT ppsa.statecd,ppsa.countycd " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd=1 " + 
						"GROUP BY ppsa.statecd,ppsa.countycd;";

			}
			else if (this.chkNonForested.Checked==true)
			{

                strSQL = "SELECT ppsa.statecd,ppsa.countycd " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd<>1 " + 
						"GROUP BY ppsa.statecd,ppsa.countycd;";
			}
			else
			{
				this.m_intError=-1;

			}

            try
            {

                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    con.Open();
                    oDataMgr.SqlQueryReader(con, strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        //load up each row in the FIADB plot input table
                        while (oDataMgr.m_DataReader.Read())
                        {
                            strState = "";
                            strCounty = "";

                            //make sure the row is not null values
                            if (oDataMgr.m_DataReader[0] != DBNull.Value &&
                                oDataMgr.m_DataReader[0].ToString().Trim().Length > 0)
                            { 
                                strState = oDataMgr.m_DataReader["statecd"].ToString();
                                strCounty = oDataMgr.m_DataReader["countycd"].ToString();
                                this.lstFilterByState.BeginUpdate();
                                System.Windows.Forms.ListViewItem listItem = new ListViewItem();
                                listItem.Checked = false;
                                listItem.SubItems.Add(strState);
                                listItem.SubItems.Add(strCounty);
                                this.lstFilterByState.Items.Add(listItem);
                                this.lstFilterByState.EndUpdate();
                                intAddedPlotRows++;
                            }
                        }
                    }
                    if (intAddedPlotRows == 0)
                    {
                        this.m_intError = -1;
                        MessageBox.Show("!!No Plots Loaded To Get State, County, Plot Information!!", "Load State, County, Plot Menus");
                    }
                    ((frmDialog)this.ParentForm).Enabled = true;
                }
            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                MessageBox.Show(caught.Message);
            }
        }

        private void mdbInputPlot()
		{
			this.m_intError=0;
			int intAddedPlotRows=0;
			string strState="";
			string strCounty="";
			string strPlot="";
			this.lstFilterByPlot.Clear();
			this.lstFilterByPlot.Columns.Add(" ", 50, HorizontalAlignment.Center); 
			this.lstFilterByPlot.Columns.Add("State", 75, HorizontalAlignment.Left);
			this.lstFilterByPlot.Columns.Add("County", 75, HorizontalAlignment.Left);
			this.lstFilterByPlot.Columns.Add("Plot", 100, HorizontalAlignment.Left);

            DataMgr oDataMgr = new DataMgr();

            string strSQLiteFile = this.txtFiadbInputFile.Text.Trim();
            string strTable = this.cmbFiadbPpsaTable.Text.Trim();
            string strConn = oDataMgr.GetConnectionString(strSQLiteFile);

            ListView oLv = (ListView)frmMain.g_oDelegate.GetListView(lstFilterByState, false);
            int intCheckedCount = (int)frmMain.g_oDelegate.GetListViewCheckedItemsCount(oLv, false);
            if (intCheckedCount == 0)
            {
                this.m_intError = -1;
                MessageBox.Show("!!No checkboxes selected. Plots cannot be loaded!!", "Load State, County, Plot Menus");
                ((frmDialog)this.ParentForm).Enabled = true;
                return;
            }

            string strSQL = "";
            if (this.chkNonForested.Checked == true && this.chkForested.Checked==true)
			{
				this.BuildFilterByStateCountyString("statecd","countycd",false);

                strSQL = "SELECT statecd,countycd,plot " + 
						"FROM " + strTable +  " " + 
						"WHERE RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"EVALID = " + this.m_strCurrFIADBEvalId + " AND " + this.m_strStateCountySQL.Trim() + " " + 
						"GROUP BY statecd,countycd,plot;";
			}
			else if (this.chkForested.Checked==true)
			{

				this.BuildFilterByStateCountyString("ppsa.statecd","ppsa.countycd",false);
                strSQL = "SELECT ppsa.statecd,ppsa.countycd,ppsa.plot " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd=1 AND " + this.m_strStateCountySQL.Trim() + " " +
						"GROUP BY ppsa.statecd,ppsa.countycd,ppsa.plot;";
			}
			else if (this.chkNonForested.Checked==true)
			{

				this.BuildFilterByStateCountyString("ppsa.statecd","ppsa.countycd",false);
                strSQL = "SELECT ppsa.statecd,ppsa.countycd,ppsa.plot " + 
						"FROM " + strTable +  " ppsa " + 
						"INNER JOIN " + this.cmbFiadbPlotTable.Text.Trim() + " p " + 
						"ON ppsa.plt_cn = p.cn " + 
						"WHERE ppsa.RSCD = " + this.m_strCurrFIADBRsCd + " AND " + 
						"ppsa.EVALID = " + this.m_strCurrFIADBEvalId + " AND " +  
						"p.plot_status_cd<>1 AND " + this.m_strStateCountySQL.Trim() + " " +
						"GROUP BY ppsa.statecd,ppsa.countycd,ppsa.plot;";
			}
			else
			{
				this.m_intError=-1;
			}

            if (this.m_intError==0)
			{
                try
                {
                    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                    {
                    con.Open();
                    oDataMgr.SqlQueryReader(con, strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        //load up each row in the FIADB plot input table
                        while (oDataMgr.m_DataReader.Read())
                        {
                            strState = "";
                            strCounty = "";
                            strPlot = "";

                            //make sure the row is not null values
                            if (oDataMgr.m_DataReader[0] != System.DBNull.Value &&
                                oDataMgr.m_DataReader[0].ToString().Trim().Length > 0)
                            {
                                strState = oDataMgr.m_DataReader["statecd"].ToString();
                                strCounty = oDataMgr.m_DataReader["countycd"].ToString();
                                strPlot = oDataMgr.m_DataReader["plot"].ToString();
                                this.lstFilterByPlot.BeginUpdate();
                                System.Windows.Forms.ListViewItem listItem = new ListViewItem();
                                listItem.Checked = false;
                                listItem.SubItems.Add(strState);
                                listItem.SubItems.Add(strCounty);
                                listItem.SubItems.Add(strPlot);
                                this.lstFilterByPlot.Items.Add(listItem);
                                this.lstFilterByPlot.EndUpdate();
                                intAddedPlotRows++;
                            }
                        }
                    }
                if (intAddedPlotRows == 0)
                {
                    this.m_intError = -1;
                    MessageBox.Show("!!No Plots Loaded To Get State, County, Plot Information!!", "Load State, County, Plot Menus");
                }
                ((frmDialog)this.ParentForm).Enabled = true;
            }
        }
        catch (Exception caught)
        {
            this.m_intError = -1;
            MessageBox.Show(caught.Message);
        }				
    }
    }


		private void btnFilterByStatePrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFilter.Visible=true;
			this.grpboxFilterByState.Visible=false;

		}

		private void btnFilterByStateCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFilterByStateNext_Click(object sender, System.EventArgs e)
		{
		    this.mdbInputPlot();
			if (this.m_intError==0)
			{
				this.grpboxFilterByPlot.Visible=true;
				this.grpboxFilterByState.Visible=false;
			} 
		}

		private void btnFilterByPlotPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFilterByPlot.Visible = false;
			this.grpboxFilterByState.Visible=true;
		}

		private void btnFilterByPlotCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();		
		}

		private void btnFilterByStateSelect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByState.Items.Count-1;x++)
			{
				this.lstFilterByState.Items[x].Checked=true;
			}
			
		}

		private void btnFilterByStateUnselect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByState.Items.Count-1;x++)
			{
				this.lstFilterByState.Items[x].Checked=false;
			}

		}

		private void btnFilterByPlotSelect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByPlot.Items.Count-1;x++)
			{
				this.lstFilterByPlot.Items[x].Checked=true;
			}

		}

		private void btnFilterByPlotUnselect_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstFilterByPlot.Items.Count-1;x++)
			{
				this.lstFilterByPlot.Items[x].Checked=false;
			}

		}

		private void btnFilterByStateFinish_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).MinimizeMainForm=true;
			this.Enabled=false;
			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";
            this.m_intError=0;
			if (this.lstFilterByState.CheckedItems.Count==0) 
			{
				MessageBox.Show("Select At Least One State, County Item","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.Enabled = true;
				return;
			}
            CalculateAdjustments_Start();
            if (m_intError == 0)
            {
                if (this.chkForested.Checked && this.chkNonForested.Checked)
                    this.BuildFilterByStateCountyString("statecd", "countycd", true);
                else
                    this.BuildFilterByStateCountyString("ppsa.statecd", "ppsa.countycd", true);
                if (this.m_intError == 0)
                {

                    this.LoadMDBPlotCondTreeData_Start();
                }
            }	
		}
		private void BuildFilterByStateCountyString(string strStateFieldAlias,string strCountyFieldAlias,bool bStringDataType)
		{
			string strCurState="";
			string strState="";
			string strCounty="";
			bool bAllCounties;
			string strStateList="";
			string strCountyList="";
			int y=0;

            System.Windows.Forms.ListView oLv = 
                (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(lstFilterByState, false);

            int intTotalCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);
            int intCheckedCount = (int)frmMain.g_oDelegate.GetListViewCheckedItemsCount(oLv, false);

			this.m_strStateCountyPlotSQL="";
			if (intTotalCount == intCheckedCount)
			{
				this.m_bAllCountiesSelected = true;
				bAllCounties=true;
			}
			else
			{
				this.m_bAllCountiesSelected = false;
				bAllCounties=false;
			}
			this.m_strStateCountySQL="";
			
			for (int x=0; x <= intTotalCount -1;x++)
			{
                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, x, "Checked", false))
				{
                    strState = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 1, "Text", false);
                    strCounty = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 2, "Text", false);
                    strState = strState.Trim(); strCounty = strCounty.Trim();

					if (strState.Trim().Length > 0)
					{
					}
					//check to see if this is a new state
					if (strState !=
						strCurState && strState.Trim().Length > 0)

					{

						if (strCurState.Trim().Length == 0)
						{
							//first time
							strCurState = strState;
						}
						if (this.m_bAllCountiesSelected == true)
						{
							if (strStateList.Trim().Length ==0)
							{
								if (bStringDataType == false)
								{
									strStateList = strState;
								}
								else
								{
									strStateList = "'" + strState.Trim() + "'";
								}
							}
							else
							{
								if (bStringDataType == false)
								{
									strStateList += "," + strState;
								}
								else
								{
									strStateList += ",'" + strState.Trim() + "'";
								}
							}
							strCurState=strState;
						}
						else
						{
							
							//current state
							//check if all counties for this state are selected
							strCountyList="";
							bAllCounties=true;

							//check to see if the previous list item is the same state and if
							//it is checked
							if (x-1 >=0)
							{
								if ((string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x-1,1,"Text",false).ToString().Trim() == strState.Trim() && 
                                    (bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x-1,"Checked",false)==false)
								{
									bAllCounties=false;
								}
							}
							
							for (y=x;y<=intTotalCount-1;y++)
							{
                                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, y, "Checked", false) == true)
								{

									if (strState.Trim() !=
                                        (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 1, "Text", false).ToString().Trim())
									{
										break;
									}
                                    strCounty = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 2, "Text", false).ToString().Trim();
									if (strCountyList.Trim().Length ==0)
									{
										if (bStringDataType == false)
										{
											strCountyList = strCounty;
										}
										else
										{
											strCountyList = "'" + strCounty.Trim() + "'";
										}
									}
									else
									{
										if (bStringDataType == false)
										{
											strCountyList += "," + strCounty;
										}
										else
										{
											strCountyList += ",'" + strCounty.Trim() + "'";
										}
									}

								}
								else
								{
                                    if (strState.Trim() == (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 1, "Text", false).ToString().Trim())
									{
										bAllCounties=false;
									}
								}
							}
							strCurState=strState;
							if (y<=intTotalCount-1)
							{
								x = y - 1;
							}
							else
							{
								x = y;
							}
							if (bAllCounties==true)
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
							}
							else
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
							}
						}

					}
				}
			}
			if (this.m_bAllCountiesSelected==true)
			{
				if (bStringDataType==false)
				{
				    this.m_strStateCountySQL = "(" + strStateFieldAlias + " IN (" + strStateList + "))";
				}
				else
				{
					this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") IN (" + strStateList + "))";
				}
			}
			else
			{
			    this.m_strStateCountySQL = "(" + this.m_strStateCountySQL + ")";
			}
		}
		private void BuildFilterByPlotString(string strStateFieldAlias,string strCountyFieldAlias,string strPlotFieldAlias, bool bStringDataType)
		{
			string strCurState="";
			string strCurCounty="";
			string strState="";
			string strCounty="";
			string strPlot="";
			bool bAllPlots;
			string strCountyList="";
			string strPlotList = "";
			int y=0;
         
			this.m_strStateCountySQL="";
            System.Windows.Forms.ListView oLv;

            oLv = (System.Windows.Forms.ListView)frmMain.g_oDelegate.GetListView(lstFilterByPlot, false);
            int intPlotCount = (int)frmMain.g_oDelegate.GetListViewItemsPropertyValue(oLv, "Count", false);

            if ((int)frmMain.g_oDelegate.GetListViewCheckedItemsCount(oLv, false) == 
                intPlotCount)
			{
				this.m_bAllPlotsSelected = true;
				bAllPlots=true;
			}
			else
			{
				this.m_bAllPlotsSelected = false;
				bAllPlots=false;
			}
			this.m_strStateCountyPlotSQL="";
			
			for (int x=0; x <= intPlotCount -1;x++)
			{
				if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x,"Checked",false))
				{
					strState = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x,1,"Text",false);
                    strCounty = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 2, "Text", false);
                    strPlot = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, x, 3, "Text", false);
                    strState = strState.Trim(); strCounty = strCounty.Trim(); strPlot = strPlot.Trim();
					if (strState.Trim().Length > 0)
					{
					}
					//check to see if this is a new state
					if ((strState !=
						strCurState && strState.Trim().Length > 0) ||
						(strCounty != strCurCounty && strCounty.Trim().Length > 0))

					{

						if (strCurState.Trim().Length == 0)
						{
							//first time
							strCurState = strState;
							strCurCounty = strCounty;
							strCountyList="";

						}
						if (this.m_bAllPlotsSelected == true)
						{
							if (strCurState.Trim() != strState.Trim())
							{
								if (this.m_strStateCountyPlotSQL.Trim().Length >0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}

								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}

								}
								strCountyList="";

							}

							if (strCountyList.Trim().Length ==0)
							{
								if (bStringDataType == false)
								{
									strCountyList = strCounty;
								}
								else
								{
									strCountyList = "'" + strCounty.Trim() + "'";
								}
							}
							else
							{
								if (bStringDataType == false)
								{
									strCountyList += "," + strCounty;
								}
								else
								{
									strCountyList += ",'" + strCounty.Trim() + "'";
								}
							}
							strCurState=strState;
							strCurCounty=strCounty;
						}
						else
						{
							
							//current state and county
							//check if all plots for this state and county are selected
							strCountyList="";
							strPlotList="";
							bAllPlots=true;

							//check to see if the previous list item is the same state and if
							//it is checked
							if (x-1 >=0)
							{
								if ((string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x-1,1,"Text",false).ToString().Trim() == strState.Trim() && 
                                    (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,x-1,2,"Text",false).ToString().Trim() == strCounty.Trim() && 
									(bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv,x-1,"Checked",false)==false)
								{
									bAllPlots=false;
								}
							}
							
							for (y=x;y<= intPlotCount-1;y++)
							{
                                if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(oLv, y, "Checked", false) == true)
								{

									if (strState.Trim() != 
                                          (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,1,"Text",false).ToString().Trim() ||
										strCounty.Trim() != (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,2,"Text",false).ToString().Trim())
									{
										break;
									}
                                    strPlot = (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv, y, 3, "Text", false).ToString().Trim();
									if (strPlotList.Trim().Length ==0)
									{
										if (bStringDataType == false)
										{
											strPlotList = strPlot;
										}
										else
										{
											strPlotList = "'" + strPlot.Trim() + "'";
										}
									}
									else
									{
										if (bStringDataType == false)
										{
											strPlotList += "," + strPlot;
										}
										else
										{
											strPlotList += ",'" + strPlot.Trim() + "'";
										}
									}

								}
								else
								{
									if (strState.Trim() == (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,1,"Text",false).ToString().Trim() && 
										strCounty.Trim() == (string)frmMain.g_oDelegate.GetListViewSubItemPropertyValue(oLv,y,2,"Text",false).ToString().Trim())
									{
										bAllPlots=false;
									}
								}
							}
							strCurState=strState;
							strCurCounty = strCounty;
							if (y<=intPlotCount-1)
							{
								x = y - 1;
							}
							else
							{
								x = y;
							}
							if (bAllPlots==true)
							{
								if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias  + " = " + strCurState + " AND " + strCountyFieldAlias + " = " + strCurCounty + ")";
									}
									else
									{
										this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "'  AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "')";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias  + " = " + strCurState + " AND " + strCountyFieldAlias + " = " + strCurCounty + ")";
									}
									else
									{
										this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "'  AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "')";
									}
								}
							}
							else
							{
								if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " = " + strCurCounty + " AND " + strPlotFieldAlias + " IN (" + strPlotList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "' AND trim(" + strPlotFieldAlias.Trim() + ") IN (" + strPlotList + "))";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " = " + strCurCounty + " AND " + strPlotFieldAlias + " IN (" + strPlotList + "))";
									}
									else
									{
										this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") = '" + strCurCounty.Trim() + "' AND trim(" + strPlotFieldAlias.Trim() + ") IN (" + strPlotList + "))";
									}
								}
							}
						}
					}
				}
			}
			if (this.m_bAllPlotsSelected == true)
			{
				
				if (this.m_strStateCountyPlotSQL.Trim().Length >0)
				{
					if (bStringDataType == false)
					{
						this.m_strStateCountyPlotSQL += " OR (" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
					}
					else
					{
						this.m_strStateCountyPlotSQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
					}

				}
				else
				{
					if (bStringDataType == false)
					{
						this.m_strStateCountyPlotSQL = "(" + strStateFieldAlias + " = " + strCurState + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
					}
					else
					{
						this.m_strStateCountyPlotSQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim() + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
					}

				}
			}
			this.m_strStateCountyPlotSQL = "(" + this.m_strStateCountyPlotSQL + ")";
			//MessageBox.Show(this.m_strStateCountyPlotSQL);
		
	
			
		}

		private void btnFilterByPlotFinish_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).MinimizeMainForm=true;
			this.Enabled=false;
            m_intError = 0;
			this.m_strStateCountyPlotSQL="";
			this.m_strStateCountySQL="";

			if (this.lstFilterByPlot.CheckedItems.Count==0) 
			{
				MessageBox.Show("Select At Least One State, County, Plot Item","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.Enabled = true;
				return;
			}
            CalculateAdjustments_Start();
            if (m_intError == 0)
            {
                this.BuildFilterByPlotString("ppsa.statecd", "ppsa.countycd", "ppsa.plot", false);
                if (this.m_intError == 0)
                {

                    this.LoadMDBPlotCondTreeData_Start();
                }
            }	
		}

		private void btnMDBInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnMDBPlotBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("SQLite Data File Containing Plot Table Data",
				"Select Plot Table",
				ref this.txtMDBPlot,
				ref this.txtMDBPlotTable);
		}

		private void btnMDBCondBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("SQLite Data File Containing Condition Table Data",
				"Select Condition Table",
				ref this.txtMDBCond,
				ref this.txtMDBCondTable);
		}

		private void btnMDBTreeBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("SQLite Data File Containing Tree Table Data",
				"Select Tree Table",
				ref this.txtMDBTree,
				ref this.txtMDBTreeTable);
		}

		private void btnIDBInvPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFilter.Visible=true;
			this.grpboxIDBInv.Visible=false;
		}

		private void btnIDBInvCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}
		private void idbLoadInv()
		{
            string strId="";
			string strDS="";
			string strDesc="";
			string strDef="";
			int intAddedRows=0;
    

			      
			this.lstIDBInv.Clear();
			this.lstIDBInv.Columns.Add("Id", 30, HorizontalAlignment.Left);
			this.lstIDBInv.Columns.Add("Data Source", 100, HorizontalAlignment.Left);
			this.lstIDBInv.Columns.Add("Description", 100, HorizontalAlignment.Left);
			this.lstIDBInv.Columns.Add("Definition", 100, HorizontalAlignment.Left);
            this.m_strIDBInv="";
			this.m_intError=0;

			

			FIA_Biosum_Manager.ado_data_access p_ado = new ado_data_access();

			
			//get all the project datasources

			string strMDBFile = ((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.ToString().Trim() + "\\db\\ref_master.mdb";
			string strConn = p_ado.getMDBConnString(strMDBFile,"","");
			p_ado.m_strSQL = "select * from inventories order by idb_data_source";

	

           
				p_ado.SqlQueryReader(strConn,p_ado.m_strSQL);
				if (p_ado.m_intError==0)
				{
					try
					{
						//load up each row in the FIADB plot input table

						while (p_ado.m_OleDbDataReader.Read())
						{
							strDS="";
							strDesc="";
							strDef="";
							
							//make sure the row is not null values
							if (p_ado.m_OleDbDataReader[0] != System.DBNull.Value &&
								p_ado.m_OleDbDataReader[0].ToString().Trim().Length > 0)
							{
								strId =p_ado.m_OleDbDataReader["inv_id"].ToString();
								strDS= p_ado.m_OleDbDataReader["idb_data_source"].ToString();
								strDesc = p_ado.m_OleDbDataReader["description"].ToString();
								strDef = p_ado.m_OleDbDataReader["inv_id_def"].ToString();
								this.lstIDBInv.BeginUpdate();
								System.Windows.Forms.ListViewItem listItem = new ListViewItem();
								listItem.Text=strId;
								listItem.SubItems.Add(strDS);
								listItem.SubItems.Add(strDesc);
								listItem.SubItems.Add(strDef);
								this.lstIDBInv.Items.Add(listItem);
								this.lstIDBInv.EndUpdate();
								intAddedRows++;
							}

						}
						p_ado.m_OleDbDataReader.Close();
						if (intAddedRows == 0 )
						{
							this.m_intError=-1;
							MessageBox.Show("!!No Inventories Loaded !","Load Inventories");
						}
						else
						{
							this.lstIDBInv.Columns[0].Width = -1;
							this.lstIDBInv.Columns[1].Width = -1;
							this.lstIDBInv.Columns[2].Width = -1;
							this.lstIDBInv.Columns[3].Width = -1;
						}
						((frmDialog)this.ParentForm).Enabled=true;
					}
					catch (Exception caught)
					{
						this.m_intError=-1;
						MessageBox.Show(caught.Message);
					}
					p_ado.m_OleDbConnection.Close();
				}
				else
				{
					this.m_intError=p_ado.m_intError;
				}
			
			
			
			p_ado=null;
		}

		private void FIADBLoadInv()
		{
			string strEvalId="";
			string strRsCd="";
			string strStateCd="";
			string strLocNm="";
			string strEvalDesc="";
			string strRptYr="";
			string strNotes="";
			int intAddedRows=0;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.FIADBLoadInv\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }


            this.lstFIADBInv.Clear();
			this.lstFIADBInv.Columns.Add("EvalId", 50, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("RsCd", 30, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("StateCd", 50, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("Location_Nm", 75, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("Eval_Descr", 400, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("ReportYear", 300, HorizontalAlignment.Left);
			this.lstFIADBInv.Columns.Add("Notes", 125, HorizontalAlignment.Left);
//			this.m_strFIADBEvalId="";
//			this.m_strFIADBRsCd="";
			this.m_intError=0;

            string strConn = SQLite.GetConnectionString(m_strTempDbFile);
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                con.Open();
                SQLite.m_strSQL = "SELECT * FROM " + this.m_strPopEvalTable + " where biosum_status_cd=9 order by statecd,evalid";
                SQLite.SqlQueryReader(strConn, SQLite.m_strSQL);
                if (SQLite.m_intError == 0)
                {
                    try
                    {
                        //load up each row in the pop_eval table
                        while (SQLite.m_DataReader.Read())
                        {
                            strEvalId = "";
                            strRsCd = "";
                            strStateCd = "";
                            strLocNm = "";
                            strEvalDesc = "";
                            strRptYr = "";
                            strNotes = "";

                            //make sure the row is not null values
                            if (SQLite.m_DataReader[0] != System.DBNull.Value &&
                                SQLite.m_DataReader[0].ToString().Trim().Length > 0)
                            {
                                strEvalId = SQLite.m_DataReader["evalid"].ToString();
                                strRsCd = SQLite.m_DataReader["RsCd"].ToString();
                                strStateCd = SQLite.m_DataReader["statecd"].ToString();
                                if (SQLite.m_DataReader["location_nm"] != System.DBNull.Value)
                                    strLocNm = SQLite.m_DataReader["location_nm"].ToString();
                                if (SQLite.m_DataReader["eval_descr"] != System.DBNull.Value)
                                    strEvalDesc = SQLite.m_DataReader["eval_descr"].ToString();
                                if (SQLite.m_DataReader["report_year_nm"] != System.DBNull.Value)
                                    strRptYr = SQLite.m_DataReader["report_year_nm"].ToString();
                                if (SQLite.m_DataReader["notes"] != System.DBNull.Value)
                                    strNotes = SQLite.m_DataReader["notes"].ToString();

                                this.lstFIADBInv.BeginUpdate();
                                System.Windows.Forms.ListViewItem listItem = new ListViewItem();
                                listItem.Text = strEvalId;
                                listItem.SubItems.Add(strRsCd);
                                listItem.SubItems.Add(strStateCd);
                                listItem.SubItems.Add(strLocNm);
                                listItem.SubItems.Add(strEvalDesc);
                                listItem.SubItems.Add(strRptYr);
                                listItem.SubItems.Add(strNotes);
                                this.lstFIADBInv.Items.Add(listItem);
                                this.lstFIADBInv.EndUpdate();
                                intAddedRows++;
                            }

                        }
                        SQLite.m_DataReader.Close();
                        if (intAddedRows == 0)
                        {
                            this.m_intError = -1;
                            MessageBox.Show("!!No Inventories Loaded !", "Load Inventories");
                        }
                        ((frmDialog)this.ParentForm).Enabled = true;
                    }
                    catch (Exception caught)
                    {
                        this.m_intError = -1;
                        MessageBox.Show(caught.Message);
                    }
                }
                else
                {
                    this.m_intError = SQLite.m_intError;
                }
            }
		}


		private void btnIDBInvSelectAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstIDBInv.Items.Count-1;x++)
			{
				this.lstIDBInv.Items[x].Checked=true;
			}
		}

		private void btnIDBInvClearAll_Click(object sender, System.EventArgs e)
		{
			for (int x=0; x<= this.lstIDBInv.Items.Count-1;x++)
			{
				this.lstIDBInv.Items[x].Checked=false;
			}
		}


		private void BuildFilterByIDBInvString(string strStateFieldAlias,string strCountyFieldAlias,bool bStringDataType)
		{
		
			//string strCurInv;
			string strCurState="";
			//string strInv="";
			string strState="";
			string strCounty="";
			bool bAllCounties;
			string strStateList="";
			string strCountyList="";
			int y=0;
         
			this.m_strIDBInv="";
			if (this.lstFilterByState.CheckedItems.Count == this.lstFilterByState.Items.Count)
			{
				this.m_bAllCountiesSelected = true;
				bAllCounties=true;
			}
			else
			{
				this.m_bAllCountiesSelected = false;
				bAllCounties=false;
			}
			this.m_strStateCountySQL="";
			
			for (int x=0; x <= this.lstFilterByState.Items.Count -1;x++)
			{
				if (this.lstFilterByState.Items[x].Checked==true)
				{
					strState = this.lstFilterByState.Items[x].SubItems[1].Text.Trim();
					strCounty = this.lstFilterByState.Items[x].SubItems[2].Text.Trim();
					if (strState.Trim().Length > 0)
					{
					}
					//check to see if this is a new state
					if (strState !=
						strCurState && strState.Trim().Length > 0)

					{

						if (strCurState.Trim().Length == 0)
						{
							//first time
							strCurState = strState;
						}
						if (this.m_bAllCountiesSelected == true)
						{
							if (strStateList.Trim().Length ==0)
							{
								if (bStringDataType == false)
								{
									strStateList = strState;
								}
								else
								{
									strStateList = "'" + strState.Trim() + "'";
								}
							}
							else
							{
								if (bStringDataType == false)
								{
									strStateList += "," + strState;
								}
								else
								{
									strStateList += ",'" + strState.Trim() + "'";
								}
							}
							strCurState=strState;
						}
						else
						{
							
							//current state
							//check if all counties for this state are selected
							strCountyList="";
							bAllCounties=true;

							//check to see if the previous list item is the same state and if
							//it is checked
							if (x-1 >=0)
							{
								if (this.lstFilterByState.Items[x-1].SubItems[1].Text.Trim() == 
									strState.Trim() && this.lstFilterByState.Items[x-1].Checked==false)
								{
									bAllCounties=false;
								}
							}
							
							for (y=x;y<=this.lstFilterByState.Items.Count-1;y++)
							{
								if (this.lstFilterByState.Items[y].Checked==true)
								{

									if (strState.Trim() != this.lstFilterByState.Items[y].SubItems[1].Text.Trim())
									{
										break;
									}
									strCounty =  this.lstFilterByState.Items[y].SubItems[2].Text.Trim();
									if (strCountyList.Trim().Length ==0)
									{
										if (bStringDataType == false)
										{
											strCountyList = strCounty;
										}
										else
										{
											strCountyList = "'" + strCounty.Trim() + "'";
										}
									}
									else
									{
										if (bStringDataType == false)
										{
											strCountyList += "," + strCounty;
										}
										else
										{
											strCountyList += ",'" + strCounty.Trim() + "'";
										}
									}

								}
								else
								{
									if (strState.Trim() == this.lstFilterByState.Items[y].SubItems[1].Text.Trim())
									{
										bAllCounties=false;
									}
								}
							}
							strCurState=strState;
							if (y<=this.lstFilterByState.Items.Count-1)
							{
								x = y - 1;
							}
							else
							{
								x = y;
							}
							if (bAllCounties==true)
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias  + " = " + strCurState + ")";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim()  + ") = '" + strCurState.Trim() + "')";
									}
								}
							}
							else
							{
								if (this.m_strStateCountySQL.Trim().Length > 0)
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL += " OR (" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL += " OR ( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
								else
								{
									if (bStringDataType == false)
									{
										this.m_strStateCountySQL = "(" + strStateFieldAlias + " = " + strCurState  + " AND " + strCountyFieldAlias + " IN (" + strCountyList + "))";
									}
									else
									{
										this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") = '" + strCurState.Trim()  + "' AND trim(" + strCountyFieldAlias.Trim() + ") IN (" + strCountyList + "))";
									}
								}
							}
						}

					}
				}
			}
			if (this.m_bAllCountiesSelected==true)
			{
				if (bStringDataType==false)
				{
					this.m_strStateCountySQL = "(" + strStateFieldAlias + " IN (" + strStateList + "))";
				}
				else
				{
					this.m_strStateCountySQL = "( trim(" + strStateFieldAlias.Trim() + ") IN (" + strStateList + "))";
				}
			}
			else
			{
				this.m_strStateCountySQL = "(" + this.m_strStateCountySQL + ")";
			}
		}

		private void lstIDBInv_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			
		}

		private void btnIDBInvNext_Click(object sender, System.EventArgs e)
		{
			if (this.lstIDBInv.SelectedItems.Count > 0)
			{
				this.m_strIDBInv = this.lstIDBInv.SelectedItems[0].Text.Trim();
				this.mdbInputStateCounty();
				this.grpboxIDBInv.Visible=false;
				this.grpboxFilterByState.Visible=true;
			}
			else
			{
				MessageBox.Show("Select A PNW IDB Inventory","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
		}

		private void btnIDBInvAppend_Click(object sender, System.EventArgs e)
		{
			if (this.lstIDBInv.SelectedItems.Count > 0)
			{
				
				
				this.m_strIDBInv = this.lstIDBInv.SelectedItems[0].Text.Trim();
                this.LoadIDBPlotCondTreeData_Start();
			}
			else
			{
				MessageBox.Show("Select A PNW IDB Inventory","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}

		
		}
        private void LoadIDBPlotCondTreeData_Process()
		{
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
			string strFields="";
			string strValues="";
			int intAddedPlotRows=0;
			int intAddedCondRows=0;
			int intAddedTreeRows=0;
			int intAddedSiteTreeRows=0;
			int x=0;
			int y=0;
			string strCol="";
			string strTime="";
            string str = "";
            string str2 = "";



			this.m_intError=0;	
						
			
            
			try
			{
				//-----------PREPARATION FOR ADDING PLOT RECORDS---------//
                    
				//instatiate the oledb data access class
				this.m_ado = new ado_data_access();


                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);




				//get all the project datasources
                FIA_Biosum_Manager.Datasource p_datasource = new Datasource();
                p_datasource.m_strDataSourceMDBFile = this.ReferenceFormDialog.m_frmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
                p_datasource.LoadTableColumnNamesAndDataTypes = false;
                p_datasource.LoadTableRecordCount = false;
                p_datasource.m_strDataSourceTableName = "datasource";
                p_datasource.m_strScenarioId = "";
                p_datasource.populate_datasource_array();

				//create a temporary mdb file with links to all the project tables
				//and return the name of the file that contains the links
				this.m_strTempMDBFile = p_datasource.CreateMDBAndTableDataSourceLinks();

			
				//instatiate dao for creating links in the temp table
				//to the fiadb plot, cond, and tree input tables
				dao_data_access p_dao1 = new dao_data_access();
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Creating Datasource Links");

				//create links to the idb input tables in the temp mdb file
                //create links to the fiadb input tables in the temp mdb file
                //plot table
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBPlot, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBPlotTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_plot_input", str.Trim(), str2.Trim());
                //cond table
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBCond, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBCondTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_cond_input", str.Trim(), str2.Trim());
                //tree table
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBTree, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBTreeTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_tree_input", str.Trim(), str2.Trim());
                //site tree
                str = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBSiteTree, "Text", false);
                str2 = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtMDBSiteTreeTable, "Text", false);
                p_dao1.CreateTableLink(this.m_strTempMDBFile, "idb_site_tree_input", str.Trim(), str2.Trim());

				//destroy the object and release it from memory
				p_dao1.m_DaoWorkspace.Close();
				p_dao1 = null;

                SetThermValue(m_frmTherm.progressBar1, "Value", 20);

				//get an ado connection string for the temp mdb file
				this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");

				//get the biosum plot, cond, and tree table names
				this.m_strPlotTable = p_datasource.getValidDataSourceTableName("PLOT");
				this.m_strCondTable = p_datasource.getValidDataSourceTableName("CONDITION");
				this.m_strTreeTable = p_datasource.getValidDataSourceTableName("TREE");
				this.m_strSiteTreeTable = p_datasource.getValidDataSourceTableName("SITE TREE");

   					
				//create a new connection to the temp MDB file
				this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();

				//open the connection to the temp mdb file 
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				//get the fiabiosum table structures
				System.Data.DataTable p_dtPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strPlotTable);
				System.Data.DataTable p_dtCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strCondTable);
				System.Data.DataTable p_dtTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strTreeTable);
				System.Data.DataTable p_dtSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from " + this.m_strSiteTreeTable);
				//get the idb table structures
				System.Data.DataTable p_dtIDBPlotSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_plot_input");
				System.Data.DataTable p_dtIDBCondSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_cond_input");
				System.Data.DataTable p_dtIDBTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_tree_input");
				System.Data.DataTable p_dtIDBSiteTreeSchema = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_site_tree_input");

				System.Data.DataTable p_dtTreeWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_tree_input");
				System.Data.DataTable p_dtSiteTreeWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile, "select * from idb_site_tree_input");



                SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                

             
				this.m_strSQL = "SELECT biosum_plot_id, statecd as cond_ttl " + 
					"FROM " + this.m_strPlotTable.Trim() + ";";


				/****************************************************************
				 **get the table structure that results from executing the sql
				 ****************************************************************/
				System.Data.DataTable p_dtPlotWorkTable = this.m_ado.getTableSchema(this.m_connTempMDBFile,this.m_strSQL);

				
                this.m_ado.m_strSQL = "SELECT * INTO plot_column_updates_work_table " + 
					                  "FROM (SELECT biosum_plot_id, statecd as cond_ttl " + 
											"FROM " + this.m_strPlotTable.Trim() + ") "  + 
					                  "WHERE 1 = 2";

				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

				this.m_ado.m_strSQL = "SELECT biosum_plot_id INTO temp_plot_table FROM " + this.m_strPlotTable + " WHERE 1=2";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

				this.m_ado.m_strSQL = "select * INTO temp_tree_input FROM idb_tree_input WHERE 1 = 2";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

				this.m_ado.m_strSQL = "select * INTO tree_work_table FROM idb_tree_input WHERE 1 = 2";
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                SetThermValue(m_frmTherm.progressBar2, "Value", 20);
                
				System.Threading.Thread.Sleep(2000);

				//-------------------------------PLOT----------------------------------//
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg,"Text","Plot Table: Inserting Plot Records...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

				//build field list string to insert sql by matching 
				//up the column names in the biosum plot table and the fiadb plot table
				strFields = "";
				for (x=0; x<=p_dtPlotSchema.Rows.Count-1;x++)
				{
					strCol = p_dtPlotSchema.Rows[x]["columnname"].ToString().Trim();
					//see if there is an equivalent FIADB column
					for (y=0; y<=p_dtIDBPlotSchema.Rows.Count-1;y++)
					{
						if (strCol.Trim().ToUpper() == p_dtIDBPlotSchema.Rows[y]["columnname"].ToString().ToUpper())
						{
							if (strFields.Trim().Length == 0)
							{
								strFields = strCol;
							}
							else
							{	
								strFields += "," + strCol;
							}
							break;
						}
					}
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                /********************************************************
				 **create plot input insert command
				 ********************************************************/
				//check the user defined filters
				if (Checked(rdoFilterNone)==true)
				{
					//forested/nonforested filters
					if (Checked(chkNonForested)==true && Checked(chkForested)==true)
					{
						//all plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " " + 
							" FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' ";
					}
					else if (Checked(chkForested)==true)
					{
						//forested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							" Plot_status_cd = 1;";

					}
					else
					{
						//nonforested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							"plot_status_cd IS NULL or plot_status_cd <> 1;";
					}
				}
				else if (Checked(rdoFilterByFile) == true)
				{
					//user defined plot_cn filter
					this.m_strSQL = "INSERT INTO  " + this.m_strPlotTable + " (" + strFields + ")" + 
						" SELECT " + strFields + " FROM idb_plot_input WHERE idb_plot_id  IN (" + this.m_strPlotIdList.Trim() + ");";
				}
				else if (Checked(rdoFilterByMenu)==true) 
				{
					if (Checked(chkNonForested)==true && Checked(chkForested)==true)
					{
						//all plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND ";
					}
					else if (Checked(chkForested)==true)
					{
						//forested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM idb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							       "(plot_status_cd = 1) AND ";
					}
					else
					{
						//nonforested plots
						this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
							" SELECT " + strFields + " FROM fiadb_plot_input " + 
							" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "' AND " + 
							       "(plot_status_cd IS NULL or plot_status_cd <> 1) AND ";
					}
					if (this.m_strStateCountyPlotSQL.Trim().Length > 0)
					{
						//state,county,plot filter
						this.m_strSQL += this.m_strStateCountyPlotSQL.Trim() + ";";
					}
					else
					{
						//state,county filter
						this.m_strSQL += this.m_strStateCountySQL.Trim() + ";";
					}

				}
				else
				{
					this.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (" + strFields + ")" + 
						" SELECT " + strFields + " FROM fiadb_plot_input " + 
						" WHERE MID(biosum_plot_id,2,4)='" + this.m_strIDBInv.Trim() + "';";
				
				}
				this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
                SetThermValue(m_frmTherm.progressBar1, "Value", 75);

				if (this.m_ado.m_intError == 0)
				{
					this.m_ado.m_strSQL="INSERT INTO temp_plot_table SELECT biosum_plot_id FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9";
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);

					intAddedPlotRows = Convert.ToInt32(this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strPlotTable + " where biosum_status_cd = 9;",this.m_strPlotTable));
				}
				else
				{
					this.m_intError=this.m_ado.m_intError;
					//error occured so remove new plot records
					this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 40);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				//-------------------------------COND----------------------------------//
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Cond Table: Inserting Condition Records...Stand By");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)m_frmTherm,"Refresh");

					//build field list string to insert sql by matching 
					//up the column names in the biosum cond table and the fiadb cond table
					strFields = "";
					strValues = "";
					for (x=0; x<=p_dtCondSchema.Rows.Count-1;x++)
					{
						strCol = p_dtCondSchema.Rows[x]["columnname"].ToString().Trim();
						//see if there is an equivalent FIADB column
						for (y=0; y<=p_dtIDBCondSchema.Rows.Count-1;y++)
						{
							if (strCol.Trim().ToUpper() == p_dtIDBCondSchema.Rows[y]["columnname"].ToString().ToUpper())
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = strCol;
									strValues = "c." + strCol.Trim();
									//strValues = strCol;
								}
								else
								{	
									strFields += "," + strCol;
									strValues += ",c." + strCol.Trim();
									//strValues += "," + strCol.Trim();
								}
								break;
							}
						}
					}
                    SetThermValue(m_frmTherm.progressBar1, "Value", 40);
				
					//create cond input insert command
				
					

					this.m_strSQL = "INSERT INTO " + this.m_strCondTable + " (" + strFields + ")" + 
						" SELECT " + strValues + " FROM idb_cond_input c INNER JOIN " + this.m_strPlotTable + " p ON c.idb_plot_id=p.idb_plot_id WHERE p.biosum_status_cd=9";

					strTime = System.DateTime.Now.ToString();
				
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 80);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
					
					this.m_ado.m_strSQL = "UPDATE " + this.m_strCondTable + " " + 
						     "SET owngrpcd = IIF(owngrpcd IS NOT NULL AND owngrpcd < 10, (owngrpcd * 10),owngrpcd)";

					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_ado.m_strSQL);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 85);

				

					if (this.m_ado.m_intError != 0)
					{
						//remove new plot and cond records since error occured
						this.m_intError=this.m_ado.m_intError;
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					}

					//delete plot records from the input table
					//that already exist in the biosum plot table
					strTime = System.DateTime.Now.ToString();

				
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				//-------------------------------TREE----------------------------------//
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Tree Table: Inserting Tree Records...Stand By");
					frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
					
					strTime = System.DateTime.Now.ToString();
					//copy the inv trees to a temp file
					if (this.m_strPlotIdList.Trim().Length > 0)
					{
						this.m_strSQL = " INSERT INTO tree_work_table SELECT * FROM idb_tree_input";
					}
					else
					{
						this.m_strSQL = " INSERT INTO tree_work_table SELECT * FROM idb_tree_input WHERE MID(biosum_cond_id,2,4) = '" + this.m_strIDBInv.Trim() + "'";
					}
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

                    SetThermValue(m_frmTherm.progressBar1, "Value", 10);

					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
				
					/***************************************************
					 **get a list of fields that are not in the 
					 **biosum tree table and delete them from the
					 **tree work table
					 ***************************************************/
					string[] strFieldList;
					strFieldList = new string[p_dtIDBTreeSchema.Rows.Count];

					for (x=0; x<=p_dtIDBTreeSchema.Rows.Count-1;x++)
					{
						strFieldList[x]="";
						strCol = p_dtIDBTreeSchema.Rows[x]["columnname"].ToString().Trim();
						//see if there is an equivalent FIADB column
						for (y=0; y<=p_dtTreeSchema.Rows.Count-1;y++)
						{
							if (strCol.Trim().ToUpper() == p_dtTreeSchema.Rows[y]["columnname"].ToString().ToUpper())
							{
								break;
							}
						}
						if (y > p_dtTreeSchema.Rows.Count-1)
							strFieldList[x]=strCol;
					}
                    SetThermValue(m_frmTherm.progressBar1, "Value", 20);

					//close the connection to the temp mdb file
					this.m_connTempMDBFile.Close();
					while (m_connTempMDBFile.State == System.Data.ConnectionState.Open)
						System.Threading.Thread.Sleep(1000);
					m_connTempMDBFile.Dispose();
					m_connTempMDBFile=null;
					


					//instatiate dao again to delete the fields from the tree work table
					//in the temp MDB file
					dao_data_access p_dao3 = new dao_data_access();

					p_dao3.DeleteField(this.m_strTempMDBFile,"tree_work_table", strFieldList);

					p_dao3 = null;
                    SetThermValue(m_frmTherm.progressBar1, "Value",30);
					

					

					//reopen the connection to the temp mdb file 
					this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();
					this.m_ado.OpenConnection(this.m_strTempMDBFileConn,ref this.m_connTempMDBFile);

					strTime = System.DateTime.Now.ToString();
					//copy the inv trees to a temp file
					this.m_strSQL = " INSERT INTO " + this.m_strTreeTable + " SELECT t.* FROM tree_work_table t INNER JOIN " + this.m_strCondTable + " c ON t.biosum_cond_id = c.biosum_cond_id WHERE c.biosum_status_cd=9;";
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					strTime += " " + System.DateTime.Now.ToString();
					//MessageBox.Show(strTime);
                    SetThermValue(m_frmTherm.progressBar1, "Value", 75);
					if (this.m_ado.m_intError != 0)
					{
						//remove new plot and cond records since error occured
						this.m_intError=this.m_ado.m_intError;
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					}
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 80);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				//-------------------------------SITE TREE----------------------------------//
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
                    SetLabelValue(m_frmTherm.lblMsg,"Text","Site Tree Table: Inserting Site Tree Records...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");


					//build field list string to insert sql by matching 
					//up the column names in the biosum cond table and the fiadb cond table
					
					strFields = "";
					strValues = "";
					for (x=0; x<=p_dtSiteTreeSchema.Rows.Count-1;x++)
					{
						strCol = p_dtSiteTreeSchema.Rows[x]["columnname"].ToString().Trim();
						//see if there is an equivalent FIADB column
						for (y=0; y<=p_dtIDBSiteTreeSchema.Rows.Count-1;y++)
						{
							if (strCol.Trim().ToUpper() == p_dtIDBSiteTreeSchema.Rows[y]["columnname"].ToString().ToUpper())
							{
								if (strFields.Trim().Length == 0)
								{
									strFields = strCol;
									strValues = "t." + strCol.Trim();
								}
								else
								{	
									strFields += "," + strCol;
									strValues += ",t." + strCol.Trim();
								}
								break;
							}
						}
					}
                    SetThermValue(m_frmTherm.progressBar1, "Value", 50);
					//create cond input insert command

					this.m_strSQL = "INSERT INTO " + this.m_strSiteTreeTable + " (" + strFields + ")" + 
						" SELECT " + strValues + " FROM idb_site_tree_input t INNER JOIN temp_plot_table p ON TRIM(t.biosum_plot_id)=p.biosum_plot_id";

					strTime = System.DateTime.Now.ToString();
				
					this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

					if (this.m_ado.m_intError != 0)
					{
						//remove new plot and cond records since error occured
						this.m_intError=this.m_ado.m_intError;
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd=9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
					}


					

					
				}
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
				System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 95);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
				if (intAddedPlotRows > 0 && this.m_intError==0 && this.m_ado.m_intError==0)
				{
				    SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Columns...Stand By");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
				

					this.UpdateColumns(this.m_ado);
					

					if (this.m_intError==0)
					{
				

						intAddedPlotRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strPlotTable + " WHERE biosum_status_cd=9",this.m_strPlotTable);
						intAddedCondRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strCondTable  + " WHERE biosum_status_cd=9",this.m_strCondTable);
						intAddedTreeRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strTreeTable  + " WHERE biosum_status_cd=9",this.m_strTreeTable);
						intAddedSiteTreeRows= (int)this.m_ado.getRecordCount(this.m_connTempMDBFile,"select count(*) from " + this.m_strSiteTreeTable  + " WHERE biosum_status_cd=9",this.m_strSiteTreeTable);


						this.m_strSQL = " UPDATE " + this.m_strPlotTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strCondTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = " UPDATE " + this.m_strSiteTreeTable + " SET biosum_status_cd=1 WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);


                        SetThermValue(m_frmTherm.progressBar1, "Value",GetThermValue(m_frmTherm.progressBar1,"Maximum"));
                        SetLabelValue(m_frmTherm.lblMsg,"Text","Done");
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

						MessageBox.Show("Successfully Appended \n" + 
							intAddedPlotRows.ToString().Trim() + " Plot Records \n" + 
							intAddedCondRows.ToString().Trim() + " Condition Records \n" + 
							intAddedTreeRows.ToString().Trim() + " Tree Records \n" + 
							intAddedSiteTreeRows.ToString().Trim() + " Site Tree Records","Add Plot Data");

					}
					else
					{

						//error occurred in the updatecolumns so delete the records
						this.m_strSQL = "DELETE FROM " + this.m_strPlotTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strCondTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						//delete added cond records since error occured
						this.m_strSQL = "DELETE FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);
						this.m_strSQL = "DELETE FROM " + this.m_strSiteTreeTable + " WHERE biosum_status_cd = 9;";
						this.m_ado.SqlNonQuery(this.m_connTempMDBFile,this.m_strSQL);

                        SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                        frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

					}
				}
				this.m_strCurrentProcess="";
                SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                SetThermValue(m_frmTherm.progressBar2, "Value", 100);		
				if (this.m_intError != 0 || this.m_ado.m_intError != 0)
					MessageBox.Show("!!Error Occured Adding Plot Records: 0 Records Added!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				
				if (this.m_connTempMDBFile.State == System.Data.ConnectionState.Open)
				{
					this.m_connTempMDBFile.Close();
					while (m_connTempMDBFile.State == System.Data.ConnectionState.Open)
						System.Threading.Thread.Sleep(1000);
					m_connTempMDBFile.Dispose();
					m_connTempMDBFile=null;
				}
				this.m_ado=null;
                
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);


                LoadMDBPlotCondTreeData_Finish();

            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (this.m_connTempMDBFile != null)
                {
                    if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(m_connTempMDBFile);
                    }
                    m_connTempMDBFile = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.CancelThreadCleanup();
                this.ThreadCleanUp();
                this.CleanupThread();

            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_plot_input:LoadIDBPlotCondTreeData_Process  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {

            }

            if (this.m_connTempMDBFile != null)
            {
                if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(m_connTempMDBFile);
                }
                m_connTempMDBFile = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);




            LoadMDBPlotCondTreeData_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
			
		
		
		}
        private bool Checked(System.Windows.Forms.RadioButton p_rdoButton)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)p_rdoButton, "Checked", false);
        }
        private bool Checked(System.Windows.Forms.CheckBox p_chkBox)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)p_chkBox, "Checked", false);
        }
        private void SetThermValue(System.Windows.Forms.ProgressBar p_oPb,string p_strPropertyName, int p_intValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)p_oPb,p_strPropertyName,(int)p_intValue);
        }
        private int GetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName)
        {
            return (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)p_oPb, p_strPropertyName, false);
        }
        private bool GetBooleanValue(System.Windows.Forms.Control p_oControl, string p_strPropertyName)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)p_oControl,p_strPropertyName,false);
        }
            
            
            
        private void SetLabelValue(System.Windows.Forms.Label p_oLabel, string p_strPropertyName, string p_strValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)p_oLabel, p_strPropertyName, p_strValue);
        }

		private void StartTherm(string p_strNumberOfTherms,string p_strTitle)
		{
			this.m_frmTherm = new frmTherm((frmDialog)this.ParentForm, p_strTitle);

			this.m_frmTherm.Text = p_strTitle;
			this.m_frmTherm.lblMsg.Text="";
			this.m_frmTherm.lblMsg2.Text="";
			this.m_frmTherm.Visible=false;
			this.m_frmTherm.btnCancel.Visible=true;
			this.m_frmTherm.lblMsg.Visible=true;
			this.m_frmTherm.progressBar1.Minimum=0;
			this.m_frmTherm.progressBar1.Visible=true;
			this.m_frmTherm.progressBar1.Maximum = 10;

			if (p_strNumberOfTherms=="2")
			{
				this.m_frmTherm.progressBar2.Size = this.m_frmTherm.progressBar1.Size;
				this.m_frmTherm.progressBar2.Left = this.m_frmTherm.progressBar1.Left;
				this.m_frmTherm.progressBar2.Top = Convert.ToInt32(this.m_frmTherm.progressBar1.Top + (this.m_frmTherm.progressBar1.Height * 3));
				this.m_frmTherm.lblMsg2.Top = this.m_frmTherm.progressBar2.Top + this.m_frmTherm.progressBar2.Height + 5;
				this.m_frmTherm.Height = this.m_frmTherm.lblMsg2.Top + this.m_frmTherm.lblMsg2.Height + this.m_frmTherm.btnCancel.Height + 50;
				this.m_frmTherm.btnCancel.Top = this.m_frmTherm.ClientSize.Height - this.m_frmTherm.btnCancel.Height - 5;
				this.m_frmTherm.lblMsg2.Show();
				this.m_frmTherm.progressBar2.Visible=true;
			}
			this.m_frmTherm.AbortProcess = false;
			this.m_frmTherm.Refresh();
			this.m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			//((frmDialog)this.ParentForm).Enabled=false;
			this.m_frmTherm.Visible=true;
			
		}
        public void StopThread()
        {

            string strMsg = "";
            
            frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel adding plot data?");

            if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"AbortProcess",true);
                this.CancelThreadCleanup();
                this.ThreadCleanUp();
            }


        }
		
		public void AddPlotRecordsFinished()
		{
			this.m_strPlotIdList="";
			this.thdProcessRecords.Abort();
			if (this.thdProcessRecords.IsAlive)
			{
				frmMain.g_oDelegate.SetControlPropertyValue(
                    (System.Windows.Forms.Label)m_frmTherm.lblMsg,"Text","Attempting To Abort Process...Stand By");
                frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Label)this.m_frmTherm.lblMsg, "Refresh");
				this.thdProcessRecords.Join(1000);
			}
			this.thdProcessRecords = null;
		}
		private void CancelThreadCleanup()
		{
			try
			{

				if (this.m_ado == null)
				{
					this.m_ado = new ado_data_access();
				}
				else
				{
					this.m_ado = null;
					this.m_ado = new ado_data_access();
				}
				if (this.m_connTempMDBFile == null)
				{

				}
				else
				{

					this.m_connTempMDBFile = null;
				}
				this.m_ado.OpenConnection(this.m_strTempMDBFileConn);


				if (this.m_strCurrentProcess=="mdbFIADBFileInput" ||
					this.m_strCurrentProcess=="txtFileInput" ||
					this.m_strCurrentProcess=="mdbIDBFileInput")
				{
                    DeleteFromTablesWhereFilter(m_ado, m_ado.m_OleDbConnection,
                        new string[]
                        {
                            //m_strPlotTable, m_strCondTable, m_strTreeTable, m_strSiteTreeTable, m_strDwmCwdTable,
                            //m_strDwmFwdTable, m_strDwmDuffLitterTable, m_strDwmTransectSegmentTable
                            m_strPlotTable, m_strCondTable, m_strTreeTable, m_strSiteTreeTable, m_strDwmCwdTable + "_master",
                            m_strDwmFwdTable+ "_master", m_strDwmDuffLitterTable+ "_master", m_strDwmTransectSegmentTable+ "_master"
                        },
                        " WHERE biosum_status_cd=9");
				}

				MessageBox.Show("!!User Canceled Adding Plot Records: 0 Records Added!!","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"Visible",false);
				//((frmDialog)this.ParentForm).m_frmMain.Visible=true;
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);



			}
			catch 
			{
			}

		}

		private void btnFIADBInvAppend_Click(object sender, System.EventArgs e)
		{
            m_intError = 0;
			((frmDialog)this.ParentForm).MinimizeMainForm=true;
			this.Enabled=false;
			if (this.lstFIADBInv.SelectedItems.Count > 0)
			{
                this.CalculateAdjustments_Start();
                if (m_intError == 0)
                {
                    this.LoadMDBFiadbPopFiles();
                    this.m_strLoadedPopEstUnitTxtInputFile = "";
                    this.m_strLoadedPopEvalTxtInputFile = "";
                    this.m_strLoadedPopStratumTxtInputFile = "";
                    this.m_strLoadedPpsaTxtInputFile = "";
                    if (this.rdoFilterNone.Checked && m_intError == 0)
                    {

                        this.LoadMDBPlotCondTreeData_Start();

                    }
                    else if (this.rdoFilterByFile.Checked && m_intError == 0)
                    {
                        this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                            this.LoadMDBPlotCondTreeData_Start();

                    }
                }

			}
			else
			{
				MessageBox.Show("Select an FIADB population evaluation","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}
			//((frmDialog)this.ParentForm).MinimizeMainForm=false;
			//this.Enabled=true;
		}

		private void btnFIADBInvPrevious_Click(object sender, System.EventArgs e)
		{
			this.grpboxFIADBInv.Hide();
			this.grpboxFilter.Show();
		}

		private void btnFIADBInvCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void btnFIADBInvNext_Click(object sender, System.EventArgs e)
		{
			if (this.lstFIADBInv.SelectedItems.Count > 0)
			{
			    if (this.m_strLoadedPopEstUnitInputTable.Trim().ToUpper() !=
					this.cmbFiadbPopEstUnitTable.Text.Trim().ToUpper() ||
					this.m_strLoadedPopStratumInputTable.Trim().ToUpper() !=
					this.cmbFiadbPopStratumTable.Text.Trim().ToUpper() ||
					this.m_strLoadedPpsaInputTable.Trim().ToUpper() !=
					this.cmbFiadbPpsaTable.Text.Trim().ToUpper() ||
					this.m_strLoadedFIADBEvalId.Trim() != 
					this.m_strCurrFIADBEvalId.Trim() ||
					this.m_strLoadedFIADBRsCd.Trim() !=
					this.m_strCurrFIADBRsCd.Trim() ||
					this.m_strLoadedFiadbInputFile.Trim().ToUpper() != 
					this.txtFiadbInputFile.Text.Trim().ToUpper())
				{
					this.LoadMDBFiadbPopFiles();
				}
				

				if (m_intError==0)
				{
					this.mdbInputStateCounty();
					this.grpboxFilterByState.Visible=true;
					this.lstFilterByState.Refresh();
					this.grpboxFIADBInv.Visible=false;
				}
			}
			else
			{
				MessageBox.Show("Select an FIADB population evaluation","Add Plot Data",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
			}

		}

        private void CalculateAdjustments_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "mdbFIADBFileInput";
            this.StartTherm("2", "Calculate Adjustment Factors");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(CalculateAdjustments_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
            while (frmMain.g_oDelegate.m_oThread != null && frmMain.g_oDelegate.m_oThread.IsAlive)
            {
                frmMain.g_oDelegate.m_oThread.Join(1000);
                System.Windows.Forms.Application.DoEvents();

            }
        }

        private void LoadMDBPlotCondTreeData_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "mdbFIADBFileInput";
            this.StartTherm("2", "Add SQLite Plot,Cond,Site Tree, & Tree Table Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(LoadMDBPlotCondTreeData_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }
        private void LoadIDBPlotCondTreeData_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "mdbIDBFileInput";
            this.StartTherm("2", "Add MS Access Plot,Cond,Site Tree, & Tree Table Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(LoadIDBPlotCondTreeData_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void CalculateAdjustments_Finish()
        {
           

            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
           
            this.m_strCurrentProcess = "";
        }

        private void LoadMDBPlotCondTreeData_Finish()
        {
            this.m_strPlotIdList = "";
            
            
            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
            if (m_intError != 0)
            {
                this.m_strLoadedPopEstUnitInputTable = "";
                this.m_strLoadedPopStratumInputTable = "";
                this.m_strLoadedPpsaInputTable = "";
                this.m_strLoadedFiadbInputFile = "";
            }
            else
            {
                this.m_strLoadedPopEstUnitTxtInputFile = "";
                this.m_strLoadedPopEvalTxtInputFile = "";
                this.m_strLoadedPopStratumTxtInputFile = "";
                this.m_strLoadedPpsaTxtInputFile = "";
            }
            this.m_strCurrentProcess = "";
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            ((frmDialog)this.ParentForm).MinimizeMainForm = false;
            
        }
		private void lstFilterByState_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			this.m_bLoadStateCountyPlotList=true;
		}

		private void btnboxMDBFiadbInputFile_Click(object sender, System.EventArgs e)
		{
			
				OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
				OpenFileDialog1.Title = "SQLite Database File Containing FIADB Tables";
				OpenFileDialog1.Filter = "SQLite Database File (*.db;*.db3)|*.db;*.db3";

                DialogResult result =  OpenFileDialog1.ShowDialog();
				if (result == DialogResult.OK) 
				{
					if (OpenFileDialog1.FileName.Trim().Length > 0) 
					{
						string strFullPath = OpenFileDialog1.FileName.Trim();
						if (strFullPath.Length > 0) 
						{
							this.txtFiadbInputFile.Text = strFullPath;
                        DataMgr oDataMgr = new DataMgr();
                        string strConn = oDataMgr.GetConnectionString(strFullPath);
                        using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            con.Open();
                            string[] arrTables = oDataMgr.getTableNames(con);
                            if (arrTables.Length > 0)
                            {
                                this.cmbFiadbCondTable.Items.Clear();
                                this.cmbFiadbPlotTable.Items.Clear();
                                this.cmbFiadbPlotGeomTable.Items.Clear();
                                this.cmbFiadbPopEstUnitTable.Items.Clear();
                                this.cmbFiadbPopEvalTable.Items.Clear();
                                this.cmbFiadbPopStratumTable.Items.Clear();
                                this.cmbFiadbPpsaTable.Items.Clear();
                                this.cmbFiadbSeedlingTable.Items.Clear();
                                this.cmbFiadbTreeTable.Items.Clear();
                                this.cmbFiadbSiteTreeTable.Items.Clear();

                                cmbFiadbSeedlingTable.Items.Add("<Optional Table>");
                                cmbFiadbSeedlingTable.Text = "<Optional Table>";

                                for (int x = 0; x <= arrTables.Length-1; x++)
                                {
                                    this.cmbFiadbCondTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbPlotTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbPlotGeomTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbPopEstUnitTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbPopEvalTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbPopStratumTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbPpsaTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbSeedlingTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbTreeTable.Items.Add(arrTables[x]);
                                    this.cmbFiadbSiteTreeTable.Items.Add(arrTables[x]);
                                    switch (arrTables[x].Trim().ToUpper())
                                    {
                                        case "COND":
                                            this.cmbFiadbCondTable.Text = "COND";
                                            break;
                                        case "TREE":
                                            this.cmbFiadbTreeTable.Text = "TREE";
                                            break;
                                        case "PLOT":
                                            this.cmbFiadbPlotTable.Text = "PLOT";
                                            break;
                                        case "PLOTGEOM":
                                            this.cmbFiadbPlotGeomTable.Text = "PLOTGEOM";
                                            break;
                                        case "POP_EVAL":
                                            this.cmbFiadbPopEvalTable.Text = "POP_EVAL";
                                            break;
                                        case "POP_ESTN_UNIT":
                                            this.cmbFiadbPopEstUnitTable.Text = "POP_ESTN_UNIT";
                                            break;
                                        case "POP_PLOT_STRATUM_ASSGN":
                                            this.cmbFiadbPpsaTable.Text = "POP_PLOT_STRATUM_ASSGN";
                                            break;
                                        case "POP_STRATUM":
                                            this.cmbFiadbPopStratumTable.Text = "POP_STRATUM";
                                            break;
                                        case "SEEDLING":
                                            this.cmbFiadbSeedlingTable.Text = "SEEDLING";
                                            break;
                                        case "SITETREE":
                                            this.cmbFiadbSiteTreeTable.Text = "SITETREE";
                                            break;
                                    }
                                }

                            }
					    }
					}
                }
            }
			OpenFileDialog1 = null;			
		}

		private void btnMDBFiadbInputNext_Click(object sender, System.EventArgs e)
		{
            
            if (this.cmbFiadbPopEvalTable.Text.Trim().Length == 0 ||
				this.cmbFiadbCondTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPlotTable.Text.Trim().Length == 0 ||
                this.cmbFiadbPlotGeomTable.Text.Trim().Length == 0 ||
                this.cmbFiadbPopEstUnitTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPopEvalTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPopStratumTable.Text.Trim().Length == 0 ||
				this.cmbFiadbPpsaTable.Text.Trim().Length == 0 ||
				this.cmbFiadbSeedlingTable.Text.Trim().Length == 0 ||
				this.cmbFiadbTreeTable.Text.Trim().Length == 0 ||
				this.cmbFiadbSiteTreeTable.Text.Trim().Length==0)
			{
				MessageBox.Show("Enter a value for each table","FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Exclamation);
				return;
			}

			this.btnFilterNext.Enabled=true;
			this.btnFilterFinish.Enabled=false;
			
			
			this.rdoFilterByFile.Text = "Filter By File (Text File Containing Plot_CN numbers)";
		

			this.grpboxFilter.Visible=true;
			this.grpboxMDBFiadbInput.Visible=false;

		}
		private bool LoadMDBFiadbPopEvalTable()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.LoadMDBFiadPopEvalTable\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            string strCN="";
			string strCNDelimited="";
			string strEvalId="";
			string strEvalIdDelimited="";
			string strRsCd="";
			string strRsCdDelimited="";
			string strStateCd="";
			string strStateCdDelimited="";
			string strLocNm="";
			string strLocNmDelimited="";
			string strEvalDesc="";
			string strEvalDescDelimited="";
			string strRptYr="";
			string strRptYrDelimited="";
			string strNotes="";
			string strNotesDelimited="";
			int x=0;
			m_intError=0;
			bool bLoad=false;
			if (m_oDatasource==null) this.InitializeDatasource();

            try
            {
                //check if the eval list box has no values
                if (this.lstFIADBInv.Items.Count == 0)
                {
                    bLoad = true;
                }
                //see if the same values in the list as the table
                string strConnection = SQLite.GetConnectionString(this.txtFiadbInputFile.Text.Trim());
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
                {
                    con.Open();
                    string strSQL = "SELECT * FROM " + this.cmbFiadbPopEvalTable.Text.Trim();
                    SQLite.SqlQueryReader(con, strSQL);
                    if (SQLite.m_DataReader.HasRows)
                    {
                        while (SQLite.m_DataReader.Read())
                        {
                            //initialize eval values
                            strCN = "";
                            strEvalId = "";
                            strRsCd = "";
                            strStateCd = "";
                            strLocNm = "";
                            strEvalDesc = "";
                            strRptYr = "";
                            strNotes = "";
                            strCN = Convert.ToString(SQLite.m_DataReader["cn"]).Trim();
                            strEvalId = Convert.ToString(SQLite.m_DataReader["evalid"]).Trim();
                            strRsCd = Convert.ToString(SQLite.m_DataReader["RsCd"]).Trim();
                            strStateCd = Convert.ToString(SQLite.m_DataReader["statecd"]).Trim();

                            if (SQLite.m_DataReader["location_nm"] != DBNull.Value)
                                strLocNm = SQLite.m_DataReader["location_nm"].ToString();
                            if (SQLite.m_DataReader["eval_descr"] != DBNull.Value)
                                strEvalDesc = SQLite.m_DataReader["eval_descr"].ToString();
                            if (SQLite.m_DataReader["report_year_nm"] != DBNull.Value)
                                strRptYr = SQLite.m_DataReader["report_year_nm"].ToString();
                            if (SQLite.m_DataReader["notes"] != DBNull.Value)
                                strNotes = SQLite.m_DataReader["notes"].ToString();
                            //string all the eval records
                            strCNDelimited = strCNDelimited + strCN + " " + "#";
                            strEvalIdDelimited = strEvalIdDelimited + strEvalId + " " + "#";
                            strRsCdDelimited = strRsCdDelimited + strRsCd + " " + "#";
                            strStateCdDelimited = strStateCdDelimited + strStateCd + " " + "#";
                            strLocNmDelimited = strLocNmDelimited + strLocNm + " " + "#";
                            strEvalDescDelimited = strEvalDescDelimited + strEvalDesc + " " + "#";
                            strRptYrDelimited = strRptYrDelimited + strRptYr + " " + "#";
                            strNotesDelimited = strNotesDelimited + strNotes + " " + "#";
                            if (!bLoad)
                            {
                                //see if the tables eval row is found in the list box
                                for (x = 0; x <= this.lstFIADBInv.Items.Count - 1; x++)
                                {
                                    if (strEvalId.Trim() == this.lstFIADBInv.Items[x].SubItems[0].Text.Trim() &&
                                        strRsCd.Trim() == this.lstFIADBInv.Items[x].SubItems[1].Text.Trim() &&
                                        strStateCd.Trim() == this.lstFIADBInv.Items[x].SubItems[2].Text.Trim() &&
                                        strLocNm.Trim() == this.lstFIADBInv.Items[x].SubItems[3].Text.Trim() &&
                                        strEvalDesc.Trim() == this.lstFIADBInv.Items[x].SubItems[4].Text.Trim() &&
                                        strRptYr.Trim() == this.lstFIADBInv.Items[x].SubItems[5].Text.Trim() &&
                                        strNotes.Trim() == this.lstFIADBInv.Items[x].SubItems[6].Text.Trim())
                                        break;
                                }
                                if (x > this.lstFIADBInv.Items.Count - 1)
                                {
                                    //the eval table record is not found in the list box
                                    bLoad = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("There are no Population Evaluations in the " + this.cmbFiadbPopEvalTable.Text + " table ",
                            "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                        m_intError = -1;
                        bLoad = false;
                    }
                    SQLite.m_DataReader.Close();
                    SQLite.CloseConnection(SQLite.m_Connection);
                } // Closing connection

                if (bLoad)
                {
                    //remove the delimiter from the end of the string list
                    if (strCNDelimited.Trim().Length > 0) strCNDelimited = strCNDelimited.Substring(0, strCNDelimited.Length - 1);
                    if (strEvalIdDelimited.Trim().Length > 0) strEvalIdDelimited = strEvalIdDelimited.Substring(0, strEvalIdDelimited.Length - 1);
                    if (strRsCdDelimited.Trim().Length > 0) strRsCdDelimited = strRsCdDelimited.Substring(0, strRsCdDelimited.Length - 1);
                    if (strStateCdDelimited.Trim().Length > 0) strStateCdDelimited = strStateCdDelimited.Substring(0, strStateCdDelimited.Length - 1);
                    if (strLocNmDelimited.Trim().Length > 0) strLocNmDelimited = strLocNmDelimited.Substring(0, strLocNmDelimited.Length - 1);
                    if (strEvalDescDelimited.Trim().Length > 0) strEvalDescDelimited = strEvalDescDelimited.Substring(0, strEvalDescDelimited.Length - 1);
                    if (strRptYrDelimited.Trim().Length > 0) strRptYrDelimited = strRptYrDelimited.Substring(0, strRptYrDelimited.Length - 1);
                    if (strNotesDelimited.Trim().Length > 0) strNotesDelimited = strNotesDelimited.Substring(0, strNotesDelimited.Length - 1);

                    //create a temporary db file for the temporary pop tables
                    this.m_strTempDbFile = frmMain.g_oUtils.getRandomFile(m_oEnv.strTempDir, "db");
                    //this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();
                    //this.m_strPopMDBFile = this.m_strTempMDBFile;
                    //if (this.m_ado == null)
                    //{
                    //    this.m_ado = new ado_data_access();
                    //}

                    //create pop tables in temporary db file
                    CreatePopTables();

                    //get a connection string for the temp mdb file
                    string strConn = SQLite.GetConnectionString(m_strTempDbFile);
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                    {
                        conn.Open();
                        //covert the string lists to arrays
                        FIA_Biosum_Manager.utils oUtils = new FIA_Biosum_Manager.utils();
                        string[] strCNArray = oUtils.ConvertListToArray(strCNDelimited, "#");
                        string[] strEvalIdArray = oUtils.ConvertListToArray(strEvalIdDelimited, "#");
                        string[] strRsCdArray = oUtils.ConvertListToArray(strRsCdDelimited, "#");
                        string[] strStateCdArray = oUtils.ConvertListToArray(strStateCdDelimited, "#");
                        string[] strLocNmArray = oUtils.ConvertListToArray(strLocNmDelimited, "#");
                        string[] strEvalDescArray = oUtils.ConvertListToArray(strEvalDescDelimited, "#");
                        string[] strRptYrArray = oUtils.ConvertListToArray(strRptYrDelimited, "#");
                        string[] strNotesArray = oUtils.ConvertListToArray(strNotesDelimited, "#");
                        oUtils = null;
                        //insert the evaluation records into the biosum evaluation table
                        for (x = 0; x <= strEvalIdArray.Length - 1; x++)
                        {
                            SQLite.m_strSQL = "INSERT INTO " + this.m_strPopEvalTable + " " +
                                "(CN,RSCD,EVALID,EVAL_DESCR,STATECD," +
                                "LOCATION_NM,REPORT_YEAR_NM,NOTES," +
                                "START_INVYR,END_INVYR,BIOSUM_STATUS_CD) VALUES " +
                                "('" + strCNArray[x].Trim() + "'," +
                                strRsCdArray[x].Trim() + "," +
                                strEvalIdArray[x].Trim() + ",'" +
                                strEvalDescArray[x] + "'," +
                                strStateCdArray[x] + ",'" +
                                strLocNmArray[x] + "','" +
                                strRptYrArray[x] + "','" +
                                strNotesArray[x] + "',null,null,9)";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                                    SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            if (SQLite.m_intError != 0)
                            {
                                this.m_intError = SQLite.m_intError;
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                this.m_intError = -1;
                MessageBox.Show(e.Message,
                    "FIA Biosum",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                return false;
            }
			return bLoad;
		}

		private bool LoadMDBFiadbPopEvalTable2()
		{
			string strCN="";
			string strCNDelimited="";
			string strEvalId="";
			string strEvalIdDelimited="";
			string strRsCd="";
			string strRsCdDelimited="";
			string strStateCd="";
			string strStateCdDelimited="";
			string strLocNm="";
			string strLocNmDelimited="";
			string strEvalDesc="";
			string strEvalDescDelimited="";
			string strRptYr="";
			string strRptYrDelimited="";
			string strNotes="";
			string strNotesDelimited="";
			int x=0;
			m_intError=0;
			bool bLoad=false;
			if (this.m_ado==null)
				this.m_ado = new ado_data_access();
			if (m_oDatasource==null) this.InitializeDatasource();
				
			try
			{
				//check if the eval list box has no values
				if (this.lstFIADBInv.Items.Count==0)
				{
					bLoad=true;
				}
				//see if the same values in the list as the table
				m_ado.SqlQueryReader(m_ado.getMDBConnString(this.txtFiadbInputFile.Text.Trim(),"",""),"SELECT * FROM " + this.cmbFiadbPopEvalTable.Text.Trim());
				if (m_ado.m_intError==0)
				{
					if (m_ado.m_OleDbDataReader.HasRows)
					{
						while (m_ado.m_OleDbDataReader.Read())
						{
							//initialize eval values
							strCN="";
							strEvalId="";
							strRsCd="";
							strStateCd="";
							strLocNm="";
							strEvalDesc="";
							strRptYr="";
							strNotes="";
							strCN =m_ado.m_OleDbDataReader["cn"].ToString();
							strEvalId =m_ado.m_OleDbDataReader["evalid"].ToString();
							strRsCd= m_ado.m_OleDbDataReader["RsCd"].ToString();
							strStateCd = m_ado.m_OleDbDataReader["statecd"].ToString();
							if (m_ado.m_OleDbDataReader["location_nm"] != System.DBNull.Value)
								strLocNm = m_ado.m_OleDbDataReader["location_nm"].ToString();
							if (m_ado.m_OleDbDataReader["eval_descr"] != System.DBNull.Value)
								strEvalDesc = m_ado.m_OleDbDataReader["eval_descr"].ToString();
							if (m_ado.m_OleDbDataReader["report_year_nm"] != System.DBNull.Value)
								strRptYr = m_ado.m_OleDbDataReader["report_year_nm"].ToString();
							if (m_ado.m_OleDbDataReader["notes"] != System.DBNull.Value)
								strNotes = m_ado.m_OleDbDataReader["notes"].ToString();	
							//string all the eval records
							strCNDelimited=strCNDelimited + strCN + " " + "#";
							strEvalIdDelimited=strEvalIdDelimited + strEvalId + " " + "#";
							strRsCdDelimited = strRsCdDelimited + strRsCd + " " + "#";
							strStateCdDelimited = strStateCdDelimited + strStateCd + " " + "#";
							strLocNmDelimited = strLocNmDelimited + strLocNm + " " + "#";
							strEvalDescDelimited = strEvalDescDelimited + strEvalDesc + " " + "#";
							strRptYrDelimited = strRptYrDelimited + strRptYr + " " + "#";
							strNotesDelimited = strNotesDelimited + strNotes + " " + "#";
							if (!bLoad)
							{
								//see if the tables eval row is found in the list box
								for (x=0;x<=this.lstFIADBInv.Items.Count-1;x++)
								{
									if (strEvalId.Trim() == this.lstFIADBInv.Items[x].SubItems[0].Text.Trim() && 
										strRsCd.Trim() == this.lstFIADBInv.Items[x].SubItems[1].Text.Trim() && 
										strStateCd.Trim() == this.lstFIADBInv.Items[x].SubItems[2].Text.Trim() &&
										strLocNm.Trim() == this.lstFIADBInv.Items[x].SubItems[3].Text.Trim() &&
										strEvalDesc.Trim() == this.lstFIADBInv.Items[x].SubItems[4].Text.Trim() &&
										strRptYr.Trim() == this.lstFIADBInv.Items[x].SubItems[5].Text.Trim() && 
										strNotes.Trim() == this.lstFIADBInv.Items[x].SubItems[6].Text.Trim())
										break;
								}
								if (x > this.lstFIADBInv.Items.Count) 
								{
									//the eval table record is not found in the list box
									bLoad=true;
								}
							}
						}
						m_ado.m_OleDbDataReader.Close();
						while (m_ado.m_OleDbDataReader.IsClosed==false)
							System.Threading.Thread.Sleep(1000);
						
						if (bLoad)
						{
							//remove the delimiter from the end of the string list
							if (strCNDelimited.Trim().Length > 0) strCNDelimited = strCNDelimited.Substring(0,strCNDelimited.Length - 1);
							if (strEvalIdDelimited.Trim().Length > 0) strEvalIdDelimited = strEvalIdDelimited.Substring(0,strEvalIdDelimited.Length - 1);
							if (strRsCdDelimited.Trim().Length > 0) strRsCdDelimited = strRsCdDelimited.Substring(0,strRsCdDelimited.Length - 1);
							if (strStateCdDelimited.Trim().Length > 0) strStateCdDelimited=strStateCdDelimited.Substring(0,strStateCdDelimited.Length - 1);
							if (strLocNmDelimited.Trim().Length > 0) strLocNmDelimited=strLocNmDelimited.Substring(0,strLocNmDelimited.Length - 1);
							if (strEvalDescDelimited.Trim().Length > 0) strEvalDescDelimited=strEvalDescDelimited.Substring(0,strEvalDescDelimited.Length - 1);
							if (strRptYrDelimited.Trim().Length > 0) strRptYrDelimited=strRptYrDelimited.Substring(0,strRptYrDelimited.Length - 1);
							if (strNotesDelimited.Trim().Length > 0) strNotesDelimited=strNotesDelimited.Substring(0,strNotesDelimited.Length - 1);

							//create a temporary mdb file with links to all the project tables
							this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();

							//get a connection string for the temp mdb file
							this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile,"","");
							this.m_ado.OpenConnection(m_strTempMDBFileConn);
							if (m_ado.m_intError==0)
							{

								//delete the current eval records that have a value of 9
								m_ado.m_strSQL = "DELETE FROM " + this.m_strPopEvalTable + " WHERE biosum_status_cd=9";
								m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);
								if (m_ado.m_intError==0)
								{
									//covert the string lists to arrays
									FIA_Biosum_Manager.utils oUtils = new FIA_Biosum_Manager.utils();
									string[] strCNArray = oUtils.ConvertListToArray(strCNDelimited,"#");
									string[] strEvalIdArray = oUtils.ConvertListToArray(strEvalIdDelimited,"#");
									string[] strRsCdArray = oUtils.ConvertListToArray(strRsCdDelimited,"#");
									string[] strStateCdArray = oUtils.ConvertListToArray(strStateCdDelimited,"#");
									string[] strLocNmArray = oUtils.ConvertListToArray(strLocNmDelimited,"#");
									string[] strEvalDescArray = oUtils.ConvertListToArray(strEvalDescDelimited,"#");
									string[] strRptYrArray = oUtils.ConvertListToArray(strRptYrDelimited,"#");
									string[] strNotesArray = oUtils.ConvertListToArray(strNotesDelimited,"#");
									oUtils=null;
									//insert the evaluation records into the biosum evaluation table
									for (x=0;x<=strEvalIdArray.Length-1;x++)
									{
										m_ado.m_strSQL="INSERT INTO " + this.m_strPopEvalTable + " " + 
											"(CN,RSCD,EVALID,EVAL_DESCR,STATECD," + 
											"LOCATION_NM,REPORT_YEAR_NM,NOTES," + 
											"START_INVYR,END_INVYR,BIOSUM_STATUS_CD) VALUES " + 
											"('" + strCNArray[x].Trim() + "'," + 
											strRsCdArray[x].Trim() + "," + 
											strEvalIdArray[x].Trim() + ",'" + 
											strEvalDescArray[x] + "'," + 
											strStateCdArray[x] + ",'" + 
											strLocNmArray[x] + "','" + 
											strRptYrArray[x] + "','" + 
											strNotesArray[x] + "',null,null,9)";
										m_ado.SqlNonQuery(m_ado.m_OleDbConnection,m_ado.m_strSQL);	
										if (m_ado.m_intError!=0)
										{
											this.m_intError=m_ado.m_intError;
											break;
										}
									}
								}          							
								else m_intError=m_ado.m_intError;
								m_ado.m_OleDbConnection.Close();
								while (this.m_ado.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
									System.Threading.Thread.Sleep(1000);
							}
							else m_intError=m_ado.m_intError;
						}
					}
					else
					{
						MessageBox.Show("There are no Population Evaluations in the " + this.cmbFiadbPopEvalTable.Text + " table ",
							"FIA Biosum",
							System.Windows.Forms.MessageBoxButtons.OK,
							System.Windows.Forms.MessageBoxIcon.Exclamation);
						m_intError=-1;
						m_ado.m_OleDbDataReader.Close();
						while (m_ado.m_OleDbDataReader.IsClosed==false)
							System.Threading.Thread.Sleep(1000);
						bLoad=false;
						
					}
				}

			}
			catch (Exception e)
			{
				this.m_intError=-1;
				MessageBox.Show(e.Message,
					"FIA Biosum",
					System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				return false;
			}
			return bLoad;


		}
		private void LoadMDBFiadbPopFiles()
		{
		    this.m_strCurrentProcess="mdbFiadbInputPopTables";	

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.LoadMDBFiadbPopFiles\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

		    this.m_bLoadStateCountyList=true;
			this.m_bLoadStateCountyPlotList=true;

			this.StartTherm("2", "Add SQLite Pop Table Data");
			this.m_frmTherm.progressBar2.Maximum=3;
			this.m_frmTherm.progressBar2.Minimum=0;
			this.m_frmTherm.progressBar2.Value=0;
			this.m_frmTherm.lblMsg2.Text = "Overall Progress";
			this.m_strTableType="POPULATION STRATUM";
			    
			this.m_strCurrentFiadbTable = this.cmbFiadbPopStratumTable.Text;
			this.m_strCurrentFiadbInputFile = this.txtFiadbInputFile.Text;
			this.m_strCurrentBiosumTable=this.m_strPopStratumTable;			

				this.thdProcessRecords = new Thread(new ThreadStart(mdbFiadbInputPopTables));
				this.thdProcessRecords.IsBackground = true;
				this.thdProcessRecords.Start();
				while (thdProcessRecords.IsAlive)
				{
					thdProcessRecords.Join(1000);
					System.Windows.Forms.Application.DoEvents();

				}
				this.m_frmTherm.progressBar2.Value=1;
				thdProcessRecords=null;
				if (m_intError==0)
				{
					this.m_strLoadedPopStratumInputTable = this.m_strCurrentFiadbTable;
					this.m_strTableType="POPULATION ESTIMATION UNIT";
					this.m_frmTherm.lblMsg.Text = "pop estimation unit table";
					this.m_strCurrentFiadbTable = this.cmbFiadbPopEstUnitTable.Text;
					this.m_strCurrentFiadbInputFile = this.txtFiadbInputFile.Text;
					this.m_strCurrentBiosumTable = this.m_strPopEstUnitTable;
					this.thdProcessRecords = new Thread(new ThreadStart(mdbFiadbInputPopTables));
					this.thdProcessRecords.IsBackground = true;
					this.thdProcessRecords.Start();
					while (thdProcessRecords.IsAlive)
					{
						thdProcessRecords.Join(1000);
						System.Windows.Forms.Application.DoEvents();

					}
					thdProcessRecords=null;

				}
				this.m_frmTherm.progressBar2.Value=2;
				if (m_intError==0)
				{
					this.m_strLoadedPopEstUnitInputTable = this.m_strCurrentFiadbTable;
					this.m_strTableType="POPULATION PLOT STRATUM ASSIGNMENT";
					this.m_frmTherm.lblMsg.Text = "ppsa table";
					this.m_strCurrentFiadbTable = this.cmbFiadbPpsaTable.Text;
					this.m_strCurrentBiosumTable = this.m_strPpsaTable;
					this.thdProcessRecords = new Thread(new ThreadStart(mdbFiadbInputPopTables));
					this.thdProcessRecords.IsBackground = true;
					this.thdProcessRecords.Start();
					while (thdProcessRecords!=null && thdProcessRecords.IsAlive)
					{
						thdProcessRecords.Join(1000);
						System.Windows.Forms.Application.DoEvents();

					}
					thdProcessRecords=null;

				}
				if (this.m_intError==0)
				{
					this.m_strLoadedPpsaInputTable=this.m_strCurrentFiadbTable;
					this.m_strLoadedFiadbInputFile=this.m_strCurrentFiadbInputFile;
				}
				this.m_frmTherm.progressBar2.Value=this.m_frmTherm.progressBar2.Maximum;
				System.Threading.Thread.Sleep(2000);
				this.m_frmTherm.Close();
				this.m_frmTherm = null;

			    this.m_strCurrentProcess="";	
			
		}
		private void mdbFiadbInputPopTables()
		{
		   
			string strFields="";
			
			
			int x=0;
			int y=0;
			string strCol="";
			
			this.m_intError=0;		
			
			string strSourceFile=this.m_strCurrentFiadbInputFile;
			string strSourceTable=this.m_strCurrentFiadbTable;
			string strDestTable=this.m_strCurrentBiosumTable;
			string strSourceTableLink="fiadb_input_" + strSourceTable;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.mdbFiadbInputPopTables\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

            try
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Maximum", 4);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Minimum", 0);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 1);

                //open the connection to the temp db file
                string strConn = SQLite.GetConnectionString(m_strTempDbFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    SQLite.m_strSQL = "DELETE FROM " + strDestTable;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 2);
                    //get the fiabiosum table structures
                    DataTable dtDestSchema = dtDestSchema = SQLite.getTableSchema(conn, "select * from " + strDestTable);
                    //attach the source schema
                    SQLite.m_strSQL = $@"ATTACH DATABASE '{strSourceFile.Trim()}' AS FIADB";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                    /****************************************************************
                     **get the table structure that results from executing the sql
                     ****************************************************************/
                    //get the fiadb table structures
                    DataTable dtSourceSchema = SQLite.getTableSchema(conn, "select * from FIADB." + strSourceTable.Trim());

                    //build field list string to insert sql by matching 
                    //up the column names in the biosum plot table and the fiadb plot table
                    strFields = "";
                    for (x = 0; x <= dtDestSchema.Rows.Count - 1; x++)
                    {
                        strCol = dtDestSchema.Rows[x]["columnname"].ToString().Trim();
                        //see if there is an equivalent FIADB column
                        for (y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                        {
                            if (strCol.Trim().ToUpper() == dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper())
                            {
                                if (strFields.Trim().Length == 0)
                                {
                                    strFields = strCol;
                                }
                                else
                                {
                                    strFields += "," + strCol;
                                }
                                break;
                            }
                        }
                    }
                    SQLite.m_strSQL = "INSERT INTO " + strDestTable + " (" + strFields + ")" +
                        " SELECT " + strFields + " FROM FIADB." + strSourceTable + "   " +
                        "WHERE rscd = " + this.m_strCurrFIADBRsCd + " AND " +
                        "evalid = " + this.m_strCurrFIADBEvalId;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 3);
                    SQLite.m_strSQL = $@"UPDATE {strDestTable} SET (biosum_status_cd) = (select 9 from FIADB.{strSourceTable} where trim({strDestTable}.cn) = trim(FIADB.{strSourceTable}.cn))";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 4);
			}
			catch 
			{
				this.m_intError=-1;
				((frmDialog)this.ParentForm).Enabled=true;
			}
			finally
			{

				((frmDialog)this.ParentForm).Enabled=true;
				
				this.AddPlotRecordsFinished();
			}
			((frmDialog)this.ParentForm).Enabled=true;
			
			this.AddPlotRecordsFinished();

		}

		private void btnMDBFiadbInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void lstFIADBInv_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (lstFIADBInv.SelectedItems.Count==0) return;
			this.m_strCurrFIADBEvalId = this.lstFIADBInv.SelectedItems[0].Text.Trim();
			this.m_strCurrFIADBRsCd = this.lstFIADBInv.SelectedItems[0].SubItems[1].Text.Trim();
		
		}

		private void btnMDBSiteTreeBrowse_Click(object sender, System.EventArgs e)
		{
			this.GetMDBFileAndTable("SQLite Data File Containing Site Tree Table Data",
									"Select Site Tree Table",
									ref this.txtMDBSiteTree,
									ref this.txtMDBSiteTreeTable);
		}
		private void GetMDBFileAndTable(string p_strDialogTitleGetMDBFile,
			string p_strDialogTitleGetMDBTable,
			ref System.Windows.Forms.TextBox p_txtMDBFile,
			ref System.Windows.Forms.TextBox p_txtMDBTable)
		{
			OpenFileDialog OpenFileDialog1 = new OpenFileDialog();
			OpenFileDialog1.Title = p_strDialogTitleGetMDBFile;
			OpenFileDialog1.Filter = "Microsoft Access Database File (*.MDB,*.MDE,*.ACCDB) |*.mdb;*.mde;*.accdb";
			DialogResult result =  OpenFileDialog1.ShowDialog();
			if (result == DialogResult.OK) 
			{
				if (OpenFileDialog1.FileName.Trim().Length > 0) 
				{
					string strFullPath = OpenFileDialog1.FileName.Trim();
					if (strFullPath.Length > 0) 
					{
						utils p_utils = new FIA_Biosum_Manager.utils();
						string strFile = p_utils.getFileName(strFullPath);
						string strDir = p_utils.getDirectory(strFullPath);
						p_utils = null;
						dao_data_access tempDao = new dao_data_access();
						tempDao.OpenDb(strFullPath);
						if (tempDao.m_intError == 0) 
						{
							frmDialog frmTemp = new frmDialog();
							frmTemp.Text = p_strDialogTitleGetMDBTable;
							frmTemp.uc_select_list_item1.lblMsg.Text= "Table contents of " + strFullPath;
							frmTemp.uc_select_list_item1.lblMsg.Visible = true;
							string strLargestString = frmTemp.uc_select_list_item1.lblMsg.Text;
						
							frmTemp.uc_project1.Visible=false;
							frmTemp.uc_select_list_item1.listBox1.Items.Clear();
							for (int x=0; x <= tempDao.m_DaoDatabase.TableDefs.Count - 1; x++)
							{
								if (tempDao.m_DaoDatabase.TableDefs[x].Name.IndexOf("MSys",0) < 0) 		
								{
									frmTemp.uc_select_list_item1.listBox1.Items.Add(tempDao.m_DaoDatabase.TableDefs[x].Name);
									if (tempDao.m_DaoDatabase.TableDefs[x].Name.Trim().Length > 
										strLargestString.Trim().Length)
										strLargestString = tempDao.m_DaoDatabase.TableDefs[x].Name;
								}

							}
                        
						
							tempDao.m_DaoDatabase.Close();
							tempDao.m_DaoDatabase = null;
						
							frmTemp.uc_select_list_item1.Initialize_Width(strLargestString);
							frmTemp.uc_select_list_item1.Visible=true;
							result = frmTemp.ShowDialog(this);
                        
						
							if (result == DialogResult.OK) 
							{
							
								p_txtMDBFile.Text = strFullPath;
								p_txtMDBTable.Text = frmTemp.uc_select_list_item1.listBox1.Text;
							}
					
							frmTemp.Close();
							frmTemp = null;
						}
						tempDao = null;
					}
				}
			}
			else 
			{
			}
			OpenFileDialog1 = null;

		}

        private void SaveSqlitePopTables(string strEvalId, string strRscd)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.SaveSqlitePopTables\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }

            DataMgr oDataMgr = new DataMgr();
            string strConnection = oDataMgr.GetConnectionString(m_strMasterDbFile);
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                string strSQL = "SELECT COUNT(*) FROM " + this.m_strPopEvalTable + " WHERE evalid=" + strEvalId +
                    " AND rscd= " + strRscd;
                int intCount = (int) oDataMgr.getRecordCount(con, strSQL, this.m_strPopEvalTable);
                if (intCount > 0 && frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "Records exist for evalid " + strEvalId + ". POP tables will not be loaded.\r\n");
                    return;
                }
                else
                {
                    //attach FIADB database so we can copy records over
                    oDataMgr.m_strSQL = "ATTACH DATABASE '" + this.txtFiadbInputFile.Text.Trim() + "' AS FIADB";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                    oDataMgr.m_strSQL = "insert into " + this.m_strPopEvalTable + "(CN,RSCD,EVALID,EVAL_DESCR,STATECD," +
                        "LOCATION_NM,REPORT_YEAR_NM,NOTES," +
                        "START_INVYR,END_INVYR,GROWTH_ACCT,LAND_ONLY, MODIFIED_DATE, BIOSUM_STATUS_CD)" +
                        "select CN,RSCD,EVALID,EVAL_DESCR,STATECD," +
                        "LOCATION_NM,REPORT_YEAR_NM,NOTES," +
                        "START_INVYR,END_INVYR,GROWTH_ACCT,LAND_ONLY, MODIFIED_DATE,1 " +
                        "from FIADB." + this.m_strPopEvalTable +
                        " where evalid = " + strEvalId + " AND rscd= " + strRscd;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                    //get the fiabiosum table structures
                    string[] arrTables = new string[] { this.m_strPopEstUnitTable, this.m_strPopStratumTable, this.m_strPpsaTable };
                    string strPopEstnUnitTableSrc = frmMain.g_oDelegate.GetControlPropertyValue(cmbFiadbPopEstUnitTable, "Text", false).ToString().Trim();
                    string strPopStratumTableSrc = frmMain.g_oDelegate.GetControlPropertyValue(cmbFiadbPopStratumTable, "Text", false).ToString().Trim();
                    string strPpsaTableSrc = frmMain.g_oDelegate.GetControlPropertyValue(cmbFiadbPpsaTable, "Text", false).ToString().Trim();
                    string[] arrSourceTables = new string[] { strPopEstnUnitTableSrc, strPopStratumTableSrc, strPpsaTableSrc };
                    int i = 0;
                    foreach (var pTable in arrTables)
                    {
                        string strSqliteFields = "";
                        DataTable dtSqliteDestSchema = oDataMgr.getTableSchema(con, "select * from " + pTable);
                        DataTable dtSourceSchema = oDataMgr.getTableSchema(con, "select * from FIADB." + arrSourceTables[i].Trim());

                        //build field list string to insert sql by matching 
                        //up the column names in the biosum plot table and the fiadb plot table
                        strSqliteFields = "";
                        for (int x = 0; x <= dtSqliteDestSchema.Rows.Count - 1; x++)
                        {
                            string strCol = dtSqliteDestSchema.Rows[x]["columnname"].ToString().Trim();
                            //see if there is an equivalent FIADB column
                            for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                            {
                                if (strCol.Trim().ToUpper() == dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper())
                                {
                                    if (strSqliteFields.Trim().Length == 0)
                                    {
                                        strSqliteFields = strCol;
                                    }
                                    else
                                    {
                                        strSqliteFields += "," + strCol;
                                    }
                                    break;
                                }
                            }
                        }
                        oDataMgr.m_strSQL = "INSERT INTO " + pTable + " (" + strSqliteFields + ")" +
                            " SELECT " + strSqliteFields + " FROM FIADB." + arrSourceTables[i] + "   " +
                            "WHERE evalid = " + strEvalId + " AND rscd= " + strRscd;
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                                oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                        oDataMgr.m_strSQL = "UPDATE " + pTable +
                            " SET biosum_status_cd = 1" +
                            " WHERE evalid = " + strEvalId;
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                                oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);
                       i++;
                    }

                    //detach FIADB database
                    oDataMgr.m_strSQL = "DETACH DATABASE FIADB";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            oDataMgr.m_strSQL + "\r\n");
                    oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);
                }
            }
        }

        private IDictionary<string, List<string>> QueryStateEvalids()
        {
            m_strMasterDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                frmMain.g_oTables.m_oFIAPlot.DefaultPopTableDbFile;
            IDictionary<string, List<String>> dictStateEvalid = new Dictionary<string, List<String>>(); //Creates new dictionary
            string strConnection = SQLite.GetConnectionString(m_strMasterDbFile);
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                string strSQL = "select distinct statecd, evalid " +
                                "from pop_eval " +
                                "group by evalid, statecd";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                        strSQL + "\r\n");
                SQLite.SqlQueryReader(con, strSQL);
                if (SQLite.m_DataReader.HasRows)
                {
                    while (SQLite.m_DataReader.Read())
                    {
                        //make sure the row is not null values
                        if (SQLite.m_DataReader[0] != DBNull.Value &&
                            SQLite.m_DataReader[0].ToString().Trim().Length > 0)
                        {
                            string strState = SQLite.m_DataReader["statecd"].ToString().Trim();
                            string strEvalid = SQLite.m_DataReader["evalid"].ToString().Trim();
                            List<string> lstEval = null;
                            if (dictStateEvalid.ContainsKey(strState))
                            {
                                lstEval = dictStateEvalid[strState];
                                lstEval.Add(strEvalid);
                            }
                            else
                            {
                                lstEval = new List<string>();
                                lstEval.Add(strEvalid);
                                dictStateEvalid.Add(strState, lstEval);
                            }
                        }
                    }
                    SQLite.m_DataReader.Close();
                }
            }
            return dictStateEvalid;
        }

        private void CreatePopTables()
        {
            string strConn = SQLite.GetConnectionString(m_strTempDbFile);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                //pop estimation unit table
                if (SQLite.TableExist(conn, this.m_strPopEstUnitTable))
                {
                    SQLite.SqlNonQuery(conn, "DELETE * FROM " + this.m_strPopEstUnitTable);
                }
                else
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEstnUnitTable(SQLite, conn, this.m_strPopEstUnitTable);
                }
                //pop eval table
                if (SQLite.TableExist(conn, this.m_strPopEvalTable))
                {
                    SQLite.SqlNonQuery(conn, "DELETE * FROM " + this.m_strPopEvalTable);
                }
                else
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopEvalTable(SQLite, conn, this.m_strPopEvalTable);
                }
                //pop plot stratum assignment table
                if (SQLite.TableExist(conn, this.m_strPpsaTable))
                {
                    SQLite.SqlNonQuery(conn, "DELETE * FROM " + this.m_strPpsaTable);
                }
                else
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopPlotStratumAssgnTable(SQLite, conn, this.m_strPpsaTable);
                }
                //pop stratum table
                if (SQLite.TableExist(conn, this.m_strPopStratumTable))
                {
                    SQLite.SqlNonQuery(conn, "DELETE * FROM " + this.m_strPopStratumTable);
                }
                else
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePopStratumTable(SQLite, conn, this.m_strPopStratumTable);
                }
            }
        }

        private List<string> QueryFiadbStates()
        {
            List<string> lstStates = new List<string>();
            for (int x = 0; x <= this.lstFIADBInv.Items.Count - 1; x++)
            {
                string strStateCd = this.lstFIADBInv.Items[x].SubItems[2].Text.Trim();
                if (!lstStates.Contains(strStateCd))
                {
                    lstStates.Add(strStateCd);
                }
            }
            return lstStates;
        }

        private void cmbCondPropPercent_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void rdoFilterByFile_CheckedChanged(object sender, EventArgs e)
        {
            //Disable Forested/Non-Forested filters if filtering by file
            chkNonForested.Enabled = !rdoFilterByFile.Checked;
            chkForested.Enabled = !rdoFilterByFile.Checked;
        }

        private void btnMDBFiadbInputHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
                m_oHelp.ShowHelp(new string[] { "DATABASE", "PLOT_DATA_MENU" });
                //m_oHelp.ShowPdfHelp(m_pdfFile, "2");
            }
        }

        private void btnFilterHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "PLOTDATA2" });
        }

        private void btnFIADBInvHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "FIADB_INVENTORY_EVAL" });
        }

        private void btnFilterByStateHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "FILTER_STATE_COUNTY" });
        }

        private void btnFilterByPlotHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "FILTER_BY_PLOT" });
        }

        /*
        public class FIADB_Adjustments
        {
            public FIADB_Adjustments()
            {
                if (!System.IO.File.Exists(
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\DB\\BIOSUM_RECALC_FIADB_ADJUSTMENTS.ACCDB"))
                {
                    dao_data_access oDao = new dao_data_access();
                    oDao.CreateMDB(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\DB\\biosum_recalc_fiadb_adjustments.accdb");
                    oDao.m_DaoWorkspace.Close();
                    oDao.m_DaoWorkspace = null;
                    oDao = null;
                }
            }
            public enum ModeValues
            {
                ADD,
                DELETE,
            }
            private ModeValues _EditMode=ModeValues.ADD;
            public ModeValues EditMode
            {
                get { return _EditMode; }
                set { _EditMode = value; }
            }
            private string _strMSAccessDbFile = "";
            public string MSAccessDbFile
            {
                get { return _strMSAccessDbFile; }
                set { _strMSAccessDbFile = value; }
            }
           
            private void ImportCSVFiles(
                string p_strPlotFile,
                string p_strPopEstUnitFile,
                string p_strPopEvalFile,
                string p_strPopStratumFile,
                string p_strPPSAFile,
                string p_strCondFile)
            {

            }
            public void FIADB_Adjustments_Process()
            {

            }

         
        }
         */
	}
    
	
}

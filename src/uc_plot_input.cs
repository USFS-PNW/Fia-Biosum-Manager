using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Data.Common;
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


        private string m_strTempDbFile;
		private System.Data.DataTable m_dtStateCounty;
		private System.Data.DataTable m_dtPlot;
		private int m_intError;
        private string m_strError;
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
        private string m_strCondPropPercent = "";

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
		private System.Windows.Forms.TextBox txtDBPlotTable;
		private System.Windows.Forms.TextBox txtDBCondTable;
		private System.Windows.Forms.TextBox txtDBTreeTable;
		private System.Windows.Forms.TextBox txtDBPlot;
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
		private System.Windows.Forms.GroupBox grpboxDBFiadbInput;
		private System.Windows.Forms.GroupBox groupBox15;
		private System.Windows.Forms.GroupBox groupBox16;
		private System.Windows.Forms.GroupBox groupBox17;
		private System.Windows.Forms.GroupBox groupBox18;
		private System.Windows.Forms.GroupBox groupBox19;
		private System.Windows.Forms.GroupBox groupBox20;
		private System.Windows.Forms.GroupBox groupBox21;
		private System.Windows.Forms.GroupBox groupBox22;
		private System.Windows.Forms.Button btnboxDBFiadbInputFile;
		private System.Windows.Forms.ComboBox cmbFiadbPlotTable;
		private System.Windows.Forms.ComboBox cmbFiadbCondTable;
		private System.Windows.Forms.GroupBox groupBox23;
		private System.Windows.Forms.ComboBox cmbFiadbTreeTable;
		private System.Windows.Forms.ComboBox cmbFiadbSeedlingTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopEvalTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopEstUnitTable;
		private System.Windows.Forms.ComboBox cmbFiadbPopStratumTable;
		private System.Windows.Forms.ComboBox cmbFiadbPpsaTable;
		private System.Windows.Forms.Button btnDBFiadbInputFinish;
		private System.Windows.Forms.Button btnDBFiadbInputHelp;
		private System.Windows.Forms.Button btnDBFiadbInputPrev;
		private System.Windows.Forms.Button btnDBFiadbInputNext;
		private System.Windows.Forms.Button btnDBFiadbInputCancel;
		private System.Windows.Forms.TextBox txtFiadbInputFile;
		private System.Windows.Forms.GroupBox groupBox24;
		private System.Windows.Forms.ComboBox cmbFiadbSiteTreeTable;
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
            this.grpboxDBFiadbInput = new System.Windows.Forms.GroupBox();
            this.groupBox24 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbSiteTreeTable = new System.Windows.Forms.ComboBox();
            this.groupBox23 = new System.Windows.Forms.GroupBox();
            this.txtFiadbInputFile = new System.Windows.Forms.TextBox();
            this.btnboxDBFiadbInputFile = new System.Windows.Forms.Button();
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
            this.btnDBFiadbInputFinish = new System.Windows.Forms.Button();
            this.btnDBFiadbInputHelp = new System.Windows.Forms.Button();
            this.btnDBFiadbInputPrev = new System.Windows.Forms.Button();
            this.btnDBFiadbInputNext = new System.Windows.Forms.Button();
            this.btnDBFiadbInputCancel = new System.Windows.Forms.Button();
            this.grpboxFIADBInv = new System.Windows.Forms.GroupBox();
            this.btnFIADBInvAppend = new System.Windows.Forms.Button();
            this.lstFIADBInv = new System.Windows.Forms.ListView();
            this.btnFIADBInvHelp = new System.Windows.Forms.Button();
            this.btnFIADBInvPrevious = new System.Windows.Forms.Button();
            this.btnFIADBInvNext = new System.Windows.Forms.Button();
            this.btnFIADBInvCancel = new System.Windows.Forms.Button();
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
            this.txtDBTreeTable = new System.Windows.Forms.TextBox();
            this.txtDBCondTable = new System.Windows.Forms.TextBox();
            this.txtDBPlotTable = new System.Windows.Forms.TextBox();
            this.txtDBPlot = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cmbFiadbPlotGeomTable = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.grpboxDBFiadbInput.SuspendLayout();
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
            this.grpboxFilterByState.SuspendLayout();
            this.grpboxFilter.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.grpboxFilterByPlot.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.grpboxDBFiadbInput);
            this.groupBox1.Controls.Add(this.grpboxFIADBInv);
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
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox2);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox24);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox23);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox15);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox16);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox17);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox18);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox19);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox20);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox21);
            this.grpboxDBFiadbInput.Controls.Add(this.groupBox22);
            this.grpboxDBFiadbInput.Controls.Add(this.btnDBFiadbInputFinish);
            this.grpboxDBFiadbInput.Controls.Add(this.btnDBFiadbInputHelp);
            this.grpboxDBFiadbInput.Controls.Add(this.btnDBFiadbInputPrev);
            this.grpboxDBFiadbInput.Controls.Add(this.btnDBFiadbInputNext);
            this.grpboxDBFiadbInput.Controls.Add(this.btnDBFiadbInputCancel);
            this.grpboxDBFiadbInput.Location = new System.Drawing.Point(19, 70);
            this.grpboxDBFiadbInput.Margin = new System.Windows.Forms.Padding(4);
            this.grpboxDBFiadbInput.Name = "grpboxMDBFiadbInput";
            this.grpboxDBFiadbInput.Padding = new System.Windows.Forms.Padding(4);
            this.grpboxDBFiadbInput.Size = new System.Drawing.Size(1344, 450);
            this.grpboxDBFiadbInput.TabIndex = 35;
            this.grpboxDBFiadbInput.TabStop = false;
            this.grpboxDBFiadbInput.Text = "FIADB SQLite Database File Input";
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
            this.groupBox23.Controls.Add(this.btnboxDBFiadbInputFile);
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
            this.btnboxDBFiadbInputFile.Image = ((System.Drawing.Image)(resources.GetObject("btnboxMDBFiadbInputFile.Image")));
            this.btnboxDBFiadbInputFile.Location = new System.Drawing.Point(740, 12);
            this.btnboxDBFiadbInputFile.Margin = new System.Windows.Forms.Padding(4);
            this.btnboxDBFiadbInputFile.Name = "btnboxMDBFiadbInputFile";
            this.btnboxDBFiadbInputFile.Size = new System.Drawing.Size(40, 40);
            this.btnboxDBFiadbInputFile.TabIndex = 1;
            this.btnboxDBFiadbInputFile.Click += new System.EventHandler(this.btnboxDBFiadbInputFile_Click);
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
            this.btnDBFiadbInputFinish.Enabled = false;
            this.btnDBFiadbInputFinish.Location = new System.Drawing.Point(730, 408);
            this.btnDBFiadbInputFinish.Margin = new System.Windows.Forms.Padding(4);
            this.btnDBFiadbInputFinish.Name = "btnMDBFiadbInputFinish";
            this.btnDBFiadbInputFinish.Size = new System.Drawing.Size(90, 30);
            this.btnDBFiadbInputFinish.TabIndex = 7;
            this.btnDBFiadbInputFinish.Text = "Append";
            // 
            // btnMDBFiadbInputHelp
            // 
            this.btnDBFiadbInputHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnDBFiadbInputHelp.Location = new System.Drawing.Point(30, 408);
            this.btnDBFiadbInputHelp.Margin = new System.Windows.Forms.Padding(4);
            this.btnDBFiadbInputHelp.Name = "btnMDBFiadbInputHelp";
            this.btnDBFiadbInputHelp.Size = new System.Drawing.Size(80, 30);
            this.btnDBFiadbInputHelp.TabIndex = 3;
            this.btnDBFiadbInputHelp.Text = "Help";
            this.btnDBFiadbInputHelp.Click += new System.EventHandler(this.btnDBFiadbInputHelp_Click);
            // 
            // btnMDBFiadbInputPrev
            // 
            this.btnDBFiadbInputPrev.Enabled = false;
            this.btnDBFiadbInputPrev.Location = new System.Drawing.Point(530, 408);
            this.btnDBFiadbInputPrev.Margin = new System.Windows.Forms.Padding(4);
            this.btnDBFiadbInputPrev.Name = "btnMDBFiadbInputPrev";
            this.btnDBFiadbInputPrev.Size = new System.Drawing.Size(90, 30);
            this.btnDBFiadbInputPrev.TabIndex = 5;
            this.btnDBFiadbInputPrev.TabStop = false;
            this.btnDBFiadbInputPrev.Text = "< Previous";
            // 
            // btnMDBFiadbInputNext
            // 
            this.btnDBFiadbInputNext.Location = new System.Drawing.Point(620, 408);
            this.btnDBFiadbInputNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnDBFiadbInputNext.Name = "btnMDBFiadbInputNext";
            this.btnDBFiadbInputNext.Size = new System.Drawing.Size(90, 30);
            this.btnDBFiadbInputNext.TabIndex = 6;
            this.btnDBFiadbInputNext.Text = "Next >";
            this.btnDBFiadbInputNext.Click += new System.EventHandler(this.btnDBFiadbInputNext_Click);
            // 
            // btnMDBFiadbInputCancel
            // 
            this.btnDBFiadbInputCancel.Location = new System.Drawing.Point(420, 408);
            this.btnDBFiadbInputCancel.Margin = new System.Windows.Forms.Padding(4);
            this.btnDBFiadbInputCancel.Name = "btnMDBFiadbInputCancel";
            this.btnDBFiadbInputCancel.Size = new System.Drawing.Size(80, 30);
            this.btnDBFiadbInputCancel.TabIndex = 4;
            this.btnDBFiadbInputCancel.Text = "Cancel";
            this.btnDBFiadbInputCancel.Click += new System.EventHandler(this.btnDBFiadbInputCancel_Click);
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
            this.grpboxDBFiadbInput.ResumeLayout(false);
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
            this.grpboxFilterByState.ResumeLayout(false);
            this.grpboxFilter.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.grpboxFilterByPlot.ResumeLayout(false);
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
			this.m_DialogHt = this.groupBox1.Top + this.grpboxDBFiadbInput.Top + this.grpboxDBFiadbInput.Height + 100 ;

		
					
			this.grpboxFilterByState.Left = this.grpboxDBFiadbInput.Left;
			this.grpboxFilterByState.Width = this.grpboxDBFiadbInput.Width;
			this.grpboxFilterByState.Height = this.grpboxDBFiadbInput.Height;
			this.grpboxFilterByState.Top = this.grpboxDBFiadbInput.Top;
            this.btnFilterByStateHelp.Location = this.btnDBFiadbInputHelp.Location;
            this.btnFilterByStateCancel.Location = this.btnDBFiadbInputCancel.Location;
            this.btnFilterByStatePrevious.Location = this.btnDBFiadbInputPrev.Location;
			this.btnFilterByStateNext.Location = this.btnDBFiadbInputNext.Location;
			this.btnFilterByStateFinish.Location = this.btnDBFiadbInputFinish.Location;
			this.grpboxFilterByState.Visible=false;	

			this.grpboxFilter.Left = this.grpboxDBFiadbInput.Left;
			this.grpboxFilter.Width = this.grpboxDBFiadbInput.Width;
			this.grpboxFilter.Height = this.grpboxDBFiadbInput.Height;
			this.grpboxFilter.Top = this.grpboxDBFiadbInput.Top;
			this.btnFilterHelp.Location = this.btnDBFiadbInputHelp.Location;
			this.btnFilterCancel.Location = this.btnDBFiadbInputCancel.Location;
			this.btnFilterPrevious.Location = this.btnDBFiadbInputPrev.Location;
			this.btnFilterNext.Location = this.btnDBFiadbInputNext.Location;
			this.btnFilterFinish.Location = this.btnDBFiadbInputFinish.Location;
			this.grpboxFilter.Visible=false;	

			this.grpboxFilterByPlot.Left = this.grpboxDBFiadbInput.Left;
			this.grpboxFilterByPlot.Width = this.grpboxDBFiadbInput.Width;
			this.grpboxFilterByPlot.Height = this.grpboxDBFiadbInput.Height;
			this.grpboxFilterByPlot.Top = this.grpboxDBFiadbInput.Top;
			this.btnFilterByPlotHelp.Location = this.btnDBFiadbInputHelp.Location;
			this.btnFilterByPlotCancel.Location = this.btnDBFiadbInputCancel.Location;
			this.btnFilterByPlotPrevious.Location = this.btnDBFiadbInputPrev.Location;
			this.btnFilterByPlotNext.Location = this.btnDBFiadbInputNext.Location;
			this.btnFilterByPlotFinish.Location = this.btnDBFiadbInputFinish.Location;
			this.grpboxFilterByPlot.Visible=false;

			this.grpboxFIADBInv.Left = this.grpboxDBFiadbInput.Left;
			this.grpboxFIADBInv.Width = this.grpboxDBFiadbInput.Width;
			this.grpboxFIADBInv.Height = this.grpboxDBFiadbInput.Height;
			this.grpboxFIADBInv.Top = this.grpboxDBFiadbInput.Top;
			this.btnFIADBInvHelp.Location = this.btnDBFiadbInputHelp.Location;
            this.btnFIADBInvHelp.Click += new System.EventHandler(this.btnFIADBInvHelp_Click);
			this.btnFIADBInvCancel.Location = this.btnDBFiadbInputCancel.Location;
			this.btnFIADBInvPrevious.Location = this.btnDBFiadbInputPrev.Location;
			this.btnFIADBInvNext.Location = this.btnDBFiadbInputNext.Location;
			this.btnFIADBInvAppend.Location = this.btnDBFiadbInputFinish.Location;
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
            this.grpboxDBFiadbInput.Visible = true;
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
                    LoadDBPlotCondTreeData_Start();
                }
                else if (this.rdoFilterByFile.Checked == true)
                {
                    if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
                    {
                        this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                        {
                            this.LoadDBPlotCondTreeData_Start();
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

            //-----------PREPARATION FOR CALCULATING ADJUSTMENTS---------//

            try
            {
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
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConnection))
                {
                    conn.Open();
                    if (!GetBooleanValue((Control)m_frmTherm, "AbortProcess"))
                    {
                        this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Drop Work Tables");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 10);
                        if (SQLite.TableExist(conn, "BIOSUM_PLOT"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_PLOT");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 20);
                        if (SQLite.TableExist(conn, "BIOSUM_COND"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_COND");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 30);
                        if (SQLite.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName))
                            SQLite.SqlNonQuery(conn, "DROP TABLE " + frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);
                        SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                        if (SQLite.TableExist(conn, "BIOSUM_PPSA"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_PPSA");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                        if (SQLite.TableExist(conn, "BIOSUM_EUS_TEMP"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_EUS_TEMP");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 60);
                        if (SQLite.TableExist(conn, "BIOSUM_PPSA_DENIED_ACCESS"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_PPSA_DENIED_ACCESS");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 70);
                        if (SQLite.TableExist(conn, "BIOSUM_PPSA_TEMP"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_PPSA_TEMP");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                        if (SQLite.TableExist(conn, "BIOSUM_EUS_TEMP"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_EUS_TEMP");
                        SetThermValue(m_frmTherm.progressBar1, "Value", 90);
                        if (SQLite.TableExist(conn, "BIOSUM_EUS_ACCESS"))
                            SQLite.SqlNonQuery(conn, "DROP TABLE BIOSUM_EUS_ACCESS");

                        SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                        System.Threading.Thread.Sleep(2000);
                        SetThermValue(m_frmTherm.progressBar2, "Value", 30);
                        strFIADBDbFile = (string)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.TextBox)txtFiadbInputFile, "Text", false);
                        strFIADBDbFile = strFIADBDbFile.Trim();
                        m_strCondPropPercent = frmMain.g_oDelegate.GetControlPropertyValue(cmbCondPropPercent, "Text", false).ToString().Trim();
                        string[] strSql = Queries.FIAPlot.FIADBPlotInput_CalculateAdjustmentFactorsSQL(
                            "POP_PLOT_STRATUM_ASSGN",
                            "POP_ESTN_UNIT",
                            "POP_STRATUM",
                            "POP_EVAL",
                            "PLOT",
                            "COND",
                             m_strCurrFIADBRsCd,
                             m_strCurrFIADBEvalId,
                             m_strCondPropPercent,
                             strFIADBDbFile);
                        SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                        this.SetLabelValue(m_frmTherm.lblMsg, "Text", "Calculate Adjustment Factors For RsCd=" + m_strCurrFIADBRsCd + " and EvalId=" + m_strCurrFIADBEvalId);

                        for (x = 0; x <= strSql.Length - 1; x++)
                        {
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, strSql[x] + "\r\n");
                            SQLite.SqlNonQuery(conn, strSql[x]);
                            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar1, "Value", 20 + x + 5);
                            if (SQLite.m_intError != 0) break;
                        }

                        // Add indexes to BIOSUM_PLOT and BIOSUM_COND tables
                        SQLite.AddIndex(conn, "BIOSUM_PLOT", "BIOSUM_PLOT_idx1", "CN");
                        SQLite.AddIndex(conn, "BIOSUM_COND", "BIOSUM_COND_idx1", "CN");
                    }
                }

                m_intError = SQLite.m_intError;
                SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                System.Threading.Thread.Sleep(2000);
                SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                
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
                if (SQLite != null)
                {
                    if (SQLite.m_DataSet != null)
                    {
                        SQLite.m_DataSet.Clear();
                        SQLite.m_DataSet.Dispose();
                    }
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
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
            if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);




            CalculateAdjustments_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }
		private void LoadDBPlotCondTreeData_Process()
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
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.LoadDBPlotCondTreeData_Process\r\n");
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
                    DeleteFromTablesWhereFilter(conn, new string[] { this.m_strPlotTable },
                        " WHERE biosum_status_cd=9 OR LENGTH(biosum_plot_id)=0;");
                    DeleteFromTablesWhereFilter(conn, new string[]
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

                        string strBiosumRefDb = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.Reference.DefaultBiosumReferenceSqliteFile;
                        SQLite.m_strSQL = "ATTACH DATABASE '" + strBiosumRefDb + "' AS REF";
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

                        if (SQLite.TableExist(conn, "tempplot") || SQLite.AttachedTableExist(conn, "tempplot"))
                        {
                            SQLite.m_strSQL = "DROP TABLE tempplot";
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }

                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Plot Table: Insert New Plot Records");
                        if (Checked(rdoFilterByFile) == true && m_strPlotIdList.Trim().Length > 0 &&
                            !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                        {
                            string strDelimiter = ",";
                            string[] strPlotIdArray = m_strPlotIdList.Split(strDelimiter.ToCharArray());
                            if (SQLite.TableExist(conn, "input_cn") || SQLite.AttachedTableExist(conn, "input_cn"))
                            {
                                SQLite.m_strSQL = "DROP TABLE input_cn";
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            }
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

                    // abort if no plots are loaded into tempplot
                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SQLite.m_strSQL = "SELECT * FROM tempplot";
                        SQLite.SqlQueryReader(conn, SQLite.m_strSQL);
                        if (!SQLite.m_DataReader.HasRows)
                        {
                            SQLite.m_strSQL = "DROP TABLE tempplot";
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            SQLite.m_strSQL = "DROP TABLE input_cn";
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            MessageBox.Show("!!No selected plots exist in selected EvalId!!", "FIA Biosum",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                            m_intError = -1;
                        }
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

                    if (m_intError == 0 & !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                        // insert the new plot records into the plot table
                        SQLite.m_strSQL = "INSERT INTO " + this.m_strPlotTable + " (biosum_plot_id, biosum_status_cd, " + strFields + ") " +
                            "SELECT TRIM(biosum_plot_id), biosum_status_cd, " + strFields + " FROM tempplot";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        // create plot column update work table
                        if (!SQLite.AttachedTableExist(conn, "plot_column_updates_work_table"))
                        {
                            SQLite.m_strSQL = "CREATE TABLE TEMPDB.plot_column_updates_work_table AS " +
                            "SELECT biosum_plot_id, statecd AS cond_ttl FROM " + this.m_strPlotTable + " WHERE 1=2";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 40);
                        SetThermValue(m_frmTherm.progressBar2, "Value", 20);
                        //-------------CONDITION TABLE----------------//
                        //build field list string to insert sql by matching FIADB and BioSum Cond columns
                        strFields = CreateStrFieldsFromDataTables(dtFIADBCondSchema, dtCondSchema);
                        /********************************************************
                         **create condition input insert command
                         ********************************************************/
                        //check the user defined filters
                        if (SQLite.TableExist(conn, "tempcond") || SQLite.AttachedTableExist(conn, "tempcond"))
                        {
                            SQLite.m_strSQL = "DELETE FROM TEMPDB.tempcond";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                            SQLite.m_strSQL = "INSERT INTO TEMPDB.tempcond " +
                                "SELECT p.biosum_plot_id, TRIM(p.biosum_plot_id) || c.condid AS biosum_cond_id, 9 AS biosum_status_cd, c.* " +
                                "FROM BIOSUM_COND AS c, " + this.m_strPlotTable + " AS p " +
                                "WHERE c.plt_cn = p.cn AND p.biosum_status_cd = 9";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        else
                        {
                            SetLabelValue(m_frmTherm.lblMsg, "Text", "Condition Table: Insert New  Records");
                            SQLite.m_strSQL = "CREATE TABLE TEMPDB.tempcond AS " +
                                "SELECT p.biosum_plot_id, TRIM(p.biosum_plot_id) || c.condid AS biosum_cond_id, 9 AS biosum_status_cd, c.* " +
                                "FROM BIOSUM_COND AS c, " + this.m_strPlotTable + " AS p " +
                                "WHERE c.plt_cn = p.cn AND p.biosum_status_cd = 9";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 50);
                        //insert the new condition records into the condition table
                        SQLite.m_strSQL = "INSERT INTO " + this.m_strCondTable + " (biosum_plot_id, biosum_cond_id, biosum_status_cd, " + strFields + ") " +
                            "SELECT TRIM(biosum_plot_id), TRIM(biosum_cond_id), biosum_status_cd, " + strFields + " FROM TEMPDB.tempcond";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS d " +
                            "SET cond_status_cd = s.cond_status_cd " +
                            "FROM TEMPDB.tempcond AS s WHERE TRIM(d.cn) = TRIM(s.cn)";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        //create cond column work table
                        if (!SQLite.AttachedTableExist(conn, "cond_column_updates_work_table"))
                        {
                            SQLite.m_strSQL = "CREATE TABLE TEMPDB.cond_column_updates_work_table AS " +
                            "SELECT biosum_cond_id, qmd_all_inch, qmd_hwd_inch, qmd_swd_inch, tpacurr, " +
                            "hwd_tpacurr, swd_tpacurr, ba_ft2_ac, hwd_ba_ft2_ac, swd_ba_ft2_ac, " +
                            "vol_ac_grs_stem_ttl_ft3, hwd_vol_ac_grs_stem_ttl_ft3, swd_vol_ac_grs_stem_ttl_ft3, " +
                            "vol_ac_grs_ft3, hwd_vol_ac_grs_ft3, swd_vol_ac_grs_ft3, volcsgrs, hwd_volcsgrs, swd_volcsgrs " +
                            "FROM " + this.m_strCondTable + " WHERE 1=2";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 55);
                        SetThermValue(m_frmTherm.progressBar2, "Value", 40);
                        //-------------TREE TABLE----------------//
                        //build field list string to insert sql by matching FIADB and BioSum Tree columns
                        strFields = CreateStrFieldsFromDataTables(dtFIADBTreeSchema, dtTreeSchema);
                        /********************************************************
                         **create tree input insert command
                         ********************************************************/
                        if (SQLite.TableExist(conn, "temptree") || SQLite.AttachedTableExist(conn, "temptree"))
                        {
                            SQLite.m_strSQL = "DROP TABLE temptree";
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        //check the user defined filters
                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Tree Table: Insert New  Records");
                        SQLite.m_strSQL = "CREATE TABLE TEMPDB.temptree AS " +
                            "SELECT TRIM(p.biosum_plot_id) || t.condid AS biosum_cond_id, 9 AS biosum_status_cd, t.* " +
                            "FROM FIADB." + strTreeSource + " AS t, " + this.m_strPlotTable + " AS p " +
                            "WHERE t.plt_cn = TRIM(p.cn) AND p.biosum_status_cd = 9 AND t.statuscd <> 0";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 60);
                        //insert the new tree records into the tree table; Note that temptree is used later for GRM processing
                        SQLite.m_strSQL = "INSERT INTO " + this.m_strTreeTable + " (biosum_cond_id, biosum_status_cd, " + strFields + ") " +
                            "SELECT TRIM(biosum_cond_id), biosum_status_cd, " + strFields + " FROM TEMPDB.temptree";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    // SEEDLINGS
                    if (m_intError == 0 && m_bLoadSeedlings == true && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 65);
                        //-------------SEEDLING TABLE----------------//
                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Seedling Table: Insert New  Records");
                        if (SQLite.TableExist(conn, "tempseedling") || SQLite.AttachedTableExist(conn, "tempseedling"))
                        {
                            SQLite.m_strSQL = "DROP TABLE tempseedling";
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        SQLite.m_strSQL = "CREATE TABLE TEMPDB.tempseedling AS " +
                            "SELECT TRIM(p.biosum_plot_id) || CAST(s.condid AS TEXT) AS biosum_cond_id, 9 AS biosum_status_cd, " +
                            "0.1 AS dia, 1 AS diahtcd, '1' || printf('%03d', SPCD) || '00' || SUBP AS fvs_tree_id, 1 AS statuscd, s.* " +
                            "FROM FIADB." + strSeedlingSource + " AS s, " + this.m_strPlotTable + " AS p " +
                            "WHERE s.plt_cn = TRIM(p.cn) AND p.biosum_status_cd = 9";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        //Set DIAHTCD for Seedlings using FIA_TREE_SPECIES_REF.WOODLAND_YN
                        SQLite.m_strSQL = "UPDATE TEMPDB.tempseedling AS t " +
                            "SET diahtcd = CASE WHEN ref.woodland_yn = 'N' THEN 1 ELSE 2 END " +
                            "FROM REF.FIA_TREE_SPECIES_REF AS ref WHERE t.spcd = ref.spcd";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        //Prepend CN with "S" to indicate seedlings
                        SQLite.m_strSQL = "UPDATE TEMPDB.tempseedling SET CN = 'S' || CN";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);

                        //build field list string to insert sql by matching FIADB and BioSum Tree columns
                        //build field list string to insert sql by matching FIADB and BioSum Tree columns
                        strFields = CreateStrFieldsFromDataTables(dtFIADBSeedlingSchema, dtTreeSchema);
                        SetThermValue(m_frmTherm.progressBar1, "Value", 70);

                        //insert the new seedling records into the tree table
                        SQLite.m_strSQL = "INSERT INTO " + this.m_strTreeTable + " (biosum_cond_id, biosum_status_cd, dia, " +
                            "diahtcd, fvs_tree_id, statuscd, " + strFields + ") " +
                            "SELECT TRIM(biosum_cond_id), biosum_status_cd, dia, diahtcd, fvs_tree_id, statuscd, " + strFields +
                            " FROM TEMPDB.tempseedling";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                        //update the cullbf column
                        SQLite.m_strSQL = "UPDATE " + this.m_strTreeTable +
                            " SET cullbf = CASE WHEN cullbf IS NULL " +
                            "THEN CASE WHEN cull IS NOT NULL AND roughcull IS NOT NULL " +
                            "THEN cull + roughcull " +
                            "ELSE CASE WHEN cull IS NOT NULL THEN cull " +
                            "ELSE CASE WHEN roughcull IS NOT NULL THEN roughcull ELSE 0 " +
                            "END END END ELSE cullbf END";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        //-------------SITE TREE TABLE----------------//
                        //build field list string to insert sql by matching biosum fiadb SiteTree table
                        strFields = CreateStrFieldsFromDataTables(dtFIADBSiteTreeSchema, dtSiteTreeSchema);
                        /********************************************************
                         **create site tree input insert command
                         ********************************************************/
                        if (SQLite.TableExist(conn, "tempsitetree") || SQLite.AttachedTableExist(conn, "tempsitetree"))
                        {
                            SQLite.m_strSQL = "DROP TABLE tempsitetree";
                            SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        }
                        //check the user defined filters
                        SetLabelValue(m_frmTherm.lblMsg, "Text", "Site Tree Table: Insert New  Records");
                        SQLite.m_strSQL = "CREATE TABLE TEMPDB.tempsitetree AS " +
                            "SELECT TRIM(p.biosum_plot_id) AS biosum_plot_id, 9 AS biosum_status_cd, t.* " +
                            "FROM FIADB." + strSiteTreeSource + " AS t, " + this.m_strPlotTable + " AS p " +
                            "WHERE t.plt_cn = TRIM(p.cn) AND p.biosum_status_cd = 9";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 80);
                        //insert the new condition records into the condition table
                        SQLite.m_strSQL = "INSERT INTO " + this.m_strSiteTreeTable + " (biosum_plot_id, biosum_status_cd, " + strFields + ") " +
                            "SELECT TRIM(biosum_plot_id), biosum_status_cd, " + strFields + " FROM TEMPDB.tempsitetree";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                        SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                        m_intError = SQLite.m_intError;
                    }

                    if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value", 100);
                        SetThermValue(m_frmTherm.progressBar2, "Value", 60);
                        m_intError = UpdateColumns(conn);
                    }

                    //Down Woody Materials Section
                    if (m_intError == 0 && Checked(chkDwmImport) &&
                        !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        SetThermValue(m_frmTherm.progressBar1, "Value",
                        GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                        SetThermValue(m_frmTherm.progressBar2, "Value", 80);

                        m_intError = ImportDownWoodyMaterials(conn);
                    }

                    //count the number of records added to each table and
                    //set their biosum_status_cd values to 1
                    if (this.m_intError == 0 &&
                        !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
                    {
                        m_intAddedPlotRows = GetNewRecordCount(conn, m_strPlotTable);
                        m_intAddedCondRows = GetNewRecordCount(conn, m_strCondTable);
                        m_intAddedTreeRows = GetNewRecordCount(conn, m_strTreeTable);
                        m_intAddedSeedlingRows = GetNewRecordCount(conn, m_strSeedlingTable);
                        m_intAddedSiteTreeRows = GetNewRecordCount(conn, m_strSiteTreeTable);

                        if (Checked(chkDwmImport))
                        {
                            m_intAddedDwmCwdRows = GetNewRecordCount(conn, m_strDwmCwdTable, "AUX.");
                            m_intAddedDwmFwdRows = GetNewRecordCount(conn, m_strDwmFwdTable, "AUX.");
                            m_intAddedDwmDuffLitterRows = GetNewRecordCount(conn, m_strDwmDuffLitterTable, "AUX.");
                            m_intAddedDwmTransectSegmentRows = GetNewRecordCount(conn, m_strDwmTransectSegmentTable, "AUX.");
                        }

                        //Successfully imported and updated plot data. Set biosum_status_cd to 1
                        string[] arrTables = new string[] {m_strPlotTable, m_strCondTable, m_strTreeTable, m_strSiteTreeTable,
                                m_strPopEvalTable, m_strPopStratumTable, m_strPpsaTable, m_strPopEstUnitTable, m_strBiosumPopStratumAdjustmentFactorsTable};
                        if (Checked(chkDwmImport))
                        {
                            arrTables = new string[] {m_strPlotTable, m_strCondTable, m_strTreeTable, m_strSiteTreeTable,
                                m_strPopEvalTable, m_strPopStratumTable, m_strPpsaTable, m_strPopEstUnitTable, m_strBiosumPopStratumAdjustmentFactorsTable,
                                m_strDwmCwdTable, m_strDwmFwdTable, m_strDwmDuffLitterTable, m_strDwmTransectSegmentTable};
                        }
                        
                        foreach (string table in arrTables)
                        {
                            if (SQLite.TableExist(conn, table) || SQLite.AttachedTableExist(conn, table))
                            {
                                if (table == m_strDwmCwdTable || table == m_strDwmFwdTable || table == m_strDwmDuffLitterTable || table == m_strDwmTransectSegmentTable)
                                {
                                    SQLite.m_strSQL = "UPDATE AUX." + table +
                                    " SET biosum_status_cd = 1 WHERE biosum_status_cd = 9";
                                }
                                else
                                {
                                    SQLite.m_strSQL = "UPDATE " + table +
                                    " SET biosum_status_cd = 1 WHERE biosum_status_cd = 9";
                                }
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                                SQLite.SqlNonQuery(conn, SQLite.m_strSQL);
                            }
                        }

                        // Save POP tables to SQLite
                        SaveSqlitePopTables(conn, this.m_strCurrFIADBEvalId, this.m_strCurrFIADBRsCd);
                        SetThermValue(m_frmTherm.progressBar1, "Value",
                        GetThermValue(m_frmTherm.progressBar1, "Maximum"));
                        SetThermValue(m_frmTherm.progressBar2, "Value",
                        GetThermValue(m_frmTherm.progressBar2, "Maximum"));
                        frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Button)m_frmTherm.btnCancel,
                            "Visible", false);

                        // Save configuration
                        SaveLoadConfigurationTxt(conn);
                        SaveLoadConfigurationTable(conn);

                        MessageBox.Show(CreateRecordsAddedMessage(), "Add Plot Data");

                        this.m_strLoadedPopEstUnitInputTable =
                            (string)frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.ComboBox)cmbFiadbPopEstUnitTable, "Text", false);
                        this.m_strLoadedPopStratumInputTable =
                            (string)frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.ComboBox)cmbFiadbPopStratumTable, "Text", false);
                        this.m_strLoadedPpsaInputTable =
                            (string)frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.ComboBox)cmbFiadbPpsaTable, "Text", false);
                        this.m_strLoadedFIADBEvalId = this.m_strCurrFIADBEvalId;
                        this.m_strLoadedFIADBRsCd = this.m_strCurrFIADBRsCd;
                        this.m_strLoadedFiadbInputFile =
                            (string)frmMain.g_oDelegate.GetControlPropertyValue(
                                (System.Windows.Forms.TextBox)txtFiadbInputFile, "Text", false);
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                    {
                        // A DataMgr error occurred so delete all the incompletely processed new records
                        DeleteFromTablesWhereFilter(conn, new string[] { m_strPlotTable, m_strCondTable, m_strTreeTable, m_strPopEvalTable, m_strPopStratumTable,
                                m_strPpsaTable, m_strPopEstUnitTable, m_strSiteTreeTable, m_strBiosumPopStratumAdjustmentFactorsTable,
                                "AUX." + m_strDwmCwdTable, "AUX." + m_strDwmFwdTable, "AUX." + m_strDwmDuffLitterTable, "AUX." + m_strDwmTransectSegmentTable },
                                " WHERE biosum_status_cd = 9");
                        MessageBox.Show("!!Error Occured Adding Plot Records: 0 Records Added!!", "FIA Biosum",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                    
                    if (SQLite.m_DataSet != null)
                    {
                        SQLite.m_DataSet.Clear();
                        SQLite.m_DataSet.Dispose();
                    }

                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible", true);
                    frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);

                    LoadDBPlotCondTreeData_Finish();
                }
            }
            catch (System.Threading.ThreadInterruptedException err)
			{
				MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
			}
			catch  (System.Threading.ThreadAbortException err)
			{
                if (SQLite != null)
                {
                    if (SQLite.m_DataSet != null)
                    {
                        SQLite.m_DataSet.Clear();
                        SQLite.m_DataSet.Dispose();
                    }
                }
                this.CancelThreadCleanup();
			    this.ThreadCleanUp();
				this.CleanupThread();
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_plot_input:dbFIADBFileInput  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
				this.m_intError=-1;
			}
			finally
			{
               
			}
            if (SQLite != null)
            {
                if (SQLite.m_DataSet != null)
                {
                    SQLite.m_DataSet.Clear();
                    SQLite.m_DataSet.Dispose();
                }
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled", true);
			if (this.m_frmTherm != null) frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm,"Visible",false);
            LoadDBPlotCondTreeData_Finish();
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
	    
        private int GetNewRecordCount(System.Data.SQLite.SQLiteConnection p_conn, string strTable, string strTablePrefix = null)
        {
            if (SQLite.TableExist(p_conn, strTable) || SQLite.AttachedTableExist(p_conn, strTable) || strTable.Equals(m_strSeedlingTable))
            {
                int intReturn = -1;
                switch (strTable)
                {
                    case "fiadb_seedling_input":
                        intReturn = (int)SQLite.getRecordCount(p_conn,
                            String.Concat("SELECT COUNT(*) FROM ", m_strTreeTable, " WHERE biosum_status_cd = 9 AND dia = 0.1"), m_strTreeTable);
                        break;
                    case "tree":
                        intReturn = (int)SQLite.getRecordCount(p_conn,
                            String.Concat("SELECT COUNT(*) FROM ", strTable, " WHERE biosum_status_cd = 9 AND dia <> 0.1"), strTable);
                        break;
                    default:
                        intReturn = (int)SQLite.getRecordCount(p_conn,
                    String.Concat("SELECT COUNT(*) FROM ", strTablePrefix, strTable, " WHERE biosum_status_cd = 9"), strTablePrefix + strTable);
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
	
        private void DeleteFromTablesWhereFilter(System.Data.SQLite.SQLiteConnection p_conn, string[] strTableNames, string strDeleteFilter)
        {
            foreach (string table in strTableNames)
            {
                if (SQLite.TableExist(p_conn, table))
                {
                    SQLite.m_strSQL = "DELETE FROM " + table + strDeleteFilter;
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            SQLite.m_strSQL + "\r\n");
                }
            }
        }

	    /// <summary>
	    /// After importing DWM table data from FIADB to Master.mdb, join the DWM table with the plot/cond tables to get the BSCID
	    /// </summary>
	    /// <param name="SQLite">Reference to the temporary database used to link FIADB and Master during plot input phase</param>
	    /// <param name="strDestTable">The DWM table to update</param>

        private void UpdateDwmBiosumCondIds(System.Data.SQLite.SQLiteConnection p_conn, string strDestTable = null)
        {
            SQLite.m_strSQL = "UPDATE " + strDestTable + " AS t " +
                "SET biosum_cond_id = c.biosum_cond_id " +
                "FROM " + m_strPlotTable + " AS p, " + m_strCondTable + " AS c " +
                "WHERE c.biosum_plot_id = p.biosum_plot_id AND TRIM(p.cn) = TRIM(t.plt_cn) AND TRIM(t.condid) = TRIM(c.condid) " +
                "AND t.biosum_cond_id = '9999999999999999999999999' AND p.biosum_status_cd=9";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            m_intError = SQLite.m_intError;
        }
        /// <summary>
        /// Generates and executes an SQL "INSERT INTO dest_tablename (columns) SELECT columns FROM source_tablename" statement
        /// strInsertFields contains a subset of the source table columns that are common with the dest table.
        /// </summary>
        /// <param name="strSourceTable"></param>
        /// <param name="strDestTable"></param>
        /// <param name="strInsertFields"></param>
        
        private void InsertIntoDestTableFromSourceTable(System.Data.SQLite.SQLiteConnection p_conn,
            string strSourceTable = null, string strDestTable = null,
            string strSourceFields = null, string strDestFields = null, bool InsertBiosumCondIdAndStatusCode = false)
        {
            string strInsertIntoValues = "INSERT INTO {0} ({1}) ";
            string strSelectColumns = "SELECT {2} FROM {3} AS f, {4} AS p ON f.plt_cn=p.cn;";
            if (InsertBiosumCondIdAndStatusCode)
            {
                strInsertIntoValues = "INSERT INTO {0} (biosum_cond_id, biosum_status_cd, {1}) ";
                strSelectColumns =
                    "SELECT '9999999999999999999999999' AS biosum_cond_id, 9 AS biosum_status_cd, {2} FROM {3} AS f, {4} AS p WHERE f.plt_cn=trim(p.cn);";
            }
            SQLite.m_strSQL = String.Format(strInsertIntoValues + strSelectColumns, strDestTable, strDestFields, strSourceFields, strSourceTable, "tempplot");
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            m_intError = SQLite.m_intError;
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

   
	    private int UpdateColumns(System.Data.SQLite.SQLiteConnection p_conn)
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
            SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                "SET condprop = CASE WHEN ps.pmh_macr IS NOT NULL AND ps.pmh_macr > 0 THEN c.condprop_unadj / ps.pmh_macr " +
                "ELSE CASE WHEN ps.pmh_sub IS NOT NULL AND ps.pmh_sub > 0 THEN c.condprop_unadj / ps.pmh_sub " +
                "ELSE CASE WHEN ps.pmh_micr IS NOT NULL AND ps.pmh_micr > 0 THEN c.condprop_unadj / ps.pmh_micr " +
                "ELSE 0 END END END " +
                "FROM TEMPDB." + this.m_strPpsaTable + " AS ppsa, " + this.m_strPlotTable + " AS p, TEMPDB." + this.m_strBiosumPopStratumAdjustmentFactorsTable + " AS ps " +
                "WHERE TRIM(ppsa.plt_cn) = p.cn AND ppsa.stratum_cn = ps.stratum_cn AND c.biosum_plot_id = p.biosum_plot_id";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
			SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Condition Acres Column...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

			//update acres column
            SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                "SET acres = CASE WHEN c.condprop IS NOT NULL AND ps.expns IS NOT NULL " +
                "THEN c.condprop * ps.expns ELSE 0 END " +
                "FROM TEMPDB." + this.m_strPpsaTable + " AS ppsa, " + this.m_strPlotTable + " AS p, TEMPDB." + this.m_strBiosumPopStratumAdjustmentFactorsTable + " AS ps " +
                "WHERE ppsa.plt_cn = p.cn AND ppsa.stratum_cn = ps.stratum_cn AND c.biosum_plot_id = p.biosum_plot_id";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
			SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            //update condprop_specific column for when plot.macro_breakpoint_dia has a value
            SetLabelValue(m_frmTherm.lblMsg, "Text", "Updating Tree Condprop Specific Column...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
            SQLite.m_strSQL = "UPDATE " + this.m_strTreeTable + " AS t " +
                "SET condprop_specific = CASE WHEN c.micrprop_unadj IS NOT NULL AND t.dia < 5 THEN c.micrprop_unadj " +
                "ELSE CASE WHEN c.subpprop_unadj IS NOT NULL AND p.MACRO_BREAKPOINT_DIA IS NOT NULL " +
                "AND t.dia >=5 AND t.dia < p.MACRO_BREAKPOINT_DIA THEN c.subpprop_unadj " +
                "ELSE CASE WHEN c.macrprop_unadj IS NOT NULL AND p.MACRO_BREAKPOINT_DIA IS NOT NULL " +
                "AND t.dia >= p.MACRO_BREAKPOINT_DIA THEN c.macrprop_unadj END END END " +
                "FROM " + this.m_strCondTable + " AS c, " + this.m_strPlotTable + " AS p " +
                "WHERE c.biosum_plot_id = p.biosum_plot_id AND t.biosum_status_cd = 9";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            //check if null values 
            SQLite.m_strSQL = "SELECT COUNT(*) AS ROWCOUNT " +
                "FROM " + this.m_strTreeTable + " AS t, " + this.m_strTreeMacroPlotBreakPointDiaTable + " AS bp " +
                "WHERE t.biosum_status_cd = 9 AND t.condprop_specific IS NULL " +
                "AND t.statecd = bp.statecd AND t.unitcd = bp.unitcd";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
            //check if condprop_specific null and exists in the tree macro plot breakpoint diameter table
            if ((double)SQLite.getSingleDoubleValueFromSQLQuery(p_conn, SQLite.m_strSQL, "temp") > 0)
            {
                //got some nulls
                SQLite.m_strSQL = "UPDATE " + m_strTreeTable + " AS t " +
                    "SET condprop_specific = CASE WHEN c.micrprop_unadj IS NOT NULL AND t.dia < 5 THEN c.micrprop_unadj " +
                    "ELSE CASE WHEN c.subpprop_unadj IS NOT NULL AND bp.MACRO_BREAKPOINT_DIA IS NOT NULL " +
                    "AND t.dia >= 5 AND t.dia < bp.MACRO_BREAKPOINT_DIA THEN c.subpprop_unadj " +
                    "ELSE CASE WHEN macrprop_unadj IS NOT NULL AND bp.MACRO_BREAKPOINT_DIA IS NOT NULL " +
                    "AND t.dia >= bp.MACRO_BREAKPOINT_DIA THEN c.macrprop_unadj END END END " +
                    "FROM " + this.m_strCondTable + " AS c, " + this.m_strPlotTable + " AS p, " + this.m_strTreeMacroPlotBreakPointDiaTable + " AS bp " +
                    "WHERE c.biosum_plot_id = p.biosum_plot_id AND p.statecd = bp.statecd AND p.unitcd = bp.unitcd " +
                    "AND t.biosum_cond_id = c.biosum_cond_id AND t.biosum_status_cd = 9";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            }
            else
            {
                //see if we have nulls and no MACRO PLOT for the unit code
                SQLite.m_strSQL = "SELECT COUNT (*) AS ROWCOUNT " +
                    "FROM " + this.m_strTreeTable + " AS a, " +
                    "(SELECT t.* FROM " + this.m_strTreeTable + " AS t " +
                    "WHERE NOT EXISTS (SELECT * FROM " + this.m_strTreeMacroPlotBreakPointDiaTable + " AS bp " +
                    "WHERE t.statecd = bp.statecd AND t.unitcd = bp.unitcd)) AS b " +
                    "WHERE a.CN = b.CN AND a.condprop_specific IS NULL AND a.biosum_status_cd = 9";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                //handle for those states and units that do not have macro plot
                if ((double)SQLite.getSingleDoubleValueFromSQLQuery(p_conn, SQLite.m_strSQL, "temp") > 0)
                {
                    SQLite.m_strSQL = "UPDATE " + this.m_strTreeTable + " AS t " +
                        "SET condprop_specific = CASE WHEN c.micrprop_unadj IS NOT NULL AND t.dia < 5 " +
                        "THEN c.micrprop_unadj ELSE CASE WHEN c.subpprop_unadj IS NOT NULL AND t.dia >= 5 " +
                        "THEN c.subpprop_unadj END END END " +
                        "FROM " + this.m_strCondTable + " AS c, " + this.m_strPlotTable + " AS p " +
                        "WHERE c.biosum_plot_id = p.biosum_plot_id AND t.biosum_cond_id = c.biosum_cond_id " +
                        "AND t.condid = c.condid AND t.biosum_status_cd = 9";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                }
            }

            //Update fvs_tree_id column for tracking a tree between BioSum and FVS for lifetime of project
            SetLabelValue(m_frmTherm.lblMsg, "Text", "Updating Tree fvs_tree_id Column...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control) this.m_frmTherm, "Refresh");
            SQLite.m_strSQL = "UPDATE " + this.m_strTreeTable + " SET fvs_tree_id = CAST((subp * 1000 + tree) AS TEXT) WHERE dia > 0.1";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Tree tpacurr Column...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
			//update tree tpacurr column
            SQLite.m_strSQL = "UPDATE " + this.m_strTreeTable +
                " SET tpacurr = CASE WHEN tpa_unadj IS NOT NULL AND condprop_specific IS NOT NULL " +
                "THEN tpa_unadj / condprop_specific ELSE 0 END";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
			SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            //Set spcd to 998 is it is 4 digits long
            SetLabelValue(m_frmTherm.lblMsg, "Text", "Updating Tree spcd Column...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
            SQLite.m_strSQL = "UPDATE " + this.m_strTreeTable + " SET spcd = 998 WHERE spcd > 999";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            //
            //drybiot,drybiom,voltsgrs processing
            //                   
            SetLabelValue(m_frmTherm.lblMsg, "Text", "Start Volume and Biomass Calculations...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

            //step 5 - delete and create work table
            string strFcsBiosumVolumesInputTable = "TEMPDB." + Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable;
            if (SQLite.TableExist(p_conn, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable) || SQLite.AttachedTableExist(p_conn, Tables.VolumeAndBiomass.FcsBiosumVolumesInputTable))
            {
                SQLite.SqlNonQuery(p_conn, "DROP TABLE " + strFcsBiosumVolumesInputTable);
            }
            frmMain.g_oTables.m_oFvs.CreateSQLiteInputFCSBiosumVolumesTable(SQLite, p_conn, strFcsBiosumVolumesInputTable);

            var treeToFcsBiosumVolumesInputTable = new List<Tuple<string, string>>
            {
                Tuple.Create("ACTUALHT", "ACTUALHT"),
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
                Tuple.Create("DIA", "CASE WHEN DIA IS NOT NULL THEN ROUND(DIA, 2) ELSE DIA END"),
                Tuple.Create("DIAHTCD", "DIAHTCD"),
                Tuple.Create("FORMCL", "FORMCL"),
                Tuple.Create("HT", "HT"),
                Tuple.Create("HTDMP", "HTDMP"),
                Tuple.Create("INVYR", "INVYR"),
                Tuple.Create("PLOT", "CAST(SUBSTR(BIOSUM_COND_ID, 16, 6) AS INTEGER) AS PLOT"),
                Tuple.Create("PLT_CN", "SUBSTR(BIOSUM_COND_ID, 1, LENGTH(BIOSUM_COND_ID) - 1) AS PLT_CN"),
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
            SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step1(strFcsBiosumVolumesInputTable, m_strTreeTable, strColumns, strValues);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step2(strFcsBiosumVolumesInputTable, m_strTreeTable,m_strPlotTable,m_strCondTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);


            SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step3(strFcsBiosumVolumesInputTable, m_strCondTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            //populate treeclcd column
            string strCullTotalWorkTable = "TEMPDB.cull_total_work_table";
            if (SQLite.TableExist(p_conn, "cull_total_work_table") || SQLite.AttachedTableExist(p_conn, "cull_total_work_table"))
            {
                SQLite.SqlNonQuery(p_conn, "DROP TABLE " + strCullTotalWorkTable);
            }
            SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.BuildInputTableForVolumeCalculation_Step4(
                                strCullTotalWorkTable, strFcsBiosumVolumesInputTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.PNWRS.BuildInputTableForVolumeCalculation_Step5(
                strCullTotalWorkTable, strFcsBiosumVolumesInputTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.PNWRS.BuildInputTableForVolumeCalculation_Step6(
                            strCullTotalWorkTable, strFcsBiosumVolumesInputTable);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            string strWorkTable = "TEMPDB." + Tables.VolumeAndBiomass.SqliteWorkTable;
            if (SQLite.TableExist(p_conn, Tables.VolumeAndBiomass.SqliteWorkTable) || SQLite.AttachedTableExist(p_conn, Tables.VolumeAndBiomass.SqliteWorkTable))
            {
                SQLite.SqlNonQuery(p_conn, "DROP TABLE " + Tables.VolumeAndBiomass.SqliteWorkTable);
            }
            frmMain.g_oTables.m_oFvs.CreateSqliteInputFCSBiosumVolumesWorkTable(SQLite, p_conn, strWorkTable);
            string strInputFields = SQLite.getFieldNames(p_conn, "SELECT * FROM " + strFcsBiosumVolumesInputTable);
            SQLite.m_strSQL = "INSERT INTO " + strWorkTable + " (" + strInputFields + ") " +
                "SELECT * FROM " + strFcsBiosumVolumesInputTable;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

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
                //Remove data from fcs_tree.db
                //
                string strFcsTreeDb = frmMain.g_oEnv.strApplicationDataDirectory + "\\FIABiosum\\" + Tables.VolumeAndBiomass.DefaultSqliteWorkDatabase;
                SQLite.m_strSQL = "ATTACH DATABASE '" + strFcsTreeDb + "' AS FCS";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "DELETE FROM FCS." + Tables.VolumeAndBiomass.BiosumVolumeCalcTable;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);

                //
                //Insert into fcs_tree.biosum_calc
                //
                SQLite.m_strSQL = "INSERT INTO FCS." + Tables.VolumeAndBiomass.BiosumVolumeCalcTable + "( " + strInputFields + ") " +
                    " SELECT " + strInputFields + " FROM " + strWorkTable;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                SetThermValue(m_frmTherm.progressBar1, "Value", 1);

                //
                //Run Java app to calculate volume/biomass
                //
                if (m_intError == 0)
                {
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
                //Update with calculated values
                //
                if (m_intError == 0)
                {
                    if (SQLite.TableExist(p_conn, Tables.VolumeAndBiomass.BiosumCalcOutputTable) || SQLite.AttachedTableExist(p_conn, Tables.VolumeAndBiomass.BiosumCalcOutputTable))
                    {
                        SQLite.m_strSQL = "DROP TABLE " + Tables.VolumeAndBiomass.BiosumCalcOutputTable;
                        SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                    }

                    SQLite.m_strSQL = "CREATE TABLE TEMPDB." + Tables.VolumeAndBiomass.BiosumCalcOutputTable +
                        " AS SELECT * FROM " + strFcsBiosumVolumesInputTable + " WHERE 1 = 2";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                    string strOutputFields = SQLite.getFieldNames(p_conn, "SELECT * FROM TEMPDB." + Tables.VolumeAndBiomass.BiosumCalcOutputTable);
                    SQLite.m_strSQL = "INSERT INTO TEMPDB." + Tables.VolumeAndBiomass.BiosumCalcOutputTable + " (" + strOutputFields + ") " +
                        "SELECT " + strOutputFields + " FROM FCS." + Tables.VolumeAndBiomass.BiosumVolumeCalcTable +
                        " WHERE VOLTSGRS_CALC IS NOT NULL AND TRE_CN IS NOT NULL";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                    //update VOLTSGRS
                    SQLite.m_strSQL = Queries.VolumeAndBiomass.FIAPlotInput.WriteCalculatedVolumeAndBiomassColumnsToTreeTable("TEMPDB." + Tables.VolumeAndBiomass.BiosumCalcOutputTable);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n\r\n");
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                }

                SetThermValue(m_frmTherm.progressBar1, "Value", m_frmTherm.progressBar1.Maximum);
            }

            SetLabelValue(m_frmTherm.lblMsg,"Text", "Updating Condition Table Columns...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");
            SetThermValue(m_frmTherm.progressBar1, "Value", 6);

            //tpa column
            //sum trees per acre on a condition 
            //for live trees >= 5 inches in diameter 
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, tpacurr) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.tottpa AS tpacurr " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(tpacurr) AS tottpa " +
                    "FROM " + this.m_strTreeTable + " WHERE statuscd = 1 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 7);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET tpacurr = u.tpacurr " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 8);

            //swd_tpa
            //sum trees per acre on a condition 
            //for softwood live trees >= 5 inches in diameter 
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, swd_tpacurr) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.totswdtpa AS swd_tpacurr " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(tpacurr) AS totswdtpa " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd < 300 AND statuscd = 1 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 9);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET swd_tpacurr = u.swd_tpacurr " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 10);

            //hwd tpa
            //sum trees per acre on a condition 
            //for hardwood live trees >= 5 inches in diameter
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, hwd_tpacurr) " +
                    "SELECT DISTINCT (a.biosum_cond_id), a.tothwdtpa AS hwd_tpacurr " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(tpacurr) AS tothwdtpa " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd > 299 AND statuscd = 1 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 11);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET hwd_tpacurr = u.hwd_tpacurr " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 12);

            //vol_ac_grs_ft3

            //total
            //for all live trees >= 5 inches in diameter 
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, vol_ac_grs_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.tot_volgrsft3 AS vol_ac_grs_ft3 " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) AS tot_volgrsft3 " +
                    "FROM " + this.m_strTreeTable + " WHERE volcfgrs IS NOT NULL AND tpacurr IS NOT NULL AND statuscd = 1 AND dia >= 5 " +
                    "GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 13);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET vol_ac_grs_ft3 = u.vol_ac_grs_ft3 " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 14);

            //hwd
            //for all live hardwood trees >= 5 inches in diameter 
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, hwd_vol_ac_grs_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.tot_volgrsft3 AS hwd_vol_ac_grs_ft3 " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) AS tot_volgrsft3 " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd > 299 AND volcfgrs IS NOT NULL AND tpacurr IS NOT NULL " +
                    "AND statuscd = 1 AND dia >= 5 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 15);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET hwd_vol_ac_grs_ft3 = u.hwd_vol_ac_grs_ft3 " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 16);

            //SWD
            //for all live softwood trees >= 5 inches in diameter 
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, swd_vol_ac_grs_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.tot_volgrsft3 AS swd_vol_ac_grs_ft3 " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(volcfgrs * tpacurr) AS tot_volgrsft3 " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd < 300 AND volcfgrs IS NOT NULL AND tpacurr IS NOT NULL " +
                    "AND statuscd = 1 AND dia >= 5 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 17);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET swd_vol_ac_grs_ft3 = u.swd_vol_ac_grs_ft3 " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value",18);

            //ba_ft2_ac basal area column
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, ba_ft2_ac) " +
                    "SELECT a.biosum_cond_id, b.tottemp AS ba_ft2_ac " +
                    "FROM " + this.m_strCondTable + " AS a, " +
                    "(SELECT biosum_cond_id, SUM((.005454154 * POW(dia, 2)) * tpacurr) AS tottemp " +
                    "FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9 " +
                    "AND statuscd = 1 GROUP BY biosum_cond_id) AS b " +
                    "WHERE a.biosum_cond_id = b.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 19);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET ba_ft2_ac = u.ba_ft2_ac " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 20);

            //swd_ba_ft2_ac softwood basal area 
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, swd_ba_ft2_ac) " +
                    "SELECT a.biosum_cond_id, b.tottemp AS swd_ba_ft2_ac " +
                    "FROM " + this.m_strCondTable + " AS a, " +
                    "(SELECT biosum_cond_id, SUM((.005454154 * POW(dia, 2)) * tpacurr) AS tottemp " +
                    "FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9 " +
                    "AND spcd < 300 AND statuscd = 1 GROUP BY biosum_cond_id) AS b " +
                    "WHERE a.biosum_cond_id = b.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 21);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET swd_ba_ft2_ac = u.swd_ba_ft2_ac " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 22);

            //hardwood ba_ft2_ac
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, hwd_ba_ft2_ac) " +
                    "SELECT a.biosum_cond_id, b.tottemp AS hwd_ba_ft2_ac " +
                    "FROM " + this.m_strCondTable + " AS a, " +
                    "(SELECT biosum_cond_id, SUM((.005454154 * POW(dia, 2)) * tpacurr) AS tottemp " +
                    "FROM " + this.m_strTreeTable + " WHERE biosum_status_cd = 9 " +
                    "AND spcd > 299 AND statuscd = 1 GROUP BY biosum_cond_id) AS b " +
                    "WHERE a.biosum_cond_id = b.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 23);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET hwd_ba_ft2_ac = u.hwd_ba_ft2_ac " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 24);

			//volcsgrs  
			//gross sawlog
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, volcsgrs) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.tt1 AS volcsgrs " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(volcsgrs) AS tt1 " +
                    "FROM " + this.m_strTreeTable + " GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 25);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET volcsgrs = u.volcsgrs " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 26);

			//swd_volcsgrs      
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, swd_volcsgrs) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.ttl AS swd_volcsgrs " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(volcsgrs) AS ttl " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd < 300 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 27);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET swd_volcsgrs = u.swd_volcsgrs " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 28);

			//hwd_volcsgrs      
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, hwd_volcsgrs) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.ttl AS hwd_volcsgrs " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(volcsgrs) AS ttl " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd > 299 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 29);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET hwd_volcsgrs = u.hwd_volcsgrs " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 30);

            //qmd_all_inch 
            // quadratic mean diameter for all the live trees on the condition
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, qmd_all_inch) " +
                    "SELECT c.biosum_cond_id, SQRT(c.ba_ft2_ac / (.005454154 * c.tpacurr)) AS qmd_all_inch " +
                    "FROM " + this.m_strCondTable + " AS c " +
                    "WHERE c.biosum_status_cd = 9 AND c.ba_ft2_ac IS NOT NULL " +
                    "AND c.ba_ft2_ac <> 0 AND c.tpacurr IS NOT NULL AND c.tpacurr <> 0";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 31);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET qmd_all_inch = u.qmd_all_inch " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 32);

            //qmd_swd_inch      
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, qmd_swd_inch) " +
                    "SELECT c.biosum_cond_id, SQRT(c.swd_ba_ft2_ac / (.005454154 * c.swd_tpacurr)) AS qmd_swd_inch " +
                    "FROM " + this.m_strCondTable + " AS c " +
                    "WHERE c.biosum_status_cd = 9 AND c.swd_ba_ft2_ac IS NOT NULL AND c.swd_ba_ft2_ac <> 0 " +
                    "AND c.swd_tpacurr IS NOT NULL AND c.swd_tpacurr <> 0";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 33);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET qmd_swd_inch = u.qmd_swd_inch " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 34);

            // qmd_hwd_inch    
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, qmd_hwd_inch) " +
                    "SELECT c.biosum_cond_id, SQRT(c.hwd_ba_ft2_ac / (.005454154 * c.hwd_tpacurr)) AS qmd_hwd_inch " +
                    "FROM " + this.m_strCondTable + " AS c " +
                    "WHERE c.biosum_status_cd = 9 AND c.hwd_ba_ft2_ac IS NOT NULL " +
                    "AND c.hwd_ba_ft2_ac <> 0 AND c.hwd_tpacurr IS NOT NULL AND c.hwd_tpacurr <> 0";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 35);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET qmd_hwd_inch = u.qmd_hwd_inch " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 36);

			//VOL_AC_GRS_STEM_TTL_FT
            //gross wood volume of the total stem from ground to tip
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, vol_ac_grs_stem_ttl_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.ttl AS vol_ac_grs_stem_ttl_ft3 " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(voltsgrs * tpacurr) AS ttl " +
                    "FROM " + this.m_strTreeTable + " WHERE voltsgrs IS NOT NULL AND tpacurr IS NOT NULL " +
                    "AND statuscd = 1 AND dia >= 1 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 37);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET vol_ac_grs_stem_ttl_ft3 = u.vol_ac_grs_stem_ttl_ft3 " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 38);

			//hwd_vol_ac_grs_stem_ttl_ft
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table;";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table (biosum_cond_id, hwd_vol_ac_grs_stem_ttl_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.ttl AS hwd_vol_ac_grs_stem_ttl_ft3 " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(voltsgrs * tpacurr) AS ttl " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd > 299 AND voltsgrs IS NOT NULL " +
                    "AND tpacurr IS NOT NULL AND statuscd = 1 AND dia >= 1 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 39);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET hwd_vol_ac_grs_stem_ttl_ft3 = u.hwd_vol_ac_grs_stem_ttl_ft3 " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn,SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 40);

			//swd_vol_ac_grs_stem_ttl_ft     
            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				SQLite.m_strSQL = "DELETE FROM TEMPDB.cond_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                SQLite.m_strSQL = "INSERT INTO TEMPDB.cond_column_updates_work_table(biosum_cond_id, swd_vol_ac_grs_stem_ttl_ft3) " +
                    "SELECT DISTINCT(a.biosum_cond_id), a.ttl AS swd_vol_ac_grs_stem_ttl_ft3 " +
                    "FROM " + this.m_strTreeTable + " AS t, " +
                    "(SELECT biosum_cond_id, SUM(voltsgrs * tpacurr) AS ttl " +
                    "FROM " + this.m_strTreeTable + " WHERE spcd < 300 AND voltsgrs IS NOT NULL " +
                    "AND tpacurr IS NOT NULL AND statuscd = 1 AND dia >= 1 GROUP BY biosum_cond_id) AS a " +
                    "WHERE t.biosum_status_cd = 9 AND a.biosum_cond_id = t.biosum_cond_id";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 41);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strCondTable + " AS c " +
                    "SET swd_vol_ac_grs_stem_ttl_ft3 = u.swd_vol_ac_grs_stem_ttl_ft3 " +
                    "FROM TEMPDB.cond_column_updates_work_table AS u " +
                    "WHERE c.biosum_cond_id = u.biosum_cond_id";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 42);


			//----------------------PLOT COLUMN UPDATES-----------------------//
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//----------------------PLOT COLUMN UPDATES-----------------------//\r\n");
            SetThermValue(m_frmTherm.progressBar1, "Maximum", 5);
            SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
            SetThermValue(m_frmTherm.progressBar1, "Value", 0);
            SetLabelValue(m_frmTherm.lblMsg,"Text","Updating Plot Table Columns...Stand By");
            frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Control)this.m_frmTherm, "Refresh");

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
            {
                /********************************************
				 **update fvs_variant and fvs_loc_cd
				 ********************************************/
                string strPlotGeomTable = "FIADB.PLOTGEOM";
                SQLite.m_strSQL = "UPDATE " + this.m_strPlotTable + " AS p " +
                    "SET fvs_variant = g.fvs_variant, fvs_loc_cd = g.fvs_loc_cd " +
                    "FROM " + strPlotGeomTable + " AS g " +
                    "WHERE TRIM(g.cn) = p.cn";
                strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                strTime += " " + System.DateTime.Now.ToString();
            }
            SetThermValue(m_frmTherm.progressBar1, "Value", 1);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				/********************************************
				 **update the plot half state field
				 ********************************************/
                SQLite.m_strSQL = "UPDATE " + this.m_strPlotTable + " AS p " +
                    "SET half_state = SUBSTR(c.vol_loc_grp, 5, LENGTH(TRIM(c.vol_loc_grp))) " +
                    "FROM " + this.m_strCondTable + " AS c " +
                    "WHERE p.biosum_plot_id = c.biosum_plot_id AND c.condid = 1";
				strTime = System.DateTime.Now.ToString();
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
				strTime += " " + System.DateTime.Now.ToString();
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 2);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
				/***************************************************
				 **update the number of conditions on each plot
				 ***************************************************/
				//use the biosum_plot_input as our work table so delete all records
				SQLite.m_strSQL = "DELETE FROM TEMPDB.plot_column_updates_work_table";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

				//insert the condition counts into the work table
                SQLite.m_strSQL = "INSERT INTO TEMPDB.plot_column_updates_work_table (biosum_plot_id, cond_ttl) " +
                    "SELECT biosum_plot_id, COUNT(biosum_plot_id) " +
                    "FROM " + this.m_strCondTable + " GROUP BY biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 3);

            if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
			{
                SQLite.m_strSQL = "UPDATE " + this.m_strPlotTable + " AS p " +
                    "SET num_cond = i.cond_ttl, one_cond_yn = CASE WHEN i.cond_ttl > 1 THEN 'N' ELSE 'Y' END " +
                    "FROM TEMPDB.plot_column_updates_work_table AS i " +
                    "WHERE p.biosum_plot_id = i.biosum_plot_id";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
				SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
			}
            SetThermValue(m_frmTherm.progressBar1, "Value", 4);

		    //Set plot.gis_yard_dist_ft based on rddistcd using crosswalk table in ref_master.db
		    if (SQLite.m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
		    {
                string strRefMasterDb = frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultRefMasterDbFile;
                SQLite.m_strSQL = "ATTACH DATABASE '" + strRefMasterDb + "' AS REFMASTER";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                string strDistXrefTable = "REFMASTER.road_class_to_ft_xref";
                SQLite.m_strSQL = "UPDATE " + this.m_strPlotTable + " AS p " +
                    "SET gis_yard_dist_ft = x.distance_ft " +
                    "FROM " + strDistXrefTable + " AS x " +
                    "WHERE x.RDDISTCD = p.RDDISTCD";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            }
		    SetThermValue(m_frmTherm.progressBar1, "Value", 5);

		    return SQLite.m_intError;
		}

        private int ImportDownWoodyMaterials(System.Data.SQLite.SQLiteConnection p_conn)
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
            String strFiaCWD = "FIADB." + m_strDwmCwdTable;
            String strFiaFWD = "FIADB." + m_strDwmFwdTable;
            String strFiaDL = "FIADB." + m_strDwmDuffLitterTable;
            String strFiaTS = "FIADB." + m_strDwmTransectSegmentTable;

            string strMasterAuxDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db";

            //If any of the FIADB source DWM tables do not exist,
            //show message, uncheck the DWM checkbox, return early
            SQLite.m_strSQL = "ATTACH DATABASE '" + strMasterAuxDb + "' AS AUX";
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

            if (!SQLite.AttachedTableExist(p_conn, m_strDwmCwdTable) || 
                !SQLite.AttachedTableExist(p_conn, m_strDwmFwdTable) || 
                !SQLite.AttachedTableExist(p_conn, m_strDwmDuffLitterTable) ||
                !SQLite.AttachedTableExist(p_conn, m_strDwmTransectSegmentTable))
            {
                Func<bool, string, string> result = (boolTableExists, tableName) =>
                {
                    if (!boolTableExists) return "\r\n - " + tableName;
                    else return "";
                };
                DialogResult dlgResult = MessageBox.Show(String.Format(
                            "!!Error!!\nModule - uc_plot_input:ImportDownWoodyMaterials\n" + "Err Msg - " +
                            "At least one FIADB Source DWM table was not found:{0}{1}{2}{3}\r\nDo you wish to continue plot data input without DWM?",
                            result(SQLite.AttachedTableExist(p_conn, m_strDwmCwdTable), m_strDwmCwdTable),
                            result(SQLite.AttachedTableExist(p_conn, m_strDwmFwdTable), m_strDwmFwdTable),
                            result(SQLite.AttachedTableExist(p_conn, m_strDwmDuffLitterTable), m_strDwmDuffLitterTable),
                            result(SQLite.AttachedTableExist(p_conn, m_strDwmTransectSegmentTable), m_strDwmTransectSegmentTable)),
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

            string strSourceFields = "";
            string strDestFields = "";
            System.Data.DataTable dtDwmCwd = SQLite.getTableSchema(p_conn, "select * from AUX." + m_strDwmCwdTable);
            System.Data.DataTable dtDwmFwd = SQLite.getTableSchema(p_conn, "select * from AUX." + m_strDwmFwdTable);
            System.Data.DataTable dtDwmDuffLitter = SQLite.getTableSchema(p_conn, "select * from AUX." + m_strDwmDuffLitterTable);
            System.Data.DataTable dtDwmTransectSegment = SQLite.getTableSchema(p_conn, "select * from AUX." + m_strDwmTransectSegmentTable);
            System.Data.DataTable dtFIADBDwmCwd = SQLite.getTableSchema(p_conn, "select * from " + strFiaCWD);
            System.Data.DataTable dtFIADBDwmFwd = SQLite.getTableSchema(p_conn, "select * from " + strFiaFWD);
            System.Data.DataTable dtFIADBDwmDuffLitter = SQLite.getTableSchema(p_conn, "select * from " + strFiaDL);
            System.Data.DataTable dtFIADBDwmTransectSegment = SQLite.getTableSchema(p_conn, "select * from " + strFiaTS);

            //Preemptively remove any records that were not imported successfully 
            DeleteFromTablesWhereFilter(p_conn,
                new string[]
                    {"AUX." + m_strDwmCwdTable, "AUX." + m_strDwmFwdTable, "AUX." + m_strDwmDuffLitterTable, "AUX." + m_strDwmTransectSegmentTable},
                " WHERE biosum_status_cd=9");

            //DWM Coarse Woody Debris FIADB into master_aux.db Table
            if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
            {
                SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM CWD...Stand By");
                SetThermValue(m_frmTherm.progressBar1, "Value", 2);
                strSourceFields = CreateStrFieldsFromDataTables(dtFIADBDwmCwd, dtDwmCwd, "f.");
                strDestFields = CreateStrFieldsFromDataTables(dtFIADBDwmCwd, dtDwmCwd);
                InsertIntoDestTableFromSourceTable(p_conn, strFiaCWD, "AUX." + m_strDwmCwdTable, strSourceFields, strDestFields, true);

                UpdateDwmBiosumCondIds(p_conn, "AUX." + m_strDwmCwdTable);
                m_intError = SQLite.m_intError;
            }

            //DWM Fine Woody Debris FIADB into master_aux.db Table
            if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
            {
                SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM FWD...Stand By");
                SetThermValue(m_frmTherm.progressBar1, "Value", 4);
                strSourceFields = CreateStrFieldsFromDataTables(dtFIADBDwmFwd, dtDwmFwd, "f.");
                strDestFields = CreateStrFieldsFromDataTables(dtFIADBDwmFwd, dtDwmFwd);
                InsertIntoDestTableFromSourceTable(p_conn, strFiaFWD, "AUX." + m_strDwmFwdTable, strSourceFields, strDestFields, true);

                UpdateDwmBiosumCondIds(p_conn, "AUX." + m_strDwmFwdTable);
                m_intError = SQLite.m_intError;
            }

            //DWM Duff Litter Fuel FIADB into master_aux.db Table
            if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
            {
                SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM DuffLitter...Stand By");
                SetThermValue(m_frmTherm.progressBar1, "Value", 6);
                strSourceFields = CreateStrFieldsFromDataTables(dtFIADBDwmDuffLitter, dtDwmDuffLitter, "f.");
                strDestFields = CreateStrFieldsFromDataTables(dtFIADBDwmDuffLitter, dtDwmDuffLitter);
                InsertIntoDestTableFromSourceTable( p_conn, strFiaDL, "AUX." + m_strDwmDuffLitterTable, strSourceFields, strDestFields, true);

                SQLite.m_strSQL = "UPDATE AUX." + m_strDwmDuffLitterTable + " SET duffdep=0 WHERE duffdep IS NULL AND duff_nonsample_reasn_cd IS NULL";
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                SQLite.m_strSQL = "UPDATE " + m_strDwmDuffLitterTable + " SET littdep=0 WHERE littdep IS NULL AND litter_nonsample_reasn_cd IS NULL";
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                UpdateDwmBiosumCondIds( p_conn, "AUX." + m_strDwmDuffLitterTable);
                m_intError = SQLite.m_intError;
            }

            //DWM Transect Segment FIADB into master_aux.db Table
            if (m_intError == 0 && !GetBooleanValue((System.Windows.Forms.Control)m_frmTherm, "AbortProcess"))
            {
                SetLabelValue(m_frmTherm.lblMsg, "Text", "Importing DWM Transect Segment...Stand By");
                SetThermValue(m_frmTherm.progressBar1, "Value", 8);
                strSourceFields = CreateStrFieldsFromDataTables(dtFIADBDwmTransectSegment, dtDwmTransectSegment, "f.");
                strDestFields = CreateStrFieldsFromDataTables(dtFIADBDwmTransectSegment, dtDwmTransectSegment);
                InsertIntoDestTableFromSourceTable(p_conn, strFiaTS, "AUX." + m_strDwmTransectSegmentTable, strSourceFields, strDestFields, true);

                UpdateDwmBiosumCondIds(p_conn, "AUX." + m_strDwmTransectSegmentTable);
                m_intError = SQLite.m_intError;
            }

            SetLabelValue(m_frmTherm.lblMsg, "Text", "");
            SetThermValue(m_frmTherm.progressBar1, "Value", GetThermValue(m_frmTherm.progressBar1, "Maximum"));
            return SQLite.m_intError;
        }

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
				
		    if (this.LoadDBFiadbPopEvalTable() && m_intError==0)
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

		private void dbInputStateCounty()
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

        private void dbInputPlot()
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
		    this.dbInputPlot();
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

                    this.LoadDBPlotCondTreeData_Start();
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

                    this.LoadDBPlotCondTreeData_Start();
                }
            }	
		}

		private void btnDBInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
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

				if (SQLite == null)
				{
                    SQLite = new DataMgr();
				}
				else
				{
					SQLite = null;
					SQLite = new DataMgr();
				}

                if (this.m_strCurrentProcess == "dbFIADBFileInput" ||
                        this.m_strCurrentProcess == "txtFileInput" ||
                        this.m_strCurrentProcess == "dbIDBFileInput")
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(SQLite.GetConnectionString(m_strMasterDbFile)))
                    {
                        conn.Open();

                        DeleteFromTablesWhereFilter(conn,
                            new string[]
                            {m_strPlotTable, m_strCondTable, m_strTreeTable, m_strSiteTreeTable},
                            " WHERE biosum_status_cd=9");
                    }

                    string strMasterAuxConn = SQLite.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\master_aux.db");
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strMasterAuxConn))
                    {
                        conn.Open();

                        DeleteFromTablesWhereFilter(conn,
                            new string[]
                            {m_strDwmCwdTable, m_strDwmFwdTable, m_strDwmDuffLitterTable, m_strDwmTransectSegmentTable},
                            " WHERE biosum_status_cd=9");
                    }
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
                    this.LoadDBFiadbPopFiles();
                    this.m_strLoadedPopEstUnitTxtInputFile = "";
                    this.m_strLoadedPopEvalTxtInputFile = "";
                    this.m_strLoadedPopStratumTxtInputFile = "";
                    this.m_strLoadedPpsaTxtInputFile = "";
                    if (this.rdoFilterNone.Checked && m_intError == 0)
                    {

                        this.LoadDBPlotCondTreeData_Start();

                    }
                    else if (this.rdoFilterByFile.Checked && m_intError == 0)
                    {
                        this.m_strPlotIdList = this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                            this.LoadDBPlotCondTreeData_Start();

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
					this.LoadDBFiadbPopFiles();
				}
				

				if (m_intError==0)
				{
					this.dbInputStateCounty();
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
            this.m_strCurrentProcess = "dbFIADBFileInput";
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

        private void LoadDBPlotCondTreeData_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "dbFIADBFileInput";
            this.StartTherm("2", "Add SQLite Plot,Cond,Site Tree, & Tree Table Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(LoadDBPlotCondTreeData_Process));
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

        private void LoadDBPlotCondTreeData_Finish()
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

		private void btnboxDBFiadbInputFile_Click(object sender, System.EventArgs e)
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

		private void btnDBFiadbInputNext_Click(object sender, System.EventArgs e)
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
			this.grpboxDBFiadbInput.Visible=false;

		}
		private bool LoadDBFiadbPopEvalTable()
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.LoadDBFiadPopEvalTable\r\n");
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
		private void LoadDBFiadbPopFiles()
		{
		    this.m_strCurrentProcess="dbFiadbInputPopTables";	

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.LoadDBFiadbPopFiles\r\n");
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

				this.thdProcessRecords = new Thread(new ThreadStart(dbFiadbInputPopTables));
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
					this.thdProcessRecords = new Thread(new ThreadStart(dbFiadbInputPopTables));
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
					this.thdProcessRecords = new Thread(new ThreadStart(dbFiadbInputPopTables));
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
		private void dbFiadbInputPopTables()
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
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.dbFiadbInputPopTables\r\n");
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
                        " SELECT " + strFields + " FROM FIADB." + strSourceTable + " " +
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

		private void btnDBFiadbInputCancel_Click(object sender, System.EventArgs e)
		{
			((frmDialog)this.ParentForm).Close();
		}

		private void lstFIADBInv_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (lstFIADBInv.SelectedItems.Count==0) return;
			this.m_strCurrFIADBEvalId = this.lstFIADBInv.SelectedItems[0].Text.Trim();
			this.m_strCurrFIADBRsCd = this.lstFIADBInv.SelectedItems[0].SubItems[1].Text.Trim();
		
		}

		

        private void SaveSqlitePopTables(System.Data.SQLite.SQLiteConnection p_conn, string strEvalId, string strRscd)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//uc_plot_input.SaveSqlitePopTables\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            SQLite.m_strSQL = "SELECT COUNT(*) FROM " + this.m_strPopEvalTable + " WHERE evalid=" + strEvalId +
                " AND rscd= " + strRscd;
            int intCount = (int) SQLite.getRecordCount(p_conn, SQLite.m_strSQL, this.m_strPopEvalTable);
            if (intCount > 0 && frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "Records exist for evalid " + strEvalId + ". POP tables will not be loaded.\r\n");
                return;
            }
            else
            {
                SQLite.m_strSQL = "INSERT INTO " + this.m_strPopEvalTable + "(CN,RSCD,EVALID,EVAL_DESCR,STATECD," +
                    "LOCATION_NM,REPORT_YEAR_NM,NOTES," +
                    "START_INVYR,END_INVYR,GROWTH_ACCT,LAND_ONLY, MODIFIED_DATE, BIOSUM_STATUS_CD)" +
                    "SELECT CN,RSCD,EVALID,EVAL_DESCR,STATECD," +
                    "LOCATION_NM,REPORT_YEAR_NM,NOTES," +
                    "START_INVYR,END_INVYR,GROWTH_ACCT,LAND_ONLY, MODIFIED_DATE,1 " +
                    "FROM FIADB." + this.m_strPopEvalTable +
                    " WHERE evalid = " + strEvalId + " AND rscd= " + strRscd;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                        SQLite.m_strSQL + "\r\n");
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                //get the fiabiosum table structures
                string[] arrTables = new string[] { this.m_strPopEstUnitTable, this.m_strPopStratumTable, this.m_strPpsaTable };
                string strPopEstnUnitTableSrc = frmMain.g_oDelegate.GetControlPropertyValue(cmbFiadbPopEstUnitTable, "Text", false).ToString().Trim();
                string strPopStratumTableSrc = frmMain.g_oDelegate.GetControlPropertyValue(cmbFiadbPopStratumTable, "Text", false).ToString().Trim();
                string strPpsaTableSrc = frmMain.g_oDelegate.GetControlPropertyValue(cmbFiadbPpsaTable, "Text", false).ToString().Trim();
                string[] arrSourceTables = new string[] { strPopEstnUnitTableSrc, strPopStratumTableSrc, strPpsaTableSrc };
                int i = 0;
                foreach (var pTable in arrTables)
                {
                    string strFields = "";
                    DataTable dtDestSchema = SQLite.getTableSchema(p_conn, "SELECT * FROM " + pTable);
                    DataTable dtSourceSchema = SQLite.getTableSchema(p_conn, "SELECT * FROM FIADB." + arrSourceTables[i].Trim());

                    //build field list string to insert sql by matching 
                    //up the column names in the biosum plot table and the fiadb plot table
                    strFields = "";
                    for (int x = 0; x <= dtDestSchema.Rows.Count - 1; x++)
                    {
                        string strCol = dtDestSchema.Rows[x]["columnname"].ToString().Trim();
                        //see if there is an equivalent FIADB column
                        for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
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
                    SQLite.m_strSQL = "INSERT INTO " + pTable + " (" + strFields + ")" +
                        " SELECT " + strFields + " FROM FIADB." + arrSourceTables[i] +
                        " WHERE evalid = " + strEvalId + " AND rscd = " + strRscd;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);

                    SQLite.m_strSQL = "UPDATE " + pTable +
                        " SET biosum_status_cd = 1 " +
                        "WHERE evalid = " + strEvalId;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile,
                            SQLite.m_strSQL + "\r\n");
                    SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
                    i++;
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

        private void btnDBFiadbInputHelp_Click(object sender, EventArgs e)
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

        private void SaveLoadConfigurationTxt(System.Data.SQLite.SQLiteConnection p_conn)
        {
            string strConfigTxtFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\biosum_plot_input_configuration.txt";
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("========================================================");
            stringBuilder.AppendLine("Plot data loaded: " + DateTime.Now.ToString());
            stringBuilder.AppendLine("========================================================");

            string strSqliteVersion = "";
            SQLite.m_strSQL = "SELECT sqlite_version() AS v";
            SQLite.SqlQueryReader(p_conn, SQLite.m_strSQL);
            if (SQLite.m_DataReader.HasRows)
            {
                while (SQLite.m_DataReader.Read())
                {
                    strSqliteVersion = SQLite.m_DataReader["v"].ToString();
                }
            }
            SQLite.m_DataReader.Close();

            stringBuilder.AppendLine("SQLite version number: " + strSqliteVersion);

            stringBuilder.AppendLine("Source FIADB database: " + m_strCurrentFiadbInputFile);
            // if using .txt file
            if (rdoFilterByFile.Checked)
            {
                stringBuilder.AppendLine("Plot_CN list .txt file: " + txtFilterByFile.Text.Trim());
            }

            string strChecked = "No";
            if (chkForested.Checked)
            {
                strChecked = "Yes";
            }
            stringBuilder.AppendLine("Forested: " + strChecked);
            strChecked = "No";
            if (chkNonForested.Checked)
            {
                strChecked = "Yes";
            }
            stringBuilder.AppendLine("Non Forested: " + strChecked);
            strChecked = "No";
            if (ckImportSeedlings.Checked)
            {
                strChecked = "Yes";
            }
            stringBuilder.AppendLine("Load seedling data: " + strChecked);
            strChecked = "No";
            if (chkDwmImport.Checked)
            {
                strChecked = "Yes";
            }
            stringBuilder.AppendLine("Load down woody material data: " + strChecked);

            stringBuilder.AppendLine("Percent condition proportion: " + m_strCondPropPercent);
            stringBuilder.AppendLine("EvalId: " + m_strCurrFIADBEvalId);
            frmMain.g_oUtils.WriteText(strConfigTxtFile, stringBuilder.ToString());
        }

        private void SaveLoadConfigurationTable(System.Data.SQLite.SQLiteConnection p_conn)
        {
            string strConfigTable = "biosum_plot_input_configuration";
            if (!SQLite.TableExist(p_conn, strConfigTable))
            {
                SQLite.m_strSQL = "CREATE TABLE " + strConfigTable +
                    " (Setting CHAR(255), Value CHAR(255))";
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            }
            else
            {
                SQLite.m_strSQL = "DELETE FROM " + strConfigTable;
                SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
            }

            string strSqliteVersion = "";
            SQLite.m_strSQL = "SELECT sqlite_version() AS v";
            SQLite.SqlQueryReader(p_conn, SQLite.m_strSQL);
            if (SQLite.m_DataReader.HasRows)
            {
                while (SQLite.m_DataReader.Read())
                {
                    strSqliteVersion = SQLite.m_DataReader["v"].ToString();
                }
            }
            SQLite.m_DataReader.Close();

            SQLite.m_strSQL = "INSERT INTO " + strConfigTable +
                " VALUES ('Load date and time', '" + DateTime.Now.ToString() + "')," +
                "('SQLite version number', '" + strSqliteVersion + "')," +
                "('Source FIADB database', '" + m_strCurrentFiadbInputFile + "'),";
            if (rdoFilterByFile.Checked)
            {
                SQLite.m_strSQL += "('Plot_CN list .txt file', '" + txtFilterByFile.Text.Trim() + "'),";
            }
            string strChecked = "No";
            if (chkForested.Checked)
            {
                strChecked = "Yes";
            }
            SQLite.m_strSQL += "('Forested', '" + strChecked + "'),";
            strChecked = "No";
            if (chkNonForested.Checked)
            {
                strChecked = "Yes";
            }
            SQLite.m_strSQL += "('Non Forested', '" + strChecked + "'),";
            strChecked = "No";
            if (ckImportSeedlings.Checked)
            {
                strChecked = "Yes";
            }
            SQLite.m_strSQL += "('Load seedling data', '" + strChecked + "'),";
            strChecked = "No";
            if (chkDwmImport.Checked)
            {
                strChecked = "Yes";
            }
            SQLite.m_strSQL += "('Load down woody material', '" + strChecked + "')," +
                "('Percent condition proportion', '" + m_strCondPropPercent + "')," +
                "('EvalId', '" + m_strCurrFIADBEvalId + "')";
            SQLite.SqlNonQuery(p_conn, SQLite.m_strSQL);
        }

        
    }
    
	
}

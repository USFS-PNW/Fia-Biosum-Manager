using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;
using SQLite.ADO;
using System.Data.SQLite;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_ffe.
	/// </summary>
    public class uc_optimizer_scenario_calculated_variables : System.Windows.Forms.UserControl
    {
        private System.Windows.Forms.GroupBox groupBox1;
        private System.ComponentModel.IContainer components;
        //private int m_intFullHt=400;
        public System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
        public System.Data.DataSet m_DataSet;
        public System.Data.OleDb.OleDbConnection m_OleDbConnectionMaster;
        public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
        public System.Data.OleDb.OleDbCommand m_OleDbCommand;
        public System.Data.DataRelation m_DataRelation;
        public System.Data.DataRow m_DataRow;
        public int m_intError = 0;
        public string m_strError = "";
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnPrev;
        private FIA_Biosum_Manager.utils m_oUtils;
        public System.Windows.Forms.Label lblTitle;
        private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario = null;
        private FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker _uc_tiebreaker;
        string m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_optimizer_calculated_variables_debug.txt";
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultTreatmentOptimizerFile;

        private int m_intCurVar = -1;
        public System.Windows.Forms.GroupBox grpboxDetails;

        public bool m_bSave = false;
        private ado_data_access m_oAdo;
        private ado_data_access m_oAdoFvs;
        private DataMgr m_oDataMgr = new DataMgr();
        private string m_strTempMDB;
        private bool m_bUsingSqlite;

        const int COLUMN_CHECKBOX = 0;
        const int COLUMN_OPTIMIZE_VARIABLE = 1;
        const int COLUMN_FVS_VARIABLE = 2;
        const int COLUMN_VALUESOURCE = 3;
        const int COLUMN_MAXMIN = 4;
        const int COLUMN_USEFILTER = 5;
        const int COLUMN_FILTER_OPERATOR = 6;
        const int COLUMN_FILTER_VALUE = 7;
        const string VARIABLE_ECON = "ECON";
        const string VARIABLE_FVS = "FVS";
        public const string PREFIX_CHIP_VOLUME = "chip_volume";
        public const string PREFIX_MERCH_VOLUME = "merchantable_volume";
        public const string PREFIX_TOTAL_VOLUME = "total_volume";
        public const string PREFIX_NET_REVENUE = "net_revenue";
        public const string PREFIX_TREATMENT_HAUL_COSTS = "treatment_haul_costs";
        public const string PREFIX_ONSITE_TREATMENT_COSTS = "onsite_treatment_costs";
        //These parallel arrays must remain in the same order
        static readonly string[] PREFIX_ECON_VALUE_ARRAY = { PREFIX_TOTAL_VOLUME, PREFIX_MERCH_VOLUME, PREFIX_CHIP_VOLUME,  
                                                             PREFIX_NET_REVENUE, PREFIX_TREATMENT_HAUL_COSTS, PREFIX_ONSITE_TREATMENT_COSTS };
        static readonly string[] PREFIX_ECON_NAME_ARRAY = { "Total Volume", "Merchantable Volume", "Chip Volume",
                                                            "Net Revenue","Treatment And Haul Costs", "OnSite Treatment Costs"};
        private bool b_FVSTableEnabled = false;
        private string m_strFvsViewTableName = "view_weights";
        string m_strCalculatedVariablesAccdb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
            "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
        string m_strCalculatedVariablesDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
            "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;

        const int TABLE_COUNT = 2;
        const int ECON_DETAILS_TABLE = 0;
        const int FVS_DETAILS_TABLE = 1;




        private FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective.Variables _oCurVar;
        public bool m_bFirstTime = true;
        private bool _bDisplayAuditMsg = true;
        private bool m_bIgnoreListViewItemCheck = false;
        private int m_intPrevColumnIdx = -1;
        private System.Windows.Forms.GroupBox grpboxSummary;
        private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
        private ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private string m_strLastValue = "";
        private FIA_Biosum_Manager.frmMain m_frmMain;
        public int m_DialogHt;
        public Panel pnlDetails;
        private Label label7;
        public Button btnFvsCalculate;
        private Button btnFvsDetailsCancel;
        private GroupBox grpBoxFvsBaseline;
        private ComboBox cboFvsVariableBaselinePkg;
        private GroupBox groupBox3;
        private ListBox lstFVSFieldsList;
        private GroupBox groupBox2;
        private ListBox lstFVSTablesList;
        private Label LblSelectedVariable;
        private Label lblSelectedFVSVariable;
        private TextBox txtFVSVariableDescr;
        private Label label8;
        public int m_DialogWd;
        private Panel pnlSummary;
        private Button btnProperties;
        private Button btnDeleteFvsVariable;
        private Button btnNewFvs;
        private ListView lstVariables;
        private Button btnCancelSummary;
        private ColumnHeader vName;
        private ColumnHeader vDescription;
        private Button BtnHelp;
        private ListViewAlternateBackgroundColors m_oLvAlternateColors = new FIA_Biosum_Manager.ListViewAlternateBackgroundColors();
        private ListViewColumnSorter lvwColumnSorter;
        private const int COL_YEAR = 1;
        private const int COL_SEQNUM = 2;
        private Button btnNewEcon;
        public GroupBox grpBoxEconomicVariable;
        public Panel panel1;
        private Button BtnHelpEconVariable;
        private TextBox txtEconVariableDescr;
        private Label label1;
        private Label label2;
        public Button BtnSaveEcon;
        private Button btnEconDetailsCancel;
        private GroupBox groupBox8;
        private ListBox lstEconVariablesList;
        private Label lblSelectedEconType;
        private Label label4;
        private TextBox txtEconVariableTotalWeight;
        private Label label6;
        private TextBox txtFvsVariableTotalWeight;
        private Label label5;
        private ColumnHeader vType;
        private DataGrid m_dg;
        private System.Data.DataTable m_dtTableSchema;
        private System.Data.DataView m_dv;
        private System.Data.DataView m_econ_dv;
        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> m_dictFVSTables;
        private Button btnFVSVariableValue;
        private ColumnHeader vId;
        private Button btnEconVariableType;
        private Label lblEconVariableName;
        private Label lblFvsVariableName;
        private DataGrid m_dgEcon;
        private ColumnHeader vBaselineRxPkg;
        private ColumnHeader vVariableSource;
        private Button BtnDeleteEconVariable;
        private Button BtnHelpCalculatedMenu;
        public Button BtnFvsImport;
        public Button BtnEconImport;
        private Button BtnRecalculateAll;
        private FIA_Biosum_Manager.OptimizerScenarioTools m_oOptimizerScenarioTools = new OptimizerScenarioTools();
        private frmTherm m_frmTherm;
        private int idxRxCycle = 0;
        private int idxWeight = 1;
        private int idxPreOrPost = 2;
        private int counter1Interval = 5;

        public uc_optimizer_scenario_calculated_variables(FIA_Biosum_Manager.frmMain p_frmMain)
        {
            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();
            this.m_oUtils = new utils();
            this.m_frmMain = p_frmMain;

            this.grpboxDetails.Top = grpboxSummary.Top;
            this.grpboxDetails.Left = this.grpboxSummary.Left;
            this.grpboxDetails.Height = this.grpboxSummary.Height;
            this.grpboxDetails.Width = this.grpboxSummary.Width;
            this.grpboxDetails.Hide();

            this.grpBoxEconomicVariable.Top = grpboxSummary.Top;
            this.grpBoxEconomicVariable.Left = this.grpboxSummary.Left;
            this.grpBoxEconomicVariable.Height = this.grpboxSummary.Height;
            this.grpBoxEconomicVariable.Width = this.grpboxSummary.Width;
            this.grpBoxEconomicVariable.Hide();

            //m_oValidate.RoundDecimalLength = 0;
            //m_oValidate.Money = false;
            //m_oValidate.NullsAllowed = false;
            //m_oValidate.TestForMaxMin = false;
            //m_oValidate.MinValue = -1000;
            //m_oValidate.TestForMin = true;

            m_oLvAlternateColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvAlternateColors.ReferenceListView = this.lstVariables;
            this.m_oLvAlternateColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvAlternateColors.CustomFullRowSelect = true;
            if (frmMain.g_oGridViewFont != null) lstVariables.Font = frmMain.g_oGridViewFont;

            // TODO: Add any initialization after the InitializeComponent call
            this.m_DialogWd = this.Width + 25;
            this.m_DialogHt = this.pnlDetails.Height + 120;
            this.Height = m_DialogHt - 40;

            this.m_oEnv = new env();
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
            this.loadvalues();
            //@ToDo
            //this.loadvalues_sqlite();
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
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
            this.grpBoxEconomicVariable = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.BtnEconImport = new System.Windows.Forms.Button();
            this.BtnDeleteEconVariable = new System.Windows.Forms.Button();
            this.m_dgEcon = new System.Windows.Forms.DataGrid();
            this.lblEconVariableName = new System.Windows.Forms.Label();
            this.btnEconVariableType = new System.Windows.Forms.Button();
            this.txtEconVariableTotalWeight = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.BtnHelpEconVariable = new System.Windows.Forms.Button();
            this.txtEconVariableDescr = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.BtnSaveEcon = new System.Windows.Forms.Button();
            this.btnEconDetailsCancel = new System.Windows.Forms.Button();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.lstEconVariablesList = new System.Windows.Forms.ListBox();
            this.lblSelectedEconType = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.grpboxSummary = new System.Windows.Forms.GroupBox();
            this.pnlSummary = new System.Windows.Forms.Panel();
            this.BtnRecalculateAll = new System.Windows.Forms.Button();
            this.BtnHelpCalculatedMenu = new System.Windows.Forms.Button();
            this.btnNewEcon = new System.Windows.Forms.Button();
            this.btnCancelSummary = new System.Windows.Forms.Button();
            this.btnProperties = new System.Windows.Forms.Button();
            this.btnNewFvs = new System.Windows.Forms.Button();
            this.lstVariables = new System.Windows.Forms.ListView();
            this.vName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vDescription = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vType = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vBaselineRxPkg = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.vVariableSource = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.grpboxDetails = new System.Windows.Forms.GroupBox();
            this.pnlDetails = new System.Windows.Forms.Panel();
            this.BtnFvsImport = new System.Windows.Forms.Button();
            this.lblFvsVariableName = new System.Windows.Forms.Label();
            this.btnFVSVariableValue = new System.Windows.Forms.Button();
            this.m_dg = new System.Windows.Forms.DataGrid();
            this.btnDeleteFvsVariable = new System.Windows.Forms.Button();
            this.txtFvsVariableTotalWeight = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.BtnHelp = new System.Windows.Forms.Button();
            this.txtFVSVariableDescr = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.btnFvsCalculate = new System.Windows.Forms.Button();
            this.btnFvsDetailsCancel = new System.Windows.Forms.Button();
            this.grpBoxFvsBaseline = new System.Windows.Forms.GroupBox();
            this.cboFvsVariableBaselinePkg = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lstFVSFieldsList = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lstFVSTablesList = new System.Windows.Forms.ListBox();
            this.LblSelectedVariable = new System.Windows.Forms.Label();
            this.lblSelectedFVSVariable = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.grpBoxEconomicVariable.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dgEcon)).BeginInit();
            this.groupBox8.SuspendLayout();
            this.grpboxSummary.SuspendLayout();
            this.pnlSummary.SuspendLayout();
            this.grpboxDetails.SuspendLayout();
            this.pnlDetails.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).BeginInit();
            this.grpBoxFvsBaseline.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.grpBoxEconomicVariable);
            this.groupBox1.Controls.Add(this.grpboxSummary);
            this.groupBox1.Controls.Add(this.grpboxDetails);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Leave += new System.EventHandler(this.groupBox1_Leave);
            // 
            // grpBoxEconomicVariable
            // 
            this.grpBoxEconomicVariable.BackColor = System.Drawing.SystemColors.Control;
            this.grpBoxEconomicVariable.Controls.Add(this.panel1);
            this.grpBoxEconomicVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxEconomicVariable.ForeColor = System.Drawing.Color.Black;
            this.grpBoxEconomicVariable.Location = new System.Drawing.Point(6, 1027);
            this.grpBoxEconomicVariable.Name = "grpBoxEconomicVariable";
            this.grpBoxEconomicVariable.Size = new System.Drawing.Size(856, 472);
            this.grpBoxEconomicVariable.TabIndex = 36;
            this.grpBoxEconomicVariable.TabStop = false;
            this.grpBoxEconomicVariable.Text = "Weighted Economic Variable";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.BtnEconImport);
            this.panel1.Controls.Add(this.BtnDeleteEconVariable);
            this.panel1.Controls.Add(this.m_dgEcon);
            this.panel1.Controls.Add(this.lblEconVariableName);
            this.panel1.Controls.Add(this.btnEconVariableType);
            this.panel1.Controls.Add(this.txtEconVariableTotalWeight);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.BtnHelpEconVariable);
            this.panel1.Controls.Add(this.txtEconVariableDescr);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.BtnSaveEcon);
            this.panel1.Controls.Add(this.btnEconDetailsCancel);
            this.panel1.Controls.Add(this.groupBox8);
            this.panel1.Controls.Add(this.lblSelectedEconType);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 22);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(850, 447);
            this.panel1.TabIndex = 70;
            // 
            // BtnEconImport
            // 
            this.BtnEconImport.Enabled = false;
            this.BtnEconImport.Location = new System.Drawing.Point(380, 171);
            this.BtnEconImport.Name = "BtnEconImport";
            this.BtnEconImport.Size = new System.Drawing.Size(140, 30);
            this.BtnEconImport.TabIndex = 97;
            this.BtnEconImport.Text = "Import weights";
            this.BtnEconImport.Click += new System.EventHandler(this.BtnEconImport_Click);
            // 
            // BtnDeleteEconVariable
            // 
            this.BtnDeleteEconVariable.Enabled = false;
            this.BtnDeleteEconVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnDeleteEconVariable.Location = new System.Drawing.Point(568, 402);
            this.BtnDeleteEconVariable.Name = "BtnDeleteEconVariable";
            this.BtnDeleteEconVariable.Size = new System.Drawing.Size(64, 30);
            this.BtnDeleteEconVariable.TabIndex = 96;
            this.BtnDeleteEconVariable.Text = "Delete";
            this.BtnDeleteEconVariable.Click += new System.EventHandler(this.BtnDeleteEconVariable_Click);
            // 
            // m_dgEcon
            // 
            this.m_dgEcon.DataMember = "";
            this.m_dgEcon.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dgEcon.Location = new System.Drawing.Point(16, 171);
            this.m_dgEcon.Name = "m_dgEcon";
            this.m_dgEcon.Size = new System.Drawing.Size(327, 177);
            this.m_dgEcon.TabIndex = 95;
            this.m_dgEcon.CurrentCellChanged += new System.EventHandler(this.m_dgEcon_CurCellChange);
            this.m_dgEcon.Leave += new System.EventHandler(this.m_dgEcon_Leave);
            // 
            // lblEconVariableName
            // 
            this.lblEconVariableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEconVariableName.Location = new System.Drawing.Point(210, 360);
            this.lblEconVariableName.Name = "lblEconVariableName";
            this.lblEconVariableName.Size = new System.Drawing.Size(302, 24);
            this.lblEconVariableName.TabIndex = 94;
            this.lblEconVariableName.Text = "Not Defined";
            // 
            // btnEconVariableType
            // 
            this.btnEconVariableType.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEconVariableType.Location = new System.Drawing.Point(257, 26);
            this.btnEconVariableType.Name = "btnEconVariableType";
            this.btnEconVariableType.Size = new System.Drawing.Size(100, 80);
            this.btnEconVariableType.TabIndex = 93;
            this.btnEconVariableType.Text = "Select";
            this.btnEconVariableType.Click += new System.EventHandler(this.btnEconVariableType_Click);
            // 
            // txtEconVariableTotalWeight
            // 
            this.txtEconVariableTotalWeight.BackColor = System.Drawing.SystemColors.Control;
            this.txtEconVariableTotalWeight.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEconVariableTotalWeight.Location = new System.Drawing.Point(393, 301);
            this.txtEconVariableTotalWeight.Name = "txtEconVariableTotalWeight";
            this.txtEconVariableTotalWeight.Size = new System.Drawing.Size(121, 26);
            this.txtEconVariableTotalWeight.TabIndex = 92;
            this.txtEconVariableTotalWeight.Text = "4.0";
            this.txtEconVariableTotalWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(387, 279);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(160, 24);
            this.label6.TabIndex = 91;
            this.label6.Text = "TOTAL WEIGHTS";
            // 
            // BtnHelpEconVariable
            // 
            this.BtnHelpEconVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnHelpEconVariable.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.BtnHelpEconVariable.Location = new System.Drawing.Point(495, 402);
            this.BtnHelpEconVariable.Name = "BtnHelpEconVariable";
            this.BtnHelpEconVariable.Size = new System.Drawing.Size(64, 30);
            this.BtnHelpEconVariable.TabIndex = 87;
            this.BtnHelpEconVariable.Text = "Help";
            this.BtnHelpEconVariable.Click += new System.EventHandler(this.BtnHelpEconVariable_Click);
            // 
            // txtEconVariableDescr
            // 
            this.txtEconVariableDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEconVariableDescr.Location = new System.Drawing.Point(173, 386);
            this.txtEconVariableDescr.Multiline = true;
            this.txtEconVariableDescr.Name = "txtEconVariableDescr";
            this.txtEconVariableDescr.Size = new System.Drawing.Size(259, 40);
            this.txtEconVariableDescr.TabIndex = 86;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(13, 389);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(212, 24);
            this.label1.TabIndex = 85;
            this.label1.Text = "Description:";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(13, 359);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(212, 24);
            this.label2.TabIndex = 79;
            this.label2.Text = "Weighted variable name:";
            // 
            // BtnSaveEcon
            // 
            this.BtnSaveEcon.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnSaveEcon.Location = new System.Drawing.Point(641, 402);
            this.BtnSaveEcon.Name = "BtnSaveEcon";
            this.BtnSaveEcon.Size = new System.Drawing.Size(76, 30);
            this.BtnSaveEcon.TabIndex = 77;
            this.BtnSaveEcon.Text = "Save";
            this.BtnSaveEcon.Click += new System.EventHandler(this.BtnSaveEcon_Click);
            // 
            // btnEconDetailsCancel
            // 
            this.btnEconDetailsCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEconDetailsCancel.Location = new System.Drawing.Point(726, 402);
            this.btnEconDetailsCancel.Name = "btnEconDetailsCancel";
            this.btnEconDetailsCancel.Size = new System.Drawing.Size(64, 30);
            this.btnEconDetailsCancel.TabIndex = 75;
            this.btnEconDetailsCancel.Text = "Cancel";
            this.btnEconDetailsCancel.Click += new System.EventHandler(this.btnEconDetailsCancel_Click);
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.lstEconVariablesList);
            this.groupBox8.Location = new System.Drawing.Point(18, 5);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(200, 133);
            this.groupBox8.TabIndex = 71;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Variable";
            // 
            // lstEconVariablesList
            // 
            this.lstEconVariablesList.FormattingEnabled = true;
            this.lstEconVariablesList.ItemHeight = 20;
            this.lstEconVariablesList.Location = new System.Drawing.Point(6, 21);
            this.lstEconVariablesList.Name = "lstEconVariablesList";
            this.lstEconVariablesList.Size = new System.Drawing.Size(181, 84);
            this.lstEconVariablesList.TabIndex = 70;
            // 
            // lblSelectedEconType
            // 
            this.lblSelectedEconType.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelectedEconType.Location = new System.Drawing.Point(236, 145);
            this.lblSelectedEconType.Name = "lblSelectedEconType";
            this.lblSelectedEconType.Size = new System.Drawing.Size(302, 24);
            this.lblSelectedEconType.TabIndex = 69;
            this.lblSelectedEconType.Text = "Not Defined";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(11, 144);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(237, 24);
            this.label4.TabIndex = 68;
            this.label4.Text = "Selected Economic Variable:";
            // 
            // grpboxSummary
            // 
            this.grpboxSummary.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxSummary.Controls.Add(this.pnlSummary);
            this.grpboxSummary.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxSummary.ForeColor = System.Drawing.Color.Black;
            this.grpboxSummary.Location = new System.Drawing.Point(8, 48);
            this.grpboxSummary.Name = "grpboxSummary";
            this.grpboxSummary.Size = new System.Drawing.Size(856, 472);
            this.grpboxSummary.TabIndex = 35;
            this.grpboxSummary.TabStop = false;
            // 
            // pnlSummary
            // 
            this.pnlSummary.AutoScroll = true;
            this.pnlSummary.Controls.Add(this.BtnRecalculateAll);
            this.pnlSummary.Controls.Add(this.BtnHelpCalculatedMenu);
            this.pnlSummary.Controls.Add(this.btnNewEcon);
            this.pnlSummary.Controls.Add(this.btnCancelSummary);
            this.pnlSummary.Controls.Add(this.btnProperties);
            this.pnlSummary.Controls.Add(this.btnNewFvs);
            this.pnlSummary.Controls.Add(this.lstVariables);
            this.pnlSummary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlSummary.Location = new System.Drawing.Point(3, 22);
            this.pnlSummary.Name = "pnlSummary";
            this.pnlSummary.Size = new System.Drawing.Size(850, 447);
            this.pnlSummary.TabIndex = 12;
            // 
            // BtnRecalculateAll
            // 
            this.BtnRecalculateAll.Enabled = false;
            this.BtnRecalculateAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnRecalculateAll.Location = new System.Drawing.Point(87, 360);
            this.BtnRecalculateAll.Name = "BtnRecalculateAll";
            this.BtnRecalculateAll.Size = new System.Drawing.Size(115, 32);
            this.BtnRecalculateAll.TabIndex = 89;
            this.BtnRecalculateAll.Text = "Recalculate All";
            this.BtnRecalculateAll.Click += new System.EventHandler(this.BtnRecalculateAll_Click);
            // 
            // BtnHelpCalculatedMenu
            // 
            this.BtnHelpCalculatedMenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnHelpCalculatedMenu.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.BtnHelpCalculatedMenu.Location = new System.Drawing.Point(17, 360);
            this.BtnHelpCalculatedMenu.Name = "BtnHelpCalculatedMenu";
            this.BtnHelpCalculatedMenu.Size = new System.Drawing.Size(64, 32);
            this.BtnHelpCalculatedMenu.TabIndex = 88;
            this.BtnHelpCalculatedMenu.Text = "Help";
            this.BtnHelpCalculatedMenu.Click += new System.EventHandler(this.BtnHelpCalculatedMenu_Click);
            // 
            // btnNewEcon
            // 
            this.btnNewEcon.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewEcon.Location = new System.Drawing.Point(207, 360);
            this.btnNewEcon.Name = "btnNewEcon";
            this.btnNewEcon.Size = new System.Drawing.Size(148, 32);
            this.btnNewEcon.TabIndex = 14;
            this.btnNewEcon.Text = "New Econ Variable";
            this.btnNewEcon.Click += new System.EventHandler(this.btnNewEcon_Click);
            // 
            // btnCancelSummary
            // 
            this.btnCancelSummary.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancelSummary.Location = new System.Drawing.Point(608, 360);
            this.btnCancelSummary.Name = "btnCancelSummary";
            this.btnCancelSummary.Size = new System.Drawing.Size(90, 32);
            this.btnCancelSummary.TabIndex = 13;
            this.btnCancelSummary.Text = "Cancel";
            this.btnCancelSummary.Click += new System.EventHandler(this.btnCancelSummary_Click);
            // 
            // btnProperties
            // 
            this.btnProperties.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProperties.Location = new System.Drawing.Point(502, 360);
            this.btnProperties.Name = "btnProperties";
            this.btnProperties.Size = new System.Drawing.Size(100, 32);
            this.btnProperties.TabIndex = 12;
            this.btnProperties.Text = "Properties";
            this.btnProperties.Click += new System.EventHandler(this.btnProperties_Click);
            // 
            // btnNewFvs
            // 
            this.btnNewFvs.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewFvs.Location = new System.Drawing.Point(361, 360);
            this.btnNewFvs.Name = "btnNewFvs";
            this.btnNewFvs.Size = new System.Drawing.Size(135, 32);
            this.btnNewFvs.TabIndex = 4;
            this.btnNewFvs.Text = "New FVS Variable";
            this.btnNewFvs.Click += new System.EventHandler(this.btnNewFvs_Click);
            // 
            // lstVariables
            // 
            this.lstVariables.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.vName,
            this.vDescription,
            this.vType,
            this.vId,
            this.vBaselineRxPkg,
            this.vVariableSource});
            this.lstVariables.FullRowSelect = true;
            this.lstVariables.GridLines = true;
            this.lstVariables.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstVariables.HideSelection = false;
            this.lstVariables.Location = new System.Drawing.Point(18, 18);
            this.lstVariables.MultiSelect = false;
            this.lstVariables.Name = "lstVariables";
            this.lstVariables.Size = new System.Drawing.Size(654, 336);
            this.lstVariables.TabIndex = 2;
            this.lstVariables.UseCompatibleStateImageBehavior = false;
            this.lstVariables.View = System.Windows.Forms.View.Details;
            this.lstVariables.SelectedIndexChanged += new System.EventHandler(this.lstVariables_SelectedIndexChanged);
            this.lstVariables.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstVariables_MouseUp);
            // 
            // vName
            // 
            this.vName.DisplayIndex = 1;
            this.vName.Text = "Variable Name";
            this.vName.Width = 200;
            // 
            // vDescription
            // 
            this.vDescription.DisplayIndex = 2;
            this.vDescription.Text = "Description";
            this.vDescription.Width = 350;
            // 
            // vType
            // 
            this.vType.DisplayIndex = 3;
            this.vType.Text = "Type";
            this.vType.Width = 100;
            // 
            // vId
            // 
            this.vId.DisplayIndex = 0;
            this.vId.Width = 0;
            // 
            // vBaselineRxPkg
            // 
            this.vBaselineRxPkg.Width = 0;
            // 
            // vVariableSource
            // 
            this.vVariableSource.Width = 0;
            // 
            // grpboxDetails
            // 
            this.grpboxDetails.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxDetails.Controls.Add(this.pnlDetails);
            this.grpboxDetails.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxDetails.ForeColor = System.Drawing.Color.Black;
            this.grpboxDetails.Location = new System.Drawing.Point(8, 536);
            this.grpboxDetails.Name = "grpboxDetails";
            this.grpboxDetails.Size = new System.Drawing.Size(856, 472);
            this.grpboxDetails.TabIndex = 32;
            this.grpboxDetails.TabStop = false;
            this.grpboxDetails.Text = "Weighted FVS Variable";
            this.grpboxDetails.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePost_Resize);
            // 
            // pnlDetails
            // 
            this.pnlDetails.AutoScroll = true;
            this.pnlDetails.Controls.Add(this.BtnFvsImport);
            this.pnlDetails.Controls.Add(this.lblFvsVariableName);
            this.pnlDetails.Controls.Add(this.btnFVSVariableValue);
            this.pnlDetails.Controls.Add(this.m_dg);
            this.pnlDetails.Controls.Add(this.btnDeleteFvsVariable);
            this.pnlDetails.Controls.Add(this.txtFvsVariableTotalWeight);
            this.pnlDetails.Controls.Add(this.label5);
            this.pnlDetails.Controls.Add(this.BtnHelp);
            this.pnlDetails.Controls.Add(this.txtFVSVariableDescr);
            this.pnlDetails.Controls.Add(this.label8);
            this.pnlDetails.Controls.Add(this.label7);
            this.pnlDetails.Controls.Add(this.btnFvsCalculate);
            this.pnlDetails.Controls.Add(this.btnFvsDetailsCancel);
            this.pnlDetails.Controls.Add(this.grpBoxFvsBaseline);
            this.pnlDetails.Controls.Add(this.groupBox3);
            this.pnlDetails.Controls.Add(this.groupBox2);
            this.pnlDetails.Controls.Add(this.LblSelectedVariable);
            this.pnlDetails.Controls.Add(this.lblSelectedFVSVariable);
            this.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDetails.Location = new System.Drawing.Point(3, 22);
            this.pnlDetails.Name = "pnlDetails";
            this.pnlDetails.Size = new System.Drawing.Size(850, 447);
            this.pnlDetails.TabIndex = 70;
            // 
            // BtnFvsImport
            // 
            this.BtnFvsImport.Enabled = false;
            this.BtnFvsImport.Location = new System.Drawing.Point(463, 165);
            this.BtnFvsImport.Name = "BtnFvsImport";
            this.BtnFvsImport.Size = new System.Drawing.Size(140, 30);
            this.BtnFvsImport.TabIndex = 94;
            this.BtnFvsImport.Text = "Import weights";
            this.BtnFvsImport.Click += new System.EventHandler(this.BtnFvsImport_Click);
            // 
            // lblFvsVariableName
            // 
            this.lblFvsVariableName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFvsVariableName.Location = new System.Drawing.Point(206, 361);
            this.lblFvsVariableName.Name = "lblFvsVariableName";
            this.lblFvsVariableName.Size = new System.Drawing.Size(264, 24);
            this.lblFvsVariableName.TabIndex = 93;
            this.lblFvsVariableName.Text = "Not Defined";
            // 
            // btnFVSVariableValue
            // 
            this.btnFVSVariableValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariableValue.Location = new System.Drawing.Point(646, 28);
            this.btnFVSVariableValue.Name = "btnFVSVariableValue";
            this.btnFVSVariableValue.Size = new System.Drawing.Size(100, 80);
            this.btnFVSVariableValue.TabIndex = 92;
            this.btnFVSVariableValue.Text = "Select";
            this.btnFVSVariableValue.Click += new System.EventHandler(this.btnFVSVariableValue_Click);
            // 
            // m_dg
            // 
            this.m_dg.DataMember = "";
            this.m_dg.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.m_dg.Location = new System.Drawing.Point(18, 165);
            this.m_dg.Name = "m_dg";
            this.m_dg.Size = new System.Drawing.Size(403, 177);
            this.m_dg.TabIndex = 91;
            this.m_dg.CurrentCellChanged += new System.EventHandler(this.m_dg_CurCellChange);
            this.m_dg.Leave += new System.EventHandler(this.m_dg_Leave);
            // 
            // btnDeleteFvsVariable
            // 
            this.btnDeleteFvsVariable.Enabled = false;
            this.btnDeleteFvsVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDeleteFvsVariable.Location = new System.Drawing.Point(564, 402);
            this.btnDeleteFvsVariable.Name = "btnDeleteFvsVariable";
            this.btnDeleteFvsVariable.Size = new System.Drawing.Size(64, 30);
            this.btnDeleteFvsVariable.TabIndex = 11;
            this.btnDeleteFvsVariable.Text = "Delete";
            this.btnDeleteFvsVariable.Click += new System.EventHandler(this.btnDeleteFvsVariable_Click);
            // 
            // txtFvsVariableTotalWeight
            // 
            this.txtFvsVariableTotalWeight.BackColor = System.Drawing.SystemColors.Control;
            this.txtFvsVariableTotalWeight.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFvsVariableTotalWeight.Location = new System.Drawing.Point(463, 297);
            this.txtFvsVariableTotalWeight.Name = "txtFvsVariableTotalWeight";
            this.txtFvsVariableTotalWeight.ReadOnly = true;
            this.txtFvsVariableTotalWeight.Size = new System.Drawing.Size(121, 26);
            this.txtFvsVariableTotalWeight.TabIndex = 90;
            this.txtFvsVariableTotalWeight.Text = "0.0";
            this.txtFvsVariableTotalWeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(459, 270);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(164, 24);
            this.label5.TabIndex = 89;
            this.label5.Text = "TOTAL WEIGHTS";
            // 
            // BtnHelp
            // 
            this.BtnHelp.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.BtnHelp.Location = new System.Drawing.Point(494, 402);
            this.BtnHelp.Name = "BtnHelp";
            this.BtnHelp.Size = new System.Drawing.Size(64, 30);
            this.BtnHelp.TabIndex = 87;
            this.BtnHelp.Text = "Help";
            this.BtnHelp.Click += new System.EventHandler(this.BtnHelp_Click);
            // 
            // txtFVSVariableDescr
            // 
            this.txtFVSVariableDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFVSVariableDescr.Location = new System.Drawing.Point(173, 386);
            this.txtFVSVariableDescr.Multiline = true;
            this.txtFVSVariableDescr.Name = "txtFVSVariableDescr";
            this.txtFVSVariableDescr.Size = new System.Drawing.Size(259, 40);
            this.txtFVSVariableDescr.TabIndex = 86;
            // 
            // label8
            // 
            this.label8.Location = new System.Drawing.Point(13, 389);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(100, 24);
            this.label8.TabIndex = 85;
            this.label8.Text = "Description:";
            // 
            // label7
            // 
            this.label7.Location = new System.Drawing.Point(13, 360);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(212, 24);
            this.label7.TabIndex = 79;
            this.label7.Text = "Weighted variable name:";
            // 
            // btnFvsCalculate
            // 
            this.btnFvsCalculate.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFvsCalculate.Location = new System.Drawing.Point(634, 402);
            this.btnFvsCalculate.Name = "btnFvsCalculate";
            this.btnFvsCalculate.Size = new System.Drawing.Size(76, 30);
            this.btnFvsCalculate.TabIndex = 77;
            this.btnFvsCalculate.Text = "Calculate";
            this.btnFvsCalculate.Click += new System.EventHandler(this.btnFvsCalculate_Click);
            // 
            // btnFvsDetailsCancel
            // 
            this.btnFvsDetailsCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFvsDetailsCancel.Location = new System.Drawing.Point(716, 402);
            this.btnFvsDetailsCancel.Name = "btnFvsDetailsCancel";
            this.btnFvsDetailsCancel.Size = new System.Drawing.Size(64, 30);
            this.btnFvsDetailsCancel.TabIndex = 75;
            this.btnFvsDetailsCancel.Text = "Cancel";
            this.btnFvsDetailsCancel.Click += new System.EventHandler(this.btnFvsDetailsCancel_Click);
            // 
            // grpBoxFvsBaseline
            // 
            this.grpBoxFvsBaseline.Controls.Add(this.cboFvsVariableBaselinePkg);
            this.grpBoxFvsBaseline.Location = new System.Drawing.Point(8, 7);
            this.grpBoxFvsBaseline.Name = "grpBoxFvsBaseline";
            this.grpBoxFvsBaseline.Size = new System.Drawing.Size(194, 68);
            this.grpBoxFvsBaseline.TabIndex = 74;
            this.grpBoxFvsBaseline.TabStop = false;
            this.grpBoxFvsBaseline.Text = "Baseline RxPackage";
            // 
            // cboFvsVariableBaselinePkg
            // 
            this.cboFvsVariableBaselinePkg.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboFvsVariableBaselinePkg.Location = new System.Drawing.Point(8, 27);
            this.cboFvsVariableBaselinePkg.Name = "cboFvsVariableBaselinePkg";
            this.cboFvsVariableBaselinePkg.Size = new System.Drawing.Size(72, 28);
            this.cboFvsVariableBaselinePkg.TabIndex = 77;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lstFVSFieldsList);
            this.groupBox3.Location = new System.Drawing.Point(440, 7);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(200, 133);
            this.groupBox3.TabIndex = 72;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "FVS Variable";
            // 
            // lstFVSFieldsList
            // 
            this.lstFVSFieldsList.FormattingEnabled = true;
            this.lstFVSFieldsList.ItemHeight = 20;
            this.lstFVSFieldsList.Location = new System.Drawing.Point(6, 21);
            this.lstFVSFieldsList.Name = "lstFVSFieldsList";
            this.lstFVSFieldsList.Size = new System.Drawing.Size(181, 84);
            this.lstFVSFieldsList.Sorted = true;
            this.lstFVSFieldsList.TabIndex = 70;
            this.lstFVSFieldsList.SelectedIndexChanged += new System.EventHandler(this.lstFVSFieldsList_SelectedIndexChanged);
            this.lstFVSFieldsList.GotFocus += new System.EventHandler(this.lstFVSTables_GotFocus);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lstFVSTablesList);
            this.groupBox2.Location = new System.Drawing.Point(208, 7);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 133);
            this.groupBox2.TabIndex = 71;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "FVS Variable Table";
            // 
            // lstFVSTablesList
            // 
            this.lstFVSTablesList.FormattingEnabled = true;
            this.lstFVSTablesList.ItemHeight = 20;
            this.lstFVSTablesList.Location = new System.Drawing.Point(6, 21);
            this.lstFVSTablesList.Name = "lstFVSTablesList";
            this.lstFVSTablesList.Size = new System.Drawing.Size(181, 84);
            this.lstFVSTablesList.TabIndex = 70;
            this.lstFVSTablesList.SelectedIndexChanged += new System.EventHandler(this.lstFVSTablesList_SelectedIndexChanged);
            this.lstFVSTablesList.GotFocus += new System.EventHandler(this.lstFVSTables_GotFocus);
            // 
            // LblSelectedVariable
            // 
            this.LblSelectedVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblSelectedVariable.Location = new System.Drawing.Point(195, 143);
            this.LblSelectedVariable.Name = "LblSelectedVariable";
            this.LblSelectedVariable.Size = new System.Drawing.Size(264, 24);
            this.LblSelectedVariable.TabIndex = 69;
            this.LblSelectedVariable.Text = "Not Defined";
            // 
            // lblSelectedFVSVariable
            // 
            this.lblSelectedFVSVariable.Location = new System.Drawing.Point(11, 144);
            this.lblSelectedFVSVariable.Name = "lblSelectedFVSVariable";
            this.lblSelectedFVSVariable.Size = new System.Drawing.Size(200, 24);
            this.lblSelectedFVSVariable.TabIndex = 68;
            this.lblSelectedFVSVariable.Text = "Selected FVS Variable:";
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(866, 32);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Calculated Variables";
            // 
            // uc_optimizer_scenario_calculated_variables
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_scenario_calculated_variables";
            this.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.ResumeLayout(false);
            this.grpBoxEconomicVariable.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dgEcon)).EndInit();
            this.groupBox8.ResumeLayout(false);
            this.grpboxSummary.ResumeLayout(false);
            this.pnlSummary.ResumeLayout(false);
            this.grpboxDetails.ResumeLayout(false);
            this.pnlDetails.ResumeLayout(false);
            this.pnlDetails.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.m_dg)).EndInit();
            this.grpBoxFvsBaseline.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }
        #endregion

        protected void loadvalues_sqlite()
        {
            this.m_intError = 0;
            this.m_strError = "";

            if (System.IO.File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: Optimizer Calculated Variables Log "
                    + System.DateTime.Now.ToString() + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Form name: " + this.Name + "\r\n\r\n");
            }

            // Check for SQLite configuration database 
            if (!System.IO.File.Exists(this.m_strCalculatedVariablesDb))
            {
                this.migrate_access_data();
            }

            SQLiteConnect();
            // One and only one transaction for this form
            m_oDataMgr.m_Transaction = m_oDataMgr.m_Connection.BeginTransaction();

            this.loadLstVariablesSqlite();

            if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, m_strFvsViewTableName))
            {
                m_oDataMgr.m_strSQL = "DROP TABLE " + m_strFvsViewTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + "\r\n");
                }
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
            }
            frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFvsVariableWeightsReferenceTable(m_oDataMgr,
                m_oDataMgr.m_Connection, m_strFvsViewTableName);
            init_sqlite_m_dg();

            //load datagrid for economic variables
            loadEconVariablesGridSqlite();

            // load listbox for economic variables
            lstEconVariablesList.Items.Clear();
            foreach (string strName in PREFIX_ECON_NAME_ARRAY)
            {
                lstEconVariablesList.Items.Add(strName);
            }

            //@ToDo: Rework this when FVS out is in Sqlite
            if (m_oAdo == null)
            {
                m_oAdo = new ado_data_access();
            }
            m_dictFVSTables = m_oOptimizerScenarioTools.LoadFvsTablesAndVariables(m_oAdo);

            foreach (string strKey in m_dictFVSTables.Keys)
            {
                // 
                if (strKey.IndexOf("_WEIGHTED") < 0)
                {
                    lstFVSTablesList.Items.Add(strKey);
                }
            }

            // Enable the refresh button if we have calculated weighted variables
            //@ToDo: Update when weighted variable output moves to sqlite
            string strPrePostWeightedAccdb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            if (System.IO.File.Exists(strPrePostWeightedAccdb))
            {
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strPrePostWeightedAccdb, "", ""));
                string[] arrTableNames = m_oAdo.getTableNames(m_oAdo.m_OleDbConnection);
                if (arrTableNames.Length > 0)
                    BtnRecalculateAll.Enabled = true;
            }

            //@ToDo: Using MS Access to query FVS Out tables
            if (String.IsNullOrEmpty(m_strTempMDB))
            {
                m_bUsingSqlite = true;
            }

            if (m_oAdo != null)
                m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
        }

        protected void loadvalues()
        {
            this.m_intError = 0;
            this.m_strError = "";

            if (System.IO.File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: Optimizer Calculated Variables Log "
                    + System.DateTime.Now.ToString() + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Form name: " + this.Name + "\r\n\r\n");
            }

            this.loadLstVariables();

            //create and set temporary mdb file
            string strDestinationLinkDir = this.m_oEnv.strTempDir;
            //get temporary mdb file
            m_strTempMDB = m_oUtils.getRandomFile(strDestinationLinkDir, "accdb");

            dao_data_access oDao = new dao_data_access();
            oDao.CreateMDB(m_strTempMDB);
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }

            m_oAdoFvs = new ado_data_access();
            m_oAdoFvs.OpenConnection(m_oAdoFvs.getMDBConnString(m_strTempMDB, "", ""));
            frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFvsVariableWeightsReferenceTable(m_oAdoFvs,
                    m_oAdoFvs.m_OleDbConnection, m_strFvsViewTableName);
            init_m_dg();

            //load datagrid for economic variables
            loadEconVariablesGrid();

            // load listbox for economic variables
            lstEconVariablesList.Items.Clear();
            foreach (string strName in PREFIX_ECON_NAME_ARRAY)
            {
                lstEconVariablesList.Items.Add(strName);
            }

            m_dictFVSTables = m_oOptimizerScenarioTools.LoadFvsTablesAndVariables(m_oAdo);
            foreach (string strKey in m_dictFVSTables.Keys)
            {
                // 
                if (strKey.IndexOf("_WEIGHTED") < 0)
                {
                    lstFVSTablesList.Items.Add(strKey);
                }
            }

            // Enable the refresh button if we have calculated weighted variables
            string strPrePostWeightedAccdb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            if (System.IO.File.Exists(strPrePostWeightedAccdb))
            {
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strPrePostWeightedAccdb, "", ""));
                string[] arrTableNames = m_oAdo.getTableNames(m_oAdo.m_OleDbConnection);
                if (arrTableNames.Length > 0)
                    BtnRecalculateAll.Enabled = true;
            }

            if (m_oAdo != null)
                m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

        }

        private void loadm_dg()
        {
            dao_data_access oDao = new dao_data_access();

            //@ToDo: Using MS Access to query FVS Out tables
            if (m_bUsingSqlite)
            {
                m_oAdoFvs = new ado_data_access();
                m_strTempMDB = m_oUtils.getRandomFile(this.m_oEnv.strTempDir, "accdb");
                oDao.CreateMDB(m_strTempMDB);
            }              

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadm_dg: Starting to load FVS calculated variable datagrid \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Temporary database path: " + m_strTempMDB + "\r\n\r\n");
            }

            string strFvsPreTableName = "";
            string strFvsPostTableName = "";
            string strFvsPrePostDb = "";
            if (this.lstFVSTablesList.SelectedItems.Count > 0)
            {
                strFvsPreTableName = "PRE_" + Convert.ToString(lstFVSTablesList.SelectedItem);
                strFvsPostTableName = "POST_" + Convert.ToString(lstFVSTablesList.SelectedItem);
                string strSourceDatabaseName = "PREPOST_" + Convert.ToString(lstFVSTablesList.SelectedItem) + ".ACCDB";
                strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                    "\\fvs\\db\\" + strSourceDatabaseName;
            }
            if (! String.IsNullOrEmpty(strFvsPreTableName))
            {
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                    frmMain.g_oFrmMain.WindowState,
                    frmMain.g_oFrmMain.Left,
                    frmMain.g_oFrmMain.Height,
                    frmMain.g_oFrmMain.Width,
                    frmMain.g_oFrmMain.Top);
                
                //Add links to FVS pre/post tables if they don't exist
                if (!oDao.TableExists(m_strTempMDB, strFvsPreTableName))
                {
                    oDao.CreateTableLink(m_strTempMDB, strFvsPreTableName, strFvsPrePostDb, strFvsPreTableName);
                }
                if (!oDao.TableExists(m_strTempMDB, strFvsPostTableName))
                {
                    oDao.CreateTableLink(m_strTempMDB, strFvsPostTableName, strFvsPrePostDb, strFvsPostTableName);
                }

                //Sleep while table link is being built
                using (var calculateConn = new OleDbConnection(m_oAdoFvs.getMDBConnString(m_strTempMDB, "", "")))
                {
                    calculateConn.Open();
                    while (!m_oAdoFvs.TableExist(calculateConn, strFvsPreTableName))
                    {
                        Thread.Sleep(1000);
                    }
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

                    //Populate baseline prescription list
                    cboFvsVariableBaselinePkg.Items.Clear();
                    string strSql = "SELECT distinct rxpackage" +
                                    " FROM " + strFvsPreTableName;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Query rxPackage for baseline package list \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql + "\r\n\r\n");
                    }
                    m_oAdoFvs.SqlQueryReader(calculateConn, strSql);
                    while (m_oAdoFvs.m_OleDbDataReader.Read())
                    {
                        if (m_oAdoFvs.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                        {
                            cboFvsVariableBaselinePkg.Items.Add(Convert.ToString(m_oAdoFvs.m_OleDbDataReader["rxpackage"]));
                        }
                    }

                    //Create temporary table to populate datagrid
                    if (! m_bUsingSqlite)
                    {
                        if (oDao.TableExists(m_strTempMDB, m_strFvsViewTableName))
                        {
                            oDao.DeleteTableFromMDB(m_strTempMDB, m_strFvsViewTableName);
                        }
                        frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFvsVariableWeightsReferenceTable(m_oAdoFvs,
                            calculateConn, m_strFvsViewTableName);
                    }
                    else
                    {
                        if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, m_strFvsViewTableName))
                        {
                            m_oDataMgr.m_strSQL = "DROP TABLE " + m_strFvsViewTableName;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + "\r\n");
                            }
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                        }
                        frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFvsVariableWeightsReferenceTable(m_oDataMgr,
                           m_oDataMgr.m_Connection, m_strFvsViewTableName);
                    }

                    strSql = "SELECT rxcycle, MIN(Year) as MinYear, \"PRE\" as pre_or_post" +
                                     " FROM " + strFvsPreTableName +
                                     " GROUP BY rxcycle, Year";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Query for year, and rxcycle from " + strFvsPreTableName + " \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql + "\r\n\r\n");
                    }

                    m_oAdoFvs.SqlQueryReader(calculateConn, strSql);
                    while (m_oAdoFvs.m_OleDbDataReader.Read())
                    {
                        string strRxCycle = "";
                        int intYear = -99;
                        if (m_oAdoFvs.m_OleDbDataReader["MinYear"] != System.DBNull.Value)
                        {
                            intYear = Convert.ToInt16(m_oAdoFvs.m_OleDbDataReader["MinYear"]);
                        }
                        if (m_oAdoFvs.m_OleDbDataReader["rxcycle"] != System.DBNull.Value)
                        {
                            strRxCycle = Convert.ToString(m_oAdoFvs.m_OleDbDataReader["rxcycle"]);
                        }

                        if (!String.IsNullOrEmpty(strRxCycle))
                        {
                            string insertSql = "INSERT INTO " + m_strFvsViewTableName +
                                               " VALUES('" + strRxCycle + "','PRE'," +
                                               intYear + ",0)";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Insert records into " + m_strFvsViewTableName + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, insertSql + "\r\n\r\n");
                            }
                            if (!m_bUsingSqlite)
                            {
                                m_oAdoFvs.SqlNonQuery(calculateConn, insertSql);
                            }
                            else
                            {
                                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, insertSql);
                            }
                        }
                    }

                    strSql = "SELECT rxcycle, MIN(Year) as MinYear, \"PRE\" as pre_or_post" +
                     " FROM " + strFvsPostTableName +
                     " GROUP BY rxcycle, Year";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Query for year, and rxcycle from " + strFvsPreTableName + " \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql + "\r\n\r\n");
                    }

                    m_oAdoFvs.SqlQueryReader(calculateConn, strSql);
                    while (m_oAdoFvs.m_OleDbDataReader.Read())
                    {
                        string strRxCycle = "";
                        int intYear = -99;
                        if (m_oAdoFvs.m_OleDbDataReader["MinYear"] != System.DBNull.Value)
                        {
                            intYear = Convert.ToInt16(m_oAdoFvs.m_OleDbDataReader["MinYear"]);
                        }
                        if (m_oAdoFvs.m_OleDbDataReader["rxcycle"] != System.DBNull.Value)
                        {
                            strRxCycle = Convert.ToString(m_oAdoFvs.m_OleDbDataReader["rxcycle"]);
                        }

                        if (!String.IsNullOrEmpty(strRxCycle))
                        {
                            string insertSql = "INSERT INTO " + m_strFvsViewTableName +
                                               " VALUES('" + strRxCycle + "','POST'," +
                                               intYear + ",0)";

                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Insert records into " + m_strFvsViewTableName + "\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, insertSql + "\r\n\r\n");
                            }
                            if (!m_bUsingSqlite)
                            {
                                m_oAdoFvs.SqlNonQuery(calculateConn, insertSql);
                            }
                            else
                            {
                                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, insertSql);
                            }
                        }
                    }
                    if (!m_bUsingSqlite)
                    {
                        m_oAdoFvs.m_DataSet.Clear();
                        m_oAdoFvs.m_strSQL = "select * from " + m_strFvsViewTableName + " order by RXYEAR";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Populating FVS data set with query " + m_oAdoFvs.m_strSQL + " \r\n");
                        }
                        m_oAdoFvs.m_OleDbCommand = calculateConn.CreateCommand();
                        m_oAdoFvs.m_OleDbCommand.CommandText = m_oAdoFvs.m_strSQL;
                        m_oAdoFvs.m_OleDbDataAdapter = new OleDbDataAdapter();
                        m_oAdoFvs.m_OleDbDataAdapter.SelectCommand = m_oAdoFvs.m_OleDbCommand;
                        m_oAdoFvs.m_OleDbDataAdapter.SelectCommand.Transaction = m_oAdoFvs.m_OleDbTransaction;
                        m_oAdoFvs.m_OleDbDataAdapter.Fill(m_oAdoFvs.m_DataSet, m_strFvsViewTableName);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "FVS dataset filled \r\n");
                        }
                    }
                    else
                    {
                        string strPrimaryKeys = "rxcycle, pre_or_post, rxyear";
                        string strColumns = "rxcycle, pre_or_post, rxyear, weight";
                        m_oDataMgr.DataAdapterArrayItemConfigureSelectCommand(FVS_DETAILS_TABLE, m_strFvsViewTableName, strColumns, strPrimaryKeys, "");
                    }
                }
                BtnFvsImport.Enabled = true;
            }
            else
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "loadm_dg: !!Unable to locate any sequence number tables\r\n\r\n");
                }
                btnNewFvs.Enabled = false;
                MessageBox.Show("!!FVS Pre/Post Tables Are Missing. FVS Weighted Variable Settings Disabled!!", "FIA Biosum",
                 System.Windows.Forms.MessageBoxButtons.OK,
                 System.Windows.Forms.MessageBoxIcon.Exclamation);
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        private void init_m_dg()
        {

            m_oAdoFvs.m_DataSet = new DataSet(m_strFvsViewTableName);
            m_oAdoFvs.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
            m_oAdoFvs.m_strSQL = "select * from " + m_strFvsViewTableName + " order by RXYEAR";
            this.m_dtTableSchema = m_oAdoFvs.getTableSchema(m_oAdoFvs.m_OleDbConnection,
                                                       m_oAdoFvs.m_OleDbTransaction,
                                                       m_oAdoFvs.m_strSQL);
            if (m_oAdoFvs.m_intError == 0)
            {
                m_oAdoFvs.m_OleDbCommand = m_oAdoFvs.m_OleDbConnection.CreateCommand();
                m_oAdoFvs.m_OleDbCommand.CommandText = m_oAdoFvs.m_strSQL;
                m_oAdoFvs.m_OleDbDataAdapter.SelectCommand = m_oAdoFvs.m_OleDbCommand;
                m_oAdoFvs.m_OleDbDataAdapter.SelectCommand.Transaction = m_oAdoFvs.m_OleDbTransaction;
            }

            try
            {

                m_oAdoFvs.m_OleDbDataAdapter.Fill(m_oAdoFvs.m_DataSet, "view_weights");
                this.m_dv = new DataView(m_oAdoFvs.m_DataSet.Tables["view_weights"]);

                this.m_dv.AllowNew = false;       //cannot append new records
                this.m_dv.AllowDelete = false;    //cannot delete records
                this.m_dv.AllowEdit = true;
                this.m_dg.CaptionText = "view_weights";
                m_dg.BackgroundColor = frmMain.g_oGridViewBackgroundColor;
                /***********************************************************************************
                 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
                 ***********************************************************************************/
                WeightedAverage_DataGridColoredTextBoxColumn aColumnTextColumn;


                /***************************************************************
                 **custom define the grid style
                 ***************************************************************/
                DataGridTableStyle tableStyle = new DataGridTableStyle();

                /***********************************************************************
                 **map the data grid table style to the scenario rx intensity dataset
                 ***********************************************************************/
                tableStyle.MappingName = "view_weights";
                tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
                tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
                tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
                tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;



                /******************************************************************************
                 **since the dataset has things like field name and number of columns,
                 **we will use those to create new columnstyles for the columns in our grid
                 ******************************************************************************/
                //get the number of columns from the view_weights data set
                int numCols = m_oAdoFvs.m_DataSet.Tables["view_weights"].Columns.Count;

                /************************************************
                 **loop through all the columns in the dataset	
                 ************************************************/
                string strColumnName;
                for (int i = 0; i < numCols; ++i)
                {
                    strColumnName = m_oAdoFvs.m_DataSet.Tables["view_weights"].Columns[i].ColumnName;

                    /***********************************
                    **all columns are read-only except weight
                    ***********************************/
                    if (strColumnName.Trim().ToUpper() == "WEIGHT")
                    {
                        /******************************************************************
                        **create a new instance of the DataGridColoredTextBoxColumn class
                        ******************************************************************/
                        aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(true, true, this);
                        aColumnTextColumn.Format = "#0.000";
                        aColumnTextColumn.ReadOnly = false;
                    }
                    else
                    {
                        /******************************************************************
                        **create a new instance of the DataGridColoredTextBoxColumn class
                        ******************************************************************/
                        aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(false, false, this);
                        aColumnTextColumn.ReadOnly = true;
                    }
                    aColumnTextColumn.HeaderText = strColumnName;

                    /********************************************************************
                     **assign the mappingname property the data sets column name
                     ********************************************************************/
                    aColumnTextColumn.MappingName = strColumnName;

                    /********************************************************************
                     **add the datagridcoloredtextboxcolumn object to the data grid 
                     **table style object
                     ********************************************************************/
                    tableStyle.GridColumnStyles.Add(aColumnTextColumn);

                    /**********************************
                     * Hide pre_or_post column
                     * *******************************
                     * if (strColumnName.Equals("pre_or_post"))
                     *   tableStyle.GridColumnStyles.Remove(aColumnTextColumn); */
                }
                /*********************************************************************
                 ** make the dataGrid use our new tablestyle and bind it to our table
                 *********************************************************************/
                if (frmMain.g_oGridViewFont != null) this.m_dg.Font = frmMain.g_oGridViewFont;
                this.m_dg.TableStyles.Clear();
                this.m_dg.TableStyles.Add(tableStyle);
                //this.m_dg.CaptionText = strCaption;
                this.m_dg.DataSource = this.m_dv;
                this.m_dg.Expand(-1);
                //sum up the weights after the grid loads
                this.SumWeights(false);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "view_weights Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.m_intError = -1;
                m_oAdoFvs.m_OleDbConnection.Close();
                m_oAdoFvs.m_OleDbConnection = null;
                m_oAdoFvs.m_DataSet.Clear();
                m_oAdoFvs.m_DataSet = null;
                m_oAdoFvs.m_OleDbDataAdapter.Dispose();
                m_oAdoFvs.m_OleDbDataAdapter = null;
                return;

            }
        }

        private void init_sqlite_m_dg()
        {
            try
            {
                string strPrimaryKeys = "rxcycle, pre_or_post, rxyear";
                string strColumns = "rxcycle, pre_or_post, rxyear, weight";
                string strWhereClause = "";
                m_oDataMgr.InitializeDataAdapterArrayItem(FVS_DETAILS_TABLE, m_strFvsViewTableName, strColumns, strPrimaryKeys, strWhereClause);
                this.m_dtTableSchema = m_oDataMgr.getTableSchema(m_oDataMgr.m_Connection,
                                             m_oDataMgr.m_Transaction,
                                             m_oDataMgr.m_strSQL);

                this.m_dv = new DataView(m_oDataMgr.m_DataSet.Tables[m_strFvsViewTableName]);

                this.m_dv.AllowNew = false;       //cannot append new records
                this.m_dv.AllowDelete = false;    //cannot delete records
                this.m_dv.AllowEdit = true;
                this.m_dv.Sort = "RXCYCLE ASC";
                this.m_dg.CaptionText = "view_weights";
                m_dg.BackgroundColor = frmMain.g_oGridViewBackgroundColor;
                /***********************************************************************************
                 **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
                 ***********************************************************************************/
                WeightedAverage_DataGridColoredTextBoxColumn aColumnTextColumn;


                /***************************************************************
                 **custom define the grid style
                 ***************************************************************/
                DataGridTableStyle tableStyle = new DataGridTableStyle();

                /***********************************************************************
                 **map the data grid table style to the scenario rx intensity dataset
                 ***********************************************************************/
                tableStyle.MappingName = m_strFvsViewTableName;
                tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
                tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
                tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
                tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;



                /******************************************************************************
                 **since the dataset has things like field name and number of columns,
                 **we will use those to create new columnstyles for the columns in our grid
                 ******************************************************************************/
                //get the number of columns from the view_weights data set
                int numCols = m_oDataMgr.m_DataSet.Tables[m_strFvsViewTableName].Columns.Count;

                /************************************************
                 **loop through all the columns in the dataset	
                 ************************************************/
                string strColumnName;
                for (int i = 0; i < numCols; ++i)
                {
                    strColumnName = m_oDataMgr.m_DataSet.Tables[m_strFvsViewTableName].Columns[i].ColumnName;

                    /***********************************
                    **all columns are read-only except weight
                    ***********************************/
                    if (strColumnName.Trim().ToUpper() == "WEIGHT")
                    {
                        /******************************************************************
                        **create a new instance of the DataGridColoredTextBoxColumn class
                        ******************************************************************/
                        aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(true, true, this);
                        aColumnTextColumn.Format = "#0.000";
                        aColumnTextColumn.ReadOnly = false;
                    }
                    else
                    {
                        /******************************************************************
                        **create a new instance of the DataGridColoredTextBoxColumn class
                        ******************************************************************/
                        aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(false, false, this);
                        aColumnTextColumn.ReadOnly = true;
                    }
                    aColumnTextColumn.HeaderText = strColumnName;

                    /********************************************************************
                     **assign the mappingname property the data sets column name
                     ********************************************************************/
                    aColumnTextColumn.MappingName = strColumnName;

                    /********************************************************************
                     **add the datagridcoloredtextboxcolumn object to the data grid 
                     **table style object
                     ********************************************************************/
                    tableStyle.GridColumnStyles.Add(aColumnTextColumn);

                    /**********************************
                     * Hide pre_or_post column
                     * *******************************
                     * if (strColumnName.Equals("pre_or_post"))
                     *   tableStyle.GridColumnStyles.Remove(aColumnTextColumn); */
                }
                /*********************************************************************
                 ** make the dataGrid use our new tablestyle and bind it to our table
                 *********************************************************************/
                if (frmMain.g_oGridViewFont != null) this.m_dg.Font = frmMain.g_oGridViewFont;
                this.m_dg.TableStyles.Clear();
                this.m_dg.TableStyles.Add(tableStyle);
                //this.m_dg.CaptionText = strCaption;
                this.m_dg.DataSource = this.m_dv;
                this.m_dg.Expand(-1);
                //sum up the weights after the grid loads
                this.SumWeights(false);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "view_weights Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.m_intError = -1;
                return;
            }
        }

        private void loadLstVariables()
        {
            //Loading the first (main) groupbox
            //Only instantiate the m_oAdo if it is null so we don't wipe everything out in subsequent refreshes
            if (m_oAdo == null)
            {
                m_oAdo = new ado_data_access();
            }
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strCalculatedVariablesAccdb, "", ""));
            m_oAdo.m_strSQL = "SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                " ORDER BY VARIABLE_NAME";
            m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            lstVariables.Items.Clear();
            if (m_oAdo.m_intError == 0)
            {
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    this.m_oLvAlternateColors.InitializeRowCollection();
                    int idxItems = 0;
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        lstVariables.Items.Add(m_oAdo.m_OleDbDataReader["VARIABLE_NAME"].ToString().Trim());
                        lstVariables.Items[idxItems].UseItemStyleForSubItems = false;
                        lstVariables.Items[idxItems].SubItems.Add(m_oAdo.m_OleDbDataReader["VARIABLE_DESCRIPTION"].ToString().Trim());
                        lstVariables.Items[idxItems].SubItems.Add(m_oAdo.m_OleDbDataReader["VARIABLE_TYPE"].ToString().Trim());
                        lstVariables.Items[idxItems].SubItems.Add(m_oAdo.m_OleDbDataReader["ID"].ToString().Trim());
                        string strBaselineRxPkg = "";
                        if (m_oAdo.m_OleDbDataReader["BASELINE_RXPACKAGE"] != System.DBNull.Value)
                        {
                            strBaselineRxPkg = m_oAdo.m_OleDbDataReader["BASELINE_RXPACKAGE"].ToString().Trim();
                        }
                        lstVariables.Items[idxItems].SubItems.Add(strBaselineRxPkg);
                        lstVariables.Items[idxItems].SubItems.Add(m_oAdo.m_OleDbDataReader["VARIABLE_SOURCE"].ToString().Trim());

                        m_oLvAlternateColors.AddRow();
                        m_oLvAlternateColors.AddColumns(idxItems, lstVariables.Columns.Count);
                        idxItems++;
                    }
                    this.m_oLvAlternateColors.ListView();
                }
            }
        }

        private void loadLstVariablesSqlite()
        {            
            //Loading the first (main) groupbox
            //Only instantiate the m_oAdo if it is null so we don't wipe everything out in subsequent refreshes
            DataMgr oDataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(this.m_strCalculatedVariablesDb)))
            {
                conn.Open();
                oDataMgr.m_strSQL = "SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                                    " ORDER BY VARIABLE_NAME COLLATE NOCASE ASC";
                oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                lstVariables.Items.Clear();
                if (oDataMgr.m_intError == 0 && oDataMgr.m_DataReader.HasRows)
                {
                    this.m_oLvAlternateColors.InitializeRowCollection();
                    int idxItems = 0;
                    while (oDataMgr.m_DataReader.Read())
                    {
                        lstVariables.Items.Add(oDataMgr.m_DataReader["VARIABLE_NAME"].ToString().Trim());
                        lstVariables.Items[idxItems].UseItemStyleForSubItems = false;
                        lstVariables.Items[idxItems].SubItems.Add(oDataMgr.m_DataReader["VARIABLE_DESCRIPTION"].ToString().Trim());
                        lstVariables.Items[idxItems].SubItems.Add(oDataMgr.m_DataReader["VARIABLE_TYPE"].ToString().Trim());
                        lstVariables.Items[idxItems].SubItems.Add(oDataMgr.m_DataReader["ID"].ToString().Trim());
                        string strBaselineRxPkg = "";
                        if (oDataMgr.m_DataReader["BASELINE_RXPACKAGE"] != DBNull.Value)
                        {
                            strBaselineRxPkg = oDataMgr.m_DataReader["BASELINE_RXPACKAGE"].ToString().Trim();
                        }
                        lstVariables.Items[idxItems].SubItems.Add(strBaselineRxPkg);
                        lstVariables.Items[idxItems].SubItems.Add(oDataMgr.m_DataReader["VARIABLE_SOURCE"].ToString().Trim());

                        m_oLvAlternateColors.AddRow();
                        m_oLvAlternateColors.AddColumns(idxItems, lstVariables.Columns.Count);
                        idxItems++;
                    }
                    this.m_oLvAlternateColors.ListView();
                }
                oDataMgr.m_DataReader.Close();
            }
        }

        public int savevalues(string strVariableType)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "savevalues BEGIN \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------------------------------------------------------\r\n");
            }

            ado_data_access oAdo = new ado_data_access();
            string strScenarioMDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB, "", ""));
            if (oAdo.m_intError == 0)
            {
                int intId = -1;
                string strSql = "";
                string strBaselinePackage = "";
                // DELETE EXISTING RECORD ON SHARED DEFINITIONS TABLE
                if (m_intCurVar > 0)
                {
                    oAdo.m_strSQL = "DELETE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                                    " WHERE ID = " + m_intCurVar;
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName;
                    if (strVariableType.Equals("ECON"))
                    {
                        strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                    }
                    oAdo.m_strSQL = "DELETE FROM " + strTableName +
                        " WHERE calculated_variables_id = " + m_intCurVar;
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    intId = m_intCurVar;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Deleted existing records for variable id: " + m_intCurVar + "\r\n\r\n");
                    }

                }
                else
                {
                    if (strVariableType.Equals("ECON"))
                    {
                        // We already calculated the next id to add it to the grid
                        DataRow oRow = this.m_econ_dv.Table.Rows[0];
                        intId = Convert.ToInt32(oRow["calculated_variables_id"]);
                    }
                    else
                    {
                        intId = GetNextId();
                    }
                }

                // SHARED BEGINNING OF INSERT STATEMENT
                strSql = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                    " (ID, VARIABLE_NAME, VARIABLE_DESCRIPTION, VARIABLE_TYPE, BASELINE_RXPACKAGE, VARIABLE_SOURCE)" +
                    " VALUES ( " + intId + ", '";

                if (strVariableType.Equals("FVS"))
                {
                    if (cboFvsVariableBaselinePkg.SelectedIndex > -1)
                    {
                        strBaselinePackage = cboFvsVariableBaselinePkg.SelectedItem.ToString();
                    }
                    string strDescription = "";
                    if (!String.IsNullOrEmpty(txtFVSVariableDescr.Text))
                        strDescription = txtFVSVariableDescr.Text.Trim();
                    strSql = strSql + lblFvsVariableName.Text.Trim() + "','" + strDescription + "','" +
                             strVariableType + "','" + strBaselinePackage + "','" + LblSelectedVariable.Text.Trim() + "')";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Add parent record for FVS weighted variable \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                    }
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                    // ADD CHILD PERCENTAGE RECORD
                    if (oAdo.m_intError == 0)
                    {
                        double[] arrPrePercents = new double[4];
                        double[] arrPostPercents = new double[4];
                        int intRxCycle;
                        double dblWeight;
                        string strPrePost = "";
                        foreach (DataRow row in this.m_dv.Table.Rows)
                        {
                            intRxCycle = Convert.ToInt32(row["rxcycle"]);
                            dblWeight = Convert.ToDouble(row["weight"]);
                            strPrePost = row["pre_or_post"].ToString().Trim();
                            if (strPrePost.Equals("PRE"))
                            {
                                arrPrePercents[intRxCycle - 1] = dblWeight;
                            }
                            else
                            {
                                arrPostPercents[intRxCycle - 1] = dblWeight;
                            }
                        }

                        strSql = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName +
                            " (calculated_variables_id, weight_1_pre, weight_1_post, weight_2_pre, weight_2_post, " +
                            "weight_3_pre, weight_3_post, weight_4_pre, weight_4_post)" +
                            " VALUES ( " + intId + ", " + arrPrePercents[0] + ", " + arrPostPercents[0] +
                            ", " + arrPrePercents[1] + ", " + arrPostPercents[1] + ", " + arrPrePercents[2] +
                            ", " + arrPostPercents[2] + ", " + arrPrePercents[3] + ", " + arrPostPercents[3] + ")";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Add child weight values entry for FVS variable id: " + intId + " \r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                        }
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                    }
                }
                else
                {
                    string strDescription = "";
                    if (!String.IsNullOrEmpty(txtEconVariableDescr.Text))
                        strDescription = txtEconVariableDescr.Text.Trim();
                    string strVariableSource = Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName +
                        "." + lblEconVariableName.Text.Trim();
                    strSql = strSql + lblEconVariableName.Text.Trim() + "','" + strDescription + "','" +
                             strVariableType + "','" + strBaselinePackage + "','" + strVariableSource + "')";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Add parent record for Economic weighted variable \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                    }
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSql);
                    // MODIFY CHILD PERCENTAGE RECORD
                    if (oAdo.m_intError == 0 && this.m_oAdo.m_intError == 0)
                    {
                        this.m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strScenarioMDB, "", ""));
                        int intCurrRow;
                        this.m_intError = 0;

                        /******************************************************
                         **save the current row, move the current row to a
                         **different row to enable getchanges() method, then
                         **move back to current row
                         ******************************************************/
                        intCurrRow = this.m_dgEcon.CurrentRowIndex;
                        if (intCurrRow == 0)
                        {
                            this.m_dgEcon.CurrentRowIndex++;
                        }
                        else
                        {
                            this.m_dgEcon.CurrentRowIndex = 0;
                        }


                        System.Data.DataTable p_dtChanges;

                        try
                        {

                            p_dtChanges = this.m_oAdo.m_DataSet.Tables["econ_variable"].GetChanges();

                            //check if any inserted rows
                            if (p_dtChanges.HasErrors)
                            {
                                this.m_oAdo.m_DataSet.Tables["econ_variable"].RejectChanges();
                                this.m_intError = -1;
                            }
                            else
                            {
                                this.m_oAdo.m_OleDbDataAdapter.Update(this.m_oAdo.m_DataSet.Tables["econ_variable"]);
                                this.m_oAdo.m_OleDbTransaction.Commit();
                                this.m_oAdo.m_DataSet.Tables["econ_variable"].AcceptChanges();
                                this.InitializeOleDbTransactionCommands();
                            }
                        }
                        catch (Exception caught)
                        {
                            this.m_intError = -1;
                            MessageBox.Show(caught.Message);
                            this.m_oAdo.m_DataSet.Tables["econ_variable"].RejectChanges();
                            //rollback the transaction to the original records 
                            this.m_oAdo.m_OleDbTransaction.Rollback();
                        }

                        p_dtChanges = null;
                        this.m_dgEcon.CurrentRowIndex = intCurrRow;
                    }
                }
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "savevalues END \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------------------------------------------------------\r\n");
            }
            return oAdo.m_intError;
        }

        public int savevalues_sqlite(string strVariableType)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "savevalues_sqlite BEGIN \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------------------------------------------------------\r\n");
            }

            int intId = -1;
            string strSql = "";
            string strBaselinePackage = "";

            if (strVariableType.Equals("ECON"))
            {
                // We already calculated the next id to add it to the grid
                intId = Convert.ToInt32(this.m_econ_dv.Table.Rows[0]["calculated_variables_id"]);
            }
            else
            {
                intId = GetNextId();
            }

            // SHARED BEGINNING OF INSERT STATEMENT
            strSql = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " (ID, VARIABLE_NAME, VARIABLE_DESCRIPTION, VARIABLE_TYPE, BASELINE_RXPACKAGE, VARIABLE_SOURCE)" +
                        " VALUES ( " + intId + ", '";

            if (strVariableType.Equals("FVS"))
            {
                if (cboFvsVariableBaselinePkg.SelectedIndex > -1)
                {
                    strBaselinePackage = cboFvsVariableBaselinePkg.SelectedItem.ToString();
                }
                string strDescription = "";
                if (!String.IsNullOrEmpty(txtFVSVariableDescr.Text))
                    strDescription = txtFVSVariableDescr.Text.Trim();
                strSql = strSql + lblFvsVariableName.Text.Trim() + "','" + strDescription + "','" +
                         strVariableType + "','" + strBaselinePackage + "','" + LblSelectedVariable.Text.Trim() + "')";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Add parent record for FVS weighted variable \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                }
                try
                {
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, strSql);
                    // ADD CHILD PERCENTAGE RECORD
                    if (m_oDataMgr.m_intError == 0)
                    {
                        double[] arrPrePercents = new double[4];
                        double[] arrPostPercents = new double[4];
                        int intRxCycle;
                        double dblWeight;
                        string strPrePost = "";
                        foreach (DataRow row in this.m_dv.Table.Rows)
                        {
                            intRxCycle = Convert.ToInt32(row["rxcycle"]);
                            dblWeight = Convert.ToDouble(row["weight"]);
                            strPrePost = row["pre_or_post"].ToString().Trim();
                            if (strPrePost.Equals("PRE"))
                            {
                                arrPrePercents[intRxCycle - 1] = dblWeight;
                            }
                            else
                            {
                                arrPostPercents[intRxCycle - 1] = dblWeight;
                            }
                        }

                        strSql = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName +
                        " (calculated_variables_id, weight_1_pre, weight_1_post, weight_2_pre, weight_2_post, " +
                        "weight_3_pre, weight_3_post, weight_4_pre, weight_4_post)" +
                        " VALUES ( " + intId + ", " + arrPrePercents[0] + ", " + arrPostPercents[0] +
                        ", " + arrPrePercents[1] + ", " + arrPostPercents[1] + ", " + arrPrePercents[2] +
                        ", " + arrPostPercents[2] + ", " + arrPrePercents[3] + ", " + arrPostPercents[3] + ")";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Add child weight values entry for FVS variable id: " + intId + " \r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                        }
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, strSql);
                        m_oDataMgr.m_Transaction.Commit();
                    }
                }
                catch (System.Data.SQLite.SQLiteException errSQLite)
                {
                    m_intError = -1;
                    MessageBox.Show(errSQLite.Message);
                }
                catch (Exception caught)
                {
                    this.m_intError = -1;
                    MessageBox.Show(caught.Message);
                }
                finally
                {
                    m_oDataMgr.ResetTransactionObjectToDataAdapterArray();
                }
            }
            else
            {
                string strDescription = "";
                if (!String.IsNullOrEmpty(txtEconVariableDescr.Text))
                    strDescription = txtEconVariableDescr.Text.Trim();
                string strVariableSource = Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName +
                    "." + lblEconVariableName.Text.Trim();
                strSql = strSql + lblEconVariableName.Text.Trim() + "','" + strDescription + "','" +
                    strVariableType + "','" + strBaselinePackage + "','" + strVariableSource + "')";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Add parent record for Economic weighted variable \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                }
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, strSql);

                // MODIFY CHILD PERCENTAGE RECORD
                if (this.m_oDataMgr.m_intError == 0)
                {
                    int intCurrRow;
                    this.m_intError = 0;

                    /******************************************************
                     **save the current row, move the current row to a
                     **different row to enable getchanges() method, then
                     **move back to current row
                     ******************************************************/
                    intCurrRow = this.m_dgEcon.CurrentRowIndex;
                    if (intCurrRow == 0)
                    {
                        this.m_dgEcon.CurrentRowIndex++;
                    }
                    else
                    {
                        this.m_dgEcon.CurrentRowIndex = 0;
                    }

                    DataTable p_dtChanges;
                    string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                    try
                    {
                        p_dtChanges = m_oDataMgr.m_DataSet.Tables[strTableName].GetChanges();

                        //check if any inserted rows
                        if (p_dtChanges.HasErrors)
                        {
                            m_oDataMgr.m_DataSet.Tables[strTableName].RejectChanges();
                            this.m_intError = -1;
                        }
                        else
                        {
                            m_oDataMgr.m_DataAdapterArray[ECON_DETAILS_TABLE].Update(m_oDataMgr.m_DataSet.Tables[strTableName]);
                            m_oDataMgr.m_Transaction.Commit();
                            m_oDataMgr.m_DataSet.Tables[strTableName].AcceptChanges();
                        }
                        p_dtChanges = null;
                        this.m_dgEcon.CurrentRowIndex = intCurrRow;
                    }
                    catch (System.Data.SQLite.SQLiteException errSQLite)
                    {
                        m_intError = -1;
                        MessageBox.Show(errSQLite.Message);
                    }
                    catch (Exception caught)
                    {
                        this.m_intError = -1;
                        MessageBox.Show(caught.Message);
                    }
                    finally
                    {
                        m_oDataMgr.ResetTransactionObjectToDataAdapterArray();
                    }
                }
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "savevalues_sqlite END \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "------------------------------------------------------------------------------------------------\r\n");
            }
            return m_oDataMgr.m_intError;
        }

        private void val_data_fvs(string strPrePostAccdb, string strPreTable, string strPostTable)
        {
            this.m_intError = 0;    // Reset error variable
            if (this.lblFvsVariableName.Text.Trim().Equals("Not Defined") ||
                this.LblSelectedVariable.Text.Trim().Equals("Not Defined"))
            {
                MessageBox.Show("!!Select An FVS Variable!!", "FIA Biosum",
                                 System.Windows.Forms.MessageBoxButtons.OK,
                                 System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.btnFVSVariableValue.Focus();
                return;
            }
            double dblTotalWeights = -1;
            bool bIsNumber = Double.TryParse(txtFvsVariableTotalWeight.Text, out dblTotalWeights);
            if (dblTotalWeights <= 0)
            {
                MessageBox.Show("!!Select Weights Totaling More Than 0!!", "FIA Biosum",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.m_dg.Focus();
                return;
            }
            if (cboFvsVariableBaselinePkg.SelectedIndex < 0)
            {
                MessageBox.Show("!!No Baseline RxPackage Selected!!", "FIA Biosum",
                               System.Windows.Forms.MessageBoxButtons.OK,
                               System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.cboFvsVariableBaselinePkg.Focus();
                return;
            }
            string strOutputAccdb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" + 
                Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            if (!System.IO.File.Exists(strOutputAccdb))
            {
                MessageBox.Show("!!FVS Weighted Variable output database missing. It should be here: " + 
                    strOutputAccdb + "!!", "FIA Biosum",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }
            string strCalculateConn = m_oAdo.getMDBConnString(frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                                      "\\fvs\\db\\" + strPrePostAccdb, "", "");
            using (var calculateConn = new OleDbConnection(strCalculateConn))
            {
                calculateConn.Open();
                m_oAdo.m_strSQL = "SELECT COUNT(*)" +
                                  " FROM " + strPreTable +
                                  " WHERE " + this.lstFVSFieldsList.SelectedItems[0].ToString() + " < 0";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Checking for negative FVS values \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                int intCount = m_oAdo.getRecordCount(calculateConn, m_oAdo.m_strSQL, strPreTable);
                if (intCount > 0)
                {
                    string strMessage = "!! BioSum found " + intCount + " negative values in the " + strPreTable + " table!!" +
                                        " Do you wish to continue with the weighted variable calculation ?";
                    DialogResult res = MessageBox.Show(strMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo,
                                                       System.Windows.Forms.MessageBoxIcon.Question);
                    if (res != DialogResult.Yes)
                    {
                        this.m_intError = -1;
                        return;
                    }
                }

                m_oAdo.m_strSQL = "SELECT COUNT(*)" +
                  " FROM " + strPostTable +
                  " WHERE " + this.lstFVSFieldsList.SelectedItems[0].ToString() + " < 0";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                intCount = m_oAdo.getRecordCount(calculateConn, m_oAdo.m_strSQL, strPreTable);
                if (intCount > 0)
                {
                    string strMessage = "!! BioSum found " + intCount + " negative values in the " + strPostTable + " table!!" +
                                        " Do you wish to continue with the weighted variable calculation ?";
                    DialogResult res = MessageBox.Show(strMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo,
                                                       System.Windows.Forms.MessageBoxIcon.Question);
                    if (res != DialogResult.Yes)
                    {
                        this.m_intError = -1;
                        return;
                    }
                }
            }
        }

        private void val_data_econ()
        {
            this.m_intError = 0;    // Reset error indicator
            if (this.lblSelectedEconType.Text.Trim().Equals("Not Defined") ||
                this.lblEconVariableName.Text.Trim().Equals("Not Defined"))
            {
                MessageBox.Show("!!Select An Economic Variable!!", "FIA Biosum",
                                 System.Windows.Forms.MessageBoxButtons.OK,
                                 System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.btnEconVariableType.Focus();
                return;
            }
            double dblTotalWeights = -1;
            bool bIsNumber = Double.TryParse(txtEconVariableTotalWeight.Text, out dblTotalWeights);
            if (dblTotalWeights <= 0)
            {
                MessageBox.Show("!!Select Weights Totaling More Than 0!!", "FIA Biosum",
                                System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                this.m_dgEcon.Focus();
                return;
            }
        }

        protected void loadEconVariablesGrid()
        {
            string strCalculatedVariablesACCDB = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(strCalculatedVariablesACCDB, "", ""));
            m_oAdo.m_DataSet = new DataSet("econ_variable");
            m_oAdo.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
            this.InitializeOleDbTransactionCommands();

            m_oAdo.m_strSQL = "SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName +
                " WHERE CALCULATED_VARIABLES_ID = -1";
            this.m_dtTableSchema = m_oAdo.getTableSchema(m_oAdo.m_OleDbConnection,
                                                       m_oAdo.m_OleDbTransaction,
                                                       m_oAdo.m_strSQL);
            if (m_oAdo.m_intError == 0)
            {
                m_oAdo.m_OleDbCommand = m_oAdo.m_OleDbConnection.CreateCommand();
                m_oAdo.m_OleDbCommand.CommandText = m_oAdo.m_strSQL;
                m_oAdo.m_OleDbDataAdapter.SelectCommand = m_oAdo.m_OleDbCommand;
                m_oAdo.m_OleDbDataAdapter.SelectCommand.Transaction = m_oAdo.m_OleDbTransaction;
                try
                {
                    m_oAdo.m_OleDbDataAdapter.Fill(m_oAdo.m_DataSet, "econ_variable");
                    this.m_econ_dv = new DataView(m_oAdo.m_DataSet.Tables["econ_variable"]);

                    this.m_econ_dv.AllowNew = false;       //cannot append new records
                    this.m_econ_dv.AllowDelete = false;    //cannot delete records
                    this.m_econ_dv.AllowEdit = true;
                    this.m_dgEcon.CaptionText = "econ_variable";
                    m_dgEcon.BackgroundColor = frmMain.g_oGridViewBackgroundColor;

                    /***********************************************************************************
                    **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
                    ***********************************************************************************/
                    WeightedAverage_DataGridColoredTextBoxColumn aColumnTextColumn;


                    /***************************************************************
                     **custom define the grid style
                     ***************************************************************/
                    DataGridTableStyle tableStyle = new DataGridTableStyle();

                    /***********************************************************************
                     **map the data grid table style to the scenario rx intensity dataset
                     ***********************************************************************/
                    tableStyle.MappingName = "econ_variable";
                    tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
                    tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
                    tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
                    tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;


                    /******************************************************************************
                     **since the dataset has things like field name and number of columns,
                     **we will use those to create new columnstyles for the columns in our grid
                     ******************************************************************************/
                    //get the number of columns from the view_weights data set
                    int numCols = m_oAdo.m_DataSet.Tables["econ_variable"].Columns.Count;

                    /************************************************
                     **loop through all the columns in the dataset	
                     ************************************************/
                    string strColumnName = ""; ;
                    for (int i = 0; i < numCols; ++i)
                    {
                        strColumnName = m_oAdo.m_DataSet.Tables["econ_variable"].Columns[i].ColumnName;

                        /***********************************
                        **all columns are read-only except weight
                        ***********************************/
                        if (strColumnName.Trim().ToUpper() == "WEIGHT")
                        {
                            /******************************************************************
                            **create a new instance of the DataGridColoredTextBoxColumn class
                            ******************************************************************/
                            aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(true, true, this);
                            aColumnTextColumn.Format = "#0.000";
                            aColumnTextColumn.ReadOnly = false;
                        }
                        else
                        {
                            /******************************************************************
                            **create a new instance of the DataGridColoredTextBoxColumn class
                            ******************************************************************/
                            aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(false, false, this);
                            aColumnTextColumn.ReadOnly = true;
                        }
                        aColumnTextColumn.HeaderText = strColumnName;

                        /********************************************************************
                         **assign the mappingname property the data sets column name
                         ********************************************************************/
                        aColumnTextColumn.MappingName = strColumnName;

                        /********************************************************************
                         **add the datagridcoloredtextboxcolumn object to the data grid 
                         **table style object
                         ********************************************************************/
                        tableStyle.GridColumnStyles.Add(aColumnTextColumn);

                        /**********************************
                         * Hide calculated_variables_id column
                         * *******************************/
                        if (strColumnName.Equals("calculated_variables_id"))
                            tableStyle.GridColumnStyles.Remove(aColumnTextColumn);


                    }
                    /*********************************************************************
                     ** make the dataGrid use our new tablestyle and bind it to our table
                     *********************************************************************/
                    if (frmMain.g_oGridViewFont != null) this.m_dgEcon.Font = frmMain.g_oGridViewFont;
                    this.m_dgEcon.TableStyles.Clear();
                    this.m_dgEcon.TableStyles.Add(tableStyle);
                    this.m_dgEcon.DataSource = this.m_econ_dv;
                    this.m_dgEcon.Expand(-1);
                    this.SumWeights(true);
                }
                catch (Exception e2)
                {
                    MessageBox.Show(e2.Message, "Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.m_intError = -1;
                    m_oAdo.m_OleDbConnection.Close();
                    m_oAdo.m_OleDbConnection = null;
                    m_oAdo.m_DataSet.Clear();
                    m_oAdo.m_DataSet = null;
                    m_oAdo.m_OleDbDataAdapter.Dispose();
                    m_oAdo.m_OleDbDataAdapter = null;
                    return;

                }
            }
        }

        protected void loadEconVariablesGridSqlite()
        {
            if (m_oDataMgr.m_intError == 0)
            {
                string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                try
                {
                    string strPrimaryKeys = "calculated_variables_id, rxcycle";
                    string strColumns = "calculated_variables_id, rxcycle, weight";
                    string strWhereClause = "calculated_variables_id = -1";
                    m_oDataMgr.InitializeDataAdapterArrayItem(ECON_DETAILS_TABLE, strTableName, strColumns, strPrimaryKeys, strWhereClause);
                    this.m_dtTableSchema = m_oDataMgr.getTableSchema(m_oDataMgr.m_Connection,
                                                 m_oDataMgr.m_Transaction,
                                                 m_oDataMgr.m_strSQL);

                    if (m_oDataMgr.m_intError == 0)
                    {
                        this.m_econ_dv = new DataView(m_oDataMgr.m_DataSet.Tables[strTableName]);

                        this.m_econ_dv.AllowNew = false;       //cannot append new records
                        this.m_econ_dv.AllowDelete = false;    //cannot delete records
                        this.m_econ_dv.AllowEdit = true;
                        this.m_dgEcon.CaptionText = "econ_variable";
                        m_dgEcon.BackgroundColor = frmMain.g_oGridViewBackgroundColor;

                        /***********************************************************************************
                        **assign the aColumnTextColumn as type DataGridColoredTextBoxColumn object class
                        ***********************************************************************************/
                        WeightedAverage_DataGridColoredTextBoxColumn aColumnTextColumn;


                        /***************************************************************
                         **custom define the grid style
                         ***************************************************************/
                        DataGridTableStyle tableStyle = new DataGridTableStyle();

                        /***********************************************************************
                         **map the data grid table style to the scenario rx intensity dataset
                         ***********************************************************************/
                        tableStyle.MappingName = strTableName;
                        tableStyle.AlternatingBackColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
                        tableStyle.BackColor = frmMain.g_oGridViewRowBackgroundColor;
                        tableStyle.ForeColor = frmMain.g_oGridViewRowForegroundColor;
                        tableStyle.SelectionBackColor = frmMain.g_oGridViewSelectedRowBackgroundColor;


                        /******************************************************************************
                         **since the dataset has things like field name and number of columns,
                         **we will use those to create new columnstyles for the columns in our grid
                         ******************************************************************************/
                        //get the number of columns from the view_weights data set
                        int numCols = m_oDataMgr.m_DataSet.Tables[strTableName].Columns.Count;

                        /************************************************
                         **loop through all the columns in the dataset	
                         ************************************************/
                        string strColumnName = ""; ;
                        for (int i = 0; i < numCols; ++i)
                        {
                            strColumnName = m_oDataMgr.m_DataSet.Tables[strTableName].Columns[i].ColumnName;

                            /***********************************
                            **all columns are read-only except weight
                            ***********************************/
                            if (strColumnName.Trim().ToUpper() == "WEIGHT")
                            {
                                /******************************************************************
                                **create a new instance of the DataGridColoredTextBoxColumn class
                                ******************************************************************/
                                aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(true, true, this);
                                aColumnTextColumn.Format = "#0.000";
                                aColumnTextColumn.ReadOnly = false;
                            }
                            else
                            {
                                /******************************************************************
                                **create a new instance of the DataGridColoredTextBoxColumn class
                                ******************************************************************/
                                aColumnTextColumn = new WeightedAverage_DataGridColoredTextBoxColumn(false, false, this);
                                aColumnTextColumn.ReadOnly = true;
                            }
                            aColumnTextColumn.HeaderText = strColumnName;

                            /********************************************************************
                             **assign the mappingname property the data sets column name
                             ********************************************************************/
                            aColumnTextColumn.MappingName = strColumnName;

                            /********************************************************************
                             **add the datagridcoloredtextboxcolumn object to the data grid 
                             **table style object
                             ********************************************************************/
                            tableStyle.GridColumnStyles.Add(aColumnTextColumn);

                            /**********************************
                             * Hide calculated_variables_id column
                             * *******************************/
                            if (strColumnName.Equals("calculated_variables_id"))
                                tableStyle.GridColumnStyles.Remove(aColumnTextColumn);


                        }
                        /*********************************************************************
                         ** make the dataGrid use our new tablestyle and bind it to our table
                         *********************************************************************/
                        if (frmMain.g_oGridViewFont != null) this.m_dgEcon.Font = frmMain.g_oGridViewFont;
                        this.m_dgEcon.TableStyles.Clear();
                        this.m_dgEcon.TableStyles.Add(tableStyle);
                        this.m_dgEcon.DataSource = this.m_econ_dv;
                        this.m_dgEcon.Expand(-1);
                        this.SumWeights(true);
                    }
                }
                catch (Exception e2)
                {
                    MessageBox.Show(e2.Message, "Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.m_intError = -1;
                    m_oDataMgr.m_DataSet.Clear();
                    m_oDataMgr.m_DataSet = null;
                    m_oDataMgr.m_DataAdapter.Dispose();
                    m_oDataMgr.m_DataAdapter = null;
                    return;
                }
            }
        }

        protected void SendKeyStrokes(System.Windows.Forms.TextBox p_oTextBox, string strKeyStrokes)
        {
            try
            {
                p_oTextBox.Focus();
                System.Windows.Forms.SendKeys.Send(strKeyStrokes);
                p_oTextBox.Refresh();
            }
            catch (Exception caught)
            {
                MessageBox.Show("SendKeyStrokes Method Failed With This Message:" + caught.Message);
            }

        }







        protected void NextButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
        {
            p_oGb.Controls.Add(p_oButton);
            p_oButton.Left = p_oGb.Width - p_oButton.Width - 5;
            p_oButton.Top = p_oGb.Height - p_oButton.Height - 5;
            p_oButton.Name = strButtonName;
        }
        protected void PrevButton(ref System.Windows.Forms.GroupBox p_oGb, ref System.Windows.Forms.Button p_oButton, string strButtonName)
        {
            p_oGb.Controls.Add(p_oButton);
            p_oButton.Top = this.btnNext.Top;
            p_oButton.Height = this.btnNext.Height;
            p_oButton.Width = this.btnNext.Width;
            p_oButton.Left = this.btnNext.Left - p_oButton.Width;
            p_oButton.Name = strButtonName;
        }

        private void ShowGroupBox(string p_strName)
        {
            int x;
            //System.Windows.Forms.Control oControl;
            for (x = 0; x <= groupBox1.Controls.Count - 1; x++)
            {
                if (groupBox1.Controls[x].Name.Substring(0, 3) == "grp")
                {
                    if (p_strName.Trim().ToUpper() ==
                        groupBox1.Controls[x].Name.Trim().ToUpper())
                    {
                        groupBox1.Controls[x].Show();
                    }
                    else
                    {
                        groupBox1.Controls[x].Hide();
                    }
                }
            }
        }

        private void grpboxFVSVariablesPrePost_Resize(object sender, System.EventArgs e)
        {

        }

        private string ValidateNumeric(string p_strValue)
        {
            string strValue = p_strValue.Replace("$", "");
            strValue = strValue.Replace(",", "");
            try
            {
                double dbl = Convert.ToDouble(strValue);
            }
            catch
            {
                return "0";
            }
            return strValue;
        }


        private void groupBox1_Leave(object sender, System.EventArgs e)
        {

        }
        private void EnableTabs(bool p_bEnable)
        {
            int x;
            ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlScenario, "tbdesc,tbnotes,tbdatasources", p_bEnable);
            ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlRules, "tbpsites,tbowners,tbcost,tbtreatmentintensity,tbfilterplots,tbrun", p_bEnable);
            ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlFVSVariables, "tbeffective,tbtiebreaker", p_bEnable);
            for (x = 0; x <= ReferenceOptimizerScenarioForm.tlbScenario.Buttons.Count - 1; x++)
            {
                ReferenceOptimizerScenarioForm.tlbScenario.Buttons[x].Enabled = p_bEnable;
            }
            frmMain.g_oFrmMain.grpboxLeft.Enabled = p_bEnable;
            frmMain.g_oFrmMain.tlbMain.Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[0].Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[1].Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[2].Enabled = p_bEnable;

        }

        private void btnOptimizationAudit_Click(object sender, System.EventArgs e)
        {
            this.DisplayAuditMessage = true;
            Audit();
        }
        public void Audit()
        {


            int x;
            this.m_intError = 0;
            m_strError = "";
            if (DisplayAuditMessage)
            {
                this.m_strError = "Audit Results \r\n";
                this.m_strError = m_strError + "-------------\r\n\r\n";
            }


            if (DisplayAuditMessage)
            {
                if (m_intError == 0) this.m_strError = m_strError + "Passed Audit";
                else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";
                MessageBox.Show(m_strError, "FIA Biosum");
            }

        }

        public bool DisplayAuditMessage
        {
            get { return _bDisplayAuditMsg; }
            set { _bDisplayAuditMsg = value; }
        }
        public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
        {
            get { return _frmScenario; }
            set { _frmScenario = value; }
        }
        public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective.Variables ReferenceFVSVariables
        {
            get { return this._oCurVar; }
            set { _oCurVar = value; }
        }
        public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker ReferenceTieBreaker
        {
            get { return _uc_tiebreaker; }
            set { _uc_tiebreaker = value; }
        }

        private void btnCancelSummary_Click(object sender, EventArgs e)
        {
            this.ParentForm.Close();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            if (m_oDataMgr != null && m_oDataMgr.m_Connection != null)
            {
                if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, m_strFvsViewTableName))
                {
                    m_oDataMgr.m_strSQL = "DROP TABLE " + m_strFvsViewTableName;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + "\r\n");
                    }
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                }
                m_oDataMgr.CloseAndDisposeConnection(m_oDataMgr.m_Connection, true);
            }
            base.OnHandleDestroyed(e);
        }

        private void btnNewFvs_Click(object sender, EventArgs e)
        {
            m_intCurVar = -1;
            this.enableFvsVariableUc(true);
            lstFVSTablesList.ClearSelected();

            if (! m_bUsingSqlite)
            {
                m_oAdoFvs.m_DataSet.Tables["view_weights"].Rows.Clear();
            }
            else
            {
                m_oDataMgr.m_DataSet.Tables[m_strFvsViewTableName].Rows.Clear();
            }

            lblFvsVariableName.Text = "Not Defined";
            txtFVSVariableDescr.Text = "";
            txtFvsVariableTotalWeight.Text = "";
            cboFvsVariableBaselinePkg.Items.Clear();
            BtnFvsImport.Enabled = false;

            this.m_dgEcon.Expand(-1);
            this.grpboxSummary.Hide();
            this.grpboxDetails.Show();
        }

        private void btnNewEcon_Click(object sender, EventArgs e)
        {
            m_intCurVar = -1;
            lstEconVariablesList.ClearSelected();
            this.enableEconVariableUc(true);
            BtnDeleteEconVariable.Enabled = false;
            lblSelectedEconType.Text = "Not Defined";
            int intNewId = GetNextId();
            this.m_econ_dv.AllowNew = true;

            //@ToDo: Refine when we complete swap to Sqlite
            if (! m_bUsingSqlite)
            {
                this.m_oAdo.m_DataSet.Clear();
                // Add row to the dataset for each cycle
                for (int i = 1; i < 5; i++)
                {
                    DataRow p_row = this.m_oAdo.m_DataSet.Tables["econ_variable"].NewRow();
                    p_row["calculated_variables_id"] = intNewId;
                    p_row["rxcycle"] = Convert.ToString(i);
                    p_row["weight"] = 0;
                    this.m_oAdo.m_DataSet.Tables["econ_variable"].Rows.Add(p_row);
                    p_row = null;
                }
            }
            else
            {
                // Replaced m_oDataMgr.m_DataSet.Tables["econ_variable"] with
                // m_oDataMgr.m_DataSet.Tables[strTableName]
                string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                this.m_oDataMgr.m_DataSet.Clear();
                for (int i = 1; i < 5; i++)
                {
                    DataRow p_row = m_oDataMgr.m_DataSet.Tables[strTableName].NewRow();
                    p_row["calculated_variables_id"] = intNewId;
                    p_row["rxcycle"] = Convert.ToString(i);
                    p_row["weight"] = 0;
                    m_oDataMgr.m_DataSet.Tables[strTableName].Rows.Add(p_row);
                    p_row = null;
                }
            }

            this.m_econ_dv.AllowNew = false;
            this.SumWeights(true);

            //Remove and re-add weight column so it is editable
            this.updateWeightColumn(VARIABLE_ECON, true);
            this.m_dgEcon.Expand(-1);

            lblEconVariableName.Text = "Not Defined";
            txtEconVariableDescr.Text = "";
            BtnEconImport.Enabled = false;
            this.grpboxSummary.Hide();
            this.grpBoxEconomicVariable.Show();
        }

        private void btnFvsDetailsCancel_Click(object sender, EventArgs e)
        {
            this.grpboxSummary.Show();
            this.grpboxDetails.Hide();
        }


        private void lstVariables_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    int intRowHt = this.lstVariables.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    this.lstVariables.Items[lstVariables.TopItem.Index + (int)dblRow - 1].Selected = true;
                    this.m_oLvAlternateColors.DelegateListViewItem(lstVariables.Items[lstVariables.TopItem.Index + (int)dblRow - 1]);
                }
            }
            catch
            {
            }
        }

        private void lstVariables_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (lstVariables.SelectedItems.Count > 0)
                m_oLvAlternateColors.DelegateListViewItem(lstVariables.SelectedItems[0]);
        }

        private void btnEconDetailsCancel_Click(object sender, EventArgs e)
        {
            this.grpboxSummary.Show();
            this.grpBoxEconomicVariable.Hide();
        }

        public void SumWeights(bool bIsEconVariable)
        {
            DataTable objDataTable;
            if (bIsEconVariable == true)
            {
                objDataTable = this.m_econ_dv.Table;
            }
            else
            {
                objDataTable = this.m_dv.Table;
            }
            double dblSum = 0;
            double dblWeight = -1;
            foreach (DataRow row in objDataTable.Rows)
            {
                string strWeight = row["weight"].ToString();
                if (Double.TryParse(strWeight, out dblWeight))
                    dblSum = dblSum + dblWeight;
            }
            if (bIsEconVariable == false)
            {
                txtFvsVariableTotalWeight.Text = String.Format("{0:0.0#}", dblSum);
            }
            else
            {
                txtEconVariableTotalWeight.Text = String.Format("{0:0.0#}", dblSum);
            }
        }

        protected void m_dg_CurCellChange(object sender, EventArgs e)
        {
            //Only recalculate if we are leaving the weight column
            if (m_intPrevColumnIdx.Equals(3))
                this.SumWeights(false);
            m_intPrevColumnIdx = m_dg.CurrentCell.ColumnNumber;
        }

        protected void m_dgEcon_CurCellChange(object sender, EventArgs e)
        {
            //Only recalculate if we are leaving the weight column
            if (m_intPrevColumnIdx.Equals(1))
                this.SumWeights(true);
            m_intPrevColumnIdx = m_dgEcon.CurrentCell.ColumnNumber;
        }

        private void m_dg_Leave(object sender, EventArgs e)
        {
            this.SumWeights(false);
        }

        private void m_dgEcon_Leave(object sender, EventArgs e)
        {
            this.SumWeights(true);
        }

        private void lstFVSTablesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstFVSFieldsList.Items.Clear();
            this.LblSelectedVariable.Text = "Not Defined";
            if (this.lstFVSTablesList.SelectedIndex > -1)
            {
                System.Collections.Generic.IList<string> lstFields =
                    m_dictFVSTables[Convert.ToString(this.lstFVSTablesList.SelectedItem)];
                if (lstFields != null)
                {
                    foreach (string strField in lstFields)
                    {
                        lstFVSFieldsList.Items.Add(strField);
                    }
                }
            }
        }

        private void lstFVSFieldsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.LblSelectedVariable.Text = "Not Defined";
            this.lblFvsVariableName.Text = "Not Defined";
            this.BtnFvsImport.Enabled = false;
            //if (this.lstFVSFieldsList.SelectedIndex > -1)
            //{
            //    this.btnFVSVariablesOptimizationVariableValues.Enabled = true;
            //}
            //else
            //{
            //    this.btnFVSVariablesOptimizationVariableValues.Enabled = false;
            //}
        }

        private void btnFVSVariableValue_Click(object sender, EventArgs e)
        {
            if (this.lstFVSTablesList.SelectedItems.Count == 0 || this.lstFVSFieldsList.SelectedItems.Count == 0) return;
            this.LblSelectedVariable.Text =
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();
            string strVariableName = "";
            bool bFoundIt = false;
            bool bExists = false;
            int intSuffix = 1;
            do
            {
                strVariableName = this.lstFVSFieldsList.SelectedItems[0].ToString() + "_" + intSuffix;
                bExists = false;
                foreach (ListViewItem oItem in this.lstVariables.Items)
                {
                    if (oItem.Text.Trim().Equals(strVariableName))
                    {
                        intSuffix = intSuffix + 1;
                        bExists = true;
                        break;
                    }
                }
                if (bExists == false)
                    bFoundIt = true;
            } 
            while (bFoundIt == false);
            lblFvsVariableName.Text = strVariableName;
            this.loadm_dg();
            //Remove and re-add weight column so it is editable
            this.updateWeightColumn(VARIABLE_FVS, true);
            this.SumWeights(false);
        }

        private void btnProperties_Click(object sender, EventArgs e)
        {
            if (lstVariables.SelectedItems.Count == 0) return;
            this.grpboxSummary.Hide();
            m_intCurVar = Convert.ToInt32(lstVariables.SelectedItems[0].SubItems[3].Text.Trim());
            string strVariableSource = lstVariables.SelectedItems[0].SubItems[5].Text.Trim();
            string strVariableName = lstVariables.SelectedItems[0].Text.Trim();

            if (!m_bUsingSqlite)
            {
                load_properties_accdb(strVariableName, strVariableSource);
            }
            else
            {
                load_properties_sqlite(strVariableName, strVariableSource);
            }

            if (lstVariables.SelectedItems[0].SubItems[2].Text.Trim().Equals(VARIABLE_ECON))
            {
                lblEconVariableName.Text = strVariableName;
                txtEconVariableDescr.Text = lstVariables.SelectedItems[0].SubItems[1].Text.Trim();
                BtnEconImport.Enabled = true;
                this.SumWeights(true);
                this.updateWeightColumn(VARIABLE_ECON, false);
            }
            else
            {
                this.LblSelectedVariable.Text =
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();
                lblFvsVariableName.Text = strVariableName;
                txtFVSVariableDescr.Text = lstVariables.SelectedItems[0].SubItems[1].Text.Trim();
                BtnFvsImport.Enabled = true;
                this.enableFvsVariableUc(false);
                this.SumWeights(false);
                this.updateWeightColumn(VARIABLE_FVS, false);
                this.grpboxDetails.Show();
            }
        }

        private void load_properties_accdb(string strVariableName, string strVariableSource)
        {
            ado_data_access oAdo = new ado_data_access();
            string strPropertiesConn = m_oAdo.getMDBConnString(m_strCalculatedVariablesAccdb, "", "");
            using (var oPropertiesConn = new OleDbConnection(strPropertiesConn))
            {
                oPropertiesConn.Open();
                if (lstVariables.SelectedItems[0].SubItems[2].Text.Trim().Equals(VARIABLE_ECON))
                {
                    string strSelectedType = getEconVariableType(strVariableName);
                    int idxType = 0;
                    foreach (string strValue in PREFIX_ECON_VALUE_ARRAY)
                    {
                        if (strValue.Equals(strSelectedType))
                        {
                            lblSelectedEconType.Text = PREFIX_ECON_NAME_ARRAY[idxType];
                            break;
                        }
                        else
                        {
                            idxType++;
                        }
                    }
                    lstEconVariablesList.SelectedIndex = idxType;
                    m_oAdo.m_DataSet.Clear();
                    oAdo.m_strSQL = "select * from " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName +
                        " where calculated_variables_id = " + m_intCurVar;
                    oAdo.m_OleDbCommand = oPropertiesConn.CreateCommand();
                    oAdo.m_OleDbCommand.CommandText = oAdo.m_strSQL;
                    oAdo.m_OleDbDataAdapter = new OleDbDataAdapter();
                    oAdo.m_OleDbDataAdapter.SelectCommand = oAdo.m_OleDbCommand;
                    oAdo.m_OleDbDataAdapter.SelectCommand.Transaction = oAdo.m_OleDbTransaction;
                    oAdo.m_OleDbDataAdapter.Fill(m_oAdo.m_DataSet, "econ_variable");
                    this.enableEconVariableUc(false);
                    BtnDeleteEconVariable.Enabled = true;
                    for (int i = 0; i < PREFIX_ECON_VALUE_ARRAY.Length; i++)
                    {
                        if (strVariableName.Equals(PREFIX_ECON_VALUE_ARRAY[i] + "_1"))
                        {
                            BtnDeleteEconVariable.Enabled = false;
                            break;
                        }
                    }
                    this.grpBoxEconomicVariable.Show();
                }
                else
                {
                    oAdo.m_strSQL = "select * from " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName +
                        " where calculated_variables_id = " + m_intCurVar;
                    oAdo.SqlQueryReader(oPropertiesConn, oAdo.m_strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            //Selected FVS table (lstFVSTablesList)
                            string[] strPieces = strVariableSource.Split('.');
                            for (int i = 0; i < lstFVSTablesList.Items.Count; i++)
                            {
                                string strTable = lstFVSTablesList.Items[i].ToString();
                                if (strPieces[0].Equals(strTable))
                                {
                                    lstFVSTablesList.SelectedIndex = i;
                                    break;
                                }
                            }
                            //Selected FVS variable (lstFVSFieldsList)
                            if (lstFVSTablesList.SelectedIndex > -1)
                            {
                                for (int i = 0; i < lstFVSFieldsList.Items.Count; i++)
                                {
                                    string strField = lstFVSFieldsList.Items[i].ToString();
                                    Console.WriteLine("field: " + strField);
                                    if (strPieces[1].Equals(strField))
                                    {
                                        lstFVSFieldsList.SelectedIndex = i;
                                        break;
                                    }
                                }
                            }
                            // weights table
                            this.loadm_dg();
                            //Baseline Rx Package
                            //This has to come after m_dg (and combobox) are loaded
                            cboFvsVariableBaselinePkg.SelectedIndex = -1;
                            string strBaselineRxPkg = Convert.ToString(lstVariables.SelectedItems[0].SubItems[4].Text.Trim());
                            for (int i = 0; i < cboFvsVariableBaselinePkg.Items.Count; i++)
                            {
                                string strRxPkg = cboFvsVariableBaselinePkg.Items[i].ToString();
                                if (strRxPkg.Equals(strBaselineRxPkg))
                                {
                                    cboFvsVariableBaselinePkg.SelectedIndex = i;
                                    break;
                                }
                            }
                            foreach (System.Data.DataRow p_row in m_oAdoFvs.m_DataSet.Tables["view_weights"].Rows)
                            {
                                string strRxCycle = Convert.ToString(p_row["rxcycle"]);
                                string strPrePost = Convert.ToString(p_row["pre_or_post"]).Trim();
                                switch (strRxCycle)
                                {
                                    case "1":
                                        if (strPrePost.Equals("PRE"))
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_1_PRE"]);
                                        }
                                        else
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_1_POST"]);
                                        }
                                        break;
                                    case "2":
                                        if (strPrePost.Equals("PRE"))
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_2_PRE"]);
                                        }
                                        else
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_2_POST"]);
                                        }
                                        break;
                                    case "3":
                                        if (strPrePost.Equals("PRE"))
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_3_PRE"]);
                                        }
                                        else
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_3_POST"]);
                                        }
                                        break;
                                    case "4":
                                        if (strPrePost.Equals("PRE"))
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_4_PRE"]);
                                        }
                                        else
                                        {
                                            p_row["weight"] = Convert.ToDouble(oAdo.m_OleDbDataReader["WEIGHT_4_POST"]);
                                        }
                                        break;
                                }

                            }

                        }
                    }
                    oAdo.m_OleDbDataReader.Close();
                    oAdo = null;
                }
            }
        }

        private void load_properties_sqlite(string strVariableName, string strVariableSource)
        {
            if (lstVariables.SelectedItems[0].SubItems[2].Text.Trim().Equals(VARIABLE_ECON))
            {
                string strSelectedType = getEconVariableType(strVariableName);
                int idxType = 0;
                foreach (string strValue in PREFIX_ECON_VALUE_ARRAY)
                {
                    if (strValue.Equals(strSelectedType))
                    {
                        lblSelectedEconType.Text = PREFIX_ECON_NAME_ARRAY[idxType];
                        break;
                    }
                    else
                    {
                        idxType++;
                    }
                }
                lstEconVariablesList.SelectedIndex = idxType;
                string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                string strPrimaryKeys = "calculated_variables_id, rxcycle";
                string strColumns = "calculated_variables_id, rxcycle, weight";
                string strWhereClause = "calculated_variables_id = " + m_intCurVar;
                //m_oDataMgr.InitializeDataAdapterArrayItem(ECON_DETAILS_TABLE, strTableName, strColumns, strPrimaryKeys, strWhereClause);
                m_oDataMgr.DataAdapterArrayItemConfigureSelectCommand(ECON_DETAILS_TABLE, strTableName, strColumns, strPrimaryKeys, strWhereClause);

                this.enableEconVariableUc(false);
                    BtnDeleteEconVariable.Enabled = true;
                    for (int i = 0; i < PREFIX_ECON_VALUE_ARRAY.Length; i++)
                    {
                        if (strVariableName.Equals(PREFIX_ECON_VALUE_ARRAY[i] + "_1"))
                        {
                            BtnDeleteEconVariable.Enabled = false;
                            break;
                        }
                    }
                    this.grpBoxEconomicVariable.Show();
                }
                else
                {

                            //Selected FVS table (lstFVSTablesList)
                            string[] strPieces = strVariableSource.Split('.');
                            for (int i = 0; i < lstFVSTablesList.Items.Count; i++)
                            {
                                string strTable = lstFVSTablesList.Items[i].ToString();
                                if (strPieces[0].Equals(strTable))
                                {
                                    lstFVSTablesList.SelectedIndex = i;
                                    break;
                                }
                            }
                            //Selected FVS variable (lstFVSFieldsList)
                            if (lstFVSTablesList.SelectedIndex > -1)
                            {
                                for (int i = 0; i < lstFVSFieldsList.Items.Count; i++)
                                {
                                    string strField = lstFVSFieldsList.Items[i].ToString();
                                    Console.WriteLine("field: " + strField);
                                    if (strPieces[1].Equals(strField))
                                    {
                                        lstFVSFieldsList.SelectedIndex = i;
                                        break;
                                    }
                                }
                            }
                            // weights table
                            this.loadm_dg();
                            //Baseline Rx Package
                            //This has to come after m_dg (and combobox) are loaded
                            cboFvsVariableBaselinePkg.SelectedIndex = -1;
                            string strBaselineRxPkg = Convert.ToString(lstVariables.SelectedItems[0].SubItems[4].Text.Trim());
                            for (int i = 0; i < cboFvsVariableBaselinePkg.Items.Count; i++)
                            {
                                string strRxPkg = cboFvsVariableBaselinePkg.Items[i].ToString();
                                if (strRxPkg.Equals(strBaselineRxPkg))
                                {
                                    cboFvsVariableBaselinePkg.SelectedIndex = i;
                                    break;
                                }
                            }
                string strSql = "select * from " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName +
                    " where calculated_variables_id = " + m_intCurVar;
                using (SQLiteCommand cmd = new SQLiteCommand(strSql, m_oDataMgr.m_Connection))
                {
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                foreach (DataRow p_row in m_oDataMgr.m_DataSet.Tables[m_strFvsViewTableName].Rows)
                                {
                                    string strRxCycle = Convert.ToString(p_row["rxcycle"]);
                                    string strPrePost = Convert.ToString(p_row["pre_or_post"]).Trim();
                                    switch (strRxCycle)
                                    {
                                        case "1":
                                            if (strPrePost.Equals("PRE"))
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_1_PRE"]);
                                            }
                                            else
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_1_POST"]);
                                            }
                                            break;
                                        case "2":
                                            if (strPrePost.Equals("PRE"))
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_2_PRE"]);
                                            }
                                            else
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_2_POST"]);
                                            }
                                            break;
                                        case "3":
                                            if (strPrePost.Equals("PRE"))
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_3_PRE"]);
                                            }
                                            else
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_3_POST"]);
                                            }
                                            break;
                                        case "4":
                                            if (strPrePost.Equals("PRE"))
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_4_PRE"]);
                                            }
                                            else
                                            {
                                                p_row["weight"] = Convert.ToDouble(reader["WEIGHT_4_POST"]);
                                            }
                                            break;
                                    }
                                }
                            }
                        }
                    } 
                } 
            }
        }

        private void btnEconVariableType_Click(object sender, EventArgs e)
        {
            if (this.lstEconVariablesList.SelectedItems.Count == 0 || this.lstEconVariablesList.SelectedItems.Count == 0) return;
            this.lblSelectedEconType.Text =
                this.lstEconVariablesList.SelectedItems[0].ToString();
            string strVariableName = "";
            int i = 0;
            foreach (string strName in PREFIX_ECON_NAME_ARRAY)
            {
                if (this.lblSelectedEconType.Text.Equals(strName))
                    break;
                i++;
            }
            bool bFoundIt = false;
            bool bExists = false;
            int intSuffix = 1;
            do
            {
                strVariableName = PREFIX_ECON_VALUE_ARRAY[i] + "_" + intSuffix;
                bExists = false;
                foreach (ListViewItem oItem in this.lstVariables.Items)
                {
                    if (oItem.Text.Trim().Equals(strVariableName))
                    {
                        intSuffix = intSuffix + 1;
                        bExists = true;
                        break;
                    }
                }
                if (bExists == false)
                    bFoundIt = true;
            }
            while (bFoundIt == false);
            lblEconVariableName.Text = strVariableName;
            this.BtnEconImport.Enabled = true;
        }

        public static string getEconVariableType(string strName)
        {
            if (strName.Contains(PREFIX_CHIP_VOLUME))
            {
                return PREFIX_CHIP_VOLUME;
            }
            else if (strName.Contains(PREFIX_MERCH_VOLUME))
            {
                return PREFIX_MERCH_VOLUME;
            }
            else if (strName.Contains(PREFIX_NET_REVENUE))
            {
                return PREFIX_NET_REVENUE;
            }
            else if (strName.Contains(PREFIX_TOTAL_VOLUME))
            {
                return PREFIX_TOTAL_VOLUME;
            }
            else if (strName.Contains(PREFIX_TREATMENT_HAUL_COSTS))
            {
                return PREFIX_TREATMENT_HAUL_COSTS;
            }
            else if (strName.Contains(PREFIX_ONSITE_TREATMENT_COSTS))
            {
                return PREFIX_ONSITE_TREATMENT_COSTS;
            }
            else
            {
                return "";
            }
        }

        private void updateWeightColumn(string strWeightType, bool bEdit)
        {
            DataGridTableStyle objTableStyle = this.m_dgEcon.TableStyles[0];
            if (strWeightType.Equals(VARIABLE_FVS))
            {
                objTableStyle = this.m_dg.TableStyles[0];
            }

            WeightedAverage_DataGridColoredTextBoxColumn objColumnWeight =
                (WeightedAverage_DataGridColoredTextBoxColumn)objTableStyle.GridColumnStyles["weight"];
            objTableStyle.GridColumnStyles.Remove(objColumnWeight);
            if (bEdit == false)
            {
                objColumnWeight = new WeightedAverage_DataGridColoredTextBoxColumn(false, true, this);
                objColumnWeight.ReadOnly = true;
            }
            else
            {
                objColumnWeight = new WeightedAverage_DataGridColoredTextBoxColumn(true, true, this);
                objColumnWeight.ReadOnly = false;
            }
            objColumnWeight.Format = "#0.000";

            objColumnWeight.HeaderText = "weight";
            objColumnWeight.MappingName = "weight";
            objTableStyle.GridColumnStyles.Add(objColumnWeight);

            if (strWeightType.Equals(VARIABLE_ECON))
            {
                this.m_dgEcon.Expand(-1);
            }
            else
            {
                this.m_dg.Expand(-1);
            }
        }

        private void btnFvsCalculate_Click(object sender, EventArgs e)
        {
            dao_data_access oDao = new dao_data_access();
            try
            {
                //Determine database and table names based on the source FVS variable
                string[] strPieces = LblSelectedVariable.Text.Split('.');
                string strSourcePreTable = "PRE_" + strPieces[0];
                string strSourcePostTable = "POST_" + strPieces[0];
                string strSourceDatabaseName = "PREPOST_" + strPieces[0] + ".ACCDB";
                string strTargetPreTable = "PRE_" + strPieces[0] + "_WEIGHTED";
                string strTargetPostTable = "POST_" + strPieces[0] + "_WEIGHTED";
                string strWeightsByRxCyclePreTable = "WEIGHTS_BY_RX_CYCLE_PRE";
                string strWeightsByRxCyclePostTable = "WEIGHTS_BY_RX_CYCLE_POST";
                string strWeightsByRxPkgPreTable = "WEIGHTS_BY_RXPACKAGE_PRE";
                string strWeightsByRxPkgPostTable = "WEIGHTS_BY_RXPACKAGE_POST";

                this.val_data_fvs(strSourceDatabaseName, strSourcePreTable, strSourcePostTable);
                if (this.m_intError == 0)
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "btnFvsCalculate_Click: Calculate weighted variable " + lblFvsVariableName.Text + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Temporary database path: " + m_strTempMDB + "\r\n\r\n");
                    }

                    this.enableFvsVariableUc(false);
                    this.btnDeleteFvsVariable.Enabled = false;
                    this.btnFvsCalculate.Visible = true;
                    frmMain.g_oFrmMain.ActivateStandByAnimation(
                       frmMain.g_oFrmMain.WindowState,
                       frmMain.g_oFrmMain.Left,
                       frmMain.g_oFrmMain.Height,
                       frmMain.g_oFrmMain.Width,
                       frmMain.g_oFrmMain.Top);
                    
                    //Save associated configuration records
                    frmMain.g_sbpInfo.Text = "Saving scenario rule definitions...Stand by";
                    if (! m_bUsingSqlite)
                    {
                        savevalues("FVS");
                    }
                    else
                    {
                        savevalues_sqlite("FVS");
                    }
                    

                    frmMain.g_sbpInfo.Text = "Calculating and saving PRE/POST values...Stand by";
                    string strPrePostWeightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                        "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
                    string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                        "\\fvs\\db\\" + strSourceDatabaseName;

                    //Link to source FVS tables in temp .mdb if they don't exist from a previous run
                    if (!oDao.TableExists(m_strTempMDB, strSourcePreTable))
                    {
                        oDao.CreateTableLink(m_strTempMDB, strSourcePreTable, strFvsPrePostDb, strSourcePreTable);
                    }
                    if (!oDao.TableExists(m_strTempMDB, strSourcePostTable))
                    {
                        oDao.CreateTableLink(m_strTempMDB, strSourcePostTable, strFvsPrePostDb, strSourcePostTable);
                    }

                    //Drop strWeightsByRxCyclePreTable if it exists so we can recreate it
                    if (oDao.TableExists(m_strTempMDB, strWeightsByRxCyclePreTable))
                    {
                        oDao.DeleteTableFromMDB(m_strTempMDB, strWeightsByRxCyclePreTable);
                    }
                    //Drop strWeightsByRxCyclePostTable if it exists so we can recreate it
                    if (oDao.TableExists(m_strTempMDB, strWeightsByRxCyclePostTable))
                    {
                        oDao.DeleteTableFromMDB(m_strTempMDB, strWeightsByRxCyclePostTable);
                    }
                    //Drop strWeightsByRxPkgPreTable if it exists so we can recreate it
                    if (oDao.TableExists(m_strTempMDB, strWeightsByRxPkgPreTable))
                    {
                        oDao.DeleteTableFromMDB(m_strTempMDB, strWeightsByRxPkgPreTable);
                    }
                    //Drop strWeightsByRxPkgPostTable if it exists so we can recreate it
                    if (oDao.TableExists(m_strTempMDB, strWeightsByRxPkgPostTable))
                    {
                        oDao.DeleteTableFromMDB(m_strTempMDB, strWeightsByRxPkgPostTable);
                    }
                    //Drop strWeightsByRxCyclePostTable if it exists so we can recreate it
                    if (oDao.TableExists(m_strTempMDB, strWeightsByRxCyclePostTable))
                    {
                        oDao.DeleteTableFromMDB(m_strTempMDB, strWeightsByRxCyclePostTable);
                    }

                    // Load the cycles and weights in a structure for CalculateVariable. This allows us to
                    // share CalculateVariable with the recalculate functions
                    IList<string[]> lstWeights = new List<string[]>();
                    foreach (DataRow row in this.m_dv.Table.Rows)
                    {
                        string[] strRow = new string[3];
                        strRow[idxRxCycle] = row["rxcycle"].ToString();
                        strRow[idxWeight] = row["weight"].ToString();
                        strRow[idxPreOrPost] = row["pre_or_post"].ToString().Trim();
                        lstWeights.Add(strRow);
                    }

                    // Create new tables if they don't exist
                    bool bNewTables = false;
                    if (!oDao.TableExists(strPrePostWeightedDb, strTargetPreTable))
                    {
                        //Link source tables to output database
                        oDao.CreateTableLink(strPrePostWeightedDb, strSourcePreTable, strFvsPrePostDb, strSourcePreTable);
                        oDao.CreateTableLink(strPrePostWeightedDb, strSourcePostTable, strFvsPrePostDb, strSourcePostTable);

                        string strConn = m_oAdo.getMDBConnString(strPrePostWeightedDb, "", "");
                        using (var conn = new System.Data.OleDb.OleDbConnection(strConn))
                        {
                            // FVS creates a record for
                            // each condition for each cycle regardless of whether there is activity
                            m_oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CDbl(0) as " +
                                  lblFvsVariableName.Text + " " +
                                  "INTO " + strTargetPreTable +
                                  " FROM " + strSourcePreTable;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Creating final pre/post tables. They did not already exist \r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + m_oAdo.m_strSQL + "\r\n\r\n");
                            }

                            m_oAdo.SqlNonQuery(strConn, m_oAdo.m_strSQL);
                            m_oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CDbl(0) as " +
                                              lblFvsVariableName.Text + " " +
                                              "INTO " + strTargetPostTable +
                                              " FROM " + strSourcePostTable;
                            m_oAdo.SqlNonQuery(strConn, m_oAdo.m_strSQL);
                            bNewTables = true;

                            oDao.DeleteTableFromMDB(strPrePostWeightedDb, strSourcePreTable);
                            oDao.DeleteTableFromMDB(strPrePostWeightedDb, strSourcePostTable);
                        }
                    }
                    
                    //Open connection to temporary database and create starting temporary tables
                    //that is table for weights by rx and rxcycle
                    string strCalculateConn = m_oAdo.getMDBConnString(m_strTempMDB, "", "");
                    int noCount = -1;
                    CalculateVariable(oDao, strCalculateConn, strWeightsByRxPkgPreTable, strSourcePreTable, strWeightsByRxPkgPostTable,
                                      strSourcePostTable, lblFvsVariableName.Text, strPieces[1], cboFvsVariableBaselinePkg.SelectedItem.ToString(),
                                      lstWeights, bNewTables, ref noCount);

                    //Reload the main grid
                    if (m_intError == 0)
                    {
                        if (!m_bUsingSqlite)
                        {
                            this.loadLstVariables();
                        }
                        else
                        {
                            this.loadLstVariablesSqlite();
                        }
                        
                    }                    

                    frmMain.g_sbpInfo.Text = "Ready";
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

                    MessageBox.Show("Variable calculation complete! Click Cancel to return to the main Calculated Variables page", "FIA Biosum");
                }
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Weighted Average", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.m_intError = -1;
                frmMain.g_sbpInfo.Text = "Ready";
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
            }
            finally
            {
                if (oDao != null)
                {
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                }
            }

        }

        private void enableFvsVariableUc(bool bEnabled)
        {
            this.cboFvsVariableBaselinePkg.Enabled = bEnabled;
            //this.lstFVSTablesList.Enabled = bEnabled;
            //this.lstFVSFieldsList.Enabled = bEnabled;
            this.b_FVSTableEnabled = bEnabled;
            this.btnFVSVariableValue.Visible = bEnabled;
            this.txtFVSVariableDescr.ReadOnly = !bEnabled;
            this.btnFvsCalculate.Enabled = bEnabled;
            this.btnDeleteFvsVariable.Enabled = !bEnabled;
            if (bEnabled == true)
            {
                BtnFvsImport.Text = "Import weights";
            }
            else
            {
                BtnFvsImport.Text = "Export weights";
            }
        }

        private void enableEconVariableUc(bool bEnabled)
        {
            this.lstEconVariablesList.Enabled = bEnabled;
            this.btnEconVariableType.Visible = bEnabled;
            this.txtEconVariableDescr.ReadOnly = !bEnabled;
            this.BtnSaveEcon.Enabled = bEnabled;
            this.BtnDeleteEconVariable.Enabled = bEnabled;
            if (bEnabled == true)
            {
                BtnEconImport.Text = "Import weights";
            }
            else
            {
                BtnEconImport.Text = "Export weights";
            }
        }

        public class VariableItem
        {
            public int intId = 0;
            public string strVariableName = "";
            public string strVariableDescr = "";
            public string strVariableType = "";
            public string strRxPackage = "";
            public string strVariableSource = "";
            public System.Collections.Generic.IList<double> lstWeights;
        }

        public class Variable_Collection : System.Collections.CollectionBase
        {
            public Variable_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.VariableItem m_oVariable)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oVariable))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oVariable);
            }
            public void Remove(int index)
            {
                // Check to see if there is a widget at the supplied index.
                if (index > Count - 1 || index < 0)
                // If no widget exists, a messagebox is shown and the operation 
                // is canColumned.
                {
                    System.Windows.Forms.MessageBox.Show("Index not valid!");
                }
                else
                {
                    List.RemoveAt(index);
                }
            }
            public FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.VariableItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.VariableItem)List[Index];
            }
        }

        private void btnDeleteFvsVariable_Click(object sender, EventArgs e)
        {
            ado_data_access oAdo = new ado_data_access();
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            string strScenarioConn = oAdo.getMDBConnString(strScenarioDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile, "", "");
            string[] strPieces = LblSelectedVariable.Text.Split('.');
            using (var oRenameConn = new OleDbConnection(strScenarioConn))
            {
                oRenameConn.Open();

                // Check for usage as Effectiveness variable
                string strWeightedVariableSource = "";
                if (strPieces.Length == 2)
                {
                    strWeightedVariableSource = "PRE_" + strPieces[0] + "_WEIGHTED." + lblFvsVariableName.Text;
                }
                else
                {
                    return;
                }
                oAdo.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName +
                    " WHERE (((UCase(Trim([PRE_FVS_VARIABLE]))) = UCase(Trim('" + strWeightedVariableSource + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oAdo.getRecordCount(oRenameConn, oAdo.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This FVS Variable Cannot Be Deleted Because It Is In Use As An Effectiveness Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
                // Check for usage as Optimization variable
                strWeightedVariableSource = strPieces[0] + "_WEIGHTED." + lblFvsVariableName.Text;
                oAdo.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName +
                    " WHERE (((UCase(Trim([fvs_variable_name]))) = UCase(Trim('" + strWeightedVariableSource + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oAdo.getRecordCount(oRenameConn, oAdo.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This FVS Variable Cannot Be Deleted Because It Is In Use As An Optimization Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
                // Check for usage as Tie-Breaker variable
                oAdo.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName +
                    " WHERE (((UCase(Trim([fvs_variable_name]))) = UCase(Trim('" + strWeightedVariableSource + "'))))";
                if ((int)oAdo.getRecordCount(oRenameConn, oAdo.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This FVS Variable Cannot Be Deleted Because It Is In Use As An Tie-Breaker Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
            }

            DialogResult objResult = MessageBox.Show("!!You are about to delete an FVS weighted variable. This action cannot be undone. Do you wish to continue?", "FIA Biosum",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        System.Windows.Forms.MessageBoxIcon.Question);
            if (objResult == DialogResult.Yes)
            {
                // Delete data entries from FVS pre/post tables
                string[] strFieldsArr = { lblFvsVariableName.Text };
                dao_data_access oDao = new dao_data_access();
                oDao.DeleteField(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile,
                    "PRE_" + strPieces[0] + "_WEIGHTED", strFieldsArr);
                oDao.DeleteField(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile,
                    "POST_" + strPieces[0] + "_WEIGHTED", strFieldsArr);

                if (! m_bUsingSqlite)
                {
                    DeleteVariableAccdb();
                }
                else
                {
                    DeleteVariableSqlite();
                }
                this.btnFvsDetailsCancel.PerformClick();

                if (oDao != null)
                {
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                }
            }
            else
            {
                return;
            }
        }

        private void DeleteVariableAccdb()
        {
            // Delete entries from configuration database
            string strCalculatedVariablesACCDB = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
            string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName;
            if (lstVariables.SelectedItems[0].SubItems[2].Text.Trim().Equals(VARIABLE_ECON))
            {
                strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
            }
            using (var oConn = new OleDbConnection(m_oAdo.getMDBConnString(strCalculatedVariablesACCDB, "", "")))
            {
                oConn.Open();
                m_oAdo.m_strSQL = "DELETE FROM " + strTableName +
                                  " WHERE calculated_variables_id = " + m_intCurVar;
                m_oAdo.SqlNonQuery(oConn, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = "DELETE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                                  " WHERE ID = " + m_intCurVar;
                m_oAdo.SqlNonQuery(oConn, m_oAdo.m_strSQL);
            }
            // Update UI
            this.loadLstVariables();
        }

        private void BtnHelpCalculatedMenu_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "TREATMENT_OPTIMIZER", "INTRODUCTION" });
        }

        private void BtnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "TREATMENT_OPTIMIZER", "FVS_VARIABLE" });
        }

        private void BtnHelpEconVariable_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "TREATMENT_OPTIMIZER", "ECONOMIC_VARIABLE" });
        }

        private void BtnDeleteEconVariable_Click(object sender, EventArgs e)
        {
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            // @ToDo: Switch this over to sqlite when migrating Optimizer
            ado_data_access oAdo = new ado_data_access();
            string strScenarioConn = oAdo.getMDBConnString(strScenarioDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile, "", "");
            using (var oScenarioConn = new OleDbConnection(strScenarioConn))
            {
                oScenarioConn.Open();

                // Check for usage as Optimization variable
                oAdo.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName +
                    " WHERE (((UCase(Trim([fvs_variable_name]))) = UCase(Trim('" + lblEconVariableName.Text + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oAdo.getRecordCount(oScenarioConn, oAdo.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This Economic Variable Cannot Be Deleted Because It Is In Use As An Optimization Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                // Check for usage as filter
                oAdo.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName +
                    " WHERE (((UCase(Trim([revenue_attribute]))) = UCase(Trim('" + lblEconVariableName.Text + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oAdo.getRecordCount(oScenarioConn, oAdo.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This Economic Variable Cannot Be Deleted Because It Is In Use As A Dollars Per Acre Filter!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                // Check for usage as tiebreaker
                oAdo.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName +
                    " WHERE (((UCase(Trim([fvs_variable_name]))) = UCase(Trim('" + lblEconVariableName.Text + "'))))";
                if ((int)oAdo.getRecordCount(oScenarioConn, oAdo.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This Economic Variable Cannot Be Deleted Because It Is In Use As An Tie-Breaker Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
            }
            if (oAdo != null)
            {
                if (oAdo.m_DataSet != null)
                {
                    oAdo.m_DataSet.Clear();
                    oAdo.m_DataSet.Dispose();
                }
                oAdo = null;
            }
            DialogResult objResult = MessageBox.Show("!!You are about to delete an Economic weighted variable. This action cannot be undone. Do you wish to continue?", "FIA Biosum",
                MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (objResult == DialogResult.Yes)
            {
                if (!m_bUsingSqlite)
                {
                    DeleteVariableAccdb();
                }
                else
                {
                    DeleteVariableSqlite();
                }
                this.btnEconDetailsCancel.PerformClick();
            }
        }

        private void DeleteVariableSqlite()
        {
            try
            {
                string strTableName = Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName;
                if (lstVariables.SelectedItems[0].SubItems[2].Text.Trim().Equals(VARIABLE_ECON))
                {
                    strTableName = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                }
                
                // Delete entries from configuration database
                m_oDataMgr.m_strSQL = "DELETE FROM " + strTableName +
                                      " WHERE calculated_variables_id = " + m_intCurVar;
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                m_oDataMgr.m_strSQL = "DELETE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                                      " WHERE ID = " + m_intCurVar;
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                m_oDataMgr.m_Transaction.Commit();
            }
            catch (SQLiteException errSQLite)
            {
                m_intError = -1;
                MessageBox.Show(errSQLite.Message);
            }
            catch (Exception caught)
            {
                m_intError = -1;
                MessageBox.Show(caught.Message);
            }
            finally
            {
                m_oDataMgr.ResetTransactionObjectToDataAdapterArray();
            }
            // Update UI
            this.loadLstVariablesSqlite();
        }

        private void InitializeOleDbTransactionCommands()
        {
            this.m_oAdo.m_strSQL = "SELECT calculated_variables_id, rxcycle, weight" +
                                   " FROM " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName +
                                   " WHERE calculated_variables_id = " + m_intCurVar + " ;";
            
            //initialize the transaction object with the connection
            this.m_oAdo.m_OleDbTransaction = this.m_oAdo.m_OleDbConnection.BeginTransaction();

            this.m_oAdo.ConfigureDataAdapterInsertCommand(this.m_oAdo.m_OleDbConnection,
                this.m_oAdo.m_OleDbDataAdapter,
                this.m_oAdo.m_OleDbTransaction,
                this.m_oAdo.m_strSQL,
                Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName);

            //Do I need to do this again? It's the same SQL as above
            //this.m_oAdo.m_strSQL = "select fvs_variant, spcd,fvs_input_spcd,common_name,genus,species,comments from " + m_oQueries.m_oFvs.m_strTreeSpcTable + " order by fvs_variant, spcd;";
            this.m_oAdo.ConfigureDataAdapterUpdateCommand(this.m_oAdo.m_OleDbConnection,
                this.m_oAdo.m_OleDbDataAdapter,
                this.m_oAdo.m_OleDbTransaction,
                this.m_oAdo.m_strSQL, "select calculated_variables_id, rxcycle from " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName,
                Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName);
        }

        private void BtnSaveEcon_Click(object sender, EventArgs e)
        {
            this.val_data_econ();
            if (this.m_intError == 0)
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "BtnSaveEcon_Click: Save weighted econ variable " + lblEconVariableName.Text + "\r\n");
                }

                this.enableEconVariableUc(false);
                this.BtnDeleteEconVariable.Enabled = false;
                this.BtnSaveEcon.Visible = true;
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                    frmMain.g_oFrmMain.WindowState,
                    frmMain.g_oFrmMain.Left,
                    frmMain.g_oFrmMain.Height,
                    frmMain.g_oFrmMain.Width,
                    frmMain.g_oFrmMain.Top);

                frmMain.g_sbpInfo.Text = "Saving scenario rule definitions...Stand by";
                if (! m_bUsingSqlite)
                {
                    //Save associated configuration records
                    savevalues("ECON");

                    //Reload the main grid
                    this.loadLstVariables();
                }
                else
                {
                    //Save associated configuration records
                    savevalues_sqlite("ECON");

                    //Reload the main grid
                    this.loadLstVariablesSqlite();
                }

                
                frmMain.g_sbpInfo.Text = "Ready";
                frmMain.g_oFrmMain.DeactivateStandByAnimation();

                MessageBox.Show("Economic variable properties saved! Click Cancel to return to the main Calculated Variables page", "FIA Biosum");

            }
        }

        private int GetNextId()
        {
            // GENERATE NEW ID NUMBER; ADD ONE TO HIGHEST EXISTING ID
            int intId = -1;
            foreach (ListViewItem oItem in this.lstVariables.Items)
            {
                int intTestId = Convert.ToInt32(oItem.SubItems[3].Text.Trim());
                if (intTestId > intId)
                    intId = intTestId;
            }
            intId = intId + 1;
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Selected new variable id: " + intId + " \r\n\r\n");
            }
            return intId;
        }

        private void lstFVSTables_GotFocus(Object sender, EventArgs e)
        {

            if (b_FVSTableEnabled == false)
            {
                // Put the focus elsewhere so the user can't change the selected index
                label7.Focus();
            }

        }

        private void lstFVSFields_GotFocus(Object sender, EventArgs e)
        {

            if (b_FVSTableEnabled == false)
            {
                // Put the focus elsewhere so the user can't change the selected index
                label7.Focus();
            }

        }

        private void BtnFvsImport_Click(object sender, EventArgs e)
        {
            if (BtnFvsImport.Text.ToUpper().Equals("IMPORT WEIGHTS"))
            {
                bool bFailed = false;
            System.Collections.Generic.IList<System.Collections.Generic.IList<Object>> lstTable = loadWeightsFromFile(out bFailed);
            if (lstTable.Count != 8)
            {
                bFailed = true;
            }
            if (bFailed == true)
            {
                MessageBox.Show("This file does not contain weights in the correct format and cannot be loaded!!", "FIA Biosum");
                return;
            }
            foreach (System.Collections.Generic.IList<Object> lstRow in lstTable)
            {
                for (int i=0; i < lstTable.Count; i++)
                {
                    int intRxCycle = Convert.ToInt16(this.m_dg[i, 0]);
                    string strPost = this.m_dg[i, 1].ToString().Trim();
                    if ((intRxCycle == Convert.ToInt16(lstRow[0])) &&                       // match rxCycle
                        (this.m_dg[i, 1].ToString().Trim().Equals(lstRow[1].ToString())))   // match PRE or POST
                    {
                        this.m_dg[i,3] = lstRow[2];
                        break;
                    }
                }
            }
                this.m_dg.SetDataBinding(this.m_dv, "");
                this.m_dg.Update();
                this.SumWeights(false);
            }
            else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Title = "Select Export Text File Name",

                    CheckPathExists = true,
                    OverwritePrompt = true,

                    DefaultExt = "txt",
                    Filter = "txt files (*.txt)|*.txt",
                    FilterIndex = 2,
                    RestoreDirectory = true,
                    
                };

                string strPath = "";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (!String.IsNullOrEmpty(saveFileDialog1.FileName))
                    {
                        strPath = saveFileDialog1.FileName;
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
                Int16 intLines = 8;
                string[] arrLines = new string[intLines];
                for (int i = 0; i < intLines; i++)
                {
                    string strLine = this.m_dg[i, 0].ToString().Trim() + ",";
                    strLine = strLine + this.m_dg[i, 1].ToString().Trim() + ",";
                    strLine = strLine + this.m_dg[i, 3].ToString().Trim();
                    arrLines[i] = strLine;
                }
                System.IO.File.WriteAllLines(strPath, arrLines);
                System.Windows.MessageBox.Show("Weights successfully saved!!", "FIA Biosum");
            }
        }

        private System.Collections.Generic.IList<System.Collections.Generic.IList<Object>> loadWeightsFromFile(out bool bFailed)
        {
            string strWeightsFile = "";
            var OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Title = "Text File With FVS variable weights";
            OpenFileDialog1.Filter = "Text File (*.TXT) |*.txt";
            var result = OpenFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (OpenFileDialog1.FileName.Trim().Length > 0)
                {
                    strWeightsFile = OpenFileDialog1.FileName.Trim();
                }
                OpenFileDialog1 = null;
            }
            System.Collections.Generic.IList<System.Collections.Generic.IList<Object>> lstTable =
                new System.Collections.Generic.List<System.Collections.Generic.IList<Object>>();
            bFailed = false;
            if (!String.IsNullOrEmpty(strWeightsFile))
            {
                //Open the file with a stream reader.
                using (System.IO.StreamReader s = new System.IO.StreamReader(strWeightsFile, System.Text.Encoding.Default))
                {
                    string strNextLine = null;
                    while ((strNextLine = s.ReadLine()) != null)
                    {
                        if (!String.IsNullOrEmpty(strNextLine))
                        {
                            System.Collections.Generic.IList<Object> lstNextLine = new System.Collections.Generic.List<Object>();
                            string[] strPieces = strNextLine.Split(',');
                            if (strPieces.Length == 3)
                            {
                                // This is an FVS variable
                                int intCycle = -1;
                                bool bSuccess = Int32.TryParse(strPieces[0], out intCycle);
                                if (bSuccess == false)
                                {
                                    bFailed = true;
                                    break;
                                }
                                string strPrePost = "";
                                if (!String.IsNullOrEmpty(strPieces[1].Trim()))
                                {
                                    strPrePost = strPieces[1].Trim().ToUpper();
                                    if (!strPrePost.Equals("PRE") &&
                                        !strPrePost.Equals("POST"))
                                    {
                                        bFailed = true;
                                        break;
                                    }
                                }
                                double dblWeight = -1.0F;
                                bSuccess = Double.TryParse(strPieces[2], out dblWeight);
                                if (bSuccess == false)
                                {
                                    bFailed = true;
                                    break;
                                }
                                lstNextLine.Add(intCycle);
                                lstNextLine.Add(strPrePost);
                                lstNextLine.Add(dblWeight);
                                lstTable.Add(lstNextLine);
                            }
                            else if (strPieces.Length == 2)
                            {
                                // This is an Economic variable
                                int intCycle = -1;
                                bool bSuccess = Int32.TryParse(strPieces[0], out intCycle);
                                if (bSuccess == false)
                                {
                                    bFailed = true;
                                    break;
                                }
                                double dblWeight = -1.0F;
                                bSuccess = Double.TryParse(strPieces[1], out dblWeight);
                                if (bSuccess == false)
                                {
                                    bFailed = true;
                                    break;
                                }
                                lstNextLine.Add(intCycle);
                                lstNextLine.Add(dblWeight);
                                lstTable.Add(lstNextLine);
                            }
                            else
                            {
                                bFailed = true;
                                break;
                            }

                        }
                    }
                }
            }
            return lstTable;
        }

        private void BtnEconImport_Click(object sender, EventArgs e)
        {
            if (BtnEconImport.Text.ToUpper().Equals("IMPORT WEIGHTS"))
            {
                bool bFailed = false;
                System.Collections.Generic.IList<System.Collections.Generic.IList<Object>> lstTable = loadWeightsFromFile(out bFailed);
                if (lstTable.Count != 4)
                {
                    bFailed = true;
                }
                if (bFailed == true)
                {
                    MessageBox.Show("This file does not contain weights in the correct format and cannot be loaded!!", "FIA Biosum");
                    return;
                }
                foreach (System.Collections.Generic.IList<Object> lstRow in lstTable)
                {
                    for (int i = 0; i < lstTable.Count; i++)
                    {
                        int intRxCycle = Convert.ToInt16(this.m_dgEcon[i, 0]);
                        if (intRxCycle == Convert.ToInt16(lstRow[0]))
                        {
                            this.m_dgEcon[i, 1] = lstRow[1];
                            break;
                        }
                    }
                }
                //this.m_dgEcon.SetDataBinding(this.m_econ_dv, "");
                //this.m_dgEcon.Update();
                this.SumWeights(true);
            }
            else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog
                {
                    Title = "Select Export Text File Name",

                    CheckPathExists = true,
                    OverwritePrompt = true,

                    DefaultExt = "txt",
                    Filter = "txt files (*.txt)|*.txt",
                    FilterIndex = 2,
                    RestoreDirectory = true,
                    
                };

                string strPath = "";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (!String.IsNullOrEmpty(saveFileDialog1.FileName))
                    {
                        strPath = saveFileDialog1.FileName;
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
                Int16 intLines = 4;
                string[] arrLines = new string[intLines];
                for (int i = 0; i < intLines; i++)
                {
                    string strLine = this.m_dgEcon[i, 0].ToString().Trim() + ",";
                    strLine = strLine + this.m_dgEcon[i, 1].ToString().Trim();
                    arrLines[i] = strLine;
                }
                System.IO.File.WriteAllLines(strPath, arrLines);
                System.Windows.MessageBox.Show("Weights successfully saved!!", "FIA Biosum");
            }
          }

        private void BtnRecalculateAll_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("The Recalculate All button overwrites the existing FVS weighted variable tables. " +
                                               "This process cannot be undone and may take several minutes. Do you wish " +
                                               "to continue?", "FIA Biosum", MessageBoxButtons.YesNo);
            if (res != DialogResult.Yes)
            {
                return;
            }
            // assemble the path for the backup database
            string strDbName = System.IO.Path.GetFileNameWithoutExtension(Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile);
            string strDbFolder = System.IO.Path.GetDirectoryName(Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile);
            string strBackupAccdb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + strDbFolder + "\\" + strDbName + "_backup.accdb";
            System.IO.File.Copy(frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile,
                strBackupAccdb, true);
            RecalculateCalculatedVariables_Start();
        }

        private void RecalculateWeightedVariables_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            m_intError = 0;
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();

            try
            {
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "RecalculateCalculatedVariables_Process: BEGIN \r\n");
                }
                
                //progress bar 1: single process
                SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);
                //progress bar 2: overall progress
                SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);              

                UpdateProgressBar2(0);

                string strPrePostWeightedAccdb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                    "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
                oAdo.OpenConnection(oAdo.getMDBConnString(strPrePostWeightedAccdb, "", ""));
                string[] arrTableNames = oAdo.getTableNames(oAdo.m_OleDbConnection);
                var counter1 = 5;
                var counter2 = 10;
                UpdateProgressBar2(counter2);
                UpdateProgressBar1("Dropping weighted variable tables", counter1);

                foreach (var strTableName in arrTableNames)
                {
                    if (!string.IsNullOrEmpty(strTableName))
                    {
                        oAdo.m_strSQL = "DROP TABLE " + strTableName;
                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, oAdo.m_strSQL + "\r\n\r\n");
                        }
                    }
                }
                int counter2Interval = 5;
                counter1 = counter1 + counter1Interval;
                counter2 = counter2 + counter2Interval;

                // Get list of variables to recalculate
                UpdateProgressBar1("Querying database for calculated variables", counter1);
                IDictionary<int, string> dictFvsWeightedVariables = new Dictionary<int,string>();
                IList<string> lstTables = new List<string>();
                if (!m_bUsingSqlite)
                {
                    oAdo.OpenConnection(oAdo.getMDBConnString(m_strCalculatedVariablesAccdb, "", ""));
                    oAdo.m_strSQL = "SELECT ID, VARIABLE_SOURCE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " WHERE VARIABLE_TYPE = 'FVS' ORDER BY VARIABLE_SOURCE";
                    oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, oAdo.m_strSQL + "\r\n\r\n");
                    }

                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            int key = Convert.ToInt32(oAdo.m_OleDbDataReader["ID"]);
                            if (oAdo.m_OleDbDataReader["VARIABLE_SOURCE"] != System.DBNull.Value && !dictFvsWeightedVariables.Keys.Contains(key))
                            {
                                dictFvsWeightedVariables.Add(key, Convert.ToString(oAdo.m_OleDbDataReader["VARIABLE_SOURCE"]));
                                // Count the number of tables so we know how to set up the step progressor
                                string[] strPieces = dictFvsWeightedVariables[key].Split('.');
                                if (strPieces.Length == 2)
                                {
                                    if (!lstTables.Contains(strPieces[0]))
                                    {
                                        lstTables.Add(strPieces[0]);
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    m_oDataMgr.m_strSQL = "SELECT ID, VARIABLE_SOURCE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " WHERE VARIABLE_TYPE = 'FVS' ORDER BY VARIABLE_SOURCE";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + "\r\n\r\n");
                    }
                    m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    if (m_oDataMgr.m_intError == 0 && m_oDataMgr.m_DataReader.HasRows)
                    {
                        while (m_oDataMgr.m_DataReader.Read())
                        {                          
                            int key = Convert.ToInt32(m_oDataMgr.m_DataReader["ID"]);
                            if (m_oDataMgr.m_DataReader["VARIABLE_SOURCE"] != System.DBNull.Value && !dictFvsWeightedVariables.Keys.Contains(key))
                            {
                                dictFvsWeightedVariables.Add(key, Convert.ToString(m_oDataMgr.m_DataReader["VARIABLE_SOURCE"]));
                                // Count the number of tables so we know how to set up the step progressor
                                string[] strPieces = dictFvsWeightedVariables[key].Split('.');
                                if (strPieces.Length == 2)
                                {
                                    if (!lstTables.Contains(strPieces[0]))
                                    {
                                        lstTables.Add(strPieces[0]);
                                    }
                                }
                            }
                        }
                    }
                }

                if (lstTables.Count > 0)
                {
                    // Reset interval for counter 2 based on number of tables
                    counter2Interval = 70 / (lstTables.Count);
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Stored FVS variable information in memory \r\n\r\n");
                }

                // Reset counter1 interval based on number of variables
                counter1Interval = 40 / dictFvsWeightedVariables.Keys.Count;

                string strCurrentDatabase = "";
                string strWeightsByRxCyclePreTable = "WEIGHTS_BY_RX_CYCLE_PRE";
                string strWeightsByRxCyclePostTable = "WEIGHTS_BY_RX_CYCLE_POST";
                string strWeightsByRxPkgPreTable = "WEIGHTS_BY_RXPACKAGE_PRE";
                string strWeightsByRxPkgPostTable = "WEIGHTS_BY_RXPACKAGE_POST";

                //create and set temporary mdb file
                string strDestinationLinkDir = this.m_oEnv.strTempDir;
                string strRecalculateAccdb = m_oUtils.getRandomFile(strDestinationLinkDir, "accdb");
                oDao.CreateMDB(strRecalculateAccdb);
                foreach (var keyId in dictFvsWeightedVariables.Keys)
                {
                    string[] strArray = frmMain.g_oUtils.ConvertListToArray(dictFvsWeightedVariables[keyId], ".");
                    string strDatabase = "";
                    string strColumn = "";
                    if (strArray.Length == 2)
                    {
                        if (strArray[0].Trim().Length > 0)
                        {
                            strDatabase = strArray[0].Trim();
                        }
                        if (strArray[1].Trim().Length > 0)
                        {
                            strColumn = strArray[1].Trim();
                        }

                        string strSourcePreTable = "PRE_" + strDatabase;
                        string strSourcePostTable = "POST_" + strDatabase;
                        if (!strDatabase.Equals(strCurrentDatabase))
                        {
                            counter2 = counter2 + counter2Interval;
                            UpdateProgressBar2(counter2);
                            // We need to create the tables
                            string strSourceDatabaseName = "PREPOST_" + strDatabase + ".ACCDB";
                            string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                                "\\fvs\\db\\" + strSourceDatabaseName;
                            string strTargetPreTable = "PRE_" + strDatabase + "_WEIGHTED";
                            string strTargetPostTable = "POST_" + strDatabase + "_WEIGHTED";

                            UpdateProgressBar1("Creating tables for " + strDatabase, counter1);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Creating tables for " + strDatabase + " \r\n\r\n");
                            }
                            counter1 = counter1 + 3;

                            //Link to source FVS tables in temp .mdb if they don't exist from a previous run
                            if (!oDao.TableExists(strRecalculateAccdb, strSourcePreTable))
                            {
                                oDao.CreateTableLink(strRecalculateAccdb, strSourcePreTable, strFvsPrePostDb, strSourcePreTable);
                            }
                            if (!oDao.TableExists(strRecalculateAccdb, strSourcePostTable))
                            {
                                oDao.CreateTableLink(strRecalculateAccdb, strSourcePostTable, strFvsPrePostDb, strSourcePostTable);
                            }
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Created table link to " + strFvsPrePostDb + "\r\n\r\n");
                            }

                            //Link source tables to output database
                            oDao.CreateTableLink(strPrePostWeightedAccdb, strSourcePreTable, strFvsPrePostDb, strSourcePreTable);
                            oDao.CreateTableLink(strPrePostWeightedAccdb, strSourcePostTable, strFvsPrePostDb, strSourcePostTable);

                            oAdo.OpenConnection(oAdo.getMDBConnString(strPrePostWeightedAccdb, "", ""));
                            // FVS creates a record for
                            // each condition for each cycle regardless of whether there is activity
                            oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant " +
                                      "INTO " + strTargetPreTable +
                                      " FROM " + strSourcePreTable;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Creating final pre/post tables. They did not already exist \r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + oAdo.m_strSQL + "\r\n\r\n");
                            }

                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant " +
                                                  "INTO " + strTargetPostTable +
                                                  " FROM " + strSourcePostTable;
                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

                            oDao.DeleteTableFromMDB(strPrePostWeightedAccdb, strSourcePreTable);
                            oDao.DeleteTableFromMDB(strPrePostWeightedAccdb, strSourcePostTable);

                            strCurrentDatabase = strDatabase;
                        }

                        //Drop strWeightsByRxCyclePreTable if it exists so we can recreate it
                        if (oDao.TableExists(strRecalculateAccdb, strWeightsByRxCyclePreTable))
                        {
                            oDao.DeleteTableFromMDB(strRecalculateAccdb, strWeightsByRxCyclePreTable);
                        }
                        //Drop strWeightsByRxCyclePostTable if it exists so we can recreate it
                        if (oDao.TableExists(strRecalculateAccdb, strWeightsByRxCyclePostTable))
                        {
                            oDao.DeleteTableFromMDB(strRecalculateAccdb, strWeightsByRxCyclePostTable);
                        }
                        //Drop strWeightsByRxPkgPreTable if it exists so we can recreate it
                        if (oDao.TableExists(strRecalculateAccdb, strWeightsByRxPkgPreTable))
                        {
                            oDao.DeleteTableFromMDB(strRecalculateAccdb, strWeightsByRxPkgPreTable);
                        }
                        //Drop strWeightsByRxPkgPostTable if it exists so we can recreate it
                        if (oDao.TableExists(strRecalculateAccdb, strWeightsByRxPkgPostTable))
                        {
                            oDao.DeleteTableFromMDB(strRecalculateAccdb, strWeightsByRxPkgPostTable);
                        }
                        //Drop strWeightsByRxCyclePostTable if it exists so we can recreate it
                        if (oDao.TableExists(strRecalculateAccdb, strWeightsByRxCyclePostTable))
                        {
                            oDao.DeleteTableFromMDB(strRecalculateAccdb, strWeightsByRxCyclePostTable);
                        }

                        IList<string[]> lstWeights = new List<string[]>();
                        string strVariableName = "";
                        string strBaselinePackage = "";

                        if (! m_bUsingSqlite)
                        {
                            // open connection to optimizer_definitions.accdb to query for weights
                            oAdo.OpenConnection(oAdo.getMDBConnString(m_strCalculatedVariablesAccdb, "", ""));
                            oAdo.m_strSQL = "SELECT ID, VARIABLE_NAME, BASELINE_RXPACKAGE, " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".* " +
                                "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " INNER JOIN " +
                                Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + " ON " +
                                Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".ID = " +
                                Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".CALCULATED_VARIABLES_ID " +
                                "WHERE CALCULATED_VARIABLES_ID = " + keyId;
                            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Querying weights for next variable \r\n\r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, oAdo.m_strSQL + "\r\n\r\n");
                            }

                            if (oAdo.m_OleDbDataReader.HasRows)
                            {
                                while (oAdo.m_OleDbDataReader.Read())
                                {
                                    strVariableName = Convert.ToString(oAdo.m_OleDbDataReader["VARIABLE_NAME"]).Trim();
                                    strBaselinePackage = Convert.ToString(oAdo.m_OleDbDataReader["BASELINE_RXPACKAGE"]).Trim();
                                    for (int rxCycle = 1; rxCycle < 5; rxCycle++)
                                    {
                                        // PRE VALUES
                                        string[] strPreRow = new string[3];
                                        string strFieldName = "weight_" + rxCycle + "_pre";
                                        strPreRow[idxRxCycle] = Convert.ToString(rxCycle);
                                        strPreRow[idxPreOrPost] = "PRE";
                                        strPreRow[idxWeight] = Convert.ToString(oAdo.m_OleDbDataReader[strFieldName]);
                                        lstWeights.Add(strPreRow);

                                        // POST VALUES
                                        string[] strPostRow = new string[3];
                                        strFieldName = "weight_" + rxCycle + "_post";
                                        strPostRow[idxRxCycle] = Convert.ToString(rxCycle);
                                        strPostRow[idxPreOrPost] = "POST";
                                        strPostRow[idxWeight] = Convert.ToString(oAdo.m_OleDbDataReader[strFieldName]);
                                        lstWeights.Add(strPostRow);
                                    }
                                }
                            }
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Saved values loaded into data structure for calculation \r\n\r\n");
                            }
                        }
                        else
                        {
                            m_oDataMgr.m_strSQL = "SELECT ID, VARIABLE_NAME, BASELINE_RXPACKAGE, " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".* " +
                                "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " INNER JOIN " +
                                Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + " ON " +
                                Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".ID = " +
                                Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".CALCULATED_VARIABLES_ID " +
                                "WHERE CALCULATED_VARIABLES_ID = " + keyId;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, m_oDataMgr.m_strSQL + "\r\n\r\n");
                            }
                            m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                            if (m_oDataMgr.m_intError == 0 && m_oDataMgr.m_DataReader.HasRows)
                            {
                                while (m_oDataMgr.m_DataReader.Read())
                                {
                                    strVariableName = Convert.ToString(m_oDataMgr.m_DataReader["VARIABLE_NAME"]).Trim();
                                    strBaselinePackage = Convert.ToString(m_oDataMgr.m_DataReader["BASELINE_RXPACKAGE"]).Trim();
                                    for (int rxCycle = 1; rxCycle < 5; rxCycle++)
                                    {
                                        // PRE VALUES
                                        string[] strPreRow = new string[3];
                                        string strFieldName = "weight_" + rxCycle + "_pre";
                                        strPreRow[idxRxCycle] = Convert.ToString(rxCycle);
                                        strPreRow[idxPreOrPost] = "PRE";
                                        strPreRow[idxWeight] = Convert.ToString(m_oDataMgr.m_DataReader[strFieldName]);
                                        lstWeights.Add(strPreRow);

                                        // POST VALUES
                                        string[] strPostRow = new string[3];
                                        strFieldName = "weight_" + rxCycle + "_post";
                                        strPostRow[idxRxCycle] = Convert.ToString(rxCycle);
                                        strPostRow[idxPreOrPost] = "POST";
                                        strPostRow[idxWeight] = Convert.ToString(m_oDataMgr.m_DataReader[strFieldName]);
                                        lstWeights.Add(strPostRow);
                                    }
                                }
                            }
                        }

                        string strCalculateConn = m_oAdo.getMDBConnString(strRecalculateAccdb, "", "");
                        CalculateVariable(oDao, strCalculateConn, strWeightsByRxPkgPreTable, strSourcePreTable,
                            strWeightsByRxPkgPostTable, strSourcePostTable, strVariableName, strColumn, strBaselinePackage,
                            lstWeights,false, ref counter1);
                    }
                }

                UpdateProgressBar1("Variables Recalculated!!", 100);
                UpdateProgressBar2(100);

                if (oAdo != null)
                {
                    if (oAdo.m_DataSet != null)
                    {
                        oAdo.m_DataSet.Clear();
                        oAdo.m_DataSet.Dispose();
                    }
                    oAdo = null;
                }
                if (oDao != null)
                {
                    oDao.m_DaoWorkspace.Close();
                    oDao.m_DaoWorkspace = null;
                    oDao = null;
                }

                MessageBox.Show("Variables Recalculated!!", "FIA Biosum");
                RecalculateCalculatedVariables_Finish();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (oAdo != null)
                {
                    if (oAdo.m_DataSet != null)
                    {
                        oAdo.m_DataSet.Clear();
                        oAdo.m_DataSet.Dispose();
                    }
                    oAdo = null;
                }
                ThreadCleanUp();
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_optimizer_scenario_calculated_variables:RecalculateCalculatedVariables_Process  \n" +
                                "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                frmMain.g_oDelegate.SetControlPropertyValue((frmDialog)ParentForm, "Enabled", true);
                m_intError = -1;
            }

            RecalculateCalculatedVariables_Finish();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }

        private void RecalculateCalculatedVariables_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            StartTherm("2", "Recalculate Calculated Variable Tables");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(RecalculateWeightedVariables_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void RecalculateCalculatedVariables_Finish()
        {
            if (m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                m_frmTherm = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((frmDialog) ParentForm, "Enabled", true);
            ((frmDialog)ParentForm).MinimizeMainForm = false;
        }

        private void CalculateVariable(dao_data_access oDao, string strCalculateConn, string strWeightsByRxPkgPreTable, 
                                       string strSourcePreTable, string strWeightsByRxPkgPostTable, string strSourcePostTable, 
                                       string strVariableName, string strFieldName, string strBaselinePkg, IList<string[]> lstRows,
                                       bool bNewTables, ref int counter1)
        {
            if (counter1 > 0)
            {
                counter1 = counter1 + counter1Interval;
                UpdateProgressBar1("Calculating values for " + strVariableName, counter1);
            }
            string strWeightsByRxCyclePreTable = "WEIGHTS_BY_RX_CYCLE_PRE";
            string strWeightsByRxCyclePostTable = "WEIGHTS_BY_RX_CYCLE_POST";
            string strTargetPreTable = strSourcePreTable + "_WEIGHTED";
            string strTargetPostTable = strSourcePostTable + "_WEIGHTED";
            string strPrePostWeightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            string strRefreshAccdb = "";
            using (var calculateConn = new OleDbConnection(strCalculateConn))
            {
                calculateConn.Open();
                strRefreshAccdb = calculateConn.DataSource;

                m_oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CDbl(0) as " +
                                  strVariableName + " " +
                                  "INTO " + strWeightsByRxCyclePreTable +
                                  " FROM " + strSourcePreTable;
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Create temporary table for weights by rx and rxcycle\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Sql: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }

                m_oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CDbl(0) as " +
                              strVariableName + " " +
                              "INTO " + strWeightsByRxCyclePostTable +
                              " FROM " + strSourcePostTable;
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);

                //Calculate values for each row in table
                double dblWeight = -1;
                string strWeight = "";
                string strRxCycle = "";
                string strPrePost = "";
                string strSourceTableName = "";
                string strTargetTableName = "";
                foreach (string[] strRow in lstRows)
                {
                    strRxCycle = strRow[idxRxCycle];
                    strWeight = strRow[idxWeight];
                    strPrePost = strRow[idxPreOrPost];
                    if (strPrePost.Equals("PRE"))
                    {
                        strTargetTableName = strWeightsByRxCyclePreTable;
                        strSourceTableName = strSourcePreTable;
                    }
                    else
                    {
                        strTargetTableName = strWeightsByRxCyclePostTable;
                        strSourceTableName = strSourcePostTable;
                    }
                    if (Double.TryParse(strWeight, out dblWeight))
                    {
                        // Apply weights to each cycle
                        m_oAdo.m_strSQL = "UPDATE " + strTargetTableName + " w " +
                                      "INNER JOIN " + strSourceTableName + " p " +
                                      "ON w.biosum_cond_id = p.biosum_cond_id " +
                                      "AND w.rxpackage = p.rxpackage AND w.rx = p.rx " +
                                      "AND w.rxcycle = p.rxcycle AND w.fvs_variant = p.fvs_variant " +
                                      "SET " + strVariableName + " = " +
                                      strFieldName + " * " + dblWeight + " " +
                                      "WHERE w.rxcycle = '" + strRxCycle + "'";
                        m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Calculate values for each row in m_dg \r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + m_oAdo.m_strSQL + "\r\n\r\n");
                        }

                    }
                }

                // Sum by rxpackage across cycles
                m_oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, \"0\" as [rx], sum(" + strVariableName + ") as [sum_pre] " +
                              "into " + strWeightsByRxPkgPreTable + " " +
                              "from " + strWeightsByRxCyclePreTable + " " +
                              "group by biosum_cond_id, rxpackage";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Sum by rxpackage across cycles \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
                // Update rx with rx from cycle 1
                m_oAdo.m_strSQL = "UPDATE " + strWeightsByRxPkgPreTable + " w " +
                              "INNER JOIN " + strWeightsByRxCyclePreTable + " r ON w.biosum_cond_id = r.biosum_cond_id " +
                              "AND w.rxpackage = r.rxpackage " +
                              "SET w.rx = r.rx " +
                              "WHERE r.rxcycle = '1'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Set rx to rx from cycle 1 \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = "SELECT biosum_cond_id, rxpackage, \"0\" as [rx], sum(" + strVariableName + ") as [sum_post] " +
                              "into " + strWeightsByRxPkgPostTable + " " +
                              "from " + strWeightsByRxCyclePostTable + " " +
                              "group by biosum_cond_id, rxpackage";
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
                // Update rx with rx from cycle 1
                m_oAdo.m_strSQL = "UPDATE " + strWeightsByRxPkgPostTable + " w " +
                              "INNER JOIN " + strWeightsByRxCyclePostTable + " r ON w.biosum_cond_id = r.biosum_cond_id " +
                              "AND w.rxpackage = r.rxpackage " +
                              "SET w.rx = r.rx " +
                              "WHERE r.rxcycle = '1'";
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
            }   // end using

            //Switch connection to the final storage location and prepare the tables to receive the output
            string strPrePostConn = m_oAdo.getMDBConnString(strPrePostWeightedDb, "", "");
            using (var prePostConn = new OleDbConnection(strPrePostConn))
            {
                prePostConn.Open();
                //Check to see if columns exists, they shouldn't, warn that values will be overwritten
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Add receiving columns to pre/post tables if they don't exist \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Warning message if they do! " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                if (m_oAdo.ColumnExist(prePostConn, strTargetPreTable, strVariableName))
                {
                    if (bNewTables == false)
                        MessageBox.Show("Values for " + strVariableName + " were previously calculated! " +
                                        "They will be overwritten!", "FIA Biosum");
                }
                else
                {
                    m_oAdo.AddColumn(prePostConn, strTargetPreTable,
                        strVariableName, "DOUBLE", "");
                    m_oAdo.AddColumn(prePostConn, strTargetPostTable,
                        strVariableName, "DOUBLE", "");
                }
            }

            //Link receiving tables to temporary database
            if (!oDao.TableExists(strRefreshAccdb, strTargetPreTable))
            {
                oDao.CreateTableLink(strRefreshAccdb, strTargetPreTable, strPrePostWeightedDb, strTargetPreTable);
            }
            if (!oDao.TableExists(strRefreshAccdb, strTargetPostTable))
            {
                oDao.CreateTableLink(strRefreshAccdb, strTargetPostTable, strPrePostWeightedDb, strTargetPostTable);
            }

            //Switch connection to temporary database
            using (var calculateConn = new OleDbConnection(strCalculateConn))
            {
                if (counter1 > 0)
                {
                    counter1 = counter1 + counter1Interval;
                    UpdateProgressBar1("Calculating values for " + strVariableName, counter1);
                }
                calculateConn.Open();
                m_oAdo.m_strSQL = "UPDATE (" + strWeightsByRxPkgPostTable + " pt " +
                                  "INNER JOIN " + strWeightsByRxPkgPreTable + " pe " +
                                  "ON (pt.biosum_cond_id = pe.biosum_cond_id)) " +
                                  "INNER JOIN " + strTargetPreTable + " f " +
                                  "ON (pe.biosum_cond_id = f.biosum_cond_id) " +
                                  "SET " + strVariableName + " = sum_pre + sum_post " +
                                  "WHERE pt.rxpackage = '" + strBaselinePkg +
                                  "' and pe.rxpackage = '" + strBaselinePkg +
                                  "' and f.rxcycle = '1'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated weighted PRE table with weighted totals from baseline scenario \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
                m_oAdo.m_strSQL = "UPDATE (" + strWeightsByRxPkgPostTable + " pt " +
                                  "INNER JOIN " + strWeightsByRxPkgPreTable + " pe " +
                                  "ON (pt.rxpackage = pe.rxpackage) AND (pt.biosum_cond_id = pe.biosum_cond_id)) " +
                                  "INNER JOIN " + strTargetPostTable + " f ON (pe.rxpackage = f.rxpackage) AND (pe.biosum_cond_id = f.biosum_cond_id) " +
                                  "SET " + strVariableName + " = sum_pre + sum_post " +
                                  "WHERE f.rxcycle = '1'";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated weighted POST table with weighted totals from baseline scenario \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + m_oAdo.m_strSQL + "\r\n\r\n");
                }
                m_oAdo.SqlNonQuery(calculateConn, m_oAdo.m_strSQL);
            }
        }

        private void SQLiteConnect()
        {
            try
            {
                if (m_oDataMgr.m_Connection != null && m_oDataMgr.m_Connection.State != ConnectionState.Closed)
                {
                    if (m_oDataMgr.m_DataReader != null) m_oDataMgr.m_DataReader.Dispose();
                    if (m_oDataMgr.m_Command != null) m_oDataMgr.m_Command.Dispose();
                    //if (m_oDataMgr.m_Transaction != null) m_oDataMgr.m_Transaction.Dispose();
                    if (m_oDataMgr.m_DataAdapter != null)
                    {
                        if (m_oDataMgr.m_DataAdapter.SelectCommand != null)
                        {
                            m_oDataMgr.m_DataAdapter.SelectCommand.Dispose();
                        }
                        if (m_oDataMgr.m_DataAdapter.UpdateCommand != null)
                        {
                            m_oDataMgr.m_DataAdapter.UpdateCommand.Dispose();
                        }
                        if (m_oDataMgr.m_DataAdapter.DeleteCommand != null)
                        {
                            m_oDataMgr.m_DataAdapter.DeleteCommand.Dispose();
                        }
                        if (m_oDataMgr.m_DataAdapter.InsertCommand != null)
                        {
                            m_oDataMgr.m_DataAdapter.InsertCommand.Dispose();
                        }
                    }
                    m_oDataMgr.CloseConnection(m_oDataMgr.m_Connection);
                }

                Thread.Sleep(2000); //sleep 5 seconds

                if (m_oDataMgr.m_Connection != null) m_oDataMgr.m_Connection.Dispose();
                m_oDataMgr.OpenConnection(m_oDataMgr.GetConnectionString(m_strCalculatedVariablesDb));
            }
            catch (System.Data.SQLite.SQLiteException errSQLite)
            {
                m_strError = errSQLite.Message;
                m_intError = -1;
            }
            catch (Exception err)
            {
                m_strError = err.Message;
                m_intError = -1;
            }
            finally
            {
                if (m_intError != 0 || m_oDataMgr.m_intError != 0)
                {
                    m_oDataMgr.CloseConnection(m_oDataMgr.m_Connection);
                }
                else
                {
                    if (m_oDataMgr.m_DataSet != null)
                    {
                        m_oDataMgr.m_DataSet.Tables.Clear();
                        m_oDataMgr.m_DataSet.Dispose();
                    }

                    m_oDataMgr.m_DataSet = new System.Data.DataSet();   // Initialize DataSet to avoid null exceptions
                    m_oDataMgr.InitializeDataAdapterArray(TABLE_COUNT);
                }
            }
        }

        private void StartTherm(string p_strNumberOfTherms, string p_strTitle)
        {
            m_frmTherm = new frmTherm((frmDialog)ParentForm, p_strTitle);

            m_frmTherm.Text = p_strTitle;
            m_frmTherm.lblMsg.Text = "";
            m_frmTherm.lblMsg2.Text = "";
            m_frmTherm.Visible = false;
            m_frmTherm.btnCancel.Visible = false;
            m_frmTherm.btnCancel.Enabled = false;
            m_frmTherm.lblMsg.Visible = true;
            m_frmTherm.progressBar1.Minimum = 0;
            m_frmTherm.progressBar1.Visible = true;
            m_frmTherm.progressBar1.Maximum = 10;

            if (p_strNumberOfTherms == "2")
            {
                m_frmTherm.progressBar2.Size = m_frmTherm.progressBar1.Size;
                m_frmTherm.progressBar2.Left = m_frmTherm.progressBar1.Left;
                m_frmTherm.progressBar2.Top =
                    Convert.ToInt32(m_frmTherm.progressBar1.Top + (m_frmTherm.progressBar1.Height * 3));
                m_frmTherm.lblMsg2.Top =
                    m_frmTherm.progressBar2.Top + m_frmTherm.progressBar2.Height + 5;
                m_frmTherm.Height = m_frmTherm.lblMsg2.Top + m_frmTherm.lblMsg2.Height +
                                         m_frmTherm.btnCancel.Height + 50;
                m_frmTherm.btnCancel.Top =
                    m_frmTherm.ClientSize.Height - m_frmTherm.btnCancel.Height - 5;
                m_frmTherm.lblMsg2.Show();
                m_frmTherm.progressBar2.Visible = true;
            }
            m_frmTherm.Width = m_frmTherm.Width + 20;
            m_frmTherm.AbortProcess = false;
            m_frmTherm.Refresh();
            m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            ((frmDialog)ParentForm).Enabled = false;
            m_frmTherm.Visible = true;
        }

        private void ThreadCleanUp()
        {
            try
            {              
                if (m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Dispose");
                    frmMain.g_oDelegate.SetControlPropertyValue((frmDialog)ParentForm, "Enabled", true);

                    m_frmTherm = null;
                }
            }
            catch
            {
            }
        }
        
        private void UpdateProgressBar1(string label, int value)
        {
            SetLabelValue(m_frmTherm.lblMsg, "Text", label);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,
                "Value", value);
        }
        
        private void UpdateProgressBar2(int value)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar2,
                "Value", value);
        }
        
        private void SetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName, int p_intValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)p_oPb, p_strPropertyName,
                (int)p_intValue);
        }

        private void SetLabelValue(System.Windows.Forms.Label p_oLabel, string p_strPropertyName, string p_strValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)p_oLabel, p_strPropertyName,
                p_strValue);
        }

        private void migrate_access_data()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "migrate_access_data \r\n");
            }

            string sourceFile = this.m_oEnv.strAppDir + "\\db\\" + 
                System.IO.Path.GetFileName(Tables.OptimizerDefinitions.DefaultSqliteDbFile);
            System.IO.File.Copy(sourceFile, this.m_strCalculatedVariablesDb);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Copied the database from the application directory \r\n");
            }

            // Check to see if the input SQLite DSN exists and if so, delete so we can add
            ODBCMgr odbcmgr = new ODBCMgr();
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, this.m_strCalculatedVariablesDb);

            // Create temporary database
            string strTempAccdb = m_oUtils.getRandomFile(this.m_oEnv.strTempDir, "accdb");
            dao_data_access oDao = new dao_data_access();
            oDao.CreateMDB(strTempAccdb);

            // Create table links for transferring data
            string targetEcon = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + "_1";
            string targetFvs = Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + "_1";
            string targetVariables = Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + "_1";
            // Link to all tables in source database
            oDao.CreateTableLinks(strTempAccdb, this.m_strCalculatedVariablesAccdb);
            oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName, targetEcon,
                ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, this.m_strCalculatedVariablesDb);
            oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName, targetFvs,
                ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, this.m_strCalculatedVariablesDb);
            oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, targetVariables,
                ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, this.m_strCalculatedVariablesDb);

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Created table links in temporary database \r\n");
            }

            if (m_oAdo == null)
            {
                m_oAdo = new ado_data_access();
            }

            string strCopyConn = m_oAdo.getMDBConnString(strTempAccdb, "", "");
            using (var copyConn = new OleDbConnection(strCopyConn))
            {
                copyConn.Open();
                m_oAdo.m_strSQL = "INSERT INTO " + targetEcon +
                                  " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n");
                }
                m_oAdo.SqlNonQuery(copyConn, this.m_oAdo.m_strSQL);

                m_oAdo.m_strSQL = "INSERT INTO " + targetFvs +
                  " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n");
                }
                m_oAdo.SqlNonQuery(copyConn, this.m_oAdo.m_strSQL);

                m_oAdo.m_strSQL = "INSERT INTO " + targetVariables +
                    " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, m_oAdo.m_strSQL + " \r\n");
                }
                m_oAdo.SqlNonQuery(copyConn, this.m_oAdo.m_strSQL);

            }


            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName);
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }
    }


    public class WeightedAverage_DataGridColoredTextBoxColumn : DataGridTextBoxColumn
    {
        bool m_bEdit = false;
        FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables uc_optimizer_scenario_calculated_variables_1;
        string m_strLastKey = "";
        bool m_bNumericOnly = false;


        public WeightedAverage_DataGridColoredTextBoxColumn(bool bEdit, bool bNumericOnly, 
            FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables p_uc)
        {
            this.m_bEdit = bEdit;
            this.m_bNumericOnly = bNumericOnly;
            this.uc_optimizer_scenario_calculated_variables_1 = p_uc;
            this.TextBox.KeyDown += new KeyEventHandler(TextBox_KeyDown);
            this.TextBox.Leave += new EventHandler(TextBox_Leave);
            this.TextBox.Enter += new EventHandler(TextBox_Enter);
        }


        protected override void Paint(System.Drawing.Graphics g, System.Drawing.Rectangle bounds, System.Windows.Forms.CurrencyManager source, int rowNum, System.Drawing.Brush backBrush, System.Drawing.Brush foreBrush, bool alignToRight)
        {

            // color only the columns that can be edited by the user
            try
            {
                if (this.m_bEdit == true)
                {
                    backBrush = new System.Drawing.Drawing2D.LinearGradientBrush(bounds,
                        Color.FromArgb(255, 200, 200),
                        Color.FromArgb(128, 20, 20),
                        System.Drawing.Drawing2D.LinearGradientMode.BackwardDiagonal);
                    foreBrush = new SolidBrush(Color.White);
                }
            }
            catch { /* empty catch */ }
            finally
            {
                try
                {
                    // make sure the base class gets called to do the drawing with
                    // the possibly changed brushes
                    base.Paint(g, bounds, source, rowNum, backBrush, foreBrush, alignToRight);
                }
                catch
                {
                }
            }
        }
        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("textchange");
        }
        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (this.m_bEdit == true)
            {

                if (this.m_bNumericOnly == true)
                {
                    if (Char.IsDigit((char)e.KeyValue) || (e.KeyCode == Keys.OemPeriod && this.Format.IndexOf(".", 0) >= 0 && this.TextBox.Text.IndexOf(".", 0) < 0))
                    {
                        this.m_strLastKey = Convert.ToString(e.KeyValue);
                        if (this.uc_optimizer_scenario_calculated_variables_1.grpboxDetails.Visible == true)
                        {
                            if (this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled = true;
                        }
                        else
                        {
                            if (this.uc_optimizer_scenario_calculated_variables_1.BtnSaveEcon.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.BtnSaveEcon.Enabled = true;

                        }
                    }
                    else
                    {
                        if (e.KeyCode == Keys.Back)
                        {
                            this.m_strLastKey = Convert.ToString(e.KeyValue);
                            if (this.uc_optimizer_scenario_calculated_variables_1.grpboxDetails.Visible == true)
                            {
                                if (this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled = true;
                            }
                            else
                            {
                                if (this.uc_optimizer_scenario_calculated_variables_1.BtnSaveEcon.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.BtnSaveEcon.Enabled = true;
      
                            }
                        }
                        else
                        {
                            e.Handled = true;
                            SendKeys.Send("{BACKSPACE}");
                        }
                    }

                }
                else
                {
                    this.m_strLastKey = Convert.ToString(e.KeyValue);

                    if (this.uc_optimizer_scenario_calculated_variables_1.grpboxDetails.Visible == true)
                    {
                        if (this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled = true;
                    }
                    else
                    {
                        if (this.uc_optimizer_scenario_calculated_variables_1.BtnSaveEcon.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.BtnSaveEcon.Enabled = true;

                    }
                }


            }



        }
        private void TextBox_Enter(object sender, EventArgs e)
        {
            this.m_strLastKey = "";
        }
        private void TextBox_Leave(object sender, EventArgs e)
        {
            if (this.m_bEdit == true)
            {
                if (this.m_strLastKey.Trim().Length > 0)
                {
                    if (this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled == false) this.uc_optimizer_scenario_calculated_variables_1.btnFvsCalculate.Enabled = true;
                }
            }
        }

    }

}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
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
        public System.Data.DataSet m_DataSet;
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
        private DataMgr m_oDataMgr = new DataMgr();
        private string m_strTempDB;
        private bool m_bUseNegatives;

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
        string m_strCalculatedVariablesDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
            "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;

        const int TABLE_COUNT = 2;
        const int ECON_DETAILS_TABLE = 0;
        const int FVS_DETAILS_TABLE = 1;

        private const int WEIGHT_SUM = 0;
        private const int NULL_COUNT = 1;
        private int intNullThreshold = 4;

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
        private GroupBox grpBoxThreshold;
        private Label lblThresholdExplanation;
        private Label lblThreshold;
        private Button btnSaveThreshold;
        private ComboBox cmbThreshold;
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
            this._frmScenario = new frmOptimizerScenario();

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
            //this.loadvalues();
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
            this.grpBoxThreshold = new System.Windows.Forms.GroupBox();
            this.lblThresholdExplanation = new System.Windows.Forms.Label();
            this.lblThreshold = new System.Windows.Forms.Label();
            this.btnSaveThreshold = new System.Windows.Forms.Button();
            this.cmbThreshold = new System.Windows.Forms.ComboBox();
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
            this.grpBoxThreshold.SuspendLayout();
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
            this.pnlSummary.Controls.Add(this.grpBoxThreshold);
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
            this.lstVariables.Size = new System.Drawing.Size(584, 326);
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
            this.vName.Width = 170;
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
            this.vType.Width = 60;
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
            // grpBoxThreshold
            //
            this.grpBoxThreshold.Controls.Add(this.lblThresholdExplanation);
            this.grpBoxThreshold.Controls.Add(this.lblThreshold);
            this.grpBoxThreshold.Controls.Add(this.btnSaveThreshold);
            this.grpBoxThreshold.Controls.Add(this.cmbThreshold);
            this.grpBoxThreshold.Location = new System.Drawing.Point(620, 18);
            this.grpBoxThreshold.Name = "grpBoxThreshold";
            this.grpBoxThreshold.Size = new System.Drawing.Size(215, 315);
            this.grpBoxThreshold.TabIndex = 75;
            this.grpBoxThreshold.TabStop = false;
            this.grpBoxThreshold.Text = "Null Threshold";
            //
            // lblThreshold
            //
            this.lblThreshold.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblThreshold.Location = new System.Drawing.Point(6, 240);
            this.lblThreshold.Name = "lblThreshold";
            this.lblThreshold.Size = new System.Drawing.Size(95, 22);
            this.lblThreshold.TabIndex = 100;
            this.lblThreshold.Text = "Null Threshold:";
            //
            // cmbThreshold
            //
            this.cmbThreshold.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbThreshold.Location = new System.Drawing.Point(104, 237);
            this.cmbThreshold.Name = "cmbThreshold";
            this.cmbThreshold.Size = new System.Drawing.Size(50, 22);
            this.cmbThreshold.TabIndex = 77;
            //
            // btnSaveThreshold
            //
            this.btnSaveThreshold.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveThreshold.Location = new System.Drawing.Point(6, 268);
            this.btnSaveThreshold.Name = "btnSaveThreshold";
            this.btnSaveThreshold.Size = new System.Drawing.Size(80, 22);
            this.btnSaveThreshold.TabIndex = 92;
            this.btnSaveThreshold.Text = "Save";
            this.btnSaveThreshold.Click += new System.EventHandler(this.btnSaveThreshold_Click);
            //
            // lblThresholdExplanation
            //
            this.lblThresholdExplanation.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblThresholdExplanation.Location = new System.Drawing.Point(6, 24);
            this.lblThresholdExplanation.Name = "lblThresholdExplanation";
            this.lblThresholdExplanation.Size = new System.Drawing.Size(187, 200);
            this.lblThresholdExplanation.TabIndex = 100;
            this.lblThresholdExplanation.Text = "Weighted variables are computed from up to " +
                "8 time points in the simulation, designated as PRE or POST " +
                "for each of the 4 BioSum Cycles. Select a threshold determining the " +
                "maximum number of null value cases for a stand-RxPackage combination, " +
                "above which the weighted value for the variable will be assigned null. " +
                "If the null count is less than or equal to the threshold, " +
                "then an adjusted weight value reflecting only the non-null cases " +
                "will be assigned in the weighted PREPOST table.";
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

            SQLiteConnect();
            // One and only one transaction for this form
            m_oDataMgr.m_Transaction = m_oDataMgr.m_Connection.BeginTransaction();

            loadLstVariables();
            for (int i = 0; i <= 7; i++)
            {
                this.cmbThreshold.Items.Add(i);
            }
            loadnullthreshold();
            
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
            init_m_dg();

            //load datagrid for economic variables
            loadEconVariablesGrid();

            // load listbox for economic variables
            lstEconVariablesList.Items.Clear();
            foreach (string strName in PREFIX_ECON_NAME_ARRAY)
            {
                lstEconVariablesList.Items.Add(strName);
            }

            if (m_oDataMgr == null)
            {
                m_oDataMgr = new DataMgr();
            }
            m_dictFVSTables = m_oOptimizerScenarioTools.LoadFvsTablesAndVariables();
 
            foreach (string strKey in m_dictFVSTables.Keys)
            {
                // 
                if (strKey.IndexOf("_WEIGHTED") < 0)
                {
                    lstFVSTablesList.Items.Add(strKey);
                }
            }

            // Enable the refresh button if we have calculated weighted variables
            string strPrePostWeightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
            if (System.IO.File.Exists(strPrePostWeightedDb))
            {
                using (SQLiteConnection conn = new SQLiteConnection(m_oDataMgr.GetConnectionString(strPrePostWeightedDb)))
                {
                    conn.Open();
                    string[] arrTableNames = m_oDataMgr.getTableNames(conn);
                    if (arrTableNames.Length > 0)
                    BtnRecalculateAll.Enabled = true;
                }
            }

        }

        private void loadm_dg()
        {
            //@ToDo: Using MS Access to query FVS Out tables          
            m_strTempDB = m_oUtils.getRandomFile(this.m_oEnv.strTempDir, "db");
            m_oDataMgr.CreateDbFile(m_strTempDB);
         
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadm_dg: Starting to load FVS calculated variable datagrid \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Temporary database path: " + m_strTempDB + "\r\n\r\n");
            }

            string strFvsPreTableName = "";
            string strFvsPostTableName = "";
            string strFvsPrePostDb = "";
            if (this.lstFVSTablesList.SelectedItems.Count > 0)
            {
                strFvsPreTableName = "PRE_" + Convert.ToString(lstFVSTablesList.SelectedItem);
                strFvsPostTableName = "POST_" + Convert.ToString(lstFVSTablesList.SelectedItem);
                strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + 
                    Tables.FVS.DefaultFVSOutPrePostDbFile;
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


                //Sleep while table link is being built
                using (var calculateConn = new SQLiteConnection(m_oDataMgr.GetConnectionString(m_strTempDB)))
                {
                    calculateConn.Open();
                    if (!m_oDataMgr.DatabaseAttached(calculateConn, strFvsPrePostDb))
                    {
                        m_oDataMgr.m_strSQL = "ATTACH DATABASE '" + strFvsPrePostDb + "' AS SOURCE";
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                        m_oDataMgr.m_strSQL = "CREATE TABLE " + strFvsPreTableName + " AS SELECT * FROM SOURCE." + strFvsPreTableName;
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                        m_oDataMgr.m_strSQL = "CREATE TABLE " + strFvsPostTableName + " AS SELECT * FROM SOURCE." + strFvsPostTableName;
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                    }
                    while (!m_oDataMgr.TableExist(calculateConn, strFvsPreTableName))
                    {
                        Thread.Sleep(1000);
                    }
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

                    //Populate baseline prescription list
                    cboFvsVariableBaselinePkg.Items.Clear();
                    string strSql = "SELECT distinct rxpackage" +
                                    " FROM " + strFvsPreTableName +
                                    " ORDER BY rxpackage ASC";
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Query rxPackage for baseline package list \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql + "\r\n\r\n");
                    }
                    m_oDataMgr.SqlQueryReader(calculateConn, strSql);
                    while(m_oDataMgr.m_DataReader.Read())
                    {
                        if (m_oDataMgr.m_DataReader["rxpackage"] != System.DBNull.Value)
                        {
                            cboFvsVariableBaselinePkg.Items.Add(Convert.ToString(m_oDataMgr.m_DataReader["rxpackage"]));
                        }
                    }
                    m_oDataMgr.m_DataReader.Close();

                    //Create temporary table to populate datagrid
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

                    strSql = "SELECT rxcycle, MIN(Year) as MinYear, \"PRE\" as pre_or_post" +
                                     " FROM " + strFvsPreTableName +
                                     " GROUP BY rxcycle, Year";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Query for year, and rxcycle from " + strFvsPreTableName + " \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql + "\r\n\r\n");
                    }

                    m_oDataMgr.SqlQueryReader(calculateConn, strSql);
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        string strRxCycle = "";
                        int intYear = -99;
                        if (m_oDataMgr.m_DataReader["MinYear"] != System.DBNull.Value)
                        {
                            intYear = Convert.ToInt16(m_oDataMgr.m_DataReader["MinYear"]);
                        }
                        if (m_oDataMgr.m_DataReader["rxcycle"] != System.DBNull.Value)
                        {
                            strRxCycle = Convert.ToString(m_oDataMgr.m_DataReader["rxcycle"]);
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

                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, insertSql);

                        }
                    }
                    m_oDataMgr.m_DataReader.Close();

                    strSql = "SELECT rxcycle, MIN(Year) as MinYear, \"PRE\" as pre_or_post" +
                     " FROM " + strFvsPostTableName +
                     " GROUP BY rxcycle, Year";

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Query for year, and rxcycle from " + strFvsPreTableName + " \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, strSql + "\r\n\r\n");
                    }

                    m_oDataMgr.SqlQueryReader(calculateConn, strSql);
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        string strRxCycle = "";
                        int intYear = -99;
                        if (m_oDataMgr.m_DataReader["MinYear"] != System.DBNull.Value)
                        {
                            intYear = Convert.ToInt16(m_oDataMgr.m_DataReader["MinYear"]);
                        }
                        if (m_oDataMgr.m_DataReader["rxcycle"] != System.DBNull.Value)
                        {
                            strRxCycle = Convert.ToString(m_oDataMgr.m_DataReader["rxcycle"]);
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
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, insertSql);
                        }
                    }
                    m_oDataMgr.m_DataReader.Close();

                    string strPrimaryKeys = "rxcycle, pre_or_post, rxyear";
                    string strColumns = "rxcycle, pre_or_post, rxyear, weight";
                    m_oDataMgr.DataAdapterArrayItemConfigureSelectCommand(FVS_DETAILS_TABLE, m_strFvsViewTableName, strColumns, strPrimaryKeys, "");

                    m_oDataMgr.m_strSQL = "DROP TABLE " + m_strFvsViewTableName;
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
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


        }

        private void init_m_dg()
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
        private void loadnullthreshold()
        {
            DataMgr oDataMgr = new DataMgr();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(this.m_strCalculatedVariablesDb)))
            {
                conn.Open();
                oDataMgr.m_strSQL = "SELECT fvs_null_threshold FROM " + Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName;
                oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);

                if (oDataMgr.m_intError == 0 && oDataMgr.m_DataReader.HasRows)
                {
                    while (oDataMgr.m_DataReader.Read())
                    {
                        string strThreshold = oDataMgr.m_DataReader["fvs_null_threshold"].ToString().Trim();
                        this.cmbThreshold.SelectedItem = strThreshold;
                        this.cmbThreshold.Text = strThreshold;
                        intNullThreshold = Convert.ToInt32(strThreshold);
                    }                
                }
                oDataMgr.m_DataReader.Close();
            }
        }
        private void savenullthreshold()
        {
            int intNewThreshold = Convert.ToInt32(this.cmbThreshold.SelectedItem);
            if (intNewThreshold != intNullThreshold)
            {
                m_oDataMgr.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName +
                    " SET fvs_null_threshold = " + intNewThreshold;
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                intNullThreshold = intNewThreshold;
                m_oDataMgr.m_Transaction.Commit();
                m_oDataMgr.m_Transaction = m_oDataMgr.m_Connection.BeginTransaction();
            }
        }

        public int savevalues(string strVariableType)
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
                        " (ID, VARIABLE_NAME, VARIABLE_DESCRIPTION, VARIABLE_TYPE, BASELINE_RXPACKAGE, VARIABLE_SOURCE, NEGATIVES_YN)" +
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
                strDescription = m_oDataMgr.FixString(strDescription, "'", "''");
                strSql = strSql + lblFvsVariableName.Text.Trim() + "','" + strDescription + "','" +
                         strVariableType + "','" + strBaselinePackage + "','" + LblSelectedVariable.Text.Trim() + "', 'N')";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Add parent record for FVS weighted variable \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + strSql + "\r\n\r\n");
                }
                try
                {
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, strSql);
                    //Set USE_NEGATIVE to Y if desired
                    if (m_bUseNegatives)
                    {
                        m_oDataMgr.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                            " SET NEGATIVES_YN = 'Y' WHERE TRIM(VARIABLE_NAME) = '" + lblFvsVariableName.Text + "'";
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, m_oDataMgr.m_strSQL);
                    }
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
                strDescription = m_oDataMgr.FixString(strDescription, "'", "''");
                string strVariableSource = Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName +
                    "." + lblEconVariableName.Text.Trim();
                strSql = strSql + lblEconVariableName.Text.Trim() + "','" + strDescription + "','" +
                    strVariableType + "','" + strBaselinePackage + "','" + strVariableSource + "', 'N')";
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

        private void val_data_fvs(string strPrePostDb, string strPreTable, string strPostTable)
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
            string strOutputDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" +
                Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
            if (!System.IO.File.Exists(strOutputDb))
            {
                MessageBox.Show("!!FVS Weighted Variable output database missing. It should be here: " +
                    strOutputDb + "!!", "FIA Biosum",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
                return;
            }
            string strCalculateConn = m_oDataMgr.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                                      "\\fvs\\db\\" + strPrePostDb);
            m_bUseNegatives = false;
            using (var calculateConn = new SQLiteConnection(strCalculateConn))
            {
                calculateConn.Open();
                m_oDataMgr.m_strSQL = "SELECT COUNT(*)" +
                                  " FROM " + strPreTable +
                                  " WHERE " + this.lstFVSFieldsList.SelectedItems[0].ToString() + " < 0";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Checking for negative FVS values \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + "\r\n\r\n");
                }
                long negCount = m_oDataMgr.getRecordCount(calculateConn, m_oDataMgr.m_strSQL, strPreTable);

                m_oDataMgr.m_strSQL = "SELECT COUNT(*)" +
                  " FROM " + strPostTable +
                  " WHERE " + this.lstFVSFieldsList.SelectedItems[0].ToString() + " < 0";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "EXECUTE SQL: " + m_oDataMgr.m_strSQL + "\r\n\r\n");
                }
                negCount += m_oDataMgr.getRecordCount(calculateConn, m_oDataMgr.m_strSQL, strPostTable);
                if (negCount > 0)
                {
                    string strMessage = "!! BioSum found " + negCount + " negative values in the PREPOST tables!!" +
                                        " Do you wish to keep the negative values? If not, they will be replaced with nulls.";
                    DialogResult res = MessageBox.Show(strMessage, "FIA Biosum", System.Windows.Forms.MessageBoxButtons.YesNo,
                                                       System.Windows.Forms.MessageBoxIcon.Question);
                    if (res == DialogResult.Yes)
                    {
                        m_bUseNegatives = true;
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

            m_oDataMgr.m_DataSet.Tables[m_strFvsViewTableName].Rows.Clear();

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
        private void btnSaveThreshold_Click(object sender, EventArgs e)
        {
            savenullthreshold();
        }

        private void btnProperties_Click(object sender, EventArgs e)
        {
            if (lstVariables.SelectedItems.Count == 0) return;
            this.grpboxSummary.Hide();
            m_intCurVar = Convert.ToInt32(lstVariables.SelectedItems[0].SubItems[3].Text.Trim());
            string strVariableSource = lstVariables.SelectedItems[0].SubItems[5].Text.Trim();
            string strVariableName = lstVariables.SelectedItems[0].Text.Trim();

            load_properties(strVariableName, strVariableSource);

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

        private void load_properties(string strVariableName, string strVariableSource)
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
            try
            {
                //Determine database and table names based on the source FVS variable
                string[] strPieces = LblSelectedVariable.Text.Split('.');
                string strSourcePreTable = "PRE_" + strPieces[0];
                string strSourcePostTable = "POST_" + strPieces[0];
                string strSourceDatabaseName = "PREPOST_FVSOUT.DB";
                string strTargetPreTable = "PRE_" + strPieces[0] + "_WEIGHTED";
                string strTargetPostTable = "POST_" + strPieces[0] + "_WEIGHTED";
                string strWeightsByRxCyclePreTable = "WEIGHTS_BY_RX_CYCLE_PRE";
                string strWeightsByRxCyclePostTable = "WEIGHTS_BY_RX_CYCLE_POST";
                string strWeightsByRxPkgPreTable = "WEIGHTS_BY_RXPACKAGE_PRE";
                string strWeightsByRxPkgPostTable = "WEIGHTS_BY_RXPACKAGE_POST";

                this.val_data_fvs(strSourceDatabaseName, strSourcePreTable, strSourcePostTable);
                savenullthreshold();
                if (this.m_intError == 0)
                {
                    string strDestinationLinkDir = this.m_oEnv.strTempDir;
                    string strTempDb = m_oUtils.getRandomFile(strDestinationLinkDir, "db");
                    m_oDataMgr.CreateDbFile(strTempDb);

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "btnFvsCalculate_Click: Calculate weighted variable " + lblFvsVariableName.Text + "\r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Temporary database path: " + strTempDb + "\r\n\r\n");
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

                    savevalues("FVS");


                    frmMain.g_sbpInfo.Text = "Calculating and saving PRE/POST values...Stand by";
                    string strPrePostWeightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                        "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
                    string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                        Tables.FVS.DefaultFVSOutPrePostDbFile;

                    string strTempConn = m_oDataMgr.GetConnectionString(strTempDb);
                    using (var tempConn = new SQLiteConnection(strTempConn))
                    {
                        tempConn.Open();
                        
                        //Drop strWeightsByRxCyclePreTable if it exists so we can recreate it
                        if (m_oDataMgr.TableExist(tempConn, strWeightsByRxCyclePreTable))
                        {
                            m_oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxCyclePreTable;
                            m_oDataMgr.SqlNonQuery(tempConn, m_oDataMgr.m_strSQL);
                        }
                        //Drop strWeightsByRxCyclePostTable if it exists so we can recreate it
                        if (m_oDataMgr.TableExist(tempConn, strWeightsByRxCyclePostTable))
                        {
                            m_oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxCyclePostTable;
                            m_oDataMgr.SqlNonQuery(tempConn, m_oDataMgr.m_strSQL);
                        }
                        //Drop strWeightsByRxPkgPreTable if it exists so we can recreate it
                        if (m_oDataMgr.TableExist(tempConn, strWeightsByRxPkgPreTable))
                        {
                            m_oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxPkgPreTable;
                            m_oDataMgr.SqlNonQuery(tempConn, m_oDataMgr.m_strSQL);
                        }
                        //Drop strWeightsByRxPkgPreTable if it exists so we can recreate it
                        if (m_oDataMgr.TableExist(tempConn, strWeightsByRxPkgPostTable))
                        {
                            m_oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxPkgPostTable;
                            m_oDataMgr.SqlNonQuery(tempConn, m_oDataMgr.m_strSQL);
                        }
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
                    string strConn = m_oDataMgr.GetConnectionString(strPrePostWeightedDb);
                    using (var conn = new SQLiteConnection(strConn))
                    {
                        conn.Open();
                        if (!m_oDataMgr.TableExist(conn, strTargetPreTable))
                        {
                            //Link source database to output database
                            if (!m_oDataMgr.DatabaseAttached(conn, strFvsPrePostDb))
                            {
                                m_oDataMgr.m_strSQL = "ATTACH DATABASE '" + strFvsPrePostDb + "' AS FVSOUT";
                                m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                            }
                            // FVS creates a record for
                            // each condition for each cycle regardless of whether there is activity
                            m_oDataMgr.m_strSQL = "CREATE TABLE " + strTargetPreTable + " AS SELECT " +
                                        "biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CAST(0 AS DOUBLE) AS " +
                                        lblFvsVariableName.Text + " FROM " + strSourcePreTable;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Creating final pre/post tables. They did not already exist \r\n");
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + m_oDataMgr.m_strSQL + "\r\n\r\n");
                            }

                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);
                            m_oDataMgr.m_strSQL = "CREATE TABLE " + strTargetPostTable + " AS SELECT " +
                                        "biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CAST(0 AS DOUBLE) AS " +
                                        lblFvsVariableName.Text + " FROM " + strSourcePostTable;
                            m_oDataMgr.SqlNonQuery(conn, m_oDataMgr.m_strSQL);

                            bNewTables = true;
                        }
                    }
                    //Open connection to temporary database and create starting temporary tables
                    //that is table for weights by rx and rxcycle
                    string strCalculateConn = m_oDataMgr.GetConnectionString(strTempDb);
                    int noCount = -1;
                    int intMissing = 0;
                    int intCorrected = 0;
                    int intCorrect = 0;
                    CalculateVariable(strCalculateConn, strWeightsByRxPkgPreTable, strSourcePreTable, strWeightsByRxPkgPostTable,
                                      strSourcePostTable, lblFvsVariableName.Text, strPieces[1], cboFvsVariableBaselinePkg.SelectedItem.ToString(),
                                      lstWeights, bNewTables, ref noCount, out intMissing, out intCorrected, out intCorrect);

                    if (m_intError == 0)
                    {
                        loadLstVariables();
                        loadnullthreshold();
                    }

                    frmMain.g_sbpInfo.Text = "Ready";
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();

                    MessageBox.Show("Variable calculation complete!\n\n" +
                        "For the variable " + lblFvsVariableName.Text + ", " + intMissing + " stand-package " +
                        "combinations had more than " + cmbThreshold.SelectedItem + " missing cases in the PREPOST data and were " +
                        "therefore attributed as NULL in the Weighted tables (" +
                        "and will need to be accounted for in Optimizer if the variable is used " +
                        "in effectiveness determination).\n\n" + intCorrected + 
                        " stand-package combinations " +
                        "had " + cmbThreshold.SelectedItem + " or fewer missing cases, so the Weighted tables " +
                        "contain a value based only on the non-null cases.\n\n" +
                        intCorrect + " stand-package combinations had no missing cases.\n\n" +
                        "Click OK to close this dialog.", "FIA Biosum");
                }
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message, "Weighted Average", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.m_intError = -1;
                frmMain.g_sbpInfo.Text = "Ready";
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
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
            DataMgr oDataMgr = new DataMgr();
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            string strScenarioConn = oDataMgr.GetConnectionString(strScenarioDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile);
            string[] strPieces = LblSelectedVariable.Text.Split('.');

            if (!System.IO.File.Exists(strScenarioDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile))
            {
                MessageBox.Show("!!Optimizer Scenario Rule Definitions database does not exist. FVS Variables cannot be deleted!!", "FIA Biosum",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            using (SQLiteConnection oRenameConn = new SQLiteConnection(strScenarioConn))
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
                oDataMgr.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName +
                    " WHERE (((upper(trim(PRE_FVS_VARIABLE))) = upper(trim('" + strWeightedVariableSource + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oDataMgr.getRecordCount(oRenameConn, oDataMgr.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This FVS Variable Cannot Be Deleted Because It Is In Use As An Effectiveness Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
                // Check for usage as Optimization variable
                strWeightedVariableSource = strPieces[0] + "_WEIGHTED." + lblFvsVariableName.Text;
                oDataMgr.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName +
                    " WHERE (((upper(trim(fvs_variable_name))) = upper(trim('" + strWeightedVariableSource + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oDataMgr.getRecordCount(oRenameConn, oDataMgr.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This FVS Variable Cannot Be Deleted Because It Is In Use As An Optimization Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
                // Check for usage as Tie-Breaker variable
                oDataMgr.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName +
                    " WHERE (((upper(trim(fvs_variable_name))) = upper(trim('" + strWeightedVariableSource + "'))))";
                if ((int)oDataMgr.getRecordCount(oRenameConn, oDataMgr.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This FVS Variable Cannot Be Deleted Because It Is In Use As An Tie-Breaker Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }
                oRenameConn.Close();
            }
            DialogResult objResult = MessageBox.Show("!!You are about to delete an FVS weighted variable. This action cannot be undone. Do you wish to continue?", "FIA Biosum",
                                        System.Windows.Forms.MessageBoxButtons.YesNo,
                                        System.Windows.Forms.MessageBoxIcon.Question);
            if (objResult == DialogResult.Yes)
            {
                // Delete data entries from FVS pre/post tables
                string strPreTable = "PRE_" + strPieces[0] + "_WEIGHTED";
                string strPostTable = "POST_" + strPieces[0] + "_WEIGHTED";
                List<string> lstFields = new List<string> {lblFvsVariableName.Text, lblFvsVariableName.Text + "_null_count"};
                string strCopyCols = "";
                string strPrePostConn = oDataMgr.GetConnectionString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile);

                using (SQLiteConnection prePostConn = new SQLiteConnection(strPrePostConn))
                {
                    prePostConn.Open();

                    string[] arrColumns = oDataMgr.getFieldNamesArray(prePostConn, "SELECT * FROM " + strPreTable);
                    foreach (string strColumn in arrColumns)
                    {
                        if (!lstFields.Contains(strColumn))
                        {
                            strCopyCols = strCopyCols + strColumn + ", ";
                        }
                    }
                    strCopyCols = strCopyCols.Substring(0, strCopyCols.Length - 2);

                    oDataMgr.m_strSQL = "CREATE TABLE " + strPreTable + "_1 AS SELECT " + strCopyCols + " FROM " + strPreTable;
                    oDataMgr.SqlNonQuery(prePostConn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = "DROP TABLE " + strPreTable;
                    oDataMgr.SqlNonQuery(prePostConn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = "ALTER TABLE " + strPreTable + "_1 RENAME TO " + strPreTable;
                    oDataMgr.SqlNonQuery(prePostConn, oDataMgr.m_strSQL);

                    oDataMgr.m_strSQL = "CREATE TABLE " + strPostTable + "_1 AS SELECT " + strCopyCols + " FROM " + strPostTable;
                    oDataMgr.SqlNonQuery(prePostConn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = "DROP TABLE " + strPostTable;
                    oDataMgr.SqlNonQuery(prePostConn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = "ALTER TABLE " + strPostTable + "_1 RENAME TO " + strPostTable;
                    oDataMgr.SqlNonQuery(prePostConn, oDataMgr.m_strSQL);

                    prePostConn.Close();
                }

                DeleteVariableSqlite();
                this.btnFvsDetailsCancel.PerformClick();
            }
            else
            {
                return;
            }
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

        public void BtnDeleteEconVariable_Click(object sender, EventArgs e)
        {
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            DataMgr oDataMgr = new DataMgr();
            string strScenarioConn = oDataMgr.GetConnectionString(strScenarioDir + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile);

            using (SQLiteConnection oScenarioConn = new SQLiteConnection(strScenarioConn))
            {
                oScenarioConn.Open();

                // Check for usage as Optimization variable
                oDataMgr.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName +
                    " WHERE (((upper(trim(fvs_variable_name))) = upper(trim('" + lblEconVariableName.Text + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oDataMgr.getRecordCount(oScenarioConn, oDataMgr.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This Economic Variable Cannot Be Deleted Because It Is In Use As An Optimization Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                // Check for usage as filter
                oDataMgr.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName +
                    " WHERE (((upper(trim(revenue_attribute))) = upper(trim('" + lblEconVariableName.Text + "'))))" +
                    " AND CURRENT_YN = 'Y'";
                if ((int)oDataMgr.getRecordCount(oScenarioConn, oDataMgr.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This Economic Variable Cannot Be Deleted Because It Is In Use As A Dollars Per Acre Filter!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                // Check for usage as tiebreaker
                oDataMgr.m_strSQL = "SELECT Count(*) FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName +
                    " WHERE (((upper(trim(fvs_variable_name))) = upper(trim('" + lblEconVariableName.Text + "'))))";
                if ((int)oDataMgr.getRecordCount(oScenarioConn, oDataMgr.m_strSQL, "TEMP") > 0)
                {
                    MessageBox.Show("!!This Economic Variable Cannot Be Deleted Because It Is In Use As An Tie-Breaker Variable!!", "FIA Biosum",
                      System.Windows.Forms.MessageBoxButtons.OK,
                      System.Windows.Forms.MessageBoxIcon.Exclamation);
                    return;
                }

                oScenarioConn.Close();
            }

            DialogResult objResult = MessageBox.Show("!!You are about to delete an Economic weighted variable. This action cannot be undone. Do you wish to continue?", "FIA Biosum",
                MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);

            if (oDataMgr != null)
            {
                oDataMgr = null;
            }
            if (objResult == DialogResult.Yes)
            {
                DeleteVariableSqlite();
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
            loadLstVariables();
            loadnullthreshold();
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
               
                //Save associated configuration records
                savevalues("ECON");

                //Reload the main grid
                loadLstVariables();
                loadnullthreshold();
                
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
            savenullthreshold();
            // assemble the path for the backup database
            string strDbName = System.IO.Path.GetFileNameWithoutExtension(Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile);
            string strDbFolder = System.IO.Path.GetDirectoryName(Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile);
            string strBackupDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + strDbFolder + "\\" + strDbName + "_backup.db";
            System.IO.File.Copy(frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile,
                strBackupDb, true);
            RecalculateCalculatedVariables_Start();
        }

        private void RecalculateWeightedVariables_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            m_intError = 0;
            DataMgr oDataMgr = new DataMgr();

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

                string strPrePostWeightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                    "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strPrePostWeightedDb)))
                {
                    conn.Open();
                    string[] arrTableNames = oDataMgr.getTableNames(conn);
                    var counter1 = 5;
                    var counter2 = 10;
                    UpdateProgressBar2(counter2);
                    UpdateProgressBar1("Dropping weighted variable tables", counter1);

                    foreach (var strTableName in arrTableNames)
                    {
                        if (!string.IsNullOrEmpty(strTableName))
                        {
                            oDataMgr.m_strSQL = "DROP TABLE " + strTableName;
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, oDataMgr.m_strSQL + "\r\n\r\n");
                            }
                        }
                    }
                    counter1Interval = 5;
                    int counter2Interval = 5;
                    counter1 = counter1 + counter1Interval;
                    counter2 = counter2 + counter2Interval;

                    // Get list of variables to recalculate
                    UpdateProgressBar1("Querying database for calculated variables", counter1);
                    IDictionary<int, string> dictFvsWeightedVariables = new Dictionary<int, string>();
                    IList<string> lstTables = new List<string>();

                    string strCalculatedVariablesDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                    "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
                    using (System.Data.SQLite.SQLiteConnection connCalcVariables = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strCalculatedVariablesDb)))
                    {
                        connCalcVariables.Open();
                        oDataMgr.m_strSQL = "SELECT ID, VARIABLE_SOURCE FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                       " WHERE TRIM(VARIABLE_TYPE) = 'FVS' ORDER BY VARIABLE_SOURCE";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, oDataMgr.m_strSQL + "\r\n\r\n");
                        }
                        oDataMgr.SqlQueryReader(connCalcVariables, oDataMgr.m_strSQL);
                        if (oDataMgr.m_intError == 0 && oDataMgr.m_DataReader.HasRows)
                        {
                            while (oDataMgr.m_DataReader.Read())
                            {
                                int key = Convert.ToInt32(oDataMgr.m_DataReader["ID"]);
                                if (oDataMgr.m_DataReader["VARIABLE_SOURCE"] != System.DBNull.Value && !dictFvsWeightedVariables.Keys.Contains(key))
                                {
                                    dictFvsWeightedVariables.Add(key, Convert.ToString(oDataMgr.m_DataReader["VARIABLE_SOURCE"]));
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

                        oDataMgr.m_DataReader.Close();

                        if (lstTables.Count > 0)
                        {
                            // Reset interval for counter 2 based on number of tables
                            counter2Interval = 70 / (lstTables.Count);
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Stored FVS variable information in memory \r\n\r\n");
                        }

                        if (dictFvsWeightedVariables.Keys.Count > 0)
                        {
                            // Reset counter1 interval based on number of variables
                            counter1Interval = 40 / dictFvsWeightedVariables.Keys.Count;
                        }
                        

                        string strCurrentTable = "";
                        string strWeightsByRxCyclePreTable = "WEIGHTS_BY_RX_CYCLE_PRE";
                        string strWeightsByRxCyclePostTable = "WEIGHTS_BY_RX_CYCLE_POST";
                        string strWeightsByRxPkgPreTable = "WEIGHTS_BY_RXPACKAGE_PRE";
                        string strWeightsByRxPkgPostTable = "WEIGHTS_BY_RXPACKAGE_POST";

                        //create and set temporary db file
                        string strDestinationLinkDir = this.m_oEnv.strTempDir;
                        string strRecalculateDb = m_oUtils.getRandomFile(strDestinationLinkDir, "db");
                        oDataMgr.CreateDbFile(strRecalculateDb);
                        foreach (var keyId in dictFvsWeightedVariables.Keys)
                        {
                            string[] strArray = frmMain.g_oUtils.ConvertListToArray(dictFvsWeightedVariables[keyId], ".");
                            string strTable = "";
                            string strColumn = "";
                            if (strArray.Length == 2)
                            {
                                if (strArray[0].Trim().Length > 0)
                                {
                                    strTable = strArray[0].Trim();
                                }
                                if (strArray[1].Trim().Length > 0)
                                {
                                    strColumn = strArray[1].Trim();
                                }

                                string strSourcePreTable = "PRE_" + strTable;
                                string strSourcePostTable = "POST_" + strTable;

                                using (SQLiteConnection connTemp = new SQLiteConnection(oDataMgr.GetConnectionString(strRecalculateDb)))
                                {
                                    connTemp.Open();

                                    if (!strTable.Equals(strCurrentTable))
                                    {
                                        counter2 = counter2 + counter2Interval;
                                        UpdateProgressBar2(counter2);
                                        // We need to create the tables
                                        string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                                             Tables.FVS.DefaultFVSOutPrePostDbFile;
                                        string strTargetPreTable = "PRE_" + strTable + "_WEIGHTED";
                                        string strTargetPostTable = "POST_" + strTable + "_WEIGHTED";

                                        UpdateProgressBar1("Creating tables for " + strTable, counter1);
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        {
                                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Creating tables for " + strTable + " \r\n\r\n");
                                        }
                                        counter1 = counter1 + 3;

                                        //Link source tables to output database
                                        if (!oDataMgr.DatabaseAttached(conn, strFvsPrePostDb))
                                        {
                                            oDataMgr.m_strSQL = "ATTACH DATABASE '" + strFvsPrePostDb + "' AS FVSSOURCE";
                                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                        }

                                        // FVS creates a record for
                                        // each condition for each cycle regardless of whether there is activity
                                        oDataMgr.m_strSQL = "CREATE TABLE " + strTargetPreTable + " (biosum_cond_id CHAR(25), rxpackage CHAR(3), rx CHAR(3), rxcycle CHAR(1), fvs_variant CHAR(2), PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
                                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                        oDataMgr.m_strSQL = "INSERT INTO " + strTargetPreTable +
                                            " SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant FROM " + strSourcePreTable;
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                        {
                                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Creating final pre/post tables. They did not already exist \r\n");
                                            frmMain.g_oUtils.WriteText(m_strDebugFile, "sql: " + oDataMgr.m_strSQL + "\r\n\r\n");
                                        }
                                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                        oDataMgr.m_strSQL = "CREATE TABLE " + strTargetPostTable + " (biosum_cond_id CHAR(25), rxpackage CHAR(3), rx CHAR(3), rxcycle CHAR(1), fvs_variant CHAR(2), PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
                                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                        oDataMgr.m_strSQL = "INSERT INTO " + strTargetPostTable +
                                            " SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant FROM " + strSourcePostTable;
                                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                                        if (oDataMgr.TableExist(conn, strSourcePreTable))
                                        {
                                            oDataMgr.m_strSQL = "DROP TABLE " + strSourcePreTable;
                                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                        }
                                        if (oDataMgr.TableExist(conn, strSourcePostTable))
                                        {
                                            oDataMgr.m_strSQL = "DROP TABLE " + strSourcePreTable;
                                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                                        }

                                        strCurrentTable = strTable;
                                    }
                                    //Drop strWeightsByRxCyclePreTable if it exists so we can recreate it
                                    if (oDataMgr.TableExist(connTemp, strWeightsByRxCyclePreTable))
                                    {
                                        oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxCyclePreTable;
                                        oDataMgr.SqlNonQuery(connTemp, oDataMgr.m_strSQL);
                                    }
                                    //Drop strWeightsByRxCyclePostTable if it exists so we can recreate it
                                    if (oDataMgr.TableExist(connTemp, strWeightsByRxCyclePostTable))
                                    {
                                        oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxCyclePostTable;
                                        oDataMgr.SqlNonQuery(connTemp, oDataMgr.m_strSQL);
                                    }
                                    //Drop strWeightsByRxPkgPreTable if it exists so we can recreate it
                                    if (oDataMgr.TableExist(connTemp, strWeightsByRxPkgPreTable))
                                    {
                                        oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxPkgPreTable;
                                        oDataMgr.SqlNonQuery(connTemp, oDataMgr.m_strSQL);
                                    }
                                    //Drop strWeightsByRxPkgPostTable if it exists so we can recreate it
                                    if (oDataMgr.TableExist(connTemp, strWeightsByRxPkgPostTable))
                                    {
                                        oDataMgr.m_strSQL = "DROP TABLE " + strWeightsByRxPkgPostTable;
                                        oDataMgr.SqlNonQuery(connTemp, oDataMgr.m_strSQL);
                                    }
                                }

                                IList<string[]> lstWeights = new List<string[]>();
                                string strVariableName = "";
                                string strBaselinePackage = "";

                                oDataMgr.m_strSQL = "SELECT ID, VARIABLE_NAME, BASELINE_RXPACKAGE, " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".* " +
                                "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + " INNER JOIN " +
                                Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + " ON " +
                                Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + ".ID = " +
                                Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + ".CALCULATED_VARIABLES_ID " +
                                "WHERE CALCULATED_VARIABLES_ID = " + keyId;
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                {
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, oDataMgr.m_strSQL + "\r\n\r\n");
                                }
                                oDataMgr.SqlQueryReader(connCalcVariables, oDataMgr.m_strSQL);
                                if (oDataMgr.m_intError == 0 && oDataMgr.m_DataReader.HasRows)
                                {
                                    while (oDataMgr.m_DataReader.Read())
                                    {
                                        strVariableName = Convert.ToString(oDataMgr.m_DataReader["VARIABLE_NAME"]).Trim();
                                        strBaselinePackage = Convert.ToString(oDataMgr.m_DataReader["BASELINE_RXPACKAGE"]).Trim();
                                        for (int rxCycle = 1; rxCycle < 5; rxCycle++)
                                        {
                                            // PRE VALUES
                                            string[] strPreRow = new string[3];
                                            string strFieldName = "weight_" + rxCycle + "_pre";
                                            strPreRow[idxRxCycle] = Convert.ToString(rxCycle);
                                            strPreRow[idxPreOrPost] = "PRE";
                                            strPreRow[idxWeight] = Convert.ToString(oDataMgr.m_DataReader[strFieldName]);
                                            lstWeights.Add(strPreRow);

                                            // POST VALUES
                                            string[] strPostRow = new string[3];
                                            strFieldName = "weight_" + rxCycle + "_post";
                                            strPostRow[idxRxCycle] = Convert.ToString(rxCycle);
                                            strPostRow[idxPreOrPost] = "POST";
                                            strPostRow[idxWeight] = Convert.ToString(oDataMgr.m_DataReader[strFieldName]);
                                            lstWeights.Add(strPostRow);
                                        }
                                    }
                                }
                                oDataMgr.m_DataReader.Close();
                                int intMissing = 0;
                                int intCorrected = 0;
                                int intCorrect = 0;
                                string strCalculateConn = oDataMgr.GetConnectionString(strRecalculateDb);
                                CalculateVariable(strCalculateConn, strWeightsByRxPkgPreTable, strSourcePreTable,
                                    strWeightsByRxPkgPostTable, strSourcePostTable, strVariableName, strColumn, strBaselinePackage,
                                    lstWeights, false, ref counter1, out intMissing, out intCorrect, out intCorrected);
                            }
                        }
                    }
                    UpdateProgressBar1("Variables Recalculated!!", 100);
                    UpdateProgressBar2(100);

                    if (conn != null)
                    {
                        conn.Close();
                    }

                    MessageBox.Show("Variables Recalculated!!", "FIA Biosum");

                    RecalculateCalculatedVariables_Finish();
                }
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }

            catch (System.Threading.ThreadAbortException err)
            {
                if (oDataMgr != null)
                {
                    if (oDataMgr.m_DataSet != null)
                    {
                        oDataMgr.m_DataSet.Clear();
                        oDataMgr.m_DataSet.Dispose();
                    }
                    oDataMgr = null;
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

        private void CalculateVariable(string strCalculateConn, string strWeightsByRxPkgPreTable,
                                       string strSourcePreTable, string strWeightsByRxPkgPostTable, string strSourcePostTable,
                                       string strVariableName, string strFieldName, string strBaselinePkg, IList<string[]> lstRows,
                                       bool bNewTables, ref int counter1, out int intMissing, out int intCorrected, out int intCorrect)
        {

            if (counter1 > 0)
            {
                counter1 = counter1 + counter1Interval;
                UpdateProgressBar1("Calculating values for " + strVariableName, counter1);
            }
            intMissing = 0;
            intCorrected = 0;
            intCorrect = 0;
            int intTotal = 0;
            string strWeightsByRxCyclePreTable = "WEIGHTS_BY_RX_CYCLE_PRE";
            string strWeightsByRxCyclePostTable = "WEIGHTS_BY_RX_CYCLE_POST";
            string strTargetPreTable = strSourcePreTable + "_WEIGHTED";
            string strTargetPostTable = strSourcePostTable + "_WEIGHTED";
            string strPrePostWeightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;


            Dictionary<string, Dictionary<string, double[]>> correctionFactors = new Dictionary<string, Dictionary<string, double[]>>();

            using (var calculateConn = new SQLiteConnection(strCalculateConn))
            {
                calculateConn.Open();

                string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                                             Tables.FVS.DefaultFVSOutPrePostDbFile;
                if (!m_oDataMgr.DatabaseAttached(calculateConn, strFvsPrePostDb))
                {
                    m_oDataMgr.m_strSQL = "ATTACH DATABASE '" + strFvsPrePostDb + "' AS FVSSOURCE";
                    _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                    m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                    _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                }

                m_oDataMgr.m_strSQL = "CREATE TABLE " + strWeightsByRxCyclePreTable +
                    " AS SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CAST(0 AS DOUBLE) " +
                    "AS " + strVariableName + " FROM " + strSourcePreTable;
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.AddIndex(calculateConn, strWeightsByRxCyclePreTable, "index_" + strWeightsByRxCyclePreTable.ToLower() + "_composite", "biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant");

                m_oDataMgr.m_strSQL = "CREATE TABLE " + strWeightsByRxCyclePostTable +
                    " AS SELECT biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant, CAST(0 AS DOUBLE) " +
                    " AS " + strVariableName + " FROM " + strSourcePostTable;
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.AddIndex(calculateConn, strWeightsByRxCyclePostTable, "index_" + strWeightsByRxCyclePostTable.ToLower() + "_composite", "biosum_cond_id, rxpackage, rx, rxcycle, fvs_variant");

                //Add column for weights and populate
                m_oDataMgr.AddColumn(calculateConn, strWeightsByRxCyclePreTable, "weight", "DOUBLE", "");
                m_oDataMgr.AddColumn(calculateConn, strWeightsByRxCyclePostTable, "weight", "DOUBLE", "");

                m_oDataMgr.m_strSQL = "ATTACH DATABASE '" + frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" +
                    Tables.OptimizerDefinitions.DefaultSqliteDbFile + "' AS variable_defs";
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);

                bool bUseNegatives = false;

                m_oDataMgr.m_strSQL = "SELECT NEGATIVES_YN FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                    " WHERE VARIABLE_NAME = '" + strVariableName + "'";
                m_oDataMgr.SqlQueryReader(calculateConn, m_oDataMgr.m_strSQL);
                if (m_oDataMgr.m_DataReader.HasRows)
                {
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        if (m_oDataMgr.m_DataReader["NEGATIVES_YN"].ToString().Trim() == "Y")
                        {
                            bUseNegatives = true;
                        }
                    }
                }
                m_oDataMgr.m_DataReader.Close();


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
                        // Populate weight column
                        m_oDataMgr.m_strSQL = "UPDATE " + strTargetTableName +
                            " SET weight = " + dblWeight +
                            " WHERE rxcycle = '" + strRxCycle + "'";
                        _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                        _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                        // Apply weights to each cycle
                        if (!bUseNegatives)
                        {
                            m_oDataMgr.m_strSQL = "UPDATE " + strTargetTableName + " AS w " +
                                "SET " + strVariableName + " = " +
                                "CASE WHEN " +
                                "(SELECT " + strFieldName + " FROM " + strSourceTableName + " AS p " +
                                "WHERE w.biosum_cond_id = p.biosum_cond_id AND w.rxpackage = p.rxpackage " +
                                "AND w.rx = p.rx AND w.rxcycle = p.rxcycle AND w.fvs_variant = p.fvs_variant) " +
                                " >= 0 THEN " +
                                "(SELECT " + strFieldName + " FROM " + strSourceTableName + " AS p " +
                                "WHERE w.biosum_cond_id = p.biosum_cond_id AND w.rxpackage = p.rxpackage " +
                                "AND w.rx = p.rx AND w.rxcycle = p.rxcycle AND w.fvs_variant = p.fvs_variant)" +
                                "* " + dblWeight +
                                " ELSE NULL END" +
                                " WHERE w.rxcycle = '" + strRxCycle + "'";
                        }
                        else
                        {
                            m_oDataMgr.m_strSQL = "UPDATE " + strTargetTableName + " AS w " +
                                "SET " + strVariableName + " = " +
                                "(SELECT " + strFieldName + " FROM " + strSourceTableName + " AS p " +
                                "WHERE w.biosum_cond_id = p.biosum_cond_id AND w.rxpackage = p.rxpackage " +
                                "AND w.rx = p.rx AND w.rxcycle = p.rxcycle AND w.fvs_variant = p.fvs_variant) " +
                                "WHERE w.rxcycle = '" + strRxCycle + "'";
                        }
                        _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                        _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                    }
                }

                // Get biosum_cond_id/rxpackage combinations where the value is null. Store the sum of the weights and the null count

                m_oDataMgr.m_strSQL = "SELECT COUNT(*) FROM (SELECT DISTINCT biosum_cond_id, rxpackage FROM " + strWeightsByRxCyclePreTable + ")";
                intTotal += Convert.ToInt32(m_oDataMgr.getRecordCount(calculateConn, m_oDataMgr.m_strSQL, strWeightsByRxCyclePreTable));

                m_oDataMgr.m_strSQL = "SELECT biosum_cond_id, rxpackage, weight FROM " + strWeightsByRxCyclePreTable +
                " WHERE " + strVariableName + " IS NULL";
                m_oDataMgr.SqlQueryReader(calculateConn, m_oDataMgr.m_strSQL);
                if (m_oDataMgr.m_DataReader.HasRows)
                {
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        double dblWt = Convert.ToDouble(m_oDataMgr.m_DataReader["weight"]);
                        string strCondId = m_oDataMgr.m_DataReader["biosum_cond_id"].ToString().Trim();
                        string strRxPkg = m_oDataMgr.m_DataReader["rxpackage"].ToString().Trim();
                        double[] entry = { dblWt, 1 };
                        if (!correctionFactors.ContainsKey(strCondId))
                        {
                            Dictionary<string, double[]> dictEntry = new Dictionary<string, double[]> { { strRxPkg, entry } };
                            correctionFactors.Add(strCondId, dictEntry);
                        }
                        else if (!correctionFactors[strCondId].ContainsKey(strRxPkg))
                        {
                            correctionFactors[strCondId].Add(strRxPkg, entry);
                        }
                        else 
                        {
                            double dblCurWtSum = correctionFactors[strCondId][strRxPkg][WEIGHT_SUM];
                            correctionFactors[strCondId][strRxPkg][WEIGHT_SUM] = dblCurWtSum + dblWt;
                            correctionFactors[strCondId][strRxPkg][NULL_COUNT]++;
                        }
                    }
                }

                m_oDataMgr.m_strSQL = "SELECT biosum_cond_id, rxpackage, weight FROM " + strWeightsByRxCyclePostTable +
                    " WHERE " + strVariableName + " IS NULL";
                m_oDataMgr.SqlQueryReader(calculateConn, m_oDataMgr.m_strSQL);
                if (m_oDataMgr.m_DataReader.HasRows)
                {
                    while (m_oDataMgr.m_DataReader.Read())
                    {
                        double dblWt = Convert.ToDouble(m_oDataMgr.m_DataReader["weight"]);
                        string strCondId = m_oDataMgr.m_DataReader["biosum_cond_id"].ToString().Trim();
                        string strRxPkg = m_oDataMgr.m_DataReader["rxpackage"].ToString().Trim();
                        double[] entry = { dblWt, 1 };

                        if (!correctionFactors.ContainsKey(strCondId))
                        {
                            Dictionary<string, double[]> dictEntry = new Dictionary<string, double[]> { { strRxPkg, entry } };
                            correctionFactors.Add(strCondId, dictEntry);
                        }
                        else if (!correctionFactors[strCondId].ContainsKey(strRxPkg))
                        {
                            correctionFactors[strCondId].Add(strRxPkg, entry);
                        }
                        else
                        {
                            double dblCurWtSum = correctionFactors[strCondId][strRxPkg][WEIGHT_SUM];
                            correctionFactors[strCondId][strRxPkg][WEIGHT_SUM] = dblCurWtSum + dblWt;
                            correctionFactors[strCondId][strRxPkg][NULL_COUNT]++;
                        }
                    }
                }

                m_oDataMgr.m_DataReader.Close();

                // Sum by rxpackage across cycles
                m_oDataMgr.m_strSQL = "CREATE TABLE " + strWeightsByRxPkgPreTable +
                " AS SELECT biosum_cond_id, rxpackage, \'0\' AS rx, " +
                "SUM(" + strVariableName + ") AS sum_pre FROM " + strWeightsByRxCyclePreTable +
                " GROUP BY biosum_cond_id, rxpackage";
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                // Update rx with rx from cycle 1
                m_oDataMgr.m_strSQL = "UPDATE " + strWeightsByRxPkgPreTable + " AS w " +
                        "SET rx = (SELECT r.rx FROM " + strWeightsByRxCyclePreTable + " AS r " +
                        "WHERE w.biosum_cond_id = r.biosum_cond_id AND w.rxpackage = r.rxpackage) " +
                        "WHERE (SELECT rxcycle FROM " + strWeightsByRxCyclePreTable + " AS r " +
                        "WHERE w.biosum_cond_id = r.biosum_cond_id AND w.rxpackage = r.rxpackage) = '1'";
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                m_oDataMgr.m_strSQL = "CREATE TABLE " + strWeightsByRxPkgPostTable +
                    " AS SELECT biosum_cond_id, rxpackage, \'0\' AS rx, " +
                    "SUM(" + strVariableName + ") AS sum_post FROM " + strWeightsByRxCyclePostTable +
                    " GROUP BY biosum_cond_id, rxpackage";
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                // Update rx with rx from cycle 1
                m_oDataMgr.m_strSQL = "UPDATE " + strWeightsByRxPkgPostTable + " AS w " +
                    "SET rx = (SELECT r.rx FROM " + strWeightsByRxCyclePostTable + " AS r " +
                    "WHERE w.biosum_cond_id = r.biosum_cond_id AND w.rxpackage = r.rxpackage) " +
                    "WHERE (SELECT rxcycle FROM " + strWeightsByRxCyclePostTable + " AS r " +
                    "WHERE w.biosum_cond_id = r.biosum_cond_id AND w.rxpackage = r.rxpackage) = '1'";
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                // end using
            }

            //Switch connection to the final storage location and prepare the tables to receive the output
            string strPrePostConn = m_oDataMgr.GetConnectionString(strPrePostWeightedDb);
            using (var prePostConn = new SQLiteConnection(strPrePostConn))
            {
                prePostConn.Open();
                //Check to see if columns exists, they shouldn't, warn that values will be overwritten
                if (m_oDataMgr.ColumnExist(prePostConn, strTargetPreTable, strVariableName))
                {
                    if (bNewTables == false)
                        MessageBox.Show("Values for " + strVariableName + " were previously calculated! " +
                                        "They will be overwritten!", "FIA Biosum");
                }
                else
                {
                    m_oDataMgr.AddColumn(prePostConn, strTargetPreTable,
                        strVariableName, "DOUBLE", "");
                    m_oDataMgr.AddColumn(prePostConn, strTargetPostTable,
                        strVariableName, "DOUBLE", "");
                }
                if (correctionFactors.Keys.Count > 0)
                {
                    m_oDataMgr.AddColumn(prePostConn, strTargetPreTable,
                    strVariableName + "_null_count", "INTEGER", "");
                    m_oDataMgr.AddColumn(prePostConn, strTargetPostTable,
                    strVariableName + "_null_count", "INTEGER", "");
                }
                prePostConn.Close();
            }

            //Switch connection to temporary database
            using (var calculateConn = new SQLiteConnection(strCalculateConn))
            {
                calculateConn.Open();
                if (!m_oDataMgr.DatabaseAttached(calculateConn, strPrePostWeightedDb))
                {
                    m_oDataMgr.m_strSQL = "ATTACH DATABASE '" + strPrePostWeightedDb + "' AS PREPOST_WEIGHTED";
                    _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                    m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                    _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                }
                if (counter1 > 0)
                {
                    counter1 = counter1 + counter1Interval;
                    if (counter1 > 100)
                    {
                        counter1 = 100;
                    }
                    UpdateProgressBar1("Calculating values for " + strVariableName, counter1);
                }

                m_oDataMgr.m_strSQL = "UPDATE " + strTargetPreTable + " AS f " +
                "SET " + strVariableName + " = (SELECT (sum_pre + sum_post) FROM " + strWeightsByRxPkgPostTable +
                " AS pt INNER JOIN " + strWeightsByRxPkgPreTable + " AS pe ON pt.biosum_cond_id = pe.biosum_cond_id " +
                "WHERE pe.biosum_cond_id = f.biosum_cond_id AND pt.rxpackage = '" + strBaselinePkg +
                "' AND pe.rxpackage = '" + strBaselinePkg + "') WHERE f.rxcycle = '1'";
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                m_oDataMgr.m_strSQL = "UPDATE " + strTargetPostTable + " AS f " +
                    "SET " + strVariableName + " = (SELECT (sum_pre + sum_post) FROM " + strWeightsByRxPkgPostTable +
                    " AS pt INNER JOIN " + strWeightsByRxPkgPreTable + " AS pe ON pt.rxpackage = pe.rxpackage AND pt.biosum_cond_id = pe.biosum_cond_id " +
                    "WHERE pe.rxpackage = f.rxpackage AND pe.biosum_cond_id = f.biosum_cond_id) " + "WHERE f.rxcycle = '1'";
                _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                foreach (string strCondId in correctionFactors.Keys)
                {
                    foreach (string strRxPkg in correctionFactors[strCondId].Keys)
                    {
                        if (correctionFactors[strCondId][strRxPkg][NULL_COUNT] > intNullThreshold)
                        {
                            intMissing++;
                            m_oDataMgr.m_strSQL = "UPDATE " + strTargetPreTable +
                                " SET " + strVariableName + " = NULL " +
                                "WHERE biosum_cond_id = '" + strCondId + "'" +
                                " AND rxpackage = '" + strRxPkg + "'";
                            _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                            m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                            _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                            m_oDataMgr.m_strSQL = "UPDATE " + strTargetPostTable +
                                " SET " + strVariableName + " = NULL " +
                                "WHERE biosum_cond_id = '" + strCondId + "'" +
                                " AND rxpackage = '" + strRxPkg + "'";
                            _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                            m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                            _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                        }
                        else
                        {
                            intCorrected++;
                            m_oDataMgr.m_strSQL = "UPDATE " + strTargetPreTable +
                               " SET " + strVariableName + " = " + strVariableName +
                               " * (1 / (1 - " + correctionFactors[strCondId][strRxPkg][WEIGHT_SUM] + "))" +
                               " WHERE biosum_cond_id = '" + strCondId + "'" +
                               " AND rxpackage = '" + strRxPkg + "'";
                            _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                            m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                            _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);

                            m_oDataMgr.m_strSQL = "UPDATE " + strTargetPostTable +
                                " SET " + strVariableName + " = " + strVariableName +
                                " * (1 / (1 - " + correctionFactors[strCondId][strRxPkg][WEIGHT_SUM] + "))" +
                                " WHERE biosum_cond_id = '" + strCondId + "'" +
                                " AND rxpackage = '" + strRxPkg + "'";
                            _frmScenario.DebugLog(true, m_strDebugFile, m_oDataMgr.m_strSQL);
                            m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                            _frmScenario.DebugLog(false, m_strDebugFile, m_oDataMgr.m_strSQL);
                        }
                        m_oDataMgr.m_strSQL = "UPDATE " + strTargetPreTable +
                            " SET " + strVariableName + "_null_count = " + correctionFactors[strCondId][strRxPkg][NULL_COUNT] + 
                            " WHERE biosum_cond_id = '" + strCondId + "' AND rxpackage = '" + strRxPkg + "' AND  rxcycle = '1'";
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                        m_oDataMgr.m_strSQL = "UPDATE " + strTargetPostTable +
                            " SET " + strVariableName + "_null_count = " + correctionFactors[strCondId][strRxPkg][NULL_COUNT] +
                            " WHERE biosum_cond_id = '" + strCondId + "' AND rxpackage = '" + strRxPkg + "' AND  rxcycle = '1'";
                        m_oDataMgr.SqlNonQuery(calculateConn, m_oDataMgr.m_strSQL);
                    }
                }
                intCorrect = intTotal - intMissing - intCorrected;
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

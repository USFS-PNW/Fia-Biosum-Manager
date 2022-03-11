using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using System.Threading;
using System.Data.OleDb;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Summary description for uc_fvs_input.
    /// </summary>
    public class uc_fvs_input : System.Windows.Forms.UserControl
    {
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnClose;
        private string m_strProjDir = "";
        private string m_strProjId = "";
        //private string m_strDsnIn="";
        private string m_strLocFile = "";
        private string m_strSlfFile = "";
        //private string m_strFvsFile="";
        private string m_strDsnOut = "";
        //private string m_strInDir="";
        // private string m_strInMDBFile="";
        //private string m_strOutDir="";
        private string m_strOutMDBFile = "";
        private string m_strRxTable = "";
        private string m_strPlotTable = "";
        private string m_strCondTable = "";
        private string m_strTreeTable = "";
        private string m_strTreeSpcTable = "";
        private string m_strOutPotFireBaseYearMDBFile = "";
        private ado_data_access m_ado;
        private dao_data_access m_dao;
        private string m_strConn = "";
        private bool m_bOverwrite = true;
        private IDictionary<string, List<string>> m_dictVariantStates = null;
        private System.Windows.Forms.Button btnHelp;
        private int m_intError = 0;
        private System.Threading.Thread m_thread;

        //list view column constants
        private const int COL_CHECKBOX = 0;
        private const int COL_VARIANT = 1;
        private const int COL_RX = 2;
        private const int COL_STANDCOUNT = 3;
        private const int COL_TREECOUNT = 4;
        private const int COL_MDBOUT = 5;
        private const int COL_SUMMARYCOUNT = 6;
        private const int COL_CUTCOUNT = 7;
        private const int COL_LEFTCOUNT = 8;
        private const int COL_POTFIRECOUNT = 9;
        private const int COL_POTFIREMDBOUT = 10;
        private const int COL_POTFIREBASEYEARCOUNT = 11;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.Button btnCancel;
        private bool bAbort = false;
        public FIA_Biosum_Manager.frmTherm m_frmTherm;
        private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
        private Queries m_oQueries = new Queries();
        private RxTools m_oRxTools = new RxTools();
        private frmMain _frmMain = null;
        private frmDialog _frmDialog = null;
        private string m_strFVSCycleLength = "10";

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;
        private string m_strDebugFile;
        private TabControl tabControl1;
        private TabPage tabPage2;
        private TextBox txtDataDir;
        private Button btnExecuteAction;
        public ListView lstFvsInput;
        private ComboBox cmbAction;
        private Label lblRxCnt;
        private Label label1;
        private Button btnChkAll;
        private Button btnRxPackage;
        private Button btnClearAll;
        private Label lblRxPackageCnt;
        private Button btnRefresh;
        private Button btnRx;
        private TabPage tabPage1;
        private GroupBox otherOptionsGroupBox;
        private CheckBox chkCull;
        private GroupBox grpGRMOptions;
        private CheckBox chkGRM;
        private GroupBox grpDWMOptions;
        private LinkLabel linkLabelFuelModel;
        private GroupBox groupBox2;
        private Label label6;
        private Label label5;
        private CheckedListBox chkLstBoxDuffYears;
        private CheckedListBox chkLstBoxLitterYears;
        private CheckBox chkDwmFuelModel;
        private CheckBox chkDwmFuelBiomass;
        private Label label4;
        private Label label2;
        private TextBox txtMinLargeFwdTL;
        private TextBox txtMinCwdTL;
        private Label label3;
        private TextBox txtMinSmallFwdTL;
        private Button btnDatamart;
        private TextBox txtFIADatamart;
        private Label lblFiaDatamartFile;
        private ComboBox cmbSelectedGroup;
        private Label lblSelectedGroup;

        delegate string[] GetListBoxItemsDlg(CheckedListBox checkedListBox);

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public uc_fvs_input()
        {
            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();
            this.m_ado = new ado_data_access();
            this.m_dao = new dao_data_access();
            this.m_oLvRowColors.ReferenceAlternateBackgroundColor = frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvRowColors.ReferenceAlternateForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
            this.m_oLvRowColors.ReferenceForegroundColor = frmMain.g_oGridViewRowForegroundColor;
            this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
            this.m_oLvRowColors.CustomFullRowSelect = true;
            this.m_oLvRowColors.ReferenceListView = this.lstFvsInput;
            if (frmMain.g_oGridViewFont != null) this.lstFvsInput.Font = frmMain.g_oGridViewFont;

            this.m_oEnv = new env();

            for (int i = 2001; i <= DateTime.Now.Year; i++)
            {
                this.chkLstBoxDuffYears.Items.Add(i.ToString());
                this.chkLstBoxLitterYears.Items.Add(i.ToString());
            }

            // TODO: Add any initialization after the InitializeComponent call

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_fvs_input));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lblSelectedGroup = new System.Windows.Forms.Label();
            this.cmbSelectedGroup = new System.Windows.Forms.ComboBox();
            this.btnDatamart = new System.Windows.Forms.Button();
            this.txtFIADatamart = new System.Windows.Forms.TextBox();
            this.lblFiaDatamartFile = new System.Windows.Forms.Label();
            this.txtDataDir = new System.Windows.Forms.TextBox();
            this.btnExecuteAction = new System.Windows.Forms.Button();
            this.lstFvsInput = new System.Windows.Forms.ListView();
            this.cmbAction = new System.Windows.Forms.ComboBox();
            this.lblRxCnt = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.btnRxPackage = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.lblRxPackageCnt = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnRx = new System.Windows.Forms.Button();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.otherOptionsGroupBox = new System.Windows.Forms.GroupBox();
            this.chkCull = new System.Windows.Forms.CheckBox();
            this.grpGRMOptions = new System.Windows.Forms.GroupBox();
            this.chkGRM = new System.Windows.Forms.CheckBox();
            this.grpDWMOptions = new System.Windows.Forms.GroupBox();
            this.linkLabelFuelModel = new System.Windows.Forms.LinkLabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.chkLstBoxDuffYears = new System.Windows.Forms.CheckedListBox();
            this.chkLstBoxLitterYears = new System.Windows.Forms.CheckedListBox();
            this.chkDwmFuelModel = new System.Windows.Forms.CheckBox();
            this.chkDwmFuelBiomass = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMinLargeFwdTL = new System.Windows.Forms.TextBox();
            this.txtMinCwdTL = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtMinSmallFwdTL = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgress = new System.Windows.Forms.Label();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.otherOptionsGroupBox.SuspendLayout();
            this.grpGRMOptions.SuspendLayout();
            this.grpDWMOptions.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tabControl1);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.lblProgress);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(800, 600);
            this.groupBox1.TabIndex = 99;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Location = new System.Drawing.Point(7, 51);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(795, 495);
            this.tabControl1.TabIndex = 100;
            this.tabControl1.TabIndexChanged += new System.EventHandler(this.tabControl1_TabIndexChanged);
            this.tabControl1.Resize += new System.EventHandler(this.tabControl1_Resize);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lblSelectedGroup);
            this.tabPage2.Controls.Add(this.cmbSelectedGroup);
            this.tabPage2.Controls.Add(this.btnDatamart);
            this.tabPage2.Controls.Add(this.txtFIADatamart);
            this.tabPage2.Controls.Add(this.lblFiaDatamartFile);
            this.tabPage2.Controls.Add(this.txtDataDir);
            this.tabPage2.Controls.Add(this.btnExecuteAction);
            this.tabPage2.Controls.Add(this.lstFvsInput);
            this.tabPage2.Controls.Add(this.cmbAction);
            this.tabPage2.Controls.Add(this.lblRxCnt);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.btnChkAll);
            this.tabPage2.Controls.Add(this.btnRxPackage);
            this.tabPage2.Controls.Add(this.btnClearAll);
            this.tabPage2.Controls.Add(this.lblRxPackageCnt);
            this.tabPage2.Controls.Add(this.btnRefresh);
            this.tabPage2.Controls.Add(this.btnRx);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(787, 466);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Main Menu";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.SizeChanged += new System.EventHandler(this.tabControl1_Resize);
            // 
            // lblSelectedGroup
            // 
            this.lblSelectedGroup.AutoSize = true;
            this.lblSelectedGroup.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelectedGroup.Location = new System.Drawing.Point(276, 385);
            this.lblSelectedGroup.Name = "lblSelectedGroup";
            this.lblSelectedGroup.Size = new System.Drawing.Size(121, 17);
            this.lblSelectedGroup.TabIndex = 104;
            this.lblSelectedGroup.Text = "Selected Group";
            // 
            // cmbSelectedGroup
            // 
            this.cmbSelectedGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSelectedGroup.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSelectedGroup.FormattingEnabled = true;
            this.cmbSelectedGroup.Location = new System.Drawing.Point(442, 382);
            this.cmbSelectedGroup.Name = "cmbSelectedGroup";
            this.cmbSelectedGroup.Size = new System.Drawing.Size(250, 24);
            this.cmbSelectedGroup.TabIndex = 103;
            // 
            // btnDatamart
            // 
            this.btnDatamart.Image = ((System.Drawing.Image)(resources.GetObject("btnDatamart.Image")));
            this.btnDatamart.Location = new System.Drawing.Point(698, 349);
            this.btnDatamart.Name = "btnDatamart";
            this.btnDatamart.Size = new System.Drawing.Size(59, 32);
            this.btnDatamart.TabIndex = 102;
            this.btnDatamart.Click += new System.EventHandler(this.btnDatamart_Click);
            // 
            // txtFIADatamart
            // 
            this.txtFIADatamart.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFIADatamart.Location = new System.Drawing.Point(279, 354);
            this.txtFIADatamart.Name = "txtFIADatamart";
            this.txtFIADatamart.Size = new System.Drawing.Size(413, 22);
            this.txtFIADatamart.TabIndex = 101;
            // 
            // lblFiaDatamartFile
            // 
            this.lblFiaDatamartFile.AutoSize = true;
            this.lblFiaDatamartFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFiaDatamartFile.Location = new System.Drawing.Point(276, 334);
            this.lblFiaDatamartFile.Name = "lblFiaDatamartFile";
            this.lblFiaDatamartFile.Size = new System.Drawing.Size(354, 17);
            this.lblFiaDatamartFile.TabIndex = 100;
            this.lblFiaDatamartFile.Text = "Path to the FIA Datamart input SQLite database";
            // 
            // txtDataDir
            // 
            this.txtDataDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataDir.Location = new System.Drawing.Point(128, 48);
            this.txtDataDir.Name = "txtDataDir";
            this.txtDataDir.Size = new System.Drawing.Size(629, 23);
            this.txtDataDir.TabIndex = 99;
            // 
            // btnExecuteAction
            // 
            this.btnExecuteAction.Location = new System.Drawing.Point(647, 424);
            this.btnExecuteAction.Name = "btnExecuteAction";
            this.btnExecuteAction.Size = new System.Drawing.Size(110, 32);
            this.btnExecuteAction.TabIndex = 5;
            this.btnExecuteAction.Text = "Execute Action";
            this.btnExecuteAction.Click += new System.EventHandler(this.btnExecuteAction_Click);
            // 
            // lstFvsInput
            // 
            this.lstFvsInput.CheckBoxes = true;
            this.lstFvsInput.GridLines = true;
            this.lstFvsInput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsInput.HideSelection = false;
            this.lstFvsInput.Location = new System.Drawing.Point(6, 84);
            this.lstFvsInput.MultiSelect = false;
            this.lstFvsInput.Name = "lstFvsInput";
            this.lstFvsInput.Size = new System.Drawing.Size(740, 239);
            this.lstFvsInput.TabIndex = 0;
            this.lstFvsInput.UseCompatibleStateImageBehavior = false;
            this.lstFvsInput.View = System.Windows.Forms.View.Details;
            this.lstFvsInput.SelectedIndexChanged += new System.EventHandler(this.lstFvsInput_SelectedIndexChanged);
            this.lstFvsInput.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFvsInput_MouseUp);
            // 
            // cmbAction
            // 
            this.cmbAction.FormattingEnabled = true;
            this.cmbAction.Items.AddRange(new object[] {
            "Create FVS Input Database Files",
            "Create FVS Input Database Files From FIA2FVS"});
            this.cmbAction.Location = new System.Drawing.Point(279, 429);
            this.cmbAction.Name = "cmbAction";
            this.cmbAction.Size = new System.Drawing.Size(362, 24);
            this.cmbAction.TabIndex = 4;
            this.cmbAction.Text = "<-------Action Items------->";
            this.cmbAction.SelectedIndexChanged += new System.EventHandler(this.cmbAction_SelectedIndexChanged);
            this.cmbAction.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbAction_KeyPress);
            // 
            // lblRxCnt
            // 
            this.lblRxCnt.BackColor = System.Drawing.Color.White;
            this.lblRxCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxCnt.Location = new System.Drawing.Point(191, 12);
            this.lblRxCnt.Name = "lblRxCnt";
            this.lblRxCnt.Size = new System.Drawing.Size(32, 16);
            this.lblRxCnt.TabIndex = 99;
            this.lblRxCnt.Text = "0";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 24);
            this.label1.TabIndex = 99;
            this.label1.Text = "Data Directory";
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(6, 340);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(75, 32);
            this.btnChkAll.TabIndex = 1;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // btnRxPackage
            // 
            this.btnRxPackage.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRxPackage.Location = new System.Drawing.Point(229, 6);
            this.btnRxPackage.Name = "btnRxPackage";
            this.btnRxPackage.Size = new System.Drawing.Size(90, 30);
            this.btnRxPackage.TabIndex = 99;
            this.btnRxPackage.Text = "Packages";
            this.btnRxPackage.Click += new System.EventHandler(this.btnRxPackage_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(85, 340);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(75, 32);
            this.btnClearAll.TabIndex = 2;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // lblRxPackageCnt
            // 
            this.lblRxPackageCnt.BackColor = System.Drawing.Color.White;
            this.lblRxPackageCnt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRxPackageCnt.Location = new System.Drawing.Point(324, 11);
            this.lblRxPackageCnt.Name = "lblRxPackageCnt";
            this.lblRxPackageCnt.Size = new System.Drawing.Size(32, 16);
            this.lblRxPackageCnt.TabIndex = 99;
            this.lblRxPackageCnt.Text = "0";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(165, 340);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(64, 32);
            this.btnRefresh.TabIndex = 3;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnRx
            // 
            this.btnRx.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRx.Location = new System.Drawing.Point(96, 6);
            this.btnRx.Name = "btnRx";
            this.btnRx.Size = new System.Drawing.Size(90, 30);
            this.btnRx.TabIndex = 99;
            this.btnRx.Text = "Treatments";
            this.btnRx.Click += new System.EventHandler(this.btnRx_Click);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.otherOptionsGroupBox);
            this.tabPage1.Controls.Add(this.grpGRMOptions);
            this.tabPage1.Controls.Add(this.grpDWMOptions);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(787, 466);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "FVSIn Options";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // otherOptionsGroupBox
            // 
            this.otherOptionsGroupBox.Controls.Add(this.chkCull);
            this.otherOptionsGroupBox.Location = new System.Drawing.Point(451, 96);
            this.otherOptionsGroupBox.Name = "otherOptionsGroupBox";
            this.otherOptionsGroupBox.Size = new System.Drawing.Size(330, 288);
            this.otherOptionsGroupBox.TabIndex = 103;
            this.otherOptionsGroupBox.TabStop = false;
            this.otherOptionsGroupBox.Text = "Other Options";
            // 
            // chkCull
            // 
            this.chkCull.AutoSize = true;
            this.chkCull.Checked = true;
            this.chkCull.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCull.Location = new System.Drawing.Point(6, 25);
            this.chkCull.Name = "chkCull";
            this.chkCull.Size = new System.Drawing.Size(295, 21);
            this.chkCull.TabIndex = 3;
            this.chkCull.Text = "Account for tree-level percent defect (cull)";
            this.chkCull.UseVisualStyleBackColor = true;
            // 
            // grpGRMOptions
            // 
            this.grpGRMOptions.Controls.Add(this.chkGRM);
            this.grpGRMOptions.Location = new System.Drawing.Point(451, 6);
            this.grpGRMOptions.Name = "grpGRMOptions";
            this.grpGRMOptions.Size = new System.Drawing.Size(330, 85);
            this.grpGRMOptions.TabIndex = 102;
            this.grpGRMOptions.TabStop = false;
            this.grpGRMOptions.Text = "Growth Removal Mortality";
            // 
            // chkGRM
            // 
            this.chkGRM.AutoSize = true;
            this.chkGRM.Location = new System.Drawing.Point(6, 20);
            this.chkGRM.Name = "chkGRM";
            this.chkGRM.Size = new System.Drawing.Size(263, 21);
            this.chkGRM.TabIndex = 2;
            this.chkGRM.Text = "Use GRM calibration data if available";
            this.chkGRM.UseVisualStyleBackColor = true;
            // 
            // grpDWMOptions
            // 
            this.grpDWMOptions.Controls.Add(this.linkLabelFuelModel);
            this.grpDWMOptions.Controls.Add(this.groupBox2);
            this.grpDWMOptions.Controls.Add(this.chkDwmFuelModel);
            this.grpDWMOptions.Controls.Add(this.chkDwmFuelBiomass);
            this.grpDWMOptions.Controls.Add(this.label4);
            this.grpDWMOptions.Controls.Add(this.label2);
            this.grpDWMOptions.Controls.Add(this.txtMinLargeFwdTL);
            this.grpDWMOptions.Controls.Add(this.txtMinCwdTL);
            this.grpDWMOptions.Controls.Add(this.label3);
            this.grpDWMOptions.Controls.Add(this.txtMinSmallFwdTL);
            this.grpDWMOptions.Location = new System.Drawing.Point(6, 6);
            this.grpDWMOptions.Name = "grpDWMOptions";
            this.grpDWMOptions.Size = new System.Drawing.Size(439, 378);
            this.grpDWMOptions.TabIndex = 101;
            this.grpDWMOptions.TabStop = false;
            this.grpDWMOptions.Text = "Down Woody Material";
            // 
            // linkLabelFuelModel
            // 
            this.linkLabelFuelModel.AutoSize = true;
            this.linkLabelFuelModel.LinkArea = new System.Windows.Forms.LinkArea(8, 23);
            this.linkLabelFuelModel.Location = new System.Drawing.Point(23, 20);
            this.linkLabelFuelModel.Name = "linkLabelFuelModel";
            this.linkLabelFuelModel.Size = new System.Drawing.Size(477, 20);
            this.linkLabelFuelModel.TabIndex = 105;
            this.linkLabelFuelModel.TabStop = true;
            this.linkLabelFuelModel.Text = "Include Scott and Burgan (2005) surface fuel model (from DWM_fuelbed_typcd)";
            this.linkLabelFuelModel.UseCompatibleTextRendering = true;
            this.linkLabelFuelModel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelFuelModel_LinkClicked);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.chkLstBoxDuffYears);
            this.groupBox2.Controls.Add(this.chkLstBoxLitterYears);
            this.groupBox2.Location = new System.Drawing.Point(7, 142);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(236, 229);
            this.groupBox2.TabIndex = 104;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Duff/Litter Years to Exclude";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(123, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(40, 17);
            this.label6.TabIndex = 105;
            this.label6.Text = "Litter";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 16);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(34, 17);
            this.label5.TabIndex = 104;
            this.label5.Text = "Duff";
            // 
            // chkLstBoxDuffYears
            // 
            this.chkLstBoxDuffYears.FormattingEnabled = true;
            this.chkLstBoxDuffYears.Location = new System.Drawing.Point(9, 35);
            this.chkLstBoxDuffYears.Name = "chkLstBoxDuffYears";
            this.chkLstBoxDuffYears.Size = new System.Drawing.Size(100, 174);
            this.chkLstBoxDuffYears.TabIndex = 6;
            // 
            // chkLstBoxLitterYears
            // 
            this.chkLstBoxLitterYears.FormattingEnabled = true;
            this.chkLstBoxLitterYears.Location = new System.Drawing.Point(126, 35);
            this.chkLstBoxLitterYears.Name = "chkLstBoxLitterYears";
            this.chkLstBoxLitterYears.Size = new System.Drawing.Size(100, 174);
            this.chkLstBoxLitterYears.TabIndex = 7;
            // 
            // chkDwmFuelModel
            // 
            this.chkDwmFuelModel.AutoSize = true;
            this.chkDwmFuelModel.Checked = true;
            this.chkDwmFuelModel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwmFuelModel.Location = new System.Drawing.Point(7, 19);
            this.chkDwmFuelModel.Name = "chkDwmFuelModel";
            this.chkDwmFuelModel.Size = new System.Drawing.Size(18, 17);
            this.chkDwmFuelModel.TabIndex = 1;
            this.chkDwmFuelModel.UseVisualStyleBackColor = true;
            // 
            // chkDwmFuelBiomass
            // 
            this.chkDwmFuelBiomass.AutoSize = true;
            this.chkDwmFuelBiomass.Location = new System.Drawing.Point(7, 39);
            this.chkDwmFuelBiomass.Name = "chkDwmFuelBiomass";
            this.chkDwmFuelBiomass.Size = new System.Drawing.Size(417, 21);
            this.chkDwmFuelBiomass.TabIndex = 2;
            this.chkDwmFuelBiomass.Text = "Calculate fuel biomasses with available DWM data (tons/acre)";
            this.chkDwmFuelBiomass.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label4.Location = new System.Drawing.Point(58, 119);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(229, 17);
            this.label4.TabIndex = 99;
            this.label4.Tag = "txtMinCwdTL";
            this.label4.Text = "Minimum CWD Transect Length (ft)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label2.Location = new System.Drawing.Point(58, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(266, 17);
            this.label2.TabIndex = 99;
            this.label2.Tag = "txtMinSmallFwdTL";
            this.label2.Text = "Minimum Small FWD Transect Length (ft)";
            // 
            // txtMinLargeFwdTL
            // 
            this.txtMinLargeFwdTL.Location = new System.Drawing.Point(7, 90);
            this.txtMinLargeFwdTL.Name = "txtMinLargeFwdTL";
            this.txtMinLargeFwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinLargeFwdTL.TabIndex = 4;
            this.txtMinLargeFwdTL.Text = "30";
            this.txtMinLargeFwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinLargeFwdTL_Validating);
            // 
            // txtMinCwdTL
            // 
            this.txtMinCwdTL.Location = new System.Drawing.Point(7, 116);
            this.txtMinCwdTL.Name = "txtMinCwdTL";
            this.txtMinCwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinCwdTL.TabIndex = 5;
            this.txtMinCwdTL.Text = "48";
            this.txtMinCwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinCwdTL_Validating);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label3.Location = new System.Drawing.Point(58, 93);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(269, 17);
            this.label3.TabIndex = 99;
            this.label3.Tag = "txtMinLargeFwdTL";
            this.label3.Text = "Minimum Large FWD Transect Length (ft)";
            // 
            // txtMinSmallFwdTL
            // 
            this.txtMinSmallFwdTL.Location = new System.Drawing.Point(7, 65);
            this.txtMinSmallFwdTL.Name = "txtMinSmallFwdTL";
            this.txtMinSmallFwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinSmallFwdTL.TabIndex = 3;
            this.txtMinSmallFwdTL.Text = "10";
            this.txtMinSmallFwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinSmallFwdTL_Validating);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(567, 559);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(56, 24);
            this.btnCancel.TabIndex = 99;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(312, 559);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(240, 8);
            this.progressBar1.TabIndex = 99;
            this.progressBar1.Visible = false;
            // 
            // lblProgress
            // 
            this.lblProgress.Location = new System.Drawing.Point(312, 575);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(239, 16);
            this.lblProgress.TabIndex = 99;
            this.lblProgress.Text = "lblProgress";
            this.lblProgress.Visible = false;
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(7, 559);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(96, 32);
            this.btnHelp.TabIndex = 99;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(678, 555);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 6;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(794, 32);
            this.lblTitle.TabIndex = 99;
            this.lblTitle.Text = "Create FVS Input";
            // 
            // uc_fvs_input
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_input";
            this.Size = new System.Drawing.Size(800, 600);
            this.Resize += new System.EventHandler(this.uc_fvs_input_Resize);
            this.groupBox1.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.otherOptionsGroupBox.ResumeLayout(false);
            this.otherOptionsGroupBox.PerformLayout();
            this.grpGRMOptions.ResumeLayout(false);
            this.grpGRMOptions.PerformLayout();
            this.grpDWMOptions.ResumeLayout(false);
            this.grpDWMOptions.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        private void uc_fvs_input_Resize(object sender, System.EventArgs e)
        {
            this.Resize_Fvs_Input();
            this.tabControl1_Resize(sender, e);
        }

        public void Resize_Fvs_Input()
        {
            try
            {
                progressBar1.Left = (int)(groupBox1.Width * .50) - (int)(progressBar1.Width * .50);
                progressBar1.Top = btnClose.Top;

                if (progressBar1.Left < (btnHelp.Left + btnHelp.Width))
                {
                    progressBar1.Left = btnHelp.Left + btnHelp.Width;
                }
                btnCancel.Left = progressBar1.Left + progressBar1.Width + 2;
                btnCancel.Top = progressBar1.Top - (int)(btnCancel.Height * .50) + (int)(progressBar1.Height * .50);
                if (btnClose.Left < (btnCancel.Left + btnCancel.Width))
                {
                    btnClose.Left = btnCancel.Left + btnCancel.Width;
                }
                lblProgress.Left = progressBar1.Left;
                lblProgress.Top = progressBar1.Top + progressBar1.Height + 2;
            }
            catch
            {
            }
        }

        public void loadvalues()
        {
            this.LoadDataSources();
            this.populate_listbox();
            this.populate_FIA2FVS_combobox();
        }

        private void populate_FIA2FVS_combobox()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string applicationDb = frmMain.g_oEnv.strAppDir + "\\db\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(applicationDb)))
            {
                con.Open();
                if (!oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputKeywordsTableName))
                {
                    MessageBox.Show("The template FVSIn.db could not be found in the BioSum installation directory!",
                        "FIA Biosum");
                    return;
                }
                else
                {
                    oDataMgr.SqlQueryReader(con, "SELECT GROUPS FROM " + Tables.FIA2FVS.DefaultFvsInputKeywordsTableName);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        while (oDataMgr.m_DataReader.Read())
                        {
                            if (oDataMgr.m_DataReader["GROUPS"] != DBNull.Value &&
                                Convert.ToString(oDataMgr.m_DataReader["GROUPS"]).Trim().Length > 0)
                            {
                                cmbSelectedGroup.Items.Add(Convert.ToString(oDataMgr.m_DataReader["GROUPS"]).Trim());
                            }
                        }
                        if (cmbSelectedGroup.Items.Count > 0)
                        {
                            cmbSelectedGroup.SelectedIndex = 0;
                        }
                    }
                }
            }
        }

        private void populate_listbox()
        {
            //bool bResult;
            string strInDirAndFile;
            string strOutDirAndFile;
            string strConn;
            string[] strValues;
            bool bFoundDsnOut;
            string strVariant = "";
            string strCurrentVariant = "";
            string strRecordCount = "";
            //bool bFoundDsnIn;
            try
            {

                m_ado.OpenConnection(m_strConn);

                this.lblRxCnt.Text = Convert.ToString((int)this.m_ado.getRecordCount(m_ado.m_OleDbConnection, "select count(*) from " + m_oQueries.m_oFvs.m_strRxTable + " where rx IS NOT NULL AND LEN(TRIM(rx)) > 0;", m_oQueries.m_oFvs.m_strRxTable));
                this.lblRxPackageCnt.Text = Convert.ToString((int)this.m_ado.getRecordCount(m_ado.m_OleDbConnection, "select count(*) from " + m_oQueries.m_oFvs.m_strRxPackageTable + " WHERE rxpackage IS NOT NULL AND LEN(TRIM(rxpackage)) > 0;", m_oQueries.m_oFvs.m_strRxTable));
                this.lstFvsInput.Clear();
                this.m_oLvRowColors.InitializeRowCollection();


                this.lstFvsInput.Columns.Add("", 2, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Variant", 55, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Package", 55, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Stands", 70, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Trees", 70, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("FVS Output DB", 230, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("SUMMARY RECS", 100, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("CUTLIST RECS", 100, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("TREELIST RECS", 100, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("POTFIRE RECS", 100, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("POTFIRE BaseYr DB", 120, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("POTFIRE BaseYr RECS", 120, HorizontalAlignment.Left);

                this.lstFvsInput.Columns[COL_CHECKBOX].Width = -2;

                this.m_ado.m_strSQL = Queries.FVS.GetFVSVariantRxPackageSQL(m_oQueries.m_oFIAPlot.m_strPlotTable, m_oQueries.m_oFvs.m_strRxPackageTable);
                this.m_ado.SqlQueryReader(m_ado.m_OleDbConnection, this.m_ado.m_strSQL);


                //declare a registry key object
                Microsoft.Win32.RegistryKey regKey; // new Microsoft.Win32 Registry Key 
                Microsoft.Win32.RegistryKey regKey2;
                // open the subkey that holds the current odbc data sources
                regKey = Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.Ini\Odbc data sources", false);
                // get string name in string array
                // then use getValue to get data value.
                //save the odbc datasources in strDsnNames
                string[] strDsnNames = regKey.GetValueNames(); // for the 1st time it's only 2 names

                string strRegKey = "";

                //Keep a count of records in FVS_StandInit and FVS_TreeInit tables in each variant
                Dictionary<string, int[]> pVariantCountsDict = new Dictionary<string, int[]>();

                while (this.m_ado.m_OleDbDataReader.Read())
                {
                    strRegKey = "";
                    bFoundDsnOut = false;

                    // Add a ListItem object to the ListView.

                    //add new row
                    System.Windows.Forms.ListViewItem entryListItem =
                        this.lstFvsInput.Items.Add("");
                    entryListItem.UseItemStyleForSubItems = false;
                    this.m_oLvRowColors.AddRow();
                    this.m_oLvRowColors.AddColumns(lstFvsInput.Items.Count - 1, lstFvsInput.Columns.Count);
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_CHECKBOX, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);

                    //fvs_variant		
                    strVariant = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                    if (pVariantCountsDict.ContainsKey(strVariant) == false)
                    {
                        pVariantCountsDict.Add(strVariant, null); //fvs_standinit, fvs_treeinit counts
                    }

                    entryListItem.SubItems.Add(this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_VARIANT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    if (!System.IO.Directory.Exists(txtDataDir.Text.Trim() + "\\" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim()))
                        System.IO.Directory.CreateDirectory(txtDataDir.Text.Trim() + "\\" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim());
                    //rx
                    entryListItem.SubItems.Add(this.m_ado.m_OleDbDataReader["rxpackage"].ToString().Trim());
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, COL_RX, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //simulation year cycle
                    if (this.m_ado.m_OleDbDataReader["rxcycle_length"] != System.DBNull.Value)
                    {
                        this.m_strFVSCycleLength = Convert.ToString(this.m_ado.m_OleDbDataReader["rxcycle_length"]).Trim();
                    }
                    //FVS_StandInit Stand_ID count
                    entryListItem.SubItems.Add(" ");
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_STANDCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //FVS_TreeInit row count
                    entryListItem.SubItems.Add(" ");
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_TREECOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //out mdb file name
                    entryListItem.SubItems.Add(" ");  //out mdb file
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_MDBOUT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //summary record count
                    entryListItem.SubItems.Add(" ");  //summary record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_SUMMARYCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //treecut list record count
                    entryListItem.SubItems.Add(" ");  //tree cut list record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_CUTCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //tree standing (uncut) record count
                    entryListItem.SubItems.Add(" ");  //tree standing record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_LEFTCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //potential fire record count
                    entryListItem.SubItems.Add(" ");  //potential fire record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_POTFIRECOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //potential fire base year MDB file
                    entryListItem.SubItems.Add(" ");  //potential fire record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_POTFIREMDBOUT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //potential fire base year table count
                    entryListItem.SubItems.Add(" ");  //potential fire record count
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_POTFIREBASEYEARCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);

                    //check to see if there is an input and output dsn name
                    this.m_strLocFile = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + ".loc";
                    this.m_strSlfFile = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + ".slf";
                    this.m_strOutMDBFile = this.m_oRxTools.GetRxPackageFvsOutDbFileName(m_ado.m_OleDbDataReader);

                    strOutDirAndFile = this.txtDataDir.Text.Trim() + "\\" + m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "\\" + this.m_strOutMDBFile.Trim();

                    frmMain.g_sbpInfo.Text = "Processing FVS Input Variant/RxPackage " +
                        this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "/" +
                        this.m_ado.m_OleDbDataReader["rxpackage"].ToString().Trim() + "...Stand By";

                    //check fvs in values
                    strInDirAndFile = this.txtDataDir.Text.Trim() + "\\" + this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim() + "\\" + "FVSIn.accdb";
                    if (frmMain.g_bSuppressFVSInputTableRowCount == false && System.IO.File.Exists(strInDirAndFile) == true)
                    {
                        if (pVariantCountsDict[strVariant] == null)
                        {
                            pVariantCountsDict[strVariant] = getFVSInputRecordCounts(strInDirAndFile);
                        }
                        entryListItem.SubItems[COL_STANDCOUNT].Text = Convert.ToString(pVariantCountsDict[strVariant][0]);
                        entryListItem.SubItems[COL_TREECOUNT].Text = Convert.ToString(pVariantCountsDict[strVariant][1]);
                    }

                    //check dsn out registry values
                    foreach (string strDsnName in strDsnNames)
                    {
                        //dsn in
                        //dsn out
                        if (this.m_strDsnOut.Trim().ToUpper() == strDsnName.Trim().ToUpper() &&
                            regKey.GetValue(strDsnName).ToString().Trim().ToUpper() ==
                            "MICROSOFT ACCESS DRIVER (*.MDB)")
                        {
                            strRegKey = "Software\\ODBC\\ODBC.ini\\" + this.m_strDsnOut.Trim();
                            regKey2 = Registry.CurrentUser.OpenSubKey(strRegKey, false);
                            strValues = regKey2.GetValueNames(); // for the 1st time it's only 2 names
                            foreach (string strValue in strValues)
                            {
                                if (strValue.Trim().ToUpper() == "DBQ")
                                {
                                    if (strOutDirAndFile.Trim().ToUpper() ==
                                        regKey2.GetValue(strValue).ToString().Trim().ToUpper())
                                    {
                                        bFoundDsnOut = true;
                                        break;
                                    }
                                }
                            }
                        }

                        if (bFoundDsnOut == true) break;
                    }


                    if (System.IO.File.Exists(strOutDirAndFile) == true)
                    {
                        strConn = this.m_ado.getMDBConnString(strOutDirAndFile, "", "");
                        entryListItem.SubItems[COL_MDBOUT].Text = this.m_strOutMDBFile;
                        if (frmMain.g_bSuppressFVSInputTableRowCount == false)
                        {
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_summary") == true)
                            {
                                entryListItem.SubItems[COL_SUMMARYCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_summary", "fvs_summary")));
                            }
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_cutlist") == true)
                            {
                                entryListItem.SubItems[COL_CUTCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_cutlist", "fvs_cutlist")));
                            }
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_treelist") == true)
                            {
                                entryListItem.SubItems[COL_LEFTCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_treelist", "fvs_treelist")));
                            }
                            if (this.m_dao.TableExists(strOutDirAndFile, "fvs_potfire") == true)
                            {
                                entryListItem.SubItems[COL_POTFIRECOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_potfire", "fvs_potfire")));
                            }
                        }
                    }
                    else
                    {
                    }
                    //
                    //POTFIRE BASE YEAR
                    //
                    m_strOutPotFireBaseYearMDBFile = "FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                    strOutDirAndFile = this.txtDataDir.Text.Trim() + "\\" + strVariant + "\\" + m_strOutPotFireBaseYearMDBFile.Trim();
                    if (strVariant != strCurrentVariant)
                    {
                        strCurrentVariant = strVariant;
                        strRecordCount = "";

                        if (System.IO.File.Exists(strOutDirAndFile) == true)
                        {
                            strConn = this.m_ado.getMDBConnString(strOutDirAndFile, "", "");
                            entryListItem.SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text = this.m_strOutPotFireBaseYearMDBFile;
                            if (frmMain.g_bSuppressFVSInputTableRowCount == false &&
                                this.m_dao.TableExists(strOutDirAndFile, "fvs_potfire") == true)
                            {
                                entryListItem.SubItems[uc_fvs_input.COL_POTFIREBASEYEARCOUNT].Text = Convert.ToString(Convert.ToInt32(this.m_ado.getRecordCount(strConn, "select count(*) from fvs_potfire", "fvs_potfire")));
                                strRecordCount = entryListItem.SubItems[uc_fvs_input.COL_POTFIREBASEYEARCOUNT].Text;
                            }
                        }
                    }
                    else
                    {
                        if (System.IO.File.Exists(strOutDirAndFile) == true)
                        {
                            entryListItem.SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text = this.m_strOutPotFireBaseYearMDBFile;
                            entryListItem.SubItems[uc_fvs_input.COL_POTFIREBASEYEARCOUNT].Text = strRecordCount;
                        }
                    }

                }
                this.m_ado.m_OleDbDataReader.Close();
                this.m_ado.CloseConnection(m_ado.m_OleDbConnection);
                Registry.CurrentUser.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_fvs_input:populate_listbox() \n" +
                                "Err Msg - " + e.Message,
                                "Create FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                                System.Windows.Forms.MessageBoxIcon.Exclamation);

                this.m_intError = -1;
            }
            this.Refresh();

        }

        private void LoadDataSources()
        {
            this.txtDataDir.Text = this.m_strProjDir + "\\fvs\\data";
            this.m_oQueries.m_oFvs.LoadDatasource = true;
            this.m_oQueries.m_oFIAPlot.LoadDatasource = true;
            this.m_oQueries.LoadDatasources(true);
            this.m_strConn = m_ado.getMDBConnString(this.m_oQueries.m_strTempDbFile, "", "");




        }

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            frmMain.g_oFrmMain.Enabled = true;
            this.ParentForm.Dispose();

        }

        private void btnRefresh_Click(object sender, System.EventArgs e)
        {
            this.populate_listbox();
        }

        private void btnChkAll_Click(object sender, System.EventArgs e)
        {
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                this.lstFvsInput.Items[x].Checked = true;
            }
        }

        private void btnClearAll_Click(object sender, System.EventArgs e)
        {
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                this.lstFvsInput.Items[x].Checked = false;
            }
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {

            string strMsg = "Do you wish to cancel process (Y/N)?";
            DialogResult result = MessageBox.Show(strMsg, "Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            switch (result)
            {
                case DialogResult.Yes:
                    this.bAbort = true;


                    return;
                case DialogResult.No:

                    return;
            }
        }
        private void CancelThreadCleanup()
        {

            this.cmbAction.Enabled = true;
            this.btnRefresh.Enabled = true;
            this.btnExecuteAction.Enabled = true;
            this.btnChkAll.Enabled = true;
            this.btnClearAll.Enabled = true;
            this.btnClose.Enabled = true;
            this.btnHelp.Enabled = true;

            this.progressBar1.Visible = false;
            this.lblProgress.Visible = false;
            this.btnCancel.Visible = false;

            this.m_thread = null;
        }

        public void CreateFia2FvsInputFiles()
        {
            // Make sure the database is chosen and that such a file exists
            bool bValidFile = true;
            if (!String.IsNullOrEmpty(txtFIADatamart.Text))
            {
                if (System.IO.File.Exists(txtFIADatamart.Text))
                {
                    // Do nothing; All is well
                }
                else
                {
                    bValidFile = false;
                }
            }
            else
            {
                bValidFile = false;
            }
            if (bValidFile == false)
            {
                MessageBox.Show("You must specify a source input database on the FVSIn FIA2FVS tab before proceeding!", "FIA Biosum");
                return;
            }

            // Make sure a group is selected
            if (cmbSelectedGroup.SelectedIndex < 0)
            {
                MessageBox.Show("You must specify a group on the FVSIn FIA2FVS tab before proceeding!", "FIA Biosum");
                return;
            }

            // Check for existing files
            List<string> lstVariants = new List<string>();
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                if (this.lstFvsInput.Items[x].Checked)
                {
                    var item = this.lstFvsInput.Items[x];
                    string strVariant = item.SubItems[1].Text.Trim();
                    string strInDirAndFile = this.txtDataDir.Text + "\\" + strVariant + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
                    if (!lstVariants.Contains(strInDirAndFile))
                    {
                        if (File.Exists(strInDirAndFile))
                        {
                            // Make sure we have both the required tables
                            string connFiadbDb = oDataMgr.GetConnectionString(strInDirAndFile);
                            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(connFiadbDb))
                            {
                                con.Open();
                                if (oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputStandTableName) &&
                                    oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputTreeTableName))
                                {
                                    // Only add to warnings if both tables exist
                                    lstVariants.Add(strInDirAndFile);
                                }
                            }
                        }                        
                    }
                }
            }

            if (lstVariants.Count > 0)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.Append("The FVS input databases listed below exist for the selected variants. ");
                sb.Append("Click 'Yes' to append data to the existing databases or 'No' to overwrite them with new data. \r\n\r\n");
                foreach (var item in lstVariants)
                {
                    sb.Append(item + "\r");
                }
                DialogResult res = MessageBox.Show(sb.ToString(), "FIA BioSum", MessageBoxButtons.YesNoCancel);
                switch (res)
                {
                    case DialogResult.Cancel:
                        return;
                    case DialogResult.Yes:
                        m_bOverwrite = false;
                        break;
                    case DialogResult.No:
                        break;
                }

                if (!m_bOverwrite)
                {
                    m_dictVariantStates = new Dictionary<string, List<string>>();    // Re-initialize dictionary
                    string connFiadbDb = oDataMgr.GetConnectionString(this.txtFIADatamart.Text);
                    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(connFiadbDb))
                    {
                        con.Open();
                        foreach (var strVariant in lstVariants)
                        {
                            oDataMgr.SqlNonQuery(con, "ATTACH '" + strVariant + "' as target");
                            string strQuery = "SELECT distinct STATE FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                              " INTERSECT SELECT distinct STATE" +
                                              " FROM target." + Tables.FIA2FVS.DefaultFvsInputStandTableName;
                            List<string> lstStates = oDataMgr.getStringList(con, strQuery);
                            if (lstStates.Count > 0)
                            {
                                m_dictVariantStates.Add(strVariant, lstStates);
                            }
                            oDataMgr.SqlNonQuery(con, "DETACH DATABASE 'target'");
                        }
                    }
                }
                if (m_dictVariantStates != null && 
                    m_dictVariantStates.Keys.Count > 0)
                {
                    sb = new System.Text.StringBuilder();
                    sb.Append("The FVS input databases listed below already contain records for state codes ");
                    sb.Append("in the FIA Datamart input SQLite database. ");
                    sb.Append("Click 'Yes' to replace the existing stands ");
                    sb.Append("and trees for those states or 'No' to stop this process.\r\n\r\n");
                    foreach (var strVariant in m_dictVariantStates.Keys)
                    {
                        sb.Append(strVariant + "\r");
                    }
                    res = MessageBox.Show(sb.ToString(), "FIA BioSum", MessageBoxButtons.YesNoCancel);
                    switch (res)
                    {
                        case DialogResult.Cancel:
                            return;
                        case DialogResult.Yes:
                            break;
                        case DialogResult.No:
                            return;
                    }
                }

            }

            this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "EXTRACT FIA2FVS FVS INPUT FILE",
                             "FIA2FVS FVS Input", "2");

            this.m_frmTherm.Visible = false;
            this.m_frmTherm.lblMsg.Text = "";
            this.Enabled = false;

            //progress bar 1: represents a single process
            this.m_frmTherm.progressBar1.Minimum = 0;
            this.m_frmTherm.progressBar1.Maximum = 100;
            this.m_frmTherm.progressBar1.Value = 0;
            this.m_frmTherm.lblMsg.Text = "";
            this.m_frmTherm.Show(this);

            //progress bar 2: represents overall progress 
            this.m_frmTherm.progressBar2.Minimum = 0;
            this.m_frmTherm.progressBar2.Maximum = 100;
            this.m_frmTherm.progressBar2.Value = 0;
            this.m_frmTherm.lblMsg2.Text = "Overall Progress";
            this.m_thread = new Thread(new ThreadStart(ExtractFIA2FVSRecords));
            this.m_thread.IsBackground = true;
            this.m_thread.Start();

        }

        private string[] GetCheckedListBoxItems(CheckedListBox chkListBox)
        {
            if (chkListBox.InvokeRequired)
            {
                var dlg = new GetListBoxItemsDlg(GetCheckedListBoxItems);
                return chkListBox.Invoke(dlg, chkListBox) as string[];
            }

            string[] items = (from object item in chkListBox.CheckedItems select item.ToString())
                .ToArray();
            return items;
        }

        private void ConfigureFvsInput(fvs_input p_fvs)
        {
            //Down Woody Materials Section
            if (chkDwmFuelModel.Checked && chkDwmFuelBiomass.Checked)
            {
                p_fvs.intDWMOption = (int)fvs_input.m_enumDWMOption.USE_FUEL_MODEL_OR_DWM_DATA;
            }
            else if (chkDwmFuelModel.Checked)
            {
                p_fvs.intDWMOption = (int)fvs_input.m_enumDWMOption.USE_FUEL_MODEL_ONLY;
            }
            else if (chkDwmFuelBiomass.Checked)
            {
                p_fvs.intDWMOption = (int)fvs_input.m_enumDWMOption.USE_DWM_DATA_ONLY;
            }
            else
            {
                p_fvs.intDWMOption = (int)fvs_input.m_enumDWMOption.SKIP_FUEL_MODEL_AND_DWM_DATA;
            }

            p_fvs.strMinSmallFwdTransectLengthTotal =
                frmMain.g_oDelegate.GetControlPropertyValue(txtMinSmallFwdTL, "Text", false).ToString();
            p_fvs.strMinLargeFwdTransectLengthTotal =
                frmMain.g_oDelegate.GetControlPropertyValue(txtMinLargeFwdTL, "Text", false).ToString();
            p_fvs.strMinCwdTransectLengthTotal =
                frmMain.g_oDelegate.GetControlPropertyValue(txtMinCwdTL, "Text", false).ToString();

            bool bFirst = true;
            foreach (var item in GetCheckedListBoxItems(chkLstBoxDuffYears))
            {
                if (bFirst)
                {
                    p_fvs.strDuffExcludedYears += item;
                    bFirst = false;
                }
                else
                {
                    p_fvs.strDuffExcludedYears += ", " + item;
                }
            }
            bFirst = true;
            foreach (var item in GetCheckedListBoxItems(chkLstBoxLitterYears))
            {
                if (bFirst)
                {
                    p_fvs.strLitterExcludedYears += item;
                    bFirst = false;
                }
                else
                {
                    p_fvs.strLitterExcludedYears += ", " + item;
                }
            }

            //Growth Removal Mortality section
            p_fvs.bUseGrmCalibrationData = (bool)frmMain.g_oDelegate.GetControlPropertyValue(chkGRM, "Checked", false);
            p_fvs.bUseCullDefect = (bool)frmMain.g_oDelegate.GetControlPropertyValue(chkCull, "Checked", false);
            p_fvs.strSourceFiaDb = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtFIADatamart, "Text", false);
            p_fvs.strSourceFiaDb = p_fvs.strSourceFiaDb.Trim();
            p_fvs.strDataDirectory = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtDataDir, "Text", false);
            p_fvs.strDataDirectory = p_fvs.strDataDirectory.Trim();
            p_fvs.strGroup = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.cmbSelectedGroup, "SelectedItem", false);
            p_fvs.strGroup = p_fvs.strGroup.Trim();
        }

        private void AppendRecords()
        {
            m_intError = 0;
            string strCurVariant = "";
            string strVariant = "";
            string strSavedir = this.m_strProjDir + "\\fvs\\data\\save";

            bAbort = false;
            try
            {
                fvs_input p_fvsinput = new fvs_input(this.m_strProjDir, this.m_frmTherm);

                if (p_fvsinput.m_intError != 0)
                {
                    return;
                }

                ConfigureFvsInput(p_fvsinput);

                for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
                {
                    int intValue = Convert.ToInt32((double)(((double)(x + 1) / (double)this.lstFvsInput.Items.Count) * 100));
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar2, "Value", intValue);
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(lstFvsInput, x, "Checked", false) == true)
                    {
                        //get the variant
                        strVariant = frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsInput, x, COL_VARIANT, "Text", false).ToString().Trim();
                        //see if this is a new variant
                        if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
                        {
                            strCurVariant = strVariant;
                            m_strDebugFile = this.strProjectDirectory + Tables.FIA2FVS.DefaultFvsInputFolderName + "\\biosum_fvs_input_debug.txt";
                            if (File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);
                            p_fvsinput.Start(strCurVariant, m_strDebugFile);
                        }
                        frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar1, "Maximum", false));

                       string strInDirAndFile = p_fvsinput.strDataDirectory + "\\" + strVariant + "\\" + "FVSIn.accdb";
                        if (System.IO.File.Exists(strInDirAndFile) == true) //redundant check here, but leaves " " instead of new "0"
                        {
                            int[] fvsInputRecordCounts = getFVSInputRecordCounts(strInDirAndFile);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_STANDCOUNT, "Text",
                                Convert.ToString(fvsInputRecordCounts[0]));
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_TREECOUNT, "Text",
                                Convert.ToString(fvsInputRecordCounts[1]));
                        }

                    }

                    frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar1, "Maximum", false));
                    System.Windows.Forms.Application.DoEvents();
                    if (bAbort == true) break;

                }


                frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar2,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar2, "Maximum", false));

            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                m_intError = -1;
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                m_intError = -1;
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_input:AppendRecords  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "Append Records", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {
                if (m_intError == 0)
                {
                    ThreadCleanUp();
                }
            }

        }

        private void ExtractFIA2FVSRecords()
        {
            m_intError = 0;
            string strCurVariant = "";
            string strVariant = "";
            //@ToDo: This will be set in AppendRecords when we merge the streams
            m_strDebugFile = this.strProjectDirectory + Tables.FIA2FVS.DefaultFvsInputFolderName + "\\biosum_fvs_input_debug.txt";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//ExtractFIA2FVSRecords\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();
            try
            {
                //string strSourceDbDir = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtFIADatamart, "Text", false);
                //strSourceDbDir = strSourceDbDir.Trim();
                fvs_input p_fvsinput = new fvs_input(this.m_strProjDir, this.m_frmTherm);
                // @ToDo: For now this is duplicated in this and the legacy process. Some day remove one of them
                ConfigureFvsInput(p_fvsinput);

                // Get a temporary database name for processing
                string strTempMDB = frmMain.g_oUtils.getRandomFile(this.m_oEnv.strTempDir, "accdb");
                // Create a temporary mdb that will contain all our required table links
                oDao.CreateMDB(strTempMDB);
                // Check to see if the input DSN exists and if so, delete so we can add
                ODBCMgr odbcmgr = new ODBCMgr();
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName, p_fvsinput.strSourceFiaDb);
                if (!string.IsNullOrEmpty(odbcmgr.m_strError))
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "ODBCMgr error: " + odbcmgr.m_strError + "\r\n");
                    return;
                }
                else
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Created DSN for " + ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName + "\r\n");
                }

                // Link to cond table
                m_strCondTable = m_oQueries.m_oFIAPlot.m_strCondTable;
                string strCondMdb = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Condition);
                oDao.CreateTableLink(strTempMDB, m_strCondTable, strCondMdb, m_strCondTable);
                // Link to tree table
                m_strTreeTable = m_oQueries.m_oFIAPlot.m_strTreeTable;
                string strTreeMdb = m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Tree);
                oDao.CreateTableLink(strTempMDB, m_strTreeTable, strTreeMdb, m_strTreeTable);

                // Link to the input SQLite table; Takes the whole path to the DB
                string strSourceStandTableAlias = Tables.FIA2FVS.DefaultFvsInputStandTableName + "_1";
                string strSourceTreeTableAlias = Tables.FIA2FVS.DefaultFvsInputTreeTableName + "_1";
                oDao.CreateSQLiteTableLink(strTempMDB, Tables.FIA2FVS.DefaultFvsInputStandTableName, strSourceStandTableAlias,
                    ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName, p_fvsinput.strSourceFiaDb);
                if (oDao.m_intError != 0)
                {
                    return;
                }

                // Set the index, required to by ODBC to update
                oDao.CreatePrimaryKeyIndex(strTempMDB, strSourceStandTableAlias, "STAND_CN");
                oDao.CreateSQLiteTableLink(strTempMDB, Tables.FIA2FVS.DefaultFvsInputTreeTableName, strSourceTreeTableAlias,
                    ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName, p_fvsinput.strSourceFiaDb);
                // Set the index, required to by ODBC to update
                oDao.CreatePrimaryKeyIndex(strTempMDB, strSourceTreeTableAlias, "TREE_CN");

                for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
                {
                    int intValue = Convert.ToInt32((double)(((double)(x + 1) / (double)this.lstFvsInput.Items.Count) * 100));
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar2, "Value", intValue);
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(lstFvsInput, x, "Checked", false) == true)
                    {
                        //get the variant
                        strVariant = frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsInput, x, COL_VARIANT, "Text", false).ToString().Trim();
                        string strInDirAndFile = p_fvsinput.strDataDirectory + "\\" + strVariant + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
                        //see if this is a new variant
                        if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(
                                m_frmTherm.progressBar1,
                                "Value", 20);
                            strCurVariant = strVariant;
                            List<string> lstStates = new List<string>();
                            if (m_dictVariantStates != null && 
                                m_dictVariantStates.ContainsKey(strInDirAndFile))
                            {
                                lstStates = m_dictVariantStates[strInDirAndFile];
                            }
                            p_fvsinput.StartFIA2FVS(odbcmgr, oDao, oAdo, strTempMDB, m_bOverwrite, m_strDebugFile, 
                                strCurVariant, lstStates, strSourceStandTableAlias, 
                                strSourceTreeTableAlias);
                        }
                        frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value", 7);

                        // This happens at the end
                        if (File.Exists(strInDirAndFile) == true) //redundant check here, but leaves " " instead of new "0"
                        {
                            int[] fvsInputRecordCounts = getFVSSQLiteInputRecordCounts(strInDirAndFile);
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_STANDCOUNT, "Text",
                                Convert.ToString(fvsInputRecordCounts[0]));
                            frmMain.g_oDelegate.SetListViewSubItemPropertyValue(this.lstFvsInput, x, COL_TREECOUNT, "Text",
                                Convert.ToString(fvsInputRecordCounts[1]));
                        }

                    }

                    frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar1, "Maximum", false));
                    Application.DoEvents();
                    if (bAbort == true) break;

                }

                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName);    // Clean up DSN
                if (!string.IsNullOrEmpty(odbcmgr.m_strError))
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "ODBCMgr error: " + odbcmgr.m_strError + "\r\n");
                }
                else
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Removed DSN for " + ODBCMgr.DSN_KEYS.Fia2FvsInputDsnName + "\r\n");
                }

                frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar2,
                            "Value",
                            frmMain.g_oDelegate.GetControlPropertyValue(
                                    m_frmTherm.progressBar2, "Maximum", false));

            }
            catch (ThreadInterruptedException err)
            {
                m_intError = -1;
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (ThreadAbortException)
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
                m_intError = -1;
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                    "Module - uc_fvs_input:ExtractFIA2FVSRecords  \n" +
                    "Err Msg - " + err.Message.ToString().Trim(),
                    "ExtractFIA2FVSRecords", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
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
                //destroy the object and release it from memory
                oDao.m_DaoWorkspace.Close();
                oDao.m_DaoWorkspace = null;
                oDao = null;
                if (m_intError == 0)
                {
                    ThreadCleanUp();
                }
            }
        }

        private void btnAppend_Click(object sender, System.EventArgs e)
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Append", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            val_data();
            if (this.m_intError == 0)
            {
                this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "APPEND FVS INPUT DATA",
                                                 "FVS Input", "2");


                this.m_frmTherm.Visible = false;
                this.m_frmTherm.lblMsg.Text = "";
                this.Enabled = false;

                //progress bar 1: represents a single process
                this.m_frmTherm.progressBar1.Minimum = 0;
                this.m_frmTherm.progressBar1.Maximum = 100;
                this.m_frmTherm.progressBar1.Value = 0;
                this.m_frmTherm.lblMsg.Text = "";
                this.m_frmTherm.Show(this);


                //progress bar 2: represents overall progress 
                this.m_frmTherm.progressBar2.Minimum = 0;
                this.m_frmTherm.progressBar2.Maximum = 100;
                this.m_frmTherm.progressBar2.Value = 0;
                this.m_frmTherm.lblMsg2.Text = "Overall Progress";
                this.m_thread = new Thread(new ThreadStart(this.AppendRecords));
                this.m_thread.IsBackground = true;
                this.m_thread.Start();
            }


        }
        private void val_data()
        {
            this.m_intError = 0;
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                if (this.lstFvsInput.Items[x].Checked == true)
                {
                }
            }
        }

        public void StopThread()
        {
            this.m_thread.Suspend();
            string strMsg = "Do you wish to cancel appending  data (Y/N)?";
            DialogResult result = MessageBox.Show(strMsg, "Cancel Process", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            switch (result)
            {
                case DialogResult.Yes:
                    this.m_thread.Resume();
                    this.m_frmTherm.AbortProcess = true;
                    this.m_thread.Abort();
                    if (this.m_thread.IsAlive)
                    {
                        this.m_frmTherm.lblMsg.Text = "Attempting To Abort Process...Stand By";
                        this.m_frmTherm.lblMsg.Refresh();
                        while (m_thread.ThreadState != ThreadState.Aborted)
                        {
                            this.m_thread.Join(2000);
                        }
                    }
                    if (this.m_frmTherm != null)
                    {
                        this.m_frmTherm.lblMsg.Text = "Cleaning Up...Stand By";
                        this.m_frmTherm.lblMsg.Refresh();
                    }
                    this.ThreadCleanUp();
                    return;
                case DialogResult.No:
                    this.m_thread.Resume();
                    return;
            }

        }
        public void InputFVSRecordsFinished()
        {

            this.m_thread.Abort();
            if (this.m_thread.IsAlive)
            {
                frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.lblMsg,
                            "Text", "Attempting To Abort Process...Stand By");

                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm.lblMsg, "Refresh");

                this.m_thread.Join();
            }
            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
            this.m_thread = null;

        }
        private void ThreadCleanUp()
        {
            try
            {
                if (this.m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                    this.m_frmTherm = null;
                }

                frmMain.g_oDelegate.SetControlPropertyValue(cmbAction, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnRefresh, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnExecuteAction, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnChkAll, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnClearAll, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnClose, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(btnHelp, "Enabled", true);
                frmMain.g_oDelegate.SetControlPropertyValue(progressBar1, "Visible", false);

                frmMain.g_oDelegate.SetControlPropertyValue(lblProgress, "Visible", false);
                frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);
                frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
                this.m_thread = null;
            }
            catch
            {
            }

        }

        private int[] getFVSInputRecordCounts(string strDirAndFile)
        {
            int stands = 0;
            int trees = 0;
            try
            {
                if (System.IO.File.Exists(strDirAndFile))
                {
                    ado_data_access p_ado = new ado_data_access();
                    string strConn = p_ado.getMDBConnString(strDirAndFile, "", "");

                    using (var pConn = new System.Data.OleDb.OleDbConnection(strConn))
                    {
                        pConn.Open();
                        if (p_ado.TableExist(pConn, "FVS_StandInit"))
                        {
                            stands = (int)p_ado.getSingleDoubleValueFromSQLQuery(pConn,
                            "SELECT COUNT(*) as StandInitCount FROM (SELECT DISTINCT Stand_ID FROM FVS_StandInit);",
                            "FVS_StandInit");
                        }
                        if (p_ado.TableExist(pConn, "FVS_TreeInit"))
                        {
                            trees = (int)p_ado.getSingleDoubleValueFromSQLQuery(pConn,
                            "SELECT COUNT(*) as TreeInitCount FROM FVS_TreeInit;",
                            "FVS_TreeInit");
                        }

                    }
                    p_ado = null;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_fvs_input:getFVSInputRecordCounts  \n" +
                                "Err Msg - " + e.Message.ToString().Trim(),
                    "FVS Input", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            return new int[] { stands, trees };
        }

        private int[] getFVSSQLiteInputRecordCounts(string strDirAndFile)
        {
            int stands = 0;
            int trees = 0;
            try
            {
                if (File.Exists(strDirAndFile))
                {
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    string connSourceDb = oDataMgr.GetConnectionString(strDirAndFile);
                    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(connSourceDb))
                    {
                        string strSql = "";
                        con.Open();
                        // Do stuff in sqlite side
                        if (oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputStandTableName))
                        {
                            strSql = "SELECT COUNT(*) as StandInitCount " +
                                     "FROM (SELECT DISTINCT STAND_CN FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName + ")";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                            stands = (int)oDataMgr.getSingleDoubleValueFromSQLQuery(con, strSql, Tables.FIA2FVS.DefaultFvsInputStandTableName);
                        }
                        if (oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputTreeTableName))
                        {
                            strSql = "SELECT COUNT(*) as TreeInitCount " +
                                     "FROM (SELECT DISTINCT TREE_CN FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName + ")";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                            trees = (int)oDataMgr.getSingleDoubleValueFromSQLQuery(con, strSql, Tables.FIA2FVS.DefaultFvsInputTreeTableName);
                        }
                    }
                    oDataMgr = null;  
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_fvs_input:getFVSSQLiteInputRecordCounts  \n" +
                                "Err Msg - " + e.Message.ToString().Trim(),
                    "FVS Input", MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            return new int[] { stands, trees };
        }

        private void btnRx_Click(object sender, System.EventArgs e)
        {

            frmMain.g_oFrmMain.StartRxDialog((frmDialog)ParentForm);
        }

        private void btnHelp_Click(object sender, System.EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            string helpPage = "INPUT_DATA";
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    helpPage = "INPUT_DATA";
                    break;
                case 1:
                    helpPage = "INPUT_OPTIONS";
                    break;
                case 2:
                    helpPage = "FIA2FVS_INPUT_OPTIONS";
                    break;
            }
            m_oHelp.ShowHelp(new string[] { "FVS", helpPage });
        }
        private void txtDataDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void txtInDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void txtOutDir_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void lstFvsInput_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            int x;
            if (e.Button == MouseButtons.Left)
            {
                try
                {
                    int intRowHt = this.lstFvsInput.Items[0].Bounds.Height;
                    double dblRow = (double)(e.Y / intRowHt);
                    this.lstFvsInput.Items[this.lstFvsInput.TopItem.Index + (int)dblRow - 1].Selected = true;
                }
                catch
                {
                }
            }

        }

        private void lstFvsInput_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            if (lstFvsInput.SelectedItems.Count > 0)
                this.m_oLvRowColors.DelegateListViewItem(lstFvsInput.SelectedItems[0]);
        }

        private void btnCreateFvsOutFiles_Click(object sender, System.EventArgs e)
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Create FVSOut DB File", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            CreateFvsOutFiles();


        }

        private void CreateFvsOutFiles()
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "Create FVSOut DB File", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            string strVariant = "";
            string strCurrentVariant = "";

            string strMsg = "";
            string strFile = "";
            string strFullPath = "";
            dao_data_access oDao = new dao_data_access();
            m_ado.OpenConnection(m_strConn);
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                if (this.lstFvsInput.Items[x].Checked == true)
                {
                    //out mdb file
                    if (this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text.Trim().Length == 0)
                    {
                        strFile = this.m_oRxTools.GetRxPackageFvsOutDbFileName(m_ado.m_OleDbConnection,
                            this.m_oQueries.m_oFvs.m_strRxPackageTable,
                            this.m_oQueries.m_oFIAPlot.m_strPlotTable,
                            "(TRIM(a.fvs_variant)='" + this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "' AND " +
                            "b.rxpackage='" + this.lstFvsInput.Items[x].SubItems[COL_RX].Text.Trim() + "')");

                        if (strFile.Trim().Length > 0)
                        {
                            strFullPath = this.txtDataDir.Text + "\\" + lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim() + "\\" + strFile;
                            if (System.IO.File.Exists(strFullPath) == false)
                            {
                                oDao.CreateMDB(strFullPath);
                                if (oDao.m_intError == 0)
                                {
                                    strMsg = strMsg + "\r\n" + strFile;
                                    this.lstFvsInput.Items[x].SubItems[COL_MDBOUT].Text = strFile;
                                }
                            }
                        }


                    }
                    strVariant = this.lstFvsInput.Items[x].SubItems[COL_VARIANT].Text.Trim();
                    if (strVariant != strCurrentVariant)
                    {
                        strCurrentVariant = strVariant;
                        if (this.lstFvsInput.Items[x].SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text.Trim().Length == 0)
                        {
                            strFile = "FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                            strFullPath = this.txtDataDir.Text + "\\" + strVariant + "\\" + strFile;
                            if (System.IO.File.Exists(strFullPath) == false)
                            {
                                oDao.CreateMDB(strFullPath);
                                if (oDao.m_intError == 0)
                                {
                                    strMsg = strMsg + "\r\n" + strFile;
                                    this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text = strFile;
                                }
                            }
                        }

                    }
                    else
                    {
                        if (this.lstFvsInput.Items[x].SubItems[uc_fvs_input.COL_POTFIREMDBOUT].Text.Trim().Length == 0)
                        {
                            strFile = "FVSOUT_" + strVariant + "_POTFIRE_BaseYr.MDB";
                            strFullPath = this.txtDataDir.Text + "\\" + strVariant + "\\" + strFile;
                            if (System.IO.File.Exists(strFullPath))
                            {
                                this.lstFvsInput.Items[x].SubItems[COL_POTFIREMDBOUT].Text = strFile;
                            }
                        }
                    }
                }
            }
            m_ado.CloseConnection(m_ado.m_OleDbConnection);
            oDao = null;
            if (strMsg.Length == 0) strMsg = "No DB Files Created";
            else
            {
                strMsg = "DB Files Created\r\n-------------------\r\n" + strMsg;
            }
            FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
            frmTemp.Text = "FIA Biosum";
            frmTemp.AutoScroll = false;
            uc_textbox uc_textbox1 = new uc_textbox();
            frmTemp.Controls.Add(uc_textbox1);
            uc_textbox1.Dock = DockStyle.Fill;
            uc_textbox1.lblTitle.Text = "Create Files";
            uc_textbox1.TextValue = strMsg;
            frmTemp.ShowDialog();
        }
        private void btnRxPackage_Click(object sender, System.EventArgs e)
        {
            frmMain.g_oFrmMain.StartRxPackageDialog((frmDialog)ParentForm);
        }


        public string strProjectDirectory
        {
            set
            {
                this.m_strProjDir = value;
            }
            get
            {
                return this.m_strProjDir;
            }
        }
        public string strProjectId
        {
            set
            {
                this.m_strProjId = value;
            }
            get
            {
                return this.m_strProjId;
            }
        }
        public frmMain ReferenceMainForm
        {
            set { _frmMain = value; }
            get { return _frmMain; }
        }
        public frmDialog ReferenceParentDialogForm
        {
            set { _frmDialog = value; }
            get { return _frmDialog; }
        }

        private void btnExecuteAction_Click(object sender, EventArgs e)
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }
            string strAction = cmbAction.Text.Trim().ToUpper();
            switch (strAction)
            {
                case "CREATE FVS INPUT DATABASE FILES":
                    btnAppend_Click(null, null);
                    break;
                case "CREATE FVS INPUT DATABASE FILES FROM FIA2FVS":
                    CreateFia2FvsInputFiles();
                    break;
                //case "CREATE FVS OUTPUT DATABASE FILES":
                //    CreateFvsOutFiles();
                //    break;
                //13-JAN-2022: No longer needed with FVSOn*
                //case "DELETE STANDARD FVS OUTPUT TABLES":
                //    BackupBeforeDelete("S");
                //    break;
                //case "DELETE POTFIRE BASE YEAR OUTPUT TABLES":
                //    BackupBeforeDelete("P");
                //    break;
                //case "DELETE BOTH STANDARD AND POTFIRE BASE YEAR OUTPUT TABLES":
                //    BackupBeforeDelete("B");
                //    break;
                //04-FEB-2022: No longer needed with FVSOn
            }
        }

        private void cmbAction_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbAction_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void txtMinSmallFwdTL_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(txtMinSmallFwdTL.Text, out temp))
            {
                txtMinSmallFwdTL.Text = "10";
            }
        }

        private void txtMinLargeFwdTL_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(txtMinLargeFwdTL.Text, out temp))
            {
                txtMinLargeFwdTL.Text = "30";
            }
        }

        private void txtMinCwdTL_Validating(object sender, CancelEventArgs e)
        {
            double temp;
            if (!double.TryParse(txtMinCwdTL.Text, out temp))
            {
                txtMinCwdTL.Text = "48";
            }
        }

        private void tabControl1_Resize(object sender, EventArgs e)
        {
            tabControl1.Top = this.lblTitle.Bottom + 5;
            tabControl1.Left = 5;
            tabControl1.Width = this.Width - 10;
            tabControl1.Height = this.Height - 100;

            btnClose.Top = groupBox1.Bottom - btnClose.Height - 5;
            btnClose.Left = tabControl1.Right - btnClose.Width;
            btnHelp.Left = tabControl1.Left;
            btnHelp.Top = btnClose.Top;


            label1.Left = tabPage2.Left + 5;
            txtDataDir.Left = label1.Right + 5;
            txtDataDir.Width = tabPage2.Width - txtDataDir.Left;
            btnRx.Left = txtDataDir.Left;
            lblRxCnt.Left = btnRx.Left + btnRx.Width + 10;
            btnRxPackage.Left = lblRxCnt.Left + lblRxCnt.Width + 10;
            lblRxPackageCnt.Left = btnRxPackage.Left + btnRxPackage.Width + 10;

            //resize all of the controls on the inside
            lstFvsInput.Left = tabPage2.Left + 5;
            lstFvsInput.Width = tabPage2.Width - 10;
            lstFvsInput.Top = txtDataDir.Bottom + 5;
            lstFvsInput.Height = tabPage2.Height - txtDataDir.Bottom - 120;

            //btns under lstFvsInput position based on tabControl perimeter
            btnExecuteAction.Top = lstFvsInput.Bottom + 80;
            btnExecuteAction.Left = lstFvsInput.Right - btnExecuteAction.Width;

            cmbAction.Top = btnExecuteAction.Top + (int)(btnExecuteAction.Height * .5) - (int)(cmbAction.Height * .5);
            cmbAction.Left = btnExecuteAction.Left - cmbAction.Width - 5;

            btnChkAll.Top = btnExecuteAction.Top;
            btnChkAll.Left = lstFvsInput.Left;
            btnClearAll.Top = btnExecuteAction.Top;
            btnClearAll.Left = btnChkAll.Right + 5;
            btnRefresh.Top = btnExecuteAction.Top;
            btnRefresh.Left = btnClearAll.Right + 5;

            lblFiaDatamartFile.Left = cmbAction.Left;
            lblFiaDatamartFile.Top = lstFvsInput.Bottom + 5;
            txtFIADatamart.Left = lblFiaDatamartFile.Left;
            txtFIADatamart.Top = lblFiaDatamartFile.Bottom + 1;
            btnDatamart.Top = txtFIADatamart.Top - 5;
            btnDatamart.Left = txtFIADatamart.Right + 5;

            lblSelectedGroup.Left = cmbAction.Left;
            lblSelectedGroup.Top = cmbAction.Top - 32;
            cmbSelectedGroup.Left = lblSelectedGroup.Right + 10;
            cmbSelectedGroup.Top = lblSelectedGroup.Top - 2;
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            tabPage1.Refresh();
            tabPage2.Refresh();
        }

        private void linkLabelFuelModel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                VisitLink();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to open link that was clicked.");
            }
        }

        private void VisitLink()
        {
            // Change the color of the link text by setting LinkVisited   
            // to true.  
            linkLabelFuelModel.LinkVisited = true;
            //Call the Process.Start method to open the default browser   
            //with a URL:  
            System.Diagnostics.Process.Start("https://www.fs.usda.gov/treesearch/pubs/download/9521.pdf");
        }

        private void btnDatamart_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog oDialog = new System.Windows.Forms.OpenFileDialog();
            oDialog.Title = "FIA Datamart database containing FIA2FVS files";
            oDialog.Filter = "SQLite databases (*.db;*.db3)|*.db;*.db3|All files (*.*)|*.*";
            DialogResult result = oDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtFIADatamart.Text = oDialog.FileName;

                // Check for the 2 required databases
                SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                string connSourceDb = oDataMgr.GetConnectionString(txtFIADatamart.Text.Trim());
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(connSourceDb))
                {
                    con.Open();
                    if (!oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputStandTableName))
                    {
                        MessageBox.Show("A valid input database was not selected. The " + Tables.FIA2FVS.DefaultFvsInputStandTableName + " table is missing!",
                            "FIA Biosum");
                        txtFIADatamart.Text = "";
                        return;
                    }
                    if (!oDataMgr.TableExist(con, Tables.FIA2FVS.DefaultFvsInputTreeTableName))
                    {
                        MessageBox.Show("A valid input database was not selected. The " + Tables.FIA2FVS.DefaultFvsInputTreeTableName + " table is missing!",
                            "FIA Biosum");
                        txtFIADatamart.Text = "";
                        return;
                    }
                }

            }
        }
    }
}
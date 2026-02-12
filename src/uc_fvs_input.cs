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
using SQLite.ADO;

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
        private DataMgr m_dataMgr = new DataMgr();
        private bool m_bOverwrite = true;
        private IDictionary<string, List<string>> m_dictVariantStates = null;
        private System.Windows.Forms.Button btnHelp;
        private int m_intError = 0;
        private System.Threading.Thread m_thread;

        //list view column constants
        private const int COL_CHECKBOX = 0;
        private const int COL_VARIANT = 1;
        private const int COL_STANDCOUNT = 2;
        private const int COL_TREECOUNT = 3;
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
        Dictionary<string, int[]> m_VariantCountsDict = null;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultFvsXPSFile;
        private string m_strDebugFile;
        private TextBox txtDataDir;
        private Button btnCreateFvsInput;
        public ListView lstFvsInput;
        private ComboBox cmbAction;
        private Label label1;
        private Button btnChkAll;
        private Button btnClearAll;
        private Button btnRefresh;
        private GroupBox otherOptionsGroupBox;
        private GroupBox grpCalibOptions;
        private CheckBox chkUsePrevHt;
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
        private CheckBox chkIncludeSeedlings;
        private CheckBox chkUsePrevDia;
        private Label lblSurfaceWarning;
        private GroupBox grpFVSlist;

        delegate string[] GetListBoxItemsDlg(CheckedListBox checkedListBox);

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public uc_fvs_input()
        {
            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();
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
            this.lblSelectedGroup = new System.Windows.Forms.Label();
            this.cmbSelectedGroup = new System.Windows.Forms.ComboBox();
            this.btnDatamart = new System.Windows.Forms.Button();
            this.txtFIADatamart = new System.Windows.Forms.TextBox();
            this.lblFiaDatamartFile = new System.Windows.Forms.Label();
            this.txtDataDir = new System.Windows.Forms.TextBox();
            this.btnCreateFvsInput = new System.Windows.Forms.Button();
            this.lstFvsInput = new System.Windows.Forms.ListView();
            this.cmbAction = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnChkAll = new System.Windows.Forms.Button();
            this.btnClearAll = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.otherOptionsGroupBox = new System.Windows.Forms.GroupBox();
            this.chkIncludeSeedlings = new System.Windows.Forms.CheckBox();
            this.grpCalibOptions = new System.Windows.Forms.GroupBox();
            this.chkUsePrevDia = new System.Windows.Forms.CheckBox();
            this.chkUsePrevHt = new System.Windows.Forms.CheckBox();
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
            this.lblSurfaceWarning = new System.Windows.Forms.Label();
            this.grpFVSlist = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.grpFVSlist.SuspendLayout();
            this.otherOptionsGroupBox.SuspendLayout();
            this.grpCalibOptions.SuspendLayout();
            this.grpDWMOptions.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.progressBar1);
            this.groupBox1.Controls.Add(this.lblProgress);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.lblSelectedGroup);
            this.groupBox1.Controls.Add(this.cmbSelectedGroup);
            this.groupBox1.Controls.Add(this.btnDatamart);
            this.groupBox1.Controls.Add(this.txtFIADatamart);
            this.groupBox1.Controls.Add(this.lblFiaDatamartFile);
            this.groupBox1.Controls.Add(this.txtDataDir);
            this.groupBox1.Controls.Add(this.btnCreateFvsInput);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.grpFVSlist);
            this.groupBox1.Controls.Add(this.grpCalibOptions);
            this.groupBox1.Controls.Add(this.grpDWMOptions);
            this.groupBox1.Controls.Add(this.otherOptionsGroupBox);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(800, 600);
            this.groupBox1.TabIndex = 99;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
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
            this.btnClose.Location = new System.Drawing.Point(697, 559);
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
            // lblSelectedGroup
            // 
            this.lblSelectedGroup.AutoSize = true;
            this.lblSelectedGroup.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelectedGroup.Location = new System.Drawing.Point(340, 520);
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
            this.cmbSelectedGroup.Location = new System.Drawing.Point(443, 517);
            this.cmbSelectedGroup.Name = "cmbSelectedGroup";
            this.cmbSelectedGroup.Size = new System.Drawing.Size(250, 24);
            this.cmbSelectedGroup.TabIndex = 103;
            // 
            // btnDatamart
            // 
            this.btnDatamart.Image = ((System.Drawing.Image)(resources.GetObject("btnDatamart.Image")));
            this.btnDatamart.Location = new System.Drawing.Point(734, 484);
            this.btnDatamart.Name = "btnDatamart";
            this.btnDatamart.Size = new System.Drawing.Size(59, 32);
            this.btnDatamart.TabIndex = 102;
            this.btnDatamart.Click += new System.EventHandler(this.btnDatamart_Click);
            // 
            // txtFIADatamart
            // 
            this.txtFIADatamart.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFIADatamart.Location = new System.Drawing.Point(315, 489);
            this.txtFIADatamart.Name = "txtFIADatamart";
            this.txtFIADatamart.Size = new System.Drawing.Size(413, 22);
            this.txtFIADatamart.TabIndex = 101;
            // 
            // lblFiaDatamartFile
            // 
            this.lblFiaDatamartFile.AutoSize = true;
            this.lblFiaDatamartFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFiaDatamartFile.Location = new System.Drawing.Point(312, 469);
            this.lblFiaDatamartFile.Name = "lblFiaDatamartFile";
            this.lblFiaDatamartFile.Size = new System.Drawing.Size(354, 17);
            this.lblFiaDatamartFile.TabIndex = 100;
            this.lblFiaDatamartFile.Text = "Path to the SQLite database sourced from FIA Datamart";
            // 
            // txtDataDir
            // 
            this.txtDataDir.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDataDir.Location = new System.Drawing.Point(173, 48);
            this.txtDataDir.Name = "txtDataDir";
            this.txtDataDir.Size = new System.Drawing.Size(589, 23);
            this.txtDataDir.TabIndex = 99;
            // 
            // btnCreateFvsInput
            // 
            this.btnCreateFvsInput.Location = new System.Drawing.Point(464, 559);
            this.btnCreateFvsInput.Name = "btnCreateFvsInput";
            this.btnCreateFvsInput.Size = new System.Drawing.Size(229, 32);
            this.btnCreateFvsInput.TabIndex = 5;
            this.btnCreateFvsInput.Text = "Create FVS Input Database File";
            this.btnCreateFvsInput.Click += new System.EventHandler(this.btnCreateFvsInput_Click);
            // 
            // cmbAction
            // 
            this.cmbAction.FormattingEnabled = true;
            this.cmbAction.Items.AddRange(new object[] {
            "Create FVS Input Database Files",
            "Create FVS Input Database Files From FIA2FVS"});
            this.cmbAction.Location = new System.Drawing.Point(9, 424);
            this.cmbAction.Name = "cmbAction";
            this.cmbAction.Size = new System.Drawing.Size(362, 24);
            this.cmbAction.TabIndex = 4;
            this.cmbAction.Text = "<-------Action Items------->";
            this.cmbAction.Visible = false;
            this.cmbAction.SelectedIndexChanged += new System.EventHandler(this.cmbAction_SelectedIndexChanged);
            this.cmbAction.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbAction_KeyPress);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 50);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 24);
            this.label1.TabIndex = 99;
            this.label1.Text = "Target directory for FVSIn.db";
            // 
            // grpFVSlist
            // 
            this.grpFVSlist.Controls.Add(this.lstFvsInput);
            this.grpFVSlist.Controls.Add(this.btnChkAll);
            this.grpFVSlist.Controls.Add(this.btnClearAll);
            this.grpFVSlist.Location = new System.Drawing.Point(28, 75);
            this.grpFVSlist.Name = "grpFVSlist";
            this.grpFVSlist.Size = new System.Drawing.Size(275, 228);
            this.grpFVSlist.TabIndex = 103;
            this.grpFVSlist.TabStop = false;
            this.grpFVSlist.Text = "Variant Selection";
            // 
            // lstFvsInput
            // 
            this.lstFvsInput.CheckBoxes = true;
            this.lstFvsInput.GridLines = true;
            this.lstFvsInput.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFvsInput.HideSelection = false;
            this.lstFvsInput.Location = new System.Drawing.Point(6, 16);
            this.lstFvsInput.MultiSelect = false;
            this.lstFvsInput.Name = "lstFvsInput";
            this.lstFvsInput.Size = new System.Drawing.Size(231, 165);
            this.lstFvsInput.TabIndex = 0;
            this.lstFvsInput.UseCompatibleStateImageBehavior = false;
            this.lstFvsInput.View = System.Windows.Forms.View.Details;
            this.lstFvsInput.SelectedIndexChanged += new System.EventHandler(this.lstFvsInput_SelectedIndexChanged);
            this.lstFvsInput.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lstFvsInput_MouseUp);
            // 
            // btnChkAll
            // 
            this.btnChkAll.Location = new System.Drawing.Point(6, 186);
            this.btnChkAll.Name = "btnChkAll";
            this.btnChkAll.Size = new System.Drawing.Size(75, 32);
            this.btnChkAll.TabIndex = 1;
            this.btnChkAll.Text = "Check All";
            this.btnChkAll.Click += new System.EventHandler(this.btnChkAll_Click);
            // 
            // btnClearAll
            // 
            this.btnClearAll.Location = new System.Drawing.Point(85, 186);
            this.btnClearAll.Name = "btnClearAll";
            this.btnClearAll.Size = new System.Drawing.Size(75, 32);
            this.btnClearAll.TabIndex = 2;
            this.btnClearAll.Text = "Clear All";
            this.btnClearAll.Click += new System.EventHandler(this.btnClearAll_Click);
            // 
            // grpDWMOptions
            // 
            this.grpDWMOptions.Controls.Add(this.lblSurfaceWarning);
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
            this.grpDWMOptions.Location = new System.Drawing.Point(308, 75);
            this.grpDWMOptions.Name = "grpDWMOptions";
            this.grpDWMOptions.Size = new System.Drawing.Size(439, 378);
            this.grpDWMOptions.TabIndex = 101;
            this.grpDWMOptions.TabStop = false;
            this.grpDWMOptions.Text = "Down Woody Material";
            // 
            // lblSurfaceWarning
            // 
            this.lblSurfaceWarning.AutoSize = true;
            this.lblSurfaceWarning.Location = new System.Drawing.Point(5, 43);
            this.lblSurfaceWarning.Name = "lblSurfaceWarning";
            this.lblSurfaceWarning.Size = new System.Drawing.Size(434, 17);
            this.lblSurfaceWarning.TabIndex = 106;
            this.lblSurfaceWarning.Text = "If surface fuel models included. FVS will ignore all DWM data loaded";
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
            // chkDwmFuelModel
            // 
            this.chkDwmFuelModel.AutoSize = true;
            this.chkDwmFuelModel.Location = new System.Drawing.Point(7, 19);
            this.chkDwmFuelModel.Name = "chkDwmFuelModel";
            this.chkDwmFuelModel.Size = new System.Drawing.Size(18, 17);
            this.chkDwmFuelModel.TabIndex = 1;
            this.chkDwmFuelModel.UseVisualStyleBackColor = true;
            // 
            // chkDwmFuelBiomass
            // 
            this.chkDwmFuelBiomass.AutoSize = true;
            this.chkDwmFuelBiomass.Checked = true;
            this.chkDwmFuelBiomass.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkDwmFuelBiomass.Location = new System.Drawing.Point(7, 73);
            this.chkDwmFuelBiomass.Name = "chkDwmFuelBiomass";
            this.chkDwmFuelBiomass.Size = new System.Drawing.Size(472, 21);
            this.chkDwmFuelBiomass.TabIndex = 2;
            this.chkDwmFuelBiomass.Text = "Include calculated fuel biomasses with available DWM data (tons/acre)";
            this.chkDwmFuelBiomass.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label4.Location = new System.Drawing.Point(58, 153);
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
            this.label2.Location = new System.Drawing.Point(58, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(347, 17);
            this.label2.TabIndex = 99;
            this.label2.Tag = "txtMinSmallFwdTL";
            this.label2.Text = "Minimum Small and Medium FWD Transect Length (ft)";
            // 
            // txtMinLargeFwdTL
            // 
            this.txtMinLargeFwdTL.Location = new System.Drawing.Point(7, 124);
            this.txtMinLargeFwdTL.Name = "txtMinLargeFwdTL";
            this.txtMinLargeFwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinLargeFwdTL.TabIndex = 4;
            this.txtMinLargeFwdTL.Text = "1";
            this.txtMinLargeFwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinLargeFwdTL_Validating);
            // 
            // txtMinCwdTL
            // 
            this.txtMinCwdTL.Location = new System.Drawing.Point(7, 150);
            this.txtMinCwdTL.Name = "txtMinCwdTL";
            this.txtMinCwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinCwdTL.TabIndex = 5;
            this.txtMinCwdTL.Text = "1";
            this.txtMinCwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinCwdTL_Validating);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.label3.Location = new System.Drawing.Point(58, 127);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(269, 17);
            this.label3.TabIndex = 99;
            this.label3.Tag = "txtMinLargeFwdTL";
            this.label3.Text = "Minimum Large FWD Transect Length (ft)";
            // 
            // txtMinSmallFwdTL
            // 
            this.txtMinSmallFwdTL.Location = new System.Drawing.Point(7, 99);
            this.txtMinSmallFwdTL.Name = "txtMinSmallFwdTL";
            this.txtMinSmallFwdTL.Size = new System.Drawing.Size(45, 22);
            this.txtMinSmallFwdTL.TabIndex = 3;
            this.txtMinSmallFwdTL.Text = "1";
            this.txtMinSmallFwdTL.Validating += new System.ComponentModel.CancelEventHandler(this.txtMinSmallFwdTL_Validating);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.chkLstBoxDuffYears);
            this.groupBox2.Controls.Add(this.chkLstBoxLitterYears);
            this.groupBox2.Location = new System.Drawing.Point(7, 176);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(236, 190);
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
            this.chkLstBoxDuffYears.Size = new System.Drawing.Size(100, 140);
            this.chkLstBoxDuffYears.TabIndex = 6;
            // 
            // chkLstBoxLitterYears
            // 
            this.chkLstBoxLitterYears.FormattingEnabled = true;
            this.chkLstBoxLitterYears.Location = new System.Drawing.Point(126, 35);
            this.chkLstBoxLitterYears.Name = "chkLstBoxLitterYears";
            this.chkLstBoxLitterYears.Size = new System.Drawing.Size(100, 140);
            this.chkLstBoxLitterYears.Sorted = true;
            this.chkLstBoxLitterYears.TabIndex = 7;
            // 
            // otherOptionsGroupBox
            // 
            this.otherOptionsGroupBox.Controls.Add(this.chkIncludeSeedlings);
            this.otherOptionsGroupBox.Location = new System.Drawing.Point(28, 407);
            this.otherOptionsGroupBox.Name = "otherOptionsGroupBox";
            this.otherOptionsGroupBox.Size = new System.Drawing.Size(275, 60);
            this.otherOptionsGroupBox.TabIndex = 103;
            this.otherOptionsGroupBox.TabStop = false;
            this.otherOptionsGroupBox.Text = "Other Options";
            // 
            // chkIncludeSeedlings
            // 
            this.chkIncludeSeedlings.AutoSize = true;
            this.chkIncludeSeedlings.Checked = true;
            this.chkIncludeSeedlings.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIncludeSeedlings.Location = new System.Drawing.Point(6, 21);
            this.chkIncludeSeedlings.Name = "chkIncludeSeedlings";
            this.chkIncludeSeedlings.Size = new System.Drawing.Size(139, 21);
            this.chkIncludeSeedlings.TabIndex = 4;
            this.chkIncludeSeedlings.Text = "Include seedlings from FIADB";
            this.chkIncludeSeedlings.UseVisualStyleBackColor = true;
            // 
            // grpCalibOptions
            // 
            this.grpCalibOptions.Controls.Add(this.chkUsePrevDia);
            this.grpCalibOptions.Controls.Add(this.chkUsePrevHt);
            this.grpCalibOptions.Location = new System.Drawing.Point(28, 307);
            this.grpCalibOptions.Name = "grpCalibOptions";
            this.grpCalibOptions.Size = new System.Drawing.Size(275, 95);
            this.grpCalibOptions.TabIndex = 102;
            this.grpCalibOptions.TabStop = false;
            this.grpCalibOptions.Text = "Tree Growth Calibration Data";
            // 
            // chkUsePrevHt
            // 
            this.chkUsePrevHt.AutoSize = true;
            this.chkUsePrevHt.Checked = true;
            this.chkUsePrevHt.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUsePrevHt.Location = new System.Drawing.Point(6, 20);
            this.chkUsePrevHt.Name = "chkUsePrevHt";
            this.chkUsePrevHt.Size = new System.Drawing.Size(302, 21);
            this.chkUsePrevHt.TabIndex = 2;
            this.chkUsePrevHt.Text = "Include previous height, if available";
            this.chkUsePrevHt.UseVisualStyleBackColor = true;
            // 
            // chkUsePrevDia
            // 
            this.chkUsePrevDia.AutoSize = true;
            this.chkUsePrevDia.Checked = true;
            this.chkUsePrevDia.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUsePrevDia.Location = new System.Drawing.Point(6, 43);
            this.chkUsePrevDia.Name = "chkUsePrevDia";
            this.chkUsePrevDia.Size = new System.Drawing.Size(318, 21);
            this.chkUsePrevDia.TabIndex = 3;
            this.chkUsePrevDia.Text = "Include previous diameter, if available";
            this.chkUsePrevDia.UseVisualStyleBackColor = true;
            // 
            // uc_fvs_input
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_input";
            this.Size = new System.Drawing.Size(800, 600);
            this.Resize += new System.EventHandler(this.uc_fvs_input_Resize);
            this.groupBox1.ResumeLayout(false);
            this.otherOptionsGroupBox.ResumeLayout(false);
            this.otherOptionsGroupBox.PerformLayout();
            this.grpCalibOptions.ResumeLayout(false);
            this.grpCalibOptions.PerformLayout();
            this.grpDWMOptions.ResumeLayout(false);
            this.grpDWMOptions.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.grpFVSlist.ResumeLayout(false);
            this.grpFVSlist.PerformLayout();
            this.ResumeLayout(false);
        }
        #endregion

        private void uc_fvs_input_Resize(object sender, System.EventArgs e)
        {
            this.Resize_Fvs_Input();
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
            PopulateDefaultFIADatamart();
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
                oDataMgr.m_DataReader.Close();
            }
        }
        private void PopulateDefaultFIADatamart()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strFVSIndb = m_strProjDir + "\\fvs\\data\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            if (System.IO.File.Exists(strFVSIndb))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFVSIndb)))
                {
                    conn.Open();
                    if (oDataMgr.TableExist(conn, "biosum_fvsin_configuration"))
                    {
                        oDataMgr.m_strSQL = "SELECT * FROM biosum_fvsin_configuration WHERE Setting = 'Source FIA Data Mart database'";
                        oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                        if (oDataMgr.m_DataReader.HasRows)
                        {
                            while (oDataMgr.m_DataReader.Read())
                            {
                                if (oDataMgr.m_DataReader["Value"] != DBNull.Value &&
                                    Convert.ToString(oDataMgr.m_DataReader["Value"]).Trim().Length > 0)
                                {
                                    string strSourceDb = Convert.ToString(oDataMgr.m_DataReader["Value"]).Trim();
                                    if (System.IO.File.Exists(strSourceDb))
                                    {
                                        txtFIADatamart.Text = strSourceDb;
                                    }
                                }
                            }
                        }
                        oDataMgr.m_DataReader.Close();
                    }
                }
            }
        }

        private void populate_listbox()
        {
            string strInDirAndFile;

            try
            {
                this.lstFvsInput.Clear();
                this.m_oLvRowColors.InitializeRowCollection();

                this.lstFvsInput.Columns.Add("", 2, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Variant", 55, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Stands", 70, HorizontalAlignment.Left);
                this.lstFvsInput.Columns.Add("Trees", 80, HorizontalAlignment.Left);

                this.lstFvsInput.Columns[COL_CHECKBOX].Width = -2;

                List<string> lstVariants = new List<string>();
                string strMasterConn = m_dataMgr.GetConnectionString(m_oQueries.m_oDataSource.getFullPathAndFile(Datasource.TableTypes.Plot));
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strMasterConn))
                {
                    conn.Open();

                    m_dataMgr.SqlQueryReader(conn, Queries.FVS.GetFVSVariantSQL(m_oQueries.m_oFIAPlot.m_strPlotTable));
                    if (m_dataMgr.m_DataReader.HasRows)
                    {
                        while (m_dataMgr.m_DataReader.Read())
                        {
                            string strCurrentVariant = m_dataMgr.m_DataReader["fvs_variant"].ToString().Trim();
                            lstVariants.Add(strCurrentVariant);
                        }
                    }
                    m_dataMgr.m_DataReader.Close();
                }

                //Keep a count of records in FVS_StandInit and FVS_TreeInit tables in each variant
                m_VariantCountsDict = new Dictionary<string, int[]>();
                foreach (string strVariant in lstVariants)
                {
                    System.Windows.Forms.ListViewItem entryListItem =
                    this.lstFvsInput.Items.Add("");
                    entryListItem.UseItemStyleForSubItems = false;
                    this.m_oLvRowColors.AddRow();
                    this.m_oLvRowColors.AddColumns(lstFvsInput.Items.Count - 1, lstFvsInput.Columns.Count);
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_CHECKBOX, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);

                    //fvs_variant
                    if (m_VariantCountsDict.ContainsKey(strVariant) == false)
                    {
                        m_VariantCountsDict.Add(strVariant, null); //fvs_standinit, fvs_treeinit counts
                    }

                    entryListItem.SubItems.Add(strVariant);
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_VARIANT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //FVS_StandInit Stand_ID count
                    entryListItem.SubItems.Add(" ");
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_STANDCOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);
                    //FVS_TreeInit row count
                    entryListItem.SubItems.Add(" ");
                    this.m_oLvRowColors.ListViewSubItem(entryListItem.Index, uc_fvs_input.COL_TREECOUNT, entryListItem.SubItems[entryListItem.SubItems.Count - 1], false);


                    frmMain.g_sbpInfo.Text = "Processing FVS Input Variant " +
                        strVariant + "...Stand By";

                    //check fvs in values
                    strInDirAndFile = $@"{this.txtDataDir.Text.Trim()}\{Tables.FIA2FVS.DefaultFvsInputFile}";
                    if (frmMain.g_bSuppressFVSInputTableRowCount == false && System.IO.File.Exists(strInDirAndFile) == true)
                    {
                        if (m_VariantCountsDict[strVariant] == null)
                        {
                            m_VariantCountsDict[strVariant] = getFVSInputRecordCounts(strInDirAndFile, strVariant);
                        }
                        entryListItem.SubItems[COL_STANDCOUNT].Text = Convert.ToString(m_VariantCountsDict[strVariant][0]);
                        entryListItem.SubItems[COL_TREECOUNT].Text = Convert.ToString(m_VariantCountsDict[strVariant][1]);
                    }

                }
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
        }

        private void LoadDataSources()
        {
            this.txtDataDir.Text = this.m_strProjDir + "\\fvs\\data";
            this.m_oQueries.m_oFvs.LoadDatasource = true;
            this.m_oQueries.m_oFIAPlot.LoadDatasource = true;
            this.m_oQueries.LoadDatasources(true);
        }

        private void btnClose_Click(object sender, System.EventArgs e)
        {
            frmMain.g_oFrmMain.Enabled = true;
            this.ParentForm.Dispose();

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

        public void CreateFia2FvsInputFiles()
        {
            m_bOverwrite = false;
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
                MessageBox.Show("You must specify a source input database before proceeding!", "FIA Biosum");
                return;
            }

            // Make sure a group is selected
            if (cmbSelectedGroup.SelectedIndex < 0)
            {
                MessageBox.Show("You must specify a group before proceeding!", "FIA Biosum");
                return;
            }
            string strInDirAndFile = this.txtDataDir.Text + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;

            // Check for existing files
            List<string> lstVariants = new List<string>();
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
            {
                if (this.lstFvsInput.Items[x].Checked)
                {
                    var item = this.lstFvsInput.Items[x];
                    string strVariant = item.SubItems[1].Text.Trim();
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
                                    oDataMgr.SqlQueryReader(con, "SELECT * FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName + " WHERE TRIM(VARIANT) = '" + strVariant + "'");
                                    if (oDataMgr.m_DataReader.HasRows)
                                    {
                                        // Only add to warnings if both tables exist and selected variant is in stand table
                                        if (!lstVariants.Contains(strVariant))
                                        {
                                            lstVariants.Add(strVariant);
                                        }
                                    }
                                    oDataMgr.m_DataReader.Close();
                                }
                            }
                        }
                    }
                }
            }

            if (lstVariants.Count > 0)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                m_dictVariantStates = new Dictionary<string, List<string>>();    // Re-initialize dictionary
                string connFiadbDb = oDataMgr.GetConnectionString(this.txtFIADatamart.Text);
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(connFiadbDb))
                {
                    con.Open();

                    oDataMgr.SqlNonQuery(con, "ATTACH '" + strInDirAndFile + "' as target");
                    foreach (var strVariant in lstVariants)
                    {
                        string strQuery = "SELECT distinct STATE FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                            " INTERSECT SELECT distinct STATE" +
                                            " FROM target." + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                            " WHERE TRIM(VARIANT) = '" + strVariant + "'";
                        List<string> lstStates = oDataMgr.getStringList(con, strQuery);
                        if (lstStates.Count > 0)
                        {
                            m_dictVariantStates.Add(strVariant, lstStates);
                        }
                    }
                    oDataMgr.SqlNonQuery(con, "DETACH DATABASE 'target'");
                }
                if (m_dictVariantStates != null &&
                    m_dictVariantStates.Keys.Count > 0)
                {
                    sb = new System.Text.StringBuilder();
                    sb.Append("FVS input data for the variant(s) listed below already contain records for state codes ");
                    sb.Append("in the FIA Datamart input SQLite database. ");
                    sb.Append("Click 'Yes' to replace the existing stands ");
                    sb.Append("and trees for those states or 'No' to stop this process.\r\n\r\n");
                    
                    foreach (var strVariant in m_dictVariantStates.Keys)
                    {
                        sb.Append(strVariant + "\r");
                    }
                    DialogResult res = MessageBox.Show(sb.ToString(), "FIA BioSum", MessageBoxButtons.YesNoCancel);
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

            // Check for existing kcp template file in fvs\data folder
            string strKcpTargetPath = this.strProjectDirectory + Tables.FIA2FVS.DefaultFvsInputFolderName + "\\" + 
                Tables.FIA2FVS.KcpFileBiosumKeywords + Tables.FIA2FVS.KcpFileExtension;

            if (System.IO.File.Exists(strKcpTargetPath))
            {
                string strMessage = "BioSum has found an existing " + Tables.FIA2FVS.KcpFileBiosumKeywords + Tables.FIA2FVS.KcpFileExtension +
                    " file. Do you wish to overwrite it?";
                DialogResult res = MessageBox.Show(strMessage, "FIA BioSum", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                {
                    strMessage = "Would you like to make a copy of your existing " + Tables.FIA2FVS.KcpFileBiosumKeywords +
                        Tables.FIA2FVS.KcpFileExtension + " file? The name of the file backup will include today's date.";
                    res = MessageBox.Show(strMessage, "FIA BioSum", MessageBoxButtons.YesNo);
                    if (res == DialogResult.Yes)
                    {
                        string strFileSuffix = "_" + DateTime.Now.ToString("MMddyyyy");
                        string strExtension = System.IO.Path.GetExtension(strKcpTargetPath);
                        string strNewFileName = System.IO.Path.GetFileNameWithoutExtension(strKcpTargetPath) + strFileSuffix + strExtension;
                        // Check to see if the backup database already exists; If it does, abort the process so user can delete
                        if (System.IO.File.Exists(this.strProjectDirectory + Tables.FIA2FVS.DefaultFvsInputFolderName + "\\" + strNewFileName))
                        {
                            MessageBox.Show("A backup file from today already exists: " + Tables.FIA2FVS.KcpFileBiosumKeywords + Tables.FIA2FVS.KcpFileExtension
                                + strFileSuffix + ". Delete this file manually if you want to " +
                                "back up today's data again!! The current file will not be overwritten.", "FIA BioSum");
                        }
                        else
                        {
                            System.IO.File.Copy(strKcpTargetPath, this.strProjectDirectory + Tables.FIA2FVS.DefaultFvsInputFolderName + "\\" + strNewFileName);
                        }
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
            p_fvs.bUsePrevHtCalibrationData = (bool)frmMain.g_oDelegate.GetControlPropertyValue(chkUsePrevHt, "Checked", false);
            p_fvs.bUsePrevDiaCalibrationData = (bool)frmMain.g_oDelegate.GetControlPropertyValue(chkUsePrevDia, "Checked", false);
            p_fvs.bIncludeSeedlings = (bool)frmMain.g_oDelegate.GetControlPropertyValue(chkIncludeSeedlings, "Checked", false);
            p_fvs.strSourceFiaDb = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtFIADatamart, "Text", false);
            p_fvs.strSourceFiaDb = p_fvs.strSourceFiaDb.Trim();
            p_fvs.strDataDirectory = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.txtDataDir, "Text", false);
            p_fvs.strDataDirectory = p_fvs.strDataDirectory.Trim();
            p_fvs.strGroup = (string)frmMain.g_oDelegate.GetControlPropertyValue((Control)this.cmbSelectedGroup, "SelectedItem", false);
            p_fvs.strGroup = p_fvs.strGroup.Trim();
        }

        
        private void ExtractFIA2FVSRecords()
        {
            m_intError = 0;
            string strCurVariant = "";
            string strVariant;
            string strMasterDb = strProjectDirectory + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableSqliteDbFile;

            m_strDebugFile = this.strProjectDirectory + Tables.FIA2FVS.DefaultFvsInputFolderName + "\\biosum_fvs_input_debug.txt";
            if (File.Exists(m_strDebugFile)) System.IO.File.Delete(m_strDebugFile);


            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//ExtractFIA2FVSRecords\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            DataMgr oDataMgr = new DataMgr();

            try
            {
                fvs_input p_fvsinput = new fvs_input(this.m_strProjDir, this.m_frmTherm);
                ConfigureFvsInput(p_fvsinput);

                int steps = m_VariantCountsDict.Keys.Count * 2;
                int interval = (int)Math.Floor((double)90 / steps);
                int intValue = interval;
                for (int x = 0; x <= this.lstFvsInput.Items.Count - 1; x++)
                {
                    frmMain.g_oDelegate.SetControlPropertyValue(m_frmTherm.progressBar2, "Value", intValue);
                    if ((bool)frmMain.g_oDelegate.GetListViewItemPropertyValue(lstFvsInput, x, "Checked", false) == true)
                    {
                        //get the variant
                        strVariant = frmMain.g_oDelegate.GetListViewSubItemPropertyValue(lstFvsInput, x, COL_VARIANT, "Text", false).ToString().Trim();
                        string strInDirAndFile = p_fvsinput.strDataDirectory + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
                        //see if this is a new variant
                        if (strVariant.Trim().ToUpper() != strCurVariant.Trim().ToUpper())
                        {
                            frmMain.g_oDelegate.SetControlPropertyValue(
                                m_frmTherm.progressBar1,
                                "Value", 1);
                            strCurVariant = strVariant;

                            List<string> lstStates = new List<string>();
                            if (m_dictVariantStates != null &&
                                m_dictVariantStates.ContainsKey(strCurVariant))
                            {
                                lstStates = m_dictVariantStates[strCurVariant];
                            }
                            p_fvsinput.StartFIA2FVSNew(oDataMgr, m_bOverwrite, strMasterDb, m_strDebugFile,
                                strCurVariant, lstStates);
                        }
                        frmMain.g_oDelegate.SetControlPropertyValue(
                            m_frmTherm.progressBar1,
                            "Value", 7);

                        // This happens at the end
                        // Populates stand and tree count on screen. Uses " " instead of 0
                        // for variants that haven't been run yet
                        if (File.Exists(strInDirAndFile) == true)
                        {
                            int[] fvsInputRecordCounts = getFVSInputRecordCounts(strInDirAndFile, strVariant);
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
                    intValue = intValue + interval;
                }
            }
            catch (ThreadInterruptedException err)
            {
                m_intError = -1;
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (ThreadAbortException)
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
                if (oDataMgr != null)
                {
                    if (oDataMgr.m_DataSet != null)
                    {
                        oDataMgr.m_DataSet.Clear();
                        oDataMgr.m_DataSet.Dispose();
                    }
                    oDataMgr = null;
                }
                if (m_intError == 0)
                {
                    ThreadCleanUp();
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
                frmMain.g_oDelegate.SetControlPropertyValue(btnCreateFvsInput, "Enabled", true);
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

        
        private int[] getFVSInputRecordCounts(string strDirAndFile, string strFVSVariant)
        {
            int stands = 0;
            int trees = 0;
            try
            {
                if (File.Exists(strDirAndFile))
                {
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    string connSourceDb = oDataMgr.GetConnectionString(strDirAndFile);
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(connSourceDb))
                    {
                        string strSql = "";
                        conn.Open();
                        // Do stuff in sqlite side
                        if (oDataMgr.TableExist(conn, Tables.FIA2FVS.DefaultFvsInputStandTableName))
                        {
                            strSql = "SELECT COUNT(*) as StandInitCount " +
                                     "FROM (SELECT DISTINCT STAND_CN FROM " + Tables.FIA2FVS.DefaultFvsInputStandTableName + 
                                     " WHERE VARIANT = '" + strFVSVariant + "')";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2 && System.IO.File.Exists(m_strDebugFile))
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                            stands = (int)oDataMgr.getSingleDoubleValueFromSQLQuery(conn, strSql, Tables.FIA2FVS.DefaultFvsInputStandTableName);
                        }
                        if (oDataMgr.TableExist(conn, Tables.FIA2FVS.DefaultFvsInputTreeTableName))
                        {
                            strSql = "SELECT COUNT(*) as TreeInitCount " +
                                     "FROM (SELECT DISTINCT TREE_CN FROM " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                     " WHERE VARIANT = '" + strFVSVariant + "')";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2 && System.IO.File.Exists(m_strDebugFile))
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                            trees = (int)oDataMgr.getSingleDoubleValueFromSQLQuery(conn, strSql, Tables.FIA2FVS.DefaultFvsInputTreeTableName);
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

        private void btnHelp_Click(object sender, System.EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            string helpPage = "INPUT_DATA";
            m_oHelp.ShowHelp(new string[] { "FVS", helpPage });
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

        private void btnCreateFvsInput_Click(object sender, EventArgs e)
        {
            if (this.lstFvsInput.CheckedItems.Count == 0)
            {
                MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            CreateFia2FvsInputFiles();
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
using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
    /// Summary description for uc_processor_scenario_tree_spc_groups.
	/// </summary>
	public class uc_processor_scenario_tree_spc_groups : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.Label lblTitle;
		public System.Windows.Forms.ListBox lstCommonName;
		private System.Windows.Forms.Button btnClose;
		public System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ListBox lstGrp1;
		private System.Windows.Forms.GroupBox grpbox1;
		private System.Windows.Forms.GroupBox grpbox2;
		private System.Windows.Forms.ListBox lstGrp2;
		private System.Windows.Forms.GroupBox grpbox3;
		private System.Windows.Forms.ListBox lstGrp3;
		private System.Windows.Forms.Button btnGrp2;
		private System.Windows.Forms.Button btnGrp3;
		private System.Windows.Forms.GroupBox grpbox4;
		private System.Windows.Forms.ListBox lstGrp4;
		private System.Windows.Forms.GroupBox grpbox5;
		private System.Windows.Forms.Button btnGrp4;
		private System.Windows.Forms.GroupBox grpbox6;
		private System.Windows.Forms.Button btnGrp5;
		private System.Windows.Forms.Button btnGrp6;
		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.TextBox txtGrp1;
		private System.Windows.Forms.TextBox txtGrp2;
		private System.Windows.Forms.TextBox txtGrp3;
		private System.Windows.Forms.TextBox txtGrp4;
		private System.Windows.Forms.TextBox txtGrp5;
		private System.Windows.Forms.ListBox lstGrp5;
		private System.Windows.Forms.TextBox txtGrp6;
		private System.Windows.Forms.ListBox lstGrp6;
		private System.Windows.Forms.Button btnRemove1;
		private System.Windows.Forms.Button btnClearAll1;
		private System.Windows.Forms.Button btnRemove2;
		private System.Windows.Forms.Button btnClearAll2;
		private System.Windows.Forms.Button btnRemove3;
		private System.Windows.Forms.Button btnClearAll3;
		private System.Windows.Forms.Button btnRemove4;
		private System.Windows.Forms.Button btnClearAll4;
		private System.Windows.Forms.Button btnRemove5;
		private System.Windows.Forms.Button btnClearAll5;
		private System.Windows.Forms.Button btnRemove6;
		private System.Windows.Forms.Button btnClearAll6;
		private System.Windows.Forms.Button btnHelp;
		public int m_intError=0;
        private FIA_Biosum_Manager.ado_data_access m_ado;
		private string m_strFvsOutTreeTable="";
        private string m_strConn = "";
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Button btnHwd;
		private System.Windows.Forms.Button btnSwd;
		private System.Windows.Forms.Button btnBoth;
		private System.ComponentModel.IContainer components;
		private FIA_Biosum_Manager.spc_groupings spc_groupings1;
		private FIA_Biosum_Manager.spc_groupings_collection spc_groupings_collection1;
		private System.Windows.Forms.Button btnNext;
		private System.Windows.Forms.Button btnPrev;
		private System.Windows.Forms.Button btnGrp1;
		private System.Windows.Forms.Button btnTreeAudit;
		private int m_intCurrGroupSet=1;
        private FIA_Biosum_Manager.spc_common_name spc_common_name1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.CheckBox chkFilterSpecies;
        private FIA_Biosum_Manager.spc_common_name_collection spc_common_name_collection1;


		Queries m_oQueries = new Queries();
        RxTools m_oRxTools = new RxTools();
        RxPackageItem_Collection m_oRxPackage_Collection = new RxPackageItem_Collection();

        string[] m_strFVSVariantsArray = null;

        // scenario-specific variables
        private string _strScenarioId = "";
        private frmProcessorScenario _frmProcessorScenario = null;
        // Help system variables
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultProcessorXPSFile;

        
		public uc_processor_scenario_tree_spc_groups()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			//create the species groupings
			this.spc_groupings_collection1 = new spc_groupings_collection();
			this.CreateSpcGrpBoxes(1);
			this.m_intCurrGroupSet=1;
			this.grpbox2.Visible=false;
			this.grpbox3.Visible=false;
			this.grpbox4.Visible=false;
			this.grpbox5.Visible=false;
			this.grpbox6.Visible=false;
			this.DisplayGroupSet(1);

            this.m_oQueries = new Queries();
            //load datasources
            //m_oQueries.m_oFvs.LoadDatasource = true;
            //m_oQueries.m_oFIAPlot.LoadDatasource = true;
            m_oQueries.LoadDatasources(true);
            m_oRxTools.LoadAllRxPackageItems(m_oRxPackage_Collection);


            //create links to all the fvstree tables
            this.m_oRxTools.CreateTableLinksToFVSOutTreeListTables(m_oQueries, m_oQueries.m_strTempDbFile);

            this.m_ado = new ado_data_access();
            spc_common_name_collection1 = new spc_common_name_collection();
            this.m_oEnv = new env();
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
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnTreeAudit = new System.Windows.Forms.Button();
            this.btnPrev = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnBoth = new System.Windows.Forms.Button();
            this.btnSwd = new System.Windows.Forms.Button();
            this.btnHwd = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnGrp6 = new System.Windows.Forms.Button();
            this.btnGrp5 = new System.Windows.Forms.Button();
            this.grpbox6 = new System.Windows.Forms.GroupBox();
            this.txtGrp6 = new System.Windows.Forms.TextBox();
            this.lstGrp6 = new System.Windows.Forms.ListBox();
            this.btnRemove6 = new System.Windows.Forms.Button();
            this.btnClearAll6 = new System.Windows.Forms.Button();
            this.btnGrp4 = new System.Windows.Forms.Button();
            this.grpbox5 = new System.Windows.Forms.GroupBox();
            this.txtGrp5 = new System.Windows.Forms.TextBox();
            this.lstGrp5 = new System.Windows.Forms.ListBox();
            this.btnRemove5 = new System.Windows.Forms.Button();
            this.btnClearAll5 = new System.Windows.Forms.Button();
            this.grpbox4 = new System.Windows.Forms.GroupBox();
            this.txtGrp4 = new System.Windows.Forms.TextBox();
            this.lstGrp4 = new System.Windows.Forms.ListBox();
            this.btnRemove4 = new System.Windows.Forms.Button();
            this.btnClearAll4 = new System.Windows.Forms.Button();
            this.btnGrp3 = new System.Windows.Forms.Button();
            this.btnGrp2 = new System.Windows.Forms.Button();
            this.grpbox3 = new System.Windows.Forms.GroupBox();
            this.txtGrp3 = new System.Windows.Forms.TextBox();
            this.lstGrp3 = new System.Windows.Forms.ListBox();
            this.btnRemove3 = new System.Windows.Forms.Button();
            this.btnClearAll3 = new System.Windows.Forms.Button();
            this.grpbox2 = new System.Windows.Forms.GroupBox();
            this.txtGrp2 = new System.Windows.Forms.TextBox();
            this.lstGrp2 = new System.Windows.Forms.ListBox();
            this.btnRemove2 = new System.Windows.Forms.Button();
            this.btnClearAll2 = new System.Windows.Forms.Button();
            this.grpbox1 = new System.Windows.Forms.GroupBox();
            this.txtGrp1 = new System.Windows.Forms.TextBox();
            this.lstGrp1 = new System.Windows.Forms.ListBox();
            this.btnRemove1 = new System.Windows.Forms.Button();
            this.btnClearAll1 = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnGrp1 = new System.Windows.Forms.Button();
            this.lstCommonName = new System.Windows.Forms.ListBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.chkFilterSpecies = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.grpbox6.SuspendLayout();
            this.grpbox5.SuspendLayout();
            this.grpbox4.SuspendLayout();
            this.grpbox3.SuspendLayout();
            this.grpbox2.SuspendLayout();
            this.grpbox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnTreeAudit);
            this.groupBox1.Controls.Add(this.btnPrev);
            this.groupBox1.Controls.Add(this.btnNext);
            this.groupBox1.Controls.Add(this.btnBoth);
            this.groupBox1.Controls.Add(this.btnSwd);
            this.groupBox1.Controls.Add(this.btnHwd);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.btnGrp6);
            this.groupBox1.Controls.Add(this.btnGrp5);
            this.groupBox1.Controls.Add(this.grpbox6);
            this.groupBox1.Controls.Add(this.btnGrp4);
            this.groupBox1.Controls.Add(this.grpbox5);
            this.groupBox1.Controls.Add(this.grpbox4);
            this.groupBox1.Controls.Add(this.btnGrp3);
            this.groupBox1.Controls.Add(this.btnGrp2);
            this.groupBox1.Controls.Add(this.grpbox3);
            this.groupBox1.Controls.Add(this.grpbox2);
            this.groupBox1.Controls.Add(this.grpbox1);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnGrp1);
            this.groupBox1.Controls.Add(this.lstCommonName);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.chkFilterSpecies);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(776, 592);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.txtLabel_Enter);
            // 
            // btnTreeAudit
            // 
            this.btnTreeAudit.Location = new System.Drawing.Point(189, 555);
            this.btnTreeAudit.Name = "btnTreeAudit";
            this.btnTreeAudit.Size = new System.Drawing.Size(140, 24);
            this.btnTreeAudit.TabIndex = 63;
            this.btnTreeAudit.Text = "Tree Audit Report";
            this.btnTreeAudit.Click += new System.EventHandler(this.btnTreeAudit_Click);
            // 
            // btnPrev
            // 
            this.btnPrev.Enabled = false;
            this.btnPrev.Location = new System.Drawing.Point(350, 555);
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.Size = new System.Drawing.Size(56, 24);
            this.btnPrev.TabIndex = 62;
            this.btnPrev.Text = "<--";
            this.btnPrev.Visible = false;
            this.btnPrev.Click += new System.EventHandler(this.btnPrev_Click);
            // 
            // btnNext
            // 
            this.btnNext.Enabled = false;
            this.btnNext.Location = new System.Drawing.Point(406, 555);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(56, 24);
            this.btnNext.TabIndex = 61;
            this.btnNext.Text = "-->";
            this.btnNext.Visible = false;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // btnBoth
            // 
            this.btnBoth.Enabled = false;
            this.btnBoth.Location = new System.Drawing.Point(93, 104);
            this.btnBoth.Name = "btnBoth";
            this.btnBoth.Size = new System.Drawing.Size(50, 24);
            this.btnBoth.TabIndex = 59;
            this.btnBoth.Text = "Both";
            this.btnBoth.Click += new System.EventHandler(this.btnBoth_Click);
            // 
            // btnSwd
            // 
            this.btnSwd.Location = new System.Drawing.Point(53, 104);
            this.btnSwd.Name = "btnSwd";
            this.btnSwd.Size = new System.Drawing.Size(40, 24);
            this.btnSwd.TabIndex = 58;
            this.btnSwd.Text = "Swd";
            this.btnSwd.Click += new System.EventHandler(this.btnSwd_Click);
            // 
            // btnHwd
            // 
            this.btnHwd.Location = new System.Drawing.Point(13, 104);
            this.btnHwd.Name = "btnHwd";
            this.btnHwd.Size = new System.Drawing.Size(40, 24);
            this.btnHwd.TabIndex = 57;
            this.btnHwd.Text = "Hwd";
            this.btnHwd.Click += new System.EventHandler(this.btnHwd_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(8, 554);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(56, 24);
            this.btnHelp.TabIndex = 56;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(468, 555);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(100, 24);
            this.btnAdd.TabIndex = 55;
            this.btnAdd.Text = "Add Groups";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnGrp6
            // 
            this.btnGrp6.Location = new System.Drawing.Point(152, 466);
            this.btnGrp6.Name = "btnGrp6";
            this.btnGrp6.Size = new System.Drawing.Size(88, 24);
            this.btnGrp6.TabIndex = 54;
            this.btnGrp6.Text = "Group 6 -->>";
            this.btnGrp6.Click += new System.EventHandler(this.btnGrp6_Click);
            // 
            // btnGrp5
            // 
            this.btnGrp5.Location = new System.Drawing.Point(152, 429);
            this.btnGrp5.Name = "btnGrp5";
            this.btnGrp5.Size = new System.Drawing.Size(88, 24);
            this.btnGrp5.TabIndex = 53;
            this.btnGrp5.Text = "Group 5 -->>";
            this.btnGrp5.Click += new System.EventHandler(this.btnGrp5_Click);
            // 
            // grpbox6
            // 
            this.grpbox6.Controls.Add(this.txtGrp6);
            this.grpbox6.Controls.Add(this.lstGrp6);
            this.grpbox6.Controls.Add(this.btnRemove6);
            this.grpbox6.Controls.Add(this.btnClearAll6);
            this.grpbox6.Location = new System.Drawing.Point(591, 345);
            this.grpbox6.Name = "grpbox6";
            this.grpbox6.Size = new System.Drawing.Size(160, 198);
            this.grpbox6.TabIndex = 52;
            this.grpbox6.TabStop = false;
            this.grpbox6.Text = "Group 6";
            this.grpbox6.Enter += new System.EventHandler(this.grpbox6_Enter);
            // 
            // txtGrp6
            // 
            this.txtGrp6.Location = new System.Drawing.Point(12, 24);
            this.txtGrp6.Name = "txtGrp6";
            this.txtGrp6.Size = new System.Drawing.Size(136, 22);
            this.txtGrp6.TabIndex = 34;
            this.txtGrp6.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGrp6_KeyPress);
            // 
            // lstGrp6
            // 
            this.lstGrp6.ItemHeight = 16;
            this.lstGrp6.Location = new System.Drawing.Point(12, 56);
            this.lstGrp6.Name = "lstGrp6";
            this.lstGrp6.Size = new System.Drawing.Size(136, 100);
            this.lstGrp6.TabIndex = 36;
            // 
            // btnRemove6
            // 
            this.btnRemove6.Location = new System.Drawing.Point(12, 168);
            this.btnRemove6.Name = "btnRemove6";
            this.btnRemove6.Size = new System.Drawing.Size(75, 24);
            this.btnRemove6.TabIndex = 38;
            this.btnRemove6.Text = "Remove";
            this.btnRemove6.Click += new System.EventHandler(this.btnRemove6_Click);
            // 
            // btnClearAll6
            // 
            this.btnClearAll6.Location = new System.Drawing.Point(91, 168);
            this.btnClearAll6.Name = "btnClearAll6";
            this.btnClearAll6.Size = new System.Drawing.Size(56, 24);
            this.btnClearAll6.TabIndex = 39;
            this.btnClearAll6.Text = "Clear All";
            this.btnClearAll6.Click += new System.EventHandler(this.btnClearAll6_Click);
            // 
            // btnGrp4
            // 
            this.btnGrp4.Location = new System.Drawing.Point(152, 390);
            this.btnGrp4.Name = "btnGrp4";
            this.btnGrp4.Size = new System.Drawing.Size(88, 24);
            this.btnGrp4.TabIndex = 51;
            this.btnGrp4.Text = "Group 4 -->>";
            this.btnGrp4.Click += new System.EventHandler(this.btnGrp4_Click);
            // 
            // grpbox5
            // 
            this.grpbox5.Controls.Add(this.txtGrp5);
            this.grpbox5.Controls.Add(this.lstGrp5);
            this.grpbox5.Controls.Add(this.btnRemove5);
            this.grpbox5.Controls.Add(this.btnClearAll5);
            this.grpbox5.Location = new System.Drawing.Point(421, 343);
            this.grpbox5.Name = "grpbox5";
            this.grpbox5.Size = new System.Drawing.Size(160, 200);
            this.grpbox5.TabIndex = 50;
            this.grpbox5.TabStop = false;
            this.grpbox5.Text = "Group 5";
            this.grpbox5.Enter += new System.EventHandler(this.grpbox5_Enter);
            // 
            // txtGrp5
            // 
            this.txtGrp5.Location = new System.Drawing.Point(12, 24);
            this.txtGrp5.Name = "txtGrp5";
            this.txtGrp5.Size = new System.Drawing.Size(136, 22);
            this.txtGrp5.TabIndex = 34;
            this.txtGrp5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGrp5_KeyPress);
            // 
            // lstGrp5
            // 
            this.lstGrp5.ItemHeight = 16;
            this.lstGrp5.Location = new System.Drawing.Point(12, 56);
            this.lstGrp5.Name = "lstGrp5";
            this.lstGrp5.Size = new System.Drawing.Size(136, 100);
            this.lstGrp5.TabIndex = 36;
            // 
            // btnRemove5
            // 
            this.btnRemove5.Location = new System.Drawing.Point(12, 168);
            this.btnRemove5.Name = "btnRemove5";
            this.btnRemove5.Size = new System.Drawing.Size(75, 24);
            this.btnRemove5.TabIndex = 38;
            this.btnRemove5.Text = "Remove";
            this.btnRemove5.Click += new System.EventHandler(this.btnRemove5_Click);
            // 
            // btnClearAll5
            // 
            this.btnClearAll5.Location = new System.Drawing.Point(91, 168);
            this.btnClearAll5.Name = "btnClearAll5";
            this.btnClearAll5.Size = new System.Drawing.Size(56, 24);
            this.btnClearAll5.TabIndex = 39;
            this.btnClearAll5.Text = "Clear All";
            this.btnClearAll5.Click += new System.EventHandler(this.btnClearAll5_Click);
            // 
            // grpbox4
            // 
            this.grpbox4.Controls.Add(this.txtGrp4);
            this.grpbox4.Controls.Add(this.lstGrp4);
            this.grpbox4.Controls.Add(this.btnRemove4);
            this.grpbox4.Controls.Add(this.btnClearAll4);
            this.grpbox4.Location = new System.Drawing.Point(247, 345);
            this.grpbox4.Name = "grpbox4";
            this.grpbox4.Size = new System.Drawing.Size(160, 198);
            this.grpbox4.TabIndex = 49;
            this.grpbox4.TabStop = false;
            this.grpbox4.Text = "Group 4";
            // 
            // txtGrp4
            // 
            this.txtGrp4.Location = new System.Drawing.Point(12, 24);
            this.txtGrp4.Name = "txtGrp4";
            this.txtGrp4.Size = new System.Drawing.Size(136, 22);
            this.txtGrp4.TabIndex = 34;
            this.txtGrp4.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGrp4_KeyPress);
            // 
            // lstGrp4
            // 
            this.lstGrp4.ItemHeight = 16;
            this.lstGrp4.Location = new System.Drawing.Point(12, 56);
            this.lstGrp4.Name = "lstGrp4";
            this.lstGrp4.Size = new System.Drawing.Size(136, 100);
            this.lstGrp4.TabIndex = 36;
            // 
            // btnRemove4
            // 
            this.btnRemove4.Location = new System.Drawing.Point(12, 168);
            this.btnRemove4.Name = "btnRemove4";
            this.btnRemove4.Size = new System.Drawing.Size(75, 24);
            this.btnRemove4.TabIndex = 38;
            this.btnRemove4.Text = "Remove";
            this.btnRemove4.Click += new System.EventHandler(this.btnRemove4_Click);
            // 
            // btnClearAll4
            // 
            this.btnClearAll4.Location = new System.Drawing.Point(91, 168);
            this.btnClearAll4.Name = "btnClearAll4";
            this.btnClearAll4.Size = new System.Drawing.Size(56, 24);
            this.btnClearAll4.TabIndex = 39;
            this.btnClearAll4.Text = "Clear All";
            this.btnClearAll4.Click += new System.EventHandler(this.btnClearAll4_Click);
            // 
            // btnGrp3
            // 
            this.btnGrp3.Location = new System.Drawing.Point(152, 256);
            this.btnGrp3.Name = "btnGrp3";
            this.btnGrp3.Size = new System.Drawing.Size(88, 24);
            this.btnGrp3.TabIndex = 48;
            this.btnGrp3.Text = "Group 3 -->>";
            this.btnGrp3.Click += new System.EventHandler(this.btnGrp3_Click);
            // 
            // btnGrp2
            // 
            this.btnGrp2.Location = new System.Drawing.Point(152, 214);
            this.btnGrp2.Name = "btnGrp2";
            this.btnGrp2.Size = new System.Drawing.Size(88, 24);
            this.btnGrp2.TabIndex = 47;
            this.btnGrp2.Text = "Group 2 -->>";
            this.btnGrp2.Click += new System.EventHandler(this.btnGrp2_Click);
            // 
            // grpbox3
            // 
            this.grpbox3.Controls.Add(this.txtGrp3);
            this.grpbox3.Controls.Add(this.lstGrp3);
            this.grpbox3.Controls.Add(this.btnRemove3);
            this.grpbox3.Controls.Add(this.btnClearAll3);
            this.grpbox3.Location = new System.Drawing.Point(591, 135);
            this.grpbox3.Name = "grpbox3";
            this.grpbox3.Size = new System.Drawing.Size(160, 201);
            this.grpbox3.TabIndex = 46;
            this.grpbox3.TabStop = false;
            this.grpbox3.Text = "Group 3";
            // 
            // txtGrp3
            // 
            this.txtGrp3.Location = new System.Drawing.Point(12, 24);
            this.txtGrp3.Name = "txtGrp3";
            this.txtGrp3.Size = new System.Drawing.Size(136, 22);
            this.txtGrp3.TabIndex = 34;
            this.txtGrp3.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGrp3_KeyPress);
            // 
            // lstGrp3
            // 
            this.lstGrp3.ItemHeight = 16;
            this.lstGrp3.Location = new System.Drawing.Point(12, 56);
            this.lstGrp3.Name = "lstGrp3";
            this.lstGrp3.Size = new System.Drawing.Size(136, 100);
            this.lstGrp3.TabIndex = 36;
            // 
            // btnRemove3
            // 
            this.btnRemove3.Location = new System.Drawing.Point(12, 168);
            this.btnRemove3.Name = "btnRemove3";
            this.btnRemove3.Size = new System.Drawing.Size(75, 24);
            this.btnRemove3.TabIndex = 38;
            this.btnRemove3.Text = "Remove";
            this.btnRemove3.Click += new System.EventHandler(this.btnRemove3_Click);
            // 
            // btnClearAll3
            // 
            this.btnClearAll3.Location = new System.Drawing.Point(91, 168);
            this.btnClearAll3.Name = "btnClearAll3";
            this.btnClearAll3.Size = new System.Drawing.Size(56, 24);
            this.btnClearAll3.TabIndex = 39;
            this.btnClearAll3.Text = "Clear All";
            this.btnClearAll3.Click += new System.EventHandler(this.btnClearAll3_Click);
            // 
            // grpbox2
            // 
            this.grpbox2.Controls.Add(this.txtGrp2);
            this.grpbox2.Controls.Add(this.lstGrp2);
            this.grpbox2.Controls.Add(this.btnRemove2);
            this.grpbox2.Controls.Add(this.btnClearAll2);
            this.grpbox2.Location = new System.Drawing.Point(423, 135);
            this.grpbox2.Name = "grpbox2";
            this.grpbox2.Size = new System.Drawing.Size(160, 201);
            this.grpbox2.TabIndex = 45;
            this.grpbox2.TabStop = false;
            this.grpbox2.Text = "Group2";
            // 
            // txtGrp2
            // 
            this.txtGrp2.Location = new System.Drawing.Point(12, 24);
            this.txtGrp2.Name = "txtGrp2";
            this.txtGrp2.Size = new System.Drawing.Size(136, 22);
            this.txtGrp2.TabIndex = 34;
            this.txtGrp2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGrp2_KeyPress);
            // 
            // lstGrp2
            // 
            this.lstGrp2.ItemHeight = 16;
            this.lstGrp2.Location = new System.Drawing.Point(12, 56);
            this.lstGrp2.Name = "lstGrp2";
            this.lstGrp2.Size = new System.Drawing.Size(136, 100);
            this.lstGrp2.TabIndex = 36;
            // 
            // btnRemove2
            // 
            this.btnRemove2.Location = new System.Drawing.Point(14, 168);
            this.btnRemove2.Name = "btnRemove2";
            this.btnRemove2.Size = new System.Drawing.Size(75, 24);
            this.btnRemove2.TabIndex = 38;
            this.btnRemove2.Text = "Remove";
            this.btnRemove2.Click += new System.EventHandler(this.btnRemove2_Click);
            // 
            // btnClearAll2
            // 
            this.btnClearAll2.Location = new System.Drawing.Point(93, 168);
            this.btnClearAll2.Name = "btnClearAll2";
            this.btnClearAll2.Size = new System.Drawing.Size(56, 24);
            this.btnClearAll2.TabIndex = 39;
            this.btnClearAll2.Text = "Clear All";
            this.btnClearAll2.Click += new System.EventHandler(this.btnClearAll2_Click);
            // 
            // grpbox1
            // 
            this.grpbox1.Controls.Add(this.txtGrp1);
            this.grpbox1.Controls.Add(this.lstGrp1);
            this.grpbox1.Controls.Add(this.btnRemove1);
            this.grpbox1.Controls.Add(this.btnClearAll1);
            this.grpbox1.Location = new System.Drawing.Point(247, 135);
            this.grpbox1.Name = "grpbox1";
            this.grpbox1.Size = new System.Drawing.Size(160, 201);
            this.grpbox1.TabIndex = 44;
            this.grpbox1.TabStop = false;
            this.grpbox1.Text = "Group 1";
            // 
            // txtGrp1
            // 
            this.txtGrp1.Location = new System.Drawing.Point(12, 24);
            this.txtGrp1.Name = "txtGrp1";
            this.txtGrp1.Size = new System.Drawing.Size(136, 22);
            this.txtGrp1.TabIndex = 34;
            this.txtGrp1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGrp1_KeyPress);
            // 
            // lstGrp1
            // 
            this.lstGrp1.ItemHeight = 16;
            this.lstGrp1.Location = new System.Drawing.Point(12, 56);
            this.lstGrp1.Name = "lstGrp1";
            this.lstGrp1.Size = new System.Drawing.Size(136, 100);
            this.lstGrp1.TabIndex = 36;
            // 
            // btnRemove1
            // 
            this.btnRemove1.Location = new System.Drawing.Point(12, 168);
            this.btnRemove1.Name = "btnRemove1";
            this.btnRemove1.Size = new System.Drawing.Size(75, 24);
            this.btnRemove1.TabIndex = 38;
            this.btnRemove1.Text = "Remove";
            this.btnRemove1.Click += new System.EventHandler(this.btnRemove1_Click);
            // 
            // btnClearAll1
            // 
            this.btnClearAll1.Location = new System.Drawing.Point(91, 168);
            this.btnClearAll1.Name = "btnClearAll1";
            this.btnClearAll1.Size = new System.Drawing.Size(56, 24);
            this.btnClearAll1.TabIndex = 39;
            this.btnClearAll1.Text = "Clear All";
            this.btnClearAll1.Click += new System.EventHandler(this.btnClearAll1_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(636, 555);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(65, 24);
            this.btnCancel.TabIndex = 43;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(572, 555);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(60, 24);
            this.btnSave.TabIndex = 41;
            this.btnSave.Text = "Save";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(712, 555);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(56, 24);
            this.btnClose.TabIndex = 37;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnGrp1
            // 
            this.btnGrp1.Location = new System.Drawing.Point(152, 176);
            this.btnGrp1.Name = "btnGrp1";
            this.btnGrp1.Size = new System.Drawing.Size(88, 24);
            this.btnGrp1.TabIndex = 35;
            this.btnGrp1.Text = "Group 1 -->>";
            this.btnGrp1.Click += new System.EventHandler(this.btnGroup1_Click);
            // 
            // lstCommonName
            // 
            this.lstCommonName.ItemHeight = 16;
            this.lstCommonName.Location = new System.Drawing.Point(8, 136);
            this.lstCommonName.Name = "lstCommonName";
            this.lstCommonName.Size = new System.Drawing.Size(136, 404);
            this.lstCommonName.TabIndex = 33;
            this.lstCommonName.SelectedIndexChanged += new System.EventHandler(this.lstCommonName_SelectedIndexChanged);
            this.lstCommonName.MouseHover += new System.EventHandler(this.lstCommonName_MouseHover);
            this.lstCommonName.MouseMove += new System.Windows.Forms.MouseEventHandler(this.lstCommonName_MouseMove);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(770, 32);
            this.lblTitle.TabIndex = 32;
            this.lblTitle.Text = "Species Groups";
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(8, 64);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(320, 16);
            this.label1.TabIndex = 66;
            this.label1.Text = "* Denotes a species found in the FVS tree tables";
            // 
            // chkFilterSpecies
            // 
            this.chkFilterSpecies.Location = new System.Drawing.Point(352, 64);
            this.chkFilterSpecies.Name = "chkFilterSpecies";
            this.chkFilterSpecies.Size = new System.Drawing.Size(400, 20);
            this.chkFilterSpecies.TabIndex = 67;
            this.chkFilterSpecies.Text = "Show only species found in the FVS Tree tables";
            this.chkFilterSpecies.CheckStateChanged += new System.EventHandler(this.chkFilterSpecies_CheckStateChanged);
            // 
            // uc_processor_scenario_tree_spc_groups
            // 
            this.AutoScroll = true;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_processor_scenario_tree_spc_groups";
            this.Size = new System.Drawing.Size(776, 592);
            this.groupBox1.ResumeLayout(false);
            this.grpbox6.ResumeLayout(false);
            this.grpbox6.PerformLayout();
            this.grpbox5.ResumeLayout(false);
            this.grpbox5.PerformLayout();
            this.grpbox4.ResumeLayout(false);
            this.grpbox4.PerformLayout();
            this.grpbox3.ResumeLayout(false);
            this.grpbox3.PerformLayout();
            this.grpbox2.ResumeLayout(false);
            this.grpbox2.PerformLayout();
            this.grpbox1.ResumeLayout(false);
            this.grpbox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void txtLabel_Enter(object sender, System.EventArgs e)
		{
		
		}

        public frmProcessorScenario ReferenceProcessorScenarioForm
        {
            get { return this._frmProcessorScenario; }
            set { this._frmProcessorScenario = value; }
        }
        public string ScenarioId
        {
            get { return _strScenarioId; }
            set { _strScenarioId = value; }
        }

        public void loadvalues()
        {

            int y = 0;
            int x = 0;
            int index = 0;

            this.lstCommonName.Sorted = true;
            this.m_intError = 0;

            this.m_ado = new ado_data_access();
            this.m_strConn = this.m_ado.getMDBConnString(this.m_oQueries.m_strTempDbFile, "", "");
            this.m_ado.OpenConnection(this.m_strConn);
            if (this.m_ado.m_intError != 0)
            {
                this.m_intError = this.m_ado.m_intError;
                this.m_ado = null;
                return;

            }
            //get all the variants in the plot table
            m_strFVSVariantsArray = frmMain.g_oUtils.ConvertListToArray(m_oRxTools.GetListOfFVSVariantsInPlotTable(m_ado, m_ado.m_OleDbConnection, m_oQueries.m_oFIAPlot.m_strPlotTable), ",");


            //create table links
            if (m_ado.TableExist(m_ado.m_OleDbConnection, "fvsouttreetemp"))
                m_ado.SqlNonQuery(m_ado.m_OleDbConnection, "DROP TABLE fvsouttreetemp");
            if (m_ado.TableExist(m_ado.m_OleDbConnection, "fvsouttreetemp2"))
                m_ado.SqlNonQuery(m_ado.m_OleDbConnection, "DROP TABLE fvsouttreetemp2");
            //append the multiple fvsout tree tables into a single fvsout tree table
            List<string> strSqlCommandList;

            strSqlCommandList = Queries.Processor.AuditFvsOut_SelectIntoUnionOfFVSTreeTablesUsingListArray(
                m_ado,
                m_ado.m_OleDbConnection,
                "fvsouttreetemp2",
                 m_oRxPackage_Collection,
                 m_strFVSVariantsArray,
                "fvs_tree_id,fvs_variant,fvs_species,FvsCreatedTree_YN,biosum_cond_id");

            for (x = 0; x <= strSqlCommandList.Count - 1; x++)
            {
                m_ado.SqlNonQuery(m_ado.m_OleDbConnection, strSqlCommandList[x]);
            }

            m_ado.m_strSQL = "SELECT DISTINCT * INTO fvsouttreetemp FROM fvsouttreetemp2";
            m_ado.SqlNonQuery(m_ado.m_OleDbConnection, m_ado.m_strSQL);

            this.m_strFvsOutTreeTable = "fvsouttreetemp";


            //GET ALL TREE SPECIES COMMON NAME
            /**********************************************************************************
             **process all tree species in the tree species table and initialize the 
             **unassigned array variable
             **********************************************************************************/
            m_ado.m_strSQL = "SELECT COUNT(*) FROM (SELECT DISTINCT common_name FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " WHERE spcd IS NOT NULL AND LEN(TRIM(common_name)) > 0 )";
            this.m_ado.m_strSQL = "SELECT DISTINCT common_name FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " WHERE spcd IS NOT NULL AND LEN(TRIM(common_name)) > 0";
            this.m_ado.SqlQueryReader(m_ado.m_OleDbConnection, m_ado.m_strSQL);
            if (this.m_ado.m_OleDbDataReader.HasRows)
            {
                while (this.m_ado.m_OleDbDataReader.Read())
                {
                    spc_common_name1 = new spc_common_name();
                    spc_common_name1.SpeciesCommonName = this.m_ado.m_OleDbDataReader["common_name"].ToString().Trim();
                    spc_common_name1.SpeciesGroupLabel = "";
                    spc_common_name1.SpeciesGroupIndex = -1;
                    spc_common_name1.FVSOutput = false;
                    this.spc_common_name_collection1.Add(spc_common_name1);

                }
            }
            this.m_ado.m_OleDbDataReader.Close();

            //ASSIGN A TREE SPECIES CODE TO THE SPECIES COMMON NAME
            /**********************************************************************************
             **process all tree species in the tree species table and initialize the 
             **unassigned array variable
             **********************************************************************************/
            this.m_ado.m_strSQL = "SELECT DISTINCT spcd,common_name FROM " + this.m_oQueries.m_oFvs.m_strTreeSpcTable + " WHERE spcd IS NOT NULL AND LEN(TRIM(common_name)) > 0";
            this.m_ado.SqlQueryReader(m_ado.m_OleDbConnection, m_ado.m_strSQL);
            if (this.m_ado.m_OleDbDataReader.HasRows)
            {
                while (this.m_ado.m_OleDbDataReader.Read())
                {
                    for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
                    {
                        if (this.spc_common_name_collection1.Item(x).SpeciesCommonName.Trim().ToUpper() ==
                            this.m_ado.m_OleDbDataReader["common_name"].ToString().Trim().ToUpper())
                        {
                            this.spc_common_name_collection1.Item(x).SpeciesCode = Convert.ToInt32(this.m_ado.m_OleDbDataReader["spcd"]);
                            break;
                        }

                    }

                }
            }
            this.m_ado.m_OleDbDataReader.Close();


            //GET FVS OUTPUT TREE SPECIES COMMON NAME
            /***************************************************************************
             **process tree species records that match up between the fvs output tree 
             **table, tree table, and tree species table
             ***************************************************************************/
            this.m_ado.m_strSQL = "SELECT DISTINCT t.spcd,f.fvs_variant " +
                                  "INTO tree_spc_groups_temp " +
                                  "FROM " + this.m_oQueries.m_oFIAPlot.m_strTreeTable + " t, " +
                                            this.m_strFvsOutTreeTable + " f " +
                                  "WHERE f.fvs_tree_id = t.fvs_tree_id AND f.biosum_cond_id = t.biosum_cond_id";
            this.m_ado.SqlNonQuery(m_ado.m_OleDbConnection, this.m_ado.m_strSQL);

            m_ado.m_strSQL = "SELECT DISTINCT s.common_name " +
                  "FROM tree_spc_groups_temp t, " + m_oQueries.m_oFvs.m_strTreeSpcTable + " s " +
                  "WHERE t.spcd = s.spcd AND " +
                  "TRIM(UCASE(t.fvs_variant)) = TRIM(UCASE(s.fvs_variant))";

            this.m_ado.SqlQueryReader(m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
            if (this.m_ado.m_OleDbDataReader.HasRows)
            {
                while (this.m_ado.m_OleDbDataReader.Read())
                {
                    for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
                    {
                        if (this.spc_common_name_collection1.Item(x).SpeciesCommonName.Trim().ToUpper() ==
                            this.m_ado.m_OleDbDataReader["common_name"].ToString().Trim().ToUpper())
                            this.spc_common_name_collection1.Item(x).FVSOutput = true;

                    }
                }
            }

            //GET FVS OUTPUT TREE SPECIES COMMON NAME FOR FVS-CREATED TREES
            /***************************************************************************
             **process tree species records that match up between the fvs output tree 
             **table and tree species table for fvs-created tree species that may be missing from the tree table
             ***************************************************************************/
            m_ado.m_strSQL = "SELECT DISTINCT s.common_name " +
                  "FROM " + this.m_strFvsOutTreeTable + " t, " + m_oQueries.m_oFvs.m_strTreeSpcTable + " s " +
                  "WHERE t.FvsCreatedTree_YN = 'Y' AND " +
                  "val(t.fvs_species) = s.spcd AND " +
                  "TRIM(UCASE(t.fvs_variant)) = TRIM(UCASE(s.fvs_variant))";

            this.m_ado.SqlQueryReader(m_ado.m_OleDbConnection, this.m_ado.m_strSQL);
            if (this.m_ado.m_OleDbDataReader.HasRows)
            {
                while (this.m_ado.m_OleDbDataReader.Read())
                {
                    for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
                    {
                        if (this.spc_common_name_collection1.Item(x).SpeciesCommonName.Trim().ToUpper() ==
                            this.m_ado.m_OleDbDataReader["common_name"].ToString().Trim().ToUpper())
                            this.spc_common_name_collection1.Item(x).FVSOutput = true;

                    }
                }
            }


            this.m_ado.m_OleDbDataReader.Close();
			
            //ScenarioId = this.ReferenceProcessorScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            //string strScenarioMDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
            //    "\\processor" + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsDbFile;
            //ReferenceProcessorScenarioForm.m_oProcessorScenarioTools.LoadTreeSpeciesGroupValues(strScenarioMDB,
            //    ScenarioId, ReferenceProcessorScenarioForm.m_oProcessorScenarioItem);

             //LOAD USER SPECIES COMMON NAME GROUPING ASSIGNMENTS
            /****************************************************************************************
             **load any previous group assignments 
             ****************************************************************************************/

            /**************************************************************************************
             **go through the species group table and assign values to the group label text box
             **************************************************************************************/
            this._strScenarioId = this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.ScenarioId;    
            if (this.ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count > 0)
                {
                    for (x = 0; x <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count - 1; x++)
                    {
                        ProcessorScenarioItem.SpcGroupItem p_oSpeciesGroupItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Item(x);
                        while (this.spc_groupings_collection1.Count <= p_oSpeciesGroupItem.SpeciesGroup)
                        {
                            this.CreateSpcGrpBoxes((this.spc_groupings_collection1.Count / 6) + 1);
                            if (this.btnNext.Enabled == false) this.btnNext.Enabled = true;
                            if (this.btnNext.Visible == false) this.btnNext.Visible = true;
                            if (this.btnPrev.Visible == false) this.btnPrev.Visible = true;
                        }
                        this.spc_groupings_collection1.Item(p_oSpeciesGroupItem.SpeciesGroup - 1).GroupLabel = p_oSpeciesGroupItem.SpeciesGroupLabel;

                    }
                //go through the species table and load its 
                //group list box with the species common name
                for (y = 0; y <= ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Count - 1; y++)
                {
                    ProcessorScenarioItem.SpcGroupListItem p_oSpeciesGroupListItem = ReferenceProcessorScenarioForm.m_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(y);

                    this.spc_groupings_collection1.Item(p_oSpeciesGroupListItem.SpeciesGroup - 1).m_lstGrp.Items.Add(p_oSpeciesGroupListItem.CommonName);
                    index = this.spc_groupings_collection1.Item(p_oSpeciesGroupListItem.SpeciesGroup - 1).m_lstGrp.Items.Count-1;
                    for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
                    {
                        if (this.spc_common_name_collection1.Item(x).SpeciesCommonName.Trim().ToUpper() ==
                            p_oSpeciesGroupListItem.CommonName.Trim().ToUpper())
                        {
                            this.spc_common_name_collection1.Item(x).SpeciesGroupIndex = p_oSpeciesGroupListItem.SpeciesGroup - 1;
                            this.spc_common_name_collection1.Item(x).SpeciesGroupLabel = this.spc_groupings_collection1.Item(p_oSpeciesGroupListItem.SpeciesGroup - 1).GroupLabel.ToString();
                            if (this.spc_common_name_collection1.Item(x).FVSOutput)
                                this.spc_groupings_collection1.Item(p_oSpeciesGroupListItem.SpeciesGroup - 1).m_lstGrp.Items[index] = this.spc_common_name_collection1.Item(x).SpeciesCommonName + "*";
                        }
                    }
					
                 }
            }
            this.loadUnassignedSpc();
        }
		private void RemoveGroupAssignment(string p_strCommonName)
		{
		}
		private void AddGroupAssignment(string p_strCommonName)
		{
		}
		public void btnClose_Click(object sender, System.EventArgs e)
		{
            if (this.btnSave.Enabled)
            {
                DialogResult result = MessageBox.Show("Do you wish to close without saving? Y/N", "FIA Biosum", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                switch (result)
                {
                    case DialogResult.No:
                        break;
                    case DialogResult.Cancel:
                        break;
                    default:
                        this.btnSave.Enabled = false;
                        this.ParentForm.Close();
                        break;
                }

            }
            else
            {
                this.ParentForm.Close();
            }
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
            if (this.btnSave.Enabled)
            {
                DialogResult result = MessageBox.Show("Do you wish to cancel without saving? Y/N", "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                switch (result)
                {
                    case DialogResult.No:
                        break;
                    default:
                        this.btnSave.Enabled = false;
                        this.ParentForm.Close();
                        break;
                }

            }
            else
            {
                this.ParentForm.Close();
            }
		}

		private void lstCommonName_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			
		}

		private void lstCommonName_MouseHover(object sender, System.EventArgs e)
		{
			
		}

		private void lstCommonName_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			string strTip="";
			int nIdx = this.lstCommonName.IndexFromPoint(e.X,e.Y);
			if ((nIdx >= 0) && (nIdx < this.lstCommonName.Items.Count))
			{
				strTip = this.lstCommonName.Items[nIdx].ToString();
				this.toolTip1.SetToolTip(this.lstCommonName,strTip);
			}
		}

		private void txtGrp1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{

			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void txtGrp2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void txtGrp3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void txtGrp4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void txtGrp5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void txtGrp6_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
		}

		private void btnGroup1_Click(object sender, System.EventArgs e)
		{
			AddSpeciesCommonNameToGroupAssignment(6 * this.m_intCurrGroupSet - 6);
		}
		public void AddSpeciesCommonNameToGroupAssignment(int p_intGroupListIndex)
		{
			if (this.lstCommonName.SelectedItems.Count > 0)
			{
				string strCommonName=this.lstCommonName.SelectedItems[0].ToString().Replace("*","");
				this.spc_groupings_collection1.Item(p_intGroupListIndex).m_lstGrp.Items.Add(this.lstCommonName.SelectedItems[0].ToString().Trim());
				this.lstCommonName.Items.Remove(this.lstCommonName.SelectedItems[0]);


                for (int x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
                {
                    if (this.spc_common_name_collection1.Item(x).SpeciesCommonName.Trim().ToUpper() ==
                        strCommonName.Trim().ToUpper())
                    {
                        this.spc_common_name_collection1.Item(x).SpeciesGroupIndex = p_intGroupListIndex;
                        this.spc_common_name_collection1.Item(x).SpeciesGroupLabel = this.spc_groupings_collection1.Item(p_intGroupListIndex).m_txtGrp.Text;
                    }
                }
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}
		public void RemoveSpeciesCommonNameFromGroupAssignment(string p_strCommonName)
		{
			string strCommonName=p_strCommonName.Replace("*","");
            for (int x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
            {
                if (this.spc_common_name_collection1.Item(x).SpeciesCommonName.Trim().ToUpper() ==
                    strCommonName.Trim().ToUpper())
                {
                    this.spc_common_name_collection1.Item(x).SpeciesGroupIndex = -1;
                    this.spc_common_name_collection1.Item(x).SpeciesGroupLabel = "";
                    AddSpeciesCommonNameToUnassignedList(x);
                }
            }
		}
		private void AddSpeciesCommonNameToUnassignedList(int index)
		{
			if (this.btnHwd.Enabled==false)
			{
                if (this.spc_common_name_collection1.Item(index).SpeciesCode > 299)
                {
                    if (this.spc_common_name_collection1.Item(index).FVSOutput)
                    {
                        this.lstCommonName.Items.Add(this.spc_common_name_collection1.Item(index).SpeciesCommonName + "*");
                    }
                    else
                    {
                        if (this.chkFilterSpecies.Checked == false)
                            this.lstCommonName.Items.Add(this.spc_common_name_collection1.Item(index).SpeciesCommonName);
                    }
                }

			}
			else if (this.btnSwd.Enabled==false)
			{
                if (this.spc_common_name_collection1.Item(index).SpeciesCode > 0 &&
                    this.spc_common_name_collection1.Item(index).SpeciesCode < 300)
                {
                    if (this.spc_common_name_collection1.Item(index).FVSOutput)
                    {
                        this.lstCommonName.Items.Add(this.spc_common_name_collection1.Item(index).SpeciesCommonName + "*");
                    }
                    else
                    {
                        if (this.chkFilterSpecies.Checked == false)
                            this.lstCommonName.Items.Add(this.spc_common_name_collection1.Item(index).SpeciesCommonName);
                    }
                }
			}
			else
			{
                if (this.spc_common_name_collection1.Item(index).FVSOutput)
                {
                    this.lstCommonName.Items.Add(this.spc_common_name_collection1.Item(index).SpeciesCommonName + "*");
                }
                else
                {
                    if (this.chkFilterSpecies.Checked == false)
                        this.lstCommonName.Items.Add(this.spc_common_name_collection1.Item(index).SpeciesCommonName);
                }
			}

		}

		private void btnGrp2_Click(object sender, System.EventArgs e)
		{
			this.AddSpeciesCommonNameToGroupAssignment(6 * this.m_intCurrGroupSet - 5);
		}

		private void btnGrp3_Click(object sender, System.EventArgs e)
		{
			this.AddSpeciesCommonNameToGroupAssignment(6 * this.m_intCurrGroupSet - 4);
		}

		private void btnGrp4_Click(object sender, System.EventArgs e)
		{
			this.AddSpeciesCommonNameToGroupAssignment(6 * this.m_intCurrGroupSet - 3);
		}

		private void btnGrp5_Click(object sender, System.EventArgs e)
		{
			this.AddSpeciesCommonNameToGroupAssignment(6 * this.m_intCurrGroupSet - 2);
		}

		private void btnGrp6_Click(object sender, System.EventArgs e)
		{
			this.AddSpeciesCommonNameToGroupAssignment(6 * this.m_intCurrGroupSet - 1);
		}

		private void btnRemove1_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp1.SelectedItems.Count > 0)
			{
				this.lstCommonName.Items.Add(this.lstGrp1.SelectedItems[0].ToString());
				this.lstGrp1.Items.Remove(this.lstGrp1.SelectedItems[0]);
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}

		private void btnHwd_Click(object sender, System.EventArgs e)
		{
			this.btnHwd.Enabled=false;
			this.btnSwd.Enabled=true;
			this.btnBoth.Enabled=true;
			this.loadUnassignedSpc();

		
		}
		private void loadUnassignedSpc()
		{
			int x;
			this.lstCommonName.Items.Clear();
            for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
            {
                if (this.spc_common_name_collection1.Item(x).SpeciesGroupIndex < 0)
                {
                    this.AddSpeciesCommonNameToUnassignedList(x);
                }
            }
		}

		private void btnBoth_Click(object sender, System.EventArgs e)
		{
			this.btnBoth.Enabled=false;
			this.btnHwd.Enabled=true;
			this.btnSwd.Enabled=true;
			this.loadUnassignedSpc();
		}

		private void btnSwd_Click(object sender, System.EventArgs e)
		{
			this.btnBoth.Enabled=true;
			this.btnHwd.Enabled=true;
			this.btnSwd.Enabled=false;
			this.loadUnassignedSpc();
		}

		private void btnRemove2_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp2.SelectedItems.Count > 0)
			{
				this.lstCommonName.Items.Add(this.lstGrp2.SelectedItems[0].ToString());
				this.lstGrp2.Items.Remove(this.lstGrp2.SelectedItems[0]);
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}

		private void btnRemove3_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp3.SelectedItems.Count > 0)
			{
				this.lstCommonName.Items.Add(this.lstGrp3.SelectedItems[0].ToString());
				this.lstGrp3.Items.Remove(this.lstGrp3.SelectedItems[0]);
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}

		private void btnRemove4_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp4.SelectedItems.Count > 0)
			{
				this.lstCommonName.Items.Add(this.lstGrp4.SelectedItems[0].ToString());
				this.lstGrp4.Items.Remove(this.lstGrp4.SelectedItems[0]);
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}

		private void btnRemove5_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp5.SelectedItems.Count > 0)
			{
				this.lstCommonName.Items.Add(this.lstGrp5.SelectedItems[0].ToString());
				this.lstGrp5.Items.Remove(this.lstGrp5.SelectedItems[0]);
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}

		private void btnRemove6_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp6.SelectedItems.Count > 0)
			{
				this.lstCommonName.Items.Add(this.lstGrp6.SelectedItems[0].ToString());
				this.lstGrp6.Items.Remove(this.lstGrp6.SelectedItems[0]);
				if (this.btnSave.Enabled==false) this.btnSave.Enabled=true;
			}
		}

		private void btnClearAll1_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp1.Items.Count > 0)
			{
				for (int x=0;x<=this.lstGrp1.Items.Count-1;x++)
				{
					this.lstCommonName.Items.Add(this.lstGrp1.Items[x].ToString());

				}
			    this.lstGrp1.Items.Clear();
			}
		}

		private void btnClearAll2_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp2.Items.Count > 0)
			{
				for (int x=0;x<=this.lstGrp2.Items.Count-1;x++)
				{
					this.lstCommonName.Items.Add(this.lstGrp2.Items[x].ToString());

				}
				this.lstGrp2.Items.Clear();
			}
		}

		private void btnClearAll3_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp3.Items.Count > 0)
			{
				for (int x=0;x<=this.lstGrp3.Items.Count-1;x++)
				{
					this.lstCommonName.Items.Add(this.lstGrp3.Items[x].ToString());

				}
				this.lstGrp3.Items.Clear();
			}
		}

		private void btnClearAll4_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp4.Items.Count > 0)
			{
				for (int x=0;x<=this.lstGrp4.Items.Count-1;x++)
				{
					this.lstCommonName.Items.Add(this.lstGrp4.Items[x].ToString());

				}
				this.lstGrp4.Items.Clear();
			}
		}

		private void btnClearAll5_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp5.Items.Count > 0)
			{
				for (int x=0;x<=this.lstGrp5.Items.Count-1;x++)
				{
					this.lstCommonName.Items.Add(this.lstGrp5.Items[x].ToString());

				}
				this.lstGrp5.Items.Clear();
			}
		}

		private void btnClearAll6_Click(object sender, System.EventArgs e)
		{
			if (this.lstGrp6.Items.Count > 0)
			{
				for (int x=0;x<=this.lstGrp6.Items.Count-1;x++)
				{
					this.lstCommonName.Items.Add(this.lstGrp6.Items[x].ToString());

				}
				this.lstGrp6.Items.Clear();
			}
		}
		public void savevalues()
		{
			val_data();
			if (this.m_intError==0)
			{
				try
				{
					int x=0;
					//save to the species groups table
				    DialogResult result = MessageBox.Show("Save Tree Species Group Data? Y/N","Tree Species Groups",System.Windows.Forms.MessageBoxButtons.YesNoCancel,System.Windows.Forms.MessageBoxIcon.Question);
					if (result == DialogResult.Yes)
					{
						string str;
						string strCommonName;
                        int intSpCd;
					    int intSpcGrp;
						string strSavedList="";
						int intGrpCollection;
						string strGrpLabel;
                        string strSql;

                        //
                        //OPEN CONNECTION TO DB FILE CONTAINING PROCESSOR SCENARIO TABLES
                        //
                        //scenario mdb connection                        
                        ado_data_access oAdo = new ado_data_access();
                        SQLite.ADO.DataMgr oDataMgr = null;
                        if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                        {
                            string strScenarioMDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                                "\\processor\\db\\scenario_processor_rule_definitions.mdb";
                            oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB, "", ""));
                            if (oAdo.m_intError == 0)
                            {
                                //delete the current groups
                                //delete all records from the tree species group table
                                oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName +
                                    " WHERE TRIM(UCASE(scenario_id))='" + _strScenarioId.Trim().ToUpper() + "'";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                if (oAdo.m_intError != 0) return;

                                //delete all records from the tree species group list table
                                oAdo.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                                    " WHERE TRIM(UCASE(scenario_id))='" + _strScenarioId.Trim().ToUpper() + "'";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                            }
                            m_intError = oAdo.m_intError;
                        }
                        else
                        {
                            oDataMgr = new SQLite.ADO.DataMgr();
                            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                            oDataMgr.OpenConnection(oDataMgr.GetConnectionString(strScenarioDB));
                            if (oDataMgr.m_intError == 0)
                            {
                                //delete the current groups
                                //delete all records from the tree species group table
                                oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName +
                                    " WHERE TRIM(UPPER(scenario_id))='" + _strScenarioId.Trim().ToUpper() + "'";
                                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                                if (oDataMgr.m_intError != 0) return;

                                //delete all records from the tree species group list table
                                oDataMgr.m_strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                                    " WHERE TRIM(UPPER(scenario_id))='" + _strScenarioId.Trim().ToUpper() + "'";
                                oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                            }
                            m_intError = oDataMgr.m_intError;
                        }

                        if (m_intError == 0)
                        {
                            for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
                            {
                                    if (this.spc_common_name_collection1.Item(x).SpeciesGroupIndex >= 0)
                                    {

                                        intSpcGrp = Convert.ToInt32(this.spc_common_name_collection1.Item(x).SpeciesGroupIndex);
                                        strCommonName = this.spc_common_name_collection1.Item(x).SpeciesCommonName;
                                        strCommonName = oAdo.FixString(strCommonName.Trim(), "'", "''");
                                        intSpCd = this.spc_common_name_collection1.Item(x).SpeciesCode;
                                        intGrpCollection = Math.Abs(intSpcGrp / 6);
                                        strGrpLabel = this.spc_groupings_collection1.Item(intSpcGrp).GroupLabel;
                                        strGrpLabel = strGrpLabel.Replace(' ', '_');
                                        str = "," + strGrpLabel.Trim() + ",";
                                    if (!ReferenceProcessorScenarioForm.m_bUsingSqlite)
                                    {
                                        if (strSavedList.IndexOf(str, 0) < 0)
                                        {
                                            oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " " +
                                                "(SPECIES_GROUP,SPECIES_LABEL,SCENARIO_ID) VALUES " +
                                                "(" + Convert.ToString(intSpcGrp + 1).Trim() + ",'" + strGrpLabel.Trim() + "','" + _strScenarioId.Trim() + "');";
                                            oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                            strSavedList += str;
                                        }

                                        oAdo.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " " +
                                            "(SPECIES_GROUP,common_name,SCENARIO_ID,SPCD) VALUES " +
                                            "(" + Convert.ToString(intSpcGrp + 1).Trim() + ",'" + strCommonName + "','" + _strScenarioId.Trim() + "', " +
                                            intSpCd + " );";
                                        oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                                    }
                                    else
                                    {
                                        if (strSavedList.IndexOf(str, 0) < 0)
                                        {
                                            oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " " +
                                                "(SPECIES_GROUP,SPECIES_LABEL,SCENARIO_ID) VALUES " +
                                                "(" + Convert.ToString(intSpcGrp + 1).Trim() + ",'" + strGrpLabel.Trim() + "','" + _strScenarioId.Trim() + "');";
                                            oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                                            strSavedList += str;
                                        }

                                        oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " " +
                                            "(SPECIES_GROUP,common_name,SCENARIO_ID,SPCD) VALUES " +
                                            "(" + Convert.ToString(intSpcGrp + 1).Trim() + ",'" + strCommonName + "','" + _strScenarioId.Trim() + "', " +
                                            intSpCd + " );";
                                        oDataMgr.SqlNonQuery(oDataMgr.m_Connection, oDataMgr.m_strSQL);
                                    }

                                }
                            }
                            this.btnSave.Enabled = false;
                        }
                        if (oAdo != null && oAdo.m_OleDbConnection != null)
                        {
                            if (oAdo.m_OleDbConnection.State == System.Data.ConnectionState.Open)
                                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                            oAdo = null;
                        }
                        else if (oDataMgr != null)
                        {
                            oDataMgr.CloseConnection(oDataMgr.m_Connection);
                            oDataMgr = null;
                        }

                    }
                    else if (result == DialogResult.Cancel)
                    {
                        this.ParentForm.DialogResult = DialogResult.Cancel;
                    }					
				}
				catch (Exception e)
				{
					MessageBox.Show("!!Error!! \n" +
                        "Module - uc_processor_scenario_tree_spc_groups:savevalues() \n" + 
						"Err Msg - " + e.Message,
						"FIA Biosum",System.Windows.Forms.MessageBoxButtons.OK,
						System.Windows.Forms.MessageBoxIcon.Exclamation);

					this.m_intError=-1;
				}
			}


		}

		private void val_data()
		{
			this.m_intError=0;

			for (int x=0; x<= this.spc_groupings_collection1.Count-1;x++)
			{
				//check if group has items in list but no group label
				if (this.spc_groupings_collection1.Item(x).m_lstGrp.Items.Count > 0)
				{
					if (this.spc_groupings_collection1.Item(x).GroupLabel.Trim().Length == 0)
					{
						MessageBox.Show("!!Enter A Group Label For Group " + Convert.ToString(x + 1).Trim() + "!!",
							              "Tree Species Groups",
							              System.Windows.Forms.MessageBoxButtons.OK,
							              System.Windows.Forms.MessageBoxIcon.Exclamation);
						this.m_intError=-1;
						return;

					}

					//make sure no duplicate group label
					string strCurGroup;
					string strGroup;
					strCurGroup = this.spc_groupings_collection1.Item(x).GroupLabel.Trim();
					strCurGroup = strCurGroup.Replace(' ','_');
					for (int y=x+1; y<= this.spc_groupings_collection1.Count-1;y++)
					{
						strGroup = this.spc_groupings_collection1.Item(y).GroupLabel.Trim();
						if (strGroup.Trim().Length > 0)
						{
							strGroup = strGroup.Replace(' ','_');
							if (strGroup.Trim().ToUpper() == strCurGroup.Trim().ToUpper())
							{
								MessageBox.Show("!!Cannot Have Duplicate Group Labels For " + strGroup + "!!",
									"Tree Species Groups",
									System.Windows.Forms.MessageBoxButtons.OK,
									System.Windows.Forms.MessageBoxIcon.Exclamation);
								this.m_intError=-1;
								return;
							}
						}

					}

				}

				


			}


		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			this.savevalues();
            // Force reload of components that use tree groups since they changed in db
            ReferenceProcessorScenarioForm.m_bTreeGroupsFirstTime = true;
            // Copied values have been saved
            ReferenceProcessorScenarioForm.m_bTreeGroupsCopied = false;
		}
		/// <summary>
		/// all items in the list box are  contantenated to a comma delimited string
		/// </summary>
		/// <param name="p_listbox">list box object</param>
		/// <returns>string object</returns>
		private string CreateCommaDelimitedString(System.Windows.Forms.ListBox p_listbox)
		{
			string str="";
			for (int x=0;x<=p_listbox.Items.Count;x++)
			{
				if (str.Trim().Length == 0)
				{
					str = "'" + p_listbox.Items[x].ToString().Trim() + "'";
				}
				else
				{
					str = ",'" + p_listbox.Items[x].ToString().Trim() + "'";
				}
			}
			return str;
		}
		private void uniquespcvalues()
		{

		}
		private void DisplayGroupSet(int intNewGroupSet)
		{
			int x=0;
			int y=1;
			//turn current groupset visibility off if changing to a 
			//different group set
			if (intNewGroupSet != this.m_intCurrGroupSet)
			{

				for (x=this.m_intCurrGroupSet * 6 - 6;x<=this.m_intCurrGroupSet*6-1;x++)
				{
                    this.spc_groupings_collection1.Item(x).Visible=false;
				}
			}
			for (x=intNewGroupSet * 6 - 6;x<=intNewGroupSet*6-1;x++)
			{
				this.spc_groupings_collection1.Item(x).Visible=true;
				switch (y)
				{
					case 1:
						this.btnGrp1.Text = "Group " + Convert.ToString(x + 1).Trim() + "-->>";
						break;
					case 2:
						this.btnGrp2.Text = "Group " + Convert.ToString(x + 1).Trim() + "-->>";
						break;
					case 3:
						this.btnGrp3.Text = "Group " + Convert.ToString(x + 1).Trim() + "-->>";
						break;
					case 4:
						this.btnGrp4.Text = "Group " + Convert.ToString(x + 1).Trim() + "-->>";
						break;
					case 5:
						this.btnGrp5.Text = "Group " + Convert.ToString(x + 1).Trim() + "-->>";
						break;
					case 6:
						this.btnGrp6.Text = "Group " + Convert.ToString(x + 1).Trim() + "-->>";
						break;
				}
				y++;
				
			}
			this.m_intCurrGroupSet = intNewGroupSet;
		}
		private void CreateSpcGrpBoxes(int intGroupSet)
		{
			int y=1;
			for (int x=intGroupSet * 6 - 6;x<=intGroupSet*6-1;x++)
			{
				this.spc_groupings1 = new spc_groupings(this);
				this.groupBox1.Controls.Add(this.spc_groupings1);
				this.spc_groupings1.Name="spcgrp" + Convert.ToString(x + 1).Trim();
				this.spc_groupings1.m_lstGrp.Name = "lstGrp" + Convert.ToString(x + 1).Trim();
				this.spc_groupings1.m_txtGrp.Name = "txtGrp" + Convert.ToString(x + 1).Trim();
				this.spc_groupings1.Size = this.grpbox1.Size;
				this.spc_groupings1.Text = "Group " + Convert.ToString(x + 1).Trim();
				this.spc_groupings1.m_txtGrp.Size = this.txtGrp1.Size;
                this.spc_groupings1.m_lstGrp.Size = this.lstGrp1.Size;
				this.spc_groupings1.m_txtGrp.Location = this.txtGrp1.Location;
				this.spc_groupings1.m_btnClearAll.Size = this.btnClearAll1.Size;
				this.spc_groupings1.m_btnClearAll.Location = this.btnClearAll1.Location;
				this.spc_groupings1.m_btnRemove.Size = this.btnRemove1.Size;
				this.spc_groupings1.m_btnRemove.Location= this.btnRemove1.Location;
				this.spc_groupings1.m_lstGrp.Location = this.lstGrp1.Location;
				switch (y)
				{
					case 1:
				        this.spc_groupings1.Location = this.grpbox1.Location;
						break;
					case 2:
						this.spc_groupings1.Top = this.grpbox1.Top;
						this.spc_groupings1.Left = this.grpbox1.Left + 
							                       this.grpbox1.Width + 2;
						break;
					case 3:
						this.spc_groupings1.Top = this.grpbox1.Top;
						this.spc_groupings1.Left = this.grpbox1.Left + 
												   (this.grpbox1.Width * 2) + 2;

						break;
					case 4:
						this.spc_groupings1.Top = this.grpbox1.Top +
							                      this.grpbox1.Height + 2;
						this.spc_groupings1.Left = this.grpbox1.Left;
							
						break;
					case 5:
						this.spc_groupings1.Top = this.grpbox1.Top + 
							                      this.grpbox1.Height + 2;
						this.spc_groupings1.Left = this.grpbox1.Left + 
							this.grpbox1.Width + 2;

						
						break;
					case 6:
						this.spc_groupings1.Top = this.grpbox1.Top + 
							                      this.grpbox1.Height + 2;
						this.spc_groupings1.Left = this.grpbox1.Left + 
							(this.grpbox1.Width * 2) + 2;

						break;
				}
				this.spc_groupings1.GroupingNumber = x + 1;

				this.spc_groupings_collection1.Add(this.spc_groupings1);
                y++;
			}
			this.grpbox1.Visible=false;
		}

		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			int intGrpSet = this.spc_groupings_collection1.Count / 6;
			this.CreateSpcGrpBoxes(intGrpSet + 1);
			this.DisplayGroupSet(intGrpSet + 1);
			if (this.btnPrev.Enabled==false) this.btnPrev.Enabled=true;
			if (this.btnNext.Enabled==true) this.btnNext.Enabled=false;
			if (this.btnPrev.Visible==false) this.btnPrev.Visible=true;
			if (this.btnNext.Visible==false) this.btnNext.Visible=true;
		
		}

		private void btnPrev_Click(object sender, System.EventArgs e)
		{
			if (this.m_intCurrGroupSet == 2) this.btnPrev.Enabled=false;
			if (this.btnNext.Enabled==false) this.btnNext.Enabled=true;
			this.DisplayGroupSet(this.m_intCurrGroupSet-1);
		}

		private void btnNext_Click(object sender, System.EventArgs e)
		{
			if (this.m_intCurrGroupSet == (this.spc_groupings_collection1.Count / 6) - 1)
			{
				this.btnNext.Enabled=false;
			}
			if (this.btnPrev.Enabled==false) this.btnPrev.Enabled=true;
			this.DisplayGroupSet(this.m_intCurrGroupSet + 1);
		}

		private void grpbox6_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void grpbox5_Enter(object sender, System.EventArgs e)
		{
		
		}
       

		private void txtMDBFile_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void txtTable_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void btnTreeAudit_Click(object sender, System.EventArgs e)
		{
			((frmMain)this.ParentForm.ParentForm).button_click("TREE SPECIES");
		}

		private void chkFilterSpecies_CheckStateChanged(object sender, System.EventArgs e)
		{
			LoadListBoxes(this.chkFilterSpecies.Checked);
		}
		private void LoadListBoxes(bool p_bFilter)
		{
			int x;
			int y;
			//
			//clear and load the group list boxes
			//
			for (x=0;x<=this.spc_groupings_collection1.Count-1;x++)
			{
				this.spc_groupings_collection1.Item(x).m_lstGrp.Items.Clear();
                for (y = 0; y <= this.spc_common_name_collection1.Count - 1; y++)
                {
                    if (this.spc_common_name_collection1.Item(y).SpeciesGroupIndex >= 0)
                    {
                        if (this.spc_groupings_collection1.Item(x).GroupingNumber - 1 ==
                            this.spc_common_name_collection1.Item(y).SpeciesGroupIndex)
                        {
                            if (p_bFilter)
                            {
                                if (this.spc_common_name_collection1.Item(y).FVSOutput)
                                {
                                    this.spc_groupings_collection1.Item(x).m_lstGrp.Items.Add(this.spc_common_name_collection1.Item(y).SpeciesCommonName + "*");
                                }
                            }
                            else
                            {
                                if (this.spc_common_name_collection1.Item(y).FVSOutput)
                                {
                                    this.spc_groupings_collection1.Item(x).m_lstGrp.Items.Add(this.spc_common_name_collection1.Item(y).SpeciesCommonName + "*");
                                }
                                else
                                    this.spc_groupings_collection1.Item(x).m_lstGrp.Items.Add(this.spc_common_name_collection1.Item(y).SpeciesCommonName);

                            }
                        }
                    }
                }
            }
            //
            //clear and load the unassigned list box
            //
            this.lstCommonName.Items.Clear();
            for (x = 0; x <= this.spc_common_name_collection1.Count - 1; x++)
            {
                if (this.spc_common_name_collection1.Item(x).SpeciesGroupIndex < 0)
                {
                    this.AddSpeciesCommonNameToUnassignedList(x);
                }
            }

		}

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "PROCESSOR", "TREE_SPECIES_GROUPS" });
        }
	}
	public class spc_groupings : System.Windows.Forms.GroupBox
	{
		public System.Windows.Forms.Button m_btnRemove;
		public System.Windows.Forms.Button m_btnClearAll;
		public System.Windows.Forms.ListBox m_lstGrp;
		public System.Windows.Forms.TextBox m_txtGrp;
        public FIA_Biosum_Manager.uc_processor_scenario_tree_spc_groups m_uc_processor_scenario_tree_spc_groups1;
		private int _intGroupingNumber=1;
        public spc_groupings(FIA_Biosum_Manager.uc_processor_scenario_tree_spc_groups p_processor_scenario_tree_spc_groups)
		{
			this.Visible=false;
            this.m_uc_processor_scenario_tree_spc_groups1 = p_processor_scenario_tree_spc_groups;
			this.m_btnRemove = new Button();
			this.m_btnClearAll = new Button();
			this.m_lstGrp = new ListBox();
			this.m_txtGrp = new TextBox();
			this.m_txtGrp.MaxLength = 50;
			this.m_btnRemove.Text = "Remove";
			this.m_btnClearAll.Text = "Clear All";

			this.Controls.Add(this.m_btnRemove);
			this.Controls.Add(this.m_btnClearAll);
			this.Controls.Add(this.m_lstGrp);
			this.Controls.Add(this.m_txtGrp);


			this.m_btnRemove.Click += new System.EventHandler(this.m_btnRemove_Click);
			this.m_btnClearAll.Click += new System.EventHandler(this.m_btnClearAll_Click);
		


		}

		private void m_btnRemove_Click(object sender, System.EventArgs e)
		{

			if (this.m_lstGrp.SelectedItems.Count > 0)
			{
                this.m_uc_processor_scenario_tree_spc_groups1.RemoveSpeciesCommonNameFromGroupAssignment(this.m_lstGrp.SelectedItems[0].ToString());

				this.m_lstGrp.Items.Remove(this.m_lstGrp.SelectedItems[0].ToString());
                if (this.m_uc_processor_scenario_tree_spc_groups1.btnSave.Enabled == false) this.m_uc_processor_scenario_tree_spc_groups1.btnSave.Enabled = true;
			}
			//MessageBox.Show("btnRemove " + this._intGroupingNumber.ToString());

		}
		private void m_btnClearAll_Click(object sender,System.EventArgs e)
		{
			if (this.m_lstGrp.Items.Count > 0)
			{
				for (int x=0;x<=this.m_lstGrp.Items.Count-1;x++)
				{
                    this.m_uc_processor_scenario_tree_spc_groups1.RemoveSpeciesCommonNameFromGroupAssignment(this.m_lstGrp.Items[x].ToString());

				}
				this.m_lstGrp.Items.Clear();
                if (this.m_uc_processor_scenario_tree_spc_groups1.btnSave.Enabled == false) this.m_uc_processor_scenario_tree_spc_groups1.btnSave.Enabled = true;
			}
			
		}
		public spc_groupings getSpcGroupObject
		{
			get
			{
				return (this);
			}
		}
		public int GroupingNumber
		{
			get
			{
				return _intGroupingNumber;
			}
			set
			{
				_intGroupingNumber = value;
			}
		}
		public string GroupLabel
		{
			get {return this.m_txtGrp.Text.Trim();}
			set {this.m_txtGrp.Text = value;}
		}







		


	}
	public class spc_groupings_collection : System.Collections.CollectionBase
	{
		public spc_groupings_collection()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public void Add(spc_groupings spc_groupings1)
		{
			// vérify if object is not already in
			if (this.List.Contains(spc_groupings1))
				throw new InvalidOperationException();
 
			// adding it
			this.List.Add(spc_groupings1);
 
			// return collection
			//return this;
		}
		public void Remove(int index)
		{
			// Check to see if there is a widget at the supplied index.
			if (index > Count - 1 || index < 0)
				// If no widget exists, a messagebox is shown and the operation 
				// is cancelled.
			{
				System.Windows.Forms.MessageBox.Show("Index not valid!");
			}
			else
			{
				List.RemoveAt(index); 
			}
		}
		public spc_groupings Item(int Index)
		{
			// The appropriate item is retrieved from the List object and
			// explicitly cast to the Widget type, then returned to the 
			// caller.
			return (spc_groupings) List[Index];
		}




	}

    public class spc_common_name
    {
        private string _strCommonName = "";
        private bool _bFVSOutput = false;
        private string _strSpeciesGroupLabel = "";
        private int _intSpeciesGroupIndex = -1;
        private int _intSpeciesCode = -1;

        public spc_common_name()
        {
        }
        public string SpeciesCommonName
        {
            get { return _strCommonName; }
            set { _strCommonName = value; }
        }
        public bool FVSOutput
        {
            get { return _bFVSOutput; }
            set { _bFVSOutput = value; }
        }
        public string SpeciesGroupLabel
        {
            get { return _strSpeciesGroupLabel; }
            set { _strSpeciesGroupLabel = value; }
        }
        public int SpeciesGroupIndex
        {
            get { return _intSpeciesGroupIndex; }
            set { _intSpeciesGroupIndex = value; }
        }
        public int SpeciesCode
        {
            get { return _intSpeciesCode; }
            set { _intSpeciesCode = value; }
        }
    }
    public class spc_common_name_collection : System.Collections.CollectionBase
    {
        public spc_common_name_collection()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public void Add(spc_common_name spc_common_name1)
        {
            // vérify if object is not already in
            if (this.List.Contains(spc_common_name1))
                throw new InvalidOperationException();

            // adding it
            this.List.Add(spc_common_name1);

            // return collection
            //return this;
        }
        public void Remove(int index)
        {
            // Check to see if there is a widget at the supplied index.
            if (index > Count - 1 || index < 0)
            // If no widget exists, a messagebox is shown and the operation 
            // is cancelled.
            {
                System.Windows.Forms.MessageBox.Show("Index not valid!");
            }
            else
            {
                List.RemoveAt(index);
            }
        }
        public spc_common_name Item(int Index)
        {
            // The appropriate item is retrieved from the List object and
            // explicitly cast to the Widget type, then returned to the 
            // caller.
            return (spc_common_name)List[Index];
        }
    }

}

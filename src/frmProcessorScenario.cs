using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmProcessorScenario.
	/// </summary>
	public class frmProcessorScenario : System.Windows.Forms.Form
	{
		public System.Windows.Forms.ToolBar tlbScenario;
		private System.Windows.Forms.ToolBarButton btnScenarioNew;
		private System.Windows.Forms.ToolBarButton btnScenarioOpen;
		private System.Windows.Forms.ToolBarButton toolBarButton1;
		private System.Windows.Forms.ToolBarButton btnScenarioSave;
		private System.Windows.Forms.ToolBarButton btnScenarioDelete;
        private System.Windows.Forms.ToolBarButton btnScenarioProperties;
		private System.Windows.Forms.ImageList imageList1;
		private System.ComponentModel.IContainer components;
		public FIA_Biosum_Manager.uc_scenario_open uc_scenario_open1;
		private System.Windows.Forms.Button btnClose;
		public FIA_Biosum_Manager.uc_scenario uc_scenario1;
		public FIA_Biosum_Manager.uc_datasource uc_datasource1;
		public FIA_Biosum_Manager.uc_scenario_notes uc_scenario_notes1;
        public FIA_Biosum_Manager.uc_processor_scenario_tree_groupings uc_scenario_tree_groupings1;
		public bool m_bSave=false;
		public bool m_bScenarioOpen = false;
		public bool m_bDataSourceFirstTime;
		public System.Windows.Forms.TabControl tabControlScenario;
		private System.Windows.Forms.TabPage tbDesc;
		private System.Windows.Forms.TabPage tbNotes;
		private System.Windows.Forms.TabPage tbDataSources;
		private System.Windows.Forms.TabPage tbRules;
        public System.Windows.Forms.TabControl tabControlRules;
        private System.Windows.Forms.TabPage tbAddHarvestCosts;
		private System.Windows.Forms.TabPage tbRun;
		public bool m_bRulesFirstTime=true;
        public bool m_bTreeGroupsFirstTime = true;
        public bool m_bTreeGroupsCopied = false;
		private System.Windows.Forms.TabPage tbHarvestMethod;
		private FIA_Biosum_Manager.uc_processor_scenario_harvest_method uc_processor_scenario_harvest_method1;
        private FIA_Biosum_Manager.uc_processor_scenario_movein_costs uc_processor_scenario_movein_costs1;
		private System.Windows.Forms.TabPage tbWoodValue;
		private FIA_Biosum_Manager.uc_processor_scenario_merch_chip_value uc_processor_scenario_merch_chip_value1;
		private System.Windows.Forms.TabPage tbEscalators;
		public FIA_Biosum_Manager.uc_processor_scenario_escalators uc_processor_scenario_escalators1;
        private uc_processor_scenario_additional_harvest_cost_columns uc_processor_scenario_additional_harvest_cost_columns1;
        private uc_processor_scenario_run uc_processor_scenario_run1=null;
		public bool m_bPopup = false;

        public FIA_Biosum_Manager.ProcessorScenarioItem m_oProcessorScenarioItem = new ProcessorScenarioItem();
        public FIA_Biosum_Manager.ProcessorScenarioTools m_oProcessorScenarioTools = new ProcessorScenarioTools();

        public int m_intError = 0;
        private ToolBarButton btnScenarioCopy;
        private TabPage tbMoveInCosts;
        private TabPage tbTreeGroupings;
        private Button btnHelp;

        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultProcessorXPSFile;
        private string m_helpChapter = "PROCESSOR_DESCRIPTION";
        private Queries m_oQueries = new Queries();
        public bool m_bUsingSqlite;

        public string m_strError = "";
		
		public frmProcessorScenario(frmMain p_frmMain)
		{
			FIA_Biosum_Manager.frmMain.g_oFrmMain=p_frmMain;

			InitializeComponent();

            if (this.Width > p_frmMain.Width - p_frmMain.grpboxLeft.Width ) 
			{
				this.Width = p_frmMain.Width - p_frmMain.grpboxLeft.Width  - 10;
			}

            this.m_oEnv = new env();
            //
            // TODO: Add any constructor code after InitializeComponent call
            //

            if (System.IO.File.Exists(
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()
                + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile))
            {
                m_bUsingSqlite = true;
            }

            //
            //scenario description and properties
            //
            this.uc_scenario1 = new uc_scenario();
				
			this.tbDesc.Controls.Add(uc_scenario1);
			uc_scenario1.btnCancel.Hide();
			uc_scenario1.btnClose.Hide();
			uc_scenario1.btnOpen.Hide();
			uc_scenario1.txtScenarioId.Enabled=false;
			uc_scenario1.txtScenarioPath.Enabled=false;
			uc_scenario1.txtDescription.Enabled=true;
			uc_scenario1.lblTitle.Text = "Scenario";
			uc_scenario1.Dock = System.Windows.Forms.DockStyle.Fill;
			uc_scenario1.ReferenceProcessorScenarioForm=(frmProcessorScenario) this;
			uc_scenario1.ScenarioType="processor";
			//
			//scenario datasource
			//
			this.uc_datasource1 = new uc_datasource();
			this.tbDataSources.Controls.Add(uc_datasource1);
			uc_datasource1.Dock = System.Windows.Forms.DockStyle.Fill;
			uc_datasource1.ReferenceProcessorScenarioForm=this;
			uc_datasource1.ScenarioType="processor";
			uc_datasource1.SetEnableToolbarButton("CLOSE",false);
			//
			//scenario notes
			//
			this.uc_scenario_notes1 = new uc_scenario_notes();
			this.tbNotes.Controls.Add(uc_scenario_notes1);
			uc_scenario_notes1.Dock = System.Windows.Forms.DockStyle.Fill;
			uc_scenario_notes1.ReferenceProcessorScenarioForm=this;
			uc_scenario_notes1.ScenarioType="processor";
			//
			//rule definitions wood market values
			//
			this.uc_processor_scenario_merch_chip_value1.m_oResizeForm.ControlToResize=this;
			this.uc_processor_scenario_merch_chip_value1.m_oResizeForm.ResizeControl();
			this.uc_processor_scenario_merch_chip_value1.ReferenceProcessorScenarioForm=this;
			this.uc_processor_scenario_merch_chip_value1.ScenarioId=uc_scenario1.txtScenarioId.Text.Trim();
			//
			//rule definitions escalators
			//
			this.uc_processor_scenario_escalators1.ReferenceProcessorScenarioForm=this;
			this.uc_processor_scenario_escalators1.ScenarioId=uc_scenario1.txtScenarioId.Text.Trim();
            //
            //additional harvest cost columns
            //
            this.uc_processor_scenario_additional_harvest_cost_columns1.ReferenceProcessorScenarioForm = this;
            this.uc_processor_scenario_additional_harvest_cost_columns1.ScenarioId = uc_scenario1.txtScenarioId.Text.Trim();
            //
            //scenario run
            //
            this.uc_processor_scenario_run1.ReferenceProcessorScenarioForm = this;
            this.uc_processor_scenario_run1.ScenarioId = uc_scenario1.txtScenarioId.Text.Trim();
            //
            //tree diameter groupings
            //
            uc_scenario_tree_groupings1.Dock = System.Windows.Forms.DockStyle.Fill;
            uc_scenario_tree_groupings1.ReferenceProcessorScenarioForm = this;
            uc_scenario_tree_groupings1.ScenarioType = "processor";
            //uc_scenario_tree_groupings1. initTreeGroupingDialogs();



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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmProcessorScenario));
            this.tlbScenario = new System.Windows.Forms.ToolBar();
            this.btnScenarioNew = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioOpen = new System.Windows.Forms.ToolBarButton();
            this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioSave = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioDelete = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioProperties = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioCopy = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnClose = new System.Windows.Forms.Button();
            this.tabControlScenario = new System.Windows.Forms.TabControl();
            this.tbDesc = new System.Windows.Forms.TabPage();
            this.tbNotes = new System.Windows.Forms.TabPage();
            this.tbDataSources = new System.Windows.Forms.TabPage();
            this.tbTreeGroupings = new System.Windows.Forms.TabPage();
            this.uc_scenario_tree_groupings1 = new FIA_Biosum_Manager.uc_processor_scenario_tree_groupings();
            this.tbRules = new System.Windows.Forms.TabPage();
            this.tabControlRules = new System.Windows.Forms.TabControl();
            this.tbHarvestMethod = new System.Windows.Forms.TabPage();
            this.uc_processor_scenario_harvest_method1 = new FIA_Biosum_Manager.uc_processor_scenario_harvest_method();
            this.tbMoveInCosts = new System.Windows.Forms.TabPage();
            this.uc_processor_scenario_movein_costs1 = new FIA_Biosum_Manager.uc_processor_scenario_movein_costs();
            this.tbWoodValue = new System.Windows.Forms.TabPage();
            this.uc_processor_scenario_merch_chip_value1 = new FIA_Biosum_Manager.uc_processor_scenario_merch_chip_value();
            this.tbEscalators = new System.Windows.Forms.TabPage();
            this.uc_processor_scenario_escalators1 = new FIA_Biosum_Manager.uc_processor_scenario_escalators();
            this.tbAddHarvestCosts = new System.Windows.Forms.TabPage();
            this.uc_processor_scenario_additional_harvest_cost_columns1 = new FIA_Biosum_Manager.uc_processor_scenario_additional_harvest_cost_columns();
            this.tbRun = new System.Windows.Forms.TabPage();
            this.uc_processor_scenario_run1 = new FIA_Biosum_Manager.uc_processor_scenario_run();
            this.btnHelp = new System.Windows.Forms.Button();
            this.tabControlScenario.SuspendLayout();
            this.tbTreeGroupings.SuspendLayout();
            this.tbRules.SuspendLayout();
            this.tabControlRules.SuspendLayout();
            this.tbHarvestMethod.SuspendLayout();
            this.tbMoveInCosts.SuspendLayout();
            this.tbWoodValue.SuspendLayout();
            this.tbEscalators.SuspendLayout();
            this.tbAddHarvestCosts.SuspendLayout();
            this.tbRun.SuspendLayout();
            this.SuspendLayout();
            // 
            // tlbScenario
            // 
            this.tlbScenario.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
            this.btnScenarioNew,
            this.btnScenarioOpen,
            this.toolBarButton1,
            this.btnScenarioSave,
            this.btnScenarioDelete,
            this.btnScenarioProperties,
            this.btnScenarioCopy});
            this.tlbScenario.ButtonSize = new System.Drawing.Size(60, 45);
            this.tlbScenario.DropDownArrows = true;
            this.tlbScenario.ImageList = this.imageList1;
            this.tlbScenario.Location = new System.Drawing.Point(0, 0);
            this.tlbScenario.Margin = new System.Windows.Forms.Padding(2);
            this.tlbScenario.Name = "tlbScenario";
            this.tlbScenario.ShowToolTips = true;
            this.tlbScenario.Size = new System.Drawing.Size(1052, 47);
            this.tlbScenario.TabIndex = 43;
            this.tlbScenario.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbScenario_ButtonClick);
            // 
            // btnScenarioNew
            // 
            this.btnScenarioNew.ImageIndex = 0;
            this.btnScenarioNew.Name = "btnScenarioNew";
            this.btnScenarioNew.Text = "New";
            this.btnScenarioNew.ToolTipText = "New Processor Scenario";
            // 
            // btnScenarioOpen
            // 
            this.btnScenarioOpen.ImageIndex = 1;
            this.btnScenarioOpen.Name = "btnScenarioOpen";
            this.btnScenarioOpen.Text = "Open";
            this.btnScenarioOpen.ToolTipText = "Open Processor Scenario";
            // 
            // toolBarButton1
            // 
            this.toolBarButton1.Name = "toolBarButton1";
            this.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnScenarioSave
            // 
            this.btnScenarioSave.ImageIndex = 2;
            this.btnScenarioSave.Name = "btnScenarioSave";
            this.btnScenarioSave.Text = "Save";
            this.btnScenarioSave.ToolTipText = "Save Processor Scenario";
            // 
            // btnScenarioDelete
            // 
            this.btnScenarioDelete.ImageIndex = 3;
            this.btnScenarioDelete.Name = "btnScenarioDelete";
            this.btnScenarioDelete.Text = "Delete";
            this.btnScenarioDelete.ToolTipText = "Delete Processor Scenario";
            // 
            // btnScenarioProperties
            // 
            this.btnScenarioProperties.ImageIndex = 4;
            this.btnScenarioProperties.Name = "btnScenarioProperties";
            this.btnScenarioProperties.Text = "Properties";
            // 
            // btnScenarioCopy
            // 
            this.btnScenarioCopy.ImageIndex = 5;
            this.btnScenarioCopy.Name = "btnScenarioCopy";
            this.btnScenarioCopy.Text = "Copy";
            this.btnScenarioCopy.ToolTipText = "Copy scenario values to this scenario";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "");
            this.imageList1.Images.SetKeyName(4, "properties.png");
            this.imageList1.Images.SetKeyName(5, "copy.bmp");
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(917, 662);
            this.btnClose.Margin = new System.Windows.Forms.Padding(2);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(127, 39);
            this.btnClose.TabIndex = 44;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // tabControlScenario
            // 
            this.tabControlScenario.Controls.Add(this.tbDesc);
            this.tabControlScenario.Controls.Add(this.tbNotes);
            this.tabControlScenario.Controls.Add(this.tbDataSources);
            this.tabControlScenario.Controls.Add(this.tbTreeGroupings);
            this.tabControlScenario.Controls.Add(this.tbRules);
            this.tabControlScenario.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlScenario.ItemSize = new System.Drawing.Size(100, 18);
            this.tabControlScenario.Location = new System.Drawing.Point(0, 68);
            this.tabControlScenario.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlScenario.Name = "tabControlScenario";
            this.tabControlScenario.SelectedIndex = 0;
            this.tabControlScenario.Size = new System.Drawing.Size(1209, 578);
            this.tabControlScenario.TabIndex = 45;
            this.tabControlScenario.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlScenario_DrawItem);
            this.tabControlScenario.SelectedIndexChanged += new System.EventHandler(this.tabControlScenario_SelectedIndexChanged);
            // 
            // tbDesc
            // 
            this.tbDesc.AutoScroll = true;
            this.tbDesc.Location = new System.Drawing.Point(4, 22);
            this.tbDesc.Margin = new System.Windows.Forms.Padding(2);
            this.tbDesc.Name = "tbDesc";
            this.tbDesc.Size = new System.Drawing.Size(1201, 552);
            this.tbDesc.TabIndex = 0;
            this.tbDesc.Text = "Description";
            // 
            // tbNotes
            // 
            this.tbNotes.AutoScroll = true;
            this.tbNotes.Location = new System.Drawing.Point(4, 22);
            this.tbNotes.Margin = new System.Windows.Forms.Padding(2);
            this.tbNotes.Name = "tbNotes";
            this.tbNotes.Size = new System.Drawing.Size(1201, 552);
            this.tbNotes.TabIndex = 1;
            this.tbNotes.Text = "Notes";
            this.tbNotes.Visible = false;
            // 
            // tbDataSources
            // 
            this.tbDataSources.AutoScroll = true;
            this.tbDataSources.Location = new System.Drawing.Point(4, 22);
            this.tbDataSources.Margin = new System.Windows.Forms.Padding(2);
            this.tbDataSources.Name = "tbDataSources";
            this.tbDataSources.Size = new System.Drawing.Size(1201, 552);
            this.tbDataSources.TabIndex = 2;
            this.tbDataSources.Text = "Data Sources";
            this.tbDataSources.Visible = false;
            // 
            // tbTreeGroupings
            // 
            this.tbTreeGroupings.Controls.Add(this.uc_scenario_tree_groupings1);
            this.tbTreeGroupings.Location = new System.Drawing.Point(4, 22);
            this.tbTreeGroupings.Margin = new System.Windows.Forms.Padding(2);
            this.tbTreeGroupings.Name = "tbTreeGroupings";
            this.tbTreeGroupings.Size = new System.Drawing.Size(1201, 552);
            this.tbTreeGroupings.TabIndex = 4;
            this.tbTreeGroupings.Text = "Tree Groupings";
            this.tbTreeGroupings.UseVisualStyleBackColor = true;
            // 
            // uc_scenario_tree_groupings1
            // 
            this.uc_scenario_tree_groupings1.Location = new System.Drawing.Point(0, 0);
            this.uc_scenario_tree_groupings1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.uc_scenario_tree_groupings1.Name = "uc_scenario_tree_groupings1";
            this.uc_scenario_tree_groupings1.ReferenceProcessorScenarioForm = null;
            this.uc_scenario_tree_groupings1.ScenarioType = "processor";
            this.uc_scenario_tree_groupings1.Size = new System.Drawing.Size(885, 517);
            this.uc_scenario_tree_groupings1.TabIndex = 0;
            // 
            // tbRules
            // 
            this.tbRules.Controls.Add(this.tabControlRules);
            this.tbRules.ForeColor = System.Drawing.Color.Red;
            this.tbRules.Location = new System.Drawing.Point(4, 22);
            this.tbRules.Margin = new System.Windows.Forms.Padding(2);
            this.tbRules.Name = "tbRules";
            this.tbRules.Size = new System.Drawing.Size(1201, 552);
            this.tbRules.TabIndex = 3;
            this.tbRules.Text = "Rule Definitions";
            this.tbRules.Visible = false;
            // 
            // tabControlRules
            // 
            this.tabControlRules.Controls.Add(this.tbHarvestMethod);
            this.tabControlRules.Controls.Add(this.tbMoveInCosts);
            this.tabControlRules.Controls.Add(this.tbWoodValue);
            this.tabControlRules.Controls.Add(this.tbEscalators);
            this.tabControlRules.Controls.Add(this.tbAddHarvestCosts);
            this.tabControlRules.Controls.Add(this.tbRun);
            this.tabControlRules.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlRules.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlRules.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlRules.Location = new System.Drawing.Point(0, 0);
            this.tabControlRules.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlRules.Name = "tabControlRules";
            this.tabControlRules.SelectedIndex = 0;
            this.tabControlRules.Size = new System.Drawing.Size(1201, 552);
            this.tabControlRules.TabIndex = 0;
            this.tabControlRules.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlRules_DrawItem);
            this.tabControlRules.SelectedIndexChanged += new System.EventHandler(this.tabControlRules_SelectedIndexChanged);
            // 
            // tbHarvestMethod
            // 
            this.tbHarvestMethod.Controls.Add(this.uc_processor_scenario_harvest_method1);
            this.tbHarvestMethod.Location = new System.Drawing.Point(4, 26);
            this.tbHarvestMethod.Margin = new System.Windows.Forms.Padding(2);
            this.tbHarvestMethod.Name = "tbHarvestMethod";
            this.tbHarvestMethod.Size = new System.Drawing.Size(1193, 522);
            this.tbHarvestMethod.TabIndex = 1;
            this.tbHarvestMethod.Text = "Harvest Method";
            this.tbHarvestMethod.UseVisualStyleBackColor = true;
            // 
            // uc_processor_scenario_harvest_method1
            // 
            this.uc_processor_scenario_harvest_method1.BackColor = System.Drawing.SystemColors.Control;
            this.uc_processor_scenario_harvest_method1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_processor_scenario_harvest_method1.Location = new System.Drawing.Point(0, 0);
            this.uc_processor_scenario_harvest_method1.Margin = new System.Windows.Forms.Padding(2);
            this.uc_processor_scenario_harvest_method1.Name = "uc_processor_scenario_harvest_method1";
            this.uc_processor_scenario_harvest_method1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_harvest_method1.ScenarioId = "";
            this.uc_processor_scenario_harvest_method1.Size = new System.Drawing.Size(1193, 522);
            this.uc_processor_scenario_harvest_method1.TabIndex = 0;
            // 
            // tbMoveInCosts
            // 
            this.tbMoveInCosts.Controls.Add(this.uc_processor_scenario_movein_costs1);
            this.tbMoveInCosts.Location = new System.Drawing.Point(4, 26);
            this.tbMoveInCosts.Margin = new System.Windows.Forms.Padding(2);
            this.tbMoveInCosts.Name = "tbMoveInCosts";
            this.tbMoveInCosts.Size = new System.Drawing.Size(192, 70);
            this.tbMoveInCosts.TabIndex = 2;
            this.tbMoveInCosts.Text = "Move-in Costs";
            this.tbMoveInCosts.UseVisualStyleBackColor = true;
            // 
            // uc_processor_scenario_movein_costs1
            // 
            this.uc_processor_scenario_movein_costs1.Location = new System.Drawing.Point(0, 0);
            this.uc_processor_scenario_movein_costs1.Margin = new System.Windows.Forms.Padding(2);
            this.uc_processor_scenario_movein_costs1.Name = "uc_processor_scenario_movein_costs1";
            this.uc_processor_scenario_movein_costs1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_movein_costs1.ScenarioId = "";
            this.uc_processor_scenario_movein_costs1.Size = new System.Drawing.Size(928, 546);
            this.uc_processor_scenario_movein_costs1.TabIndex = 0;
            // 
            // tbWoodValue
            // 
            this.tbWoodValue.Controls.Add(this.uc_processor_scenario_merch_chip_value1);
            this.tbWoodValue.Location = new System.Drawing.Point(4, 26);
            this.tbWoodValue.Margin = new System.Windows.Forms.Padding(2);
            this.tbWoodValue.Name = "tbWoodValue";
            this.tbWoodValue.Size = new System.Drawing.Size(192, 70);
            this.tbWoodValue.TabIndex = 7;
            this.tbWoodValue.Text = "Wood Value";
            this.tbWoodValue.UseVisualStyleBackColor = true;
            this.tbWoodValue.Visible = false;
            // 
            // uc_processor_scenario_merch_chip_value1
            // 
            this.uc_processor_scenario_merch_chip_value1.BackColor = System.Drawing.SystemColors.Control;
            this.uc_processor_scenario_merch_chip_value1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_processor_scenario_merch_chip_value1.Location = new System.Drawing.Point(0, 0);
            this.uc_processor_scenario_merch_chip_value1.Margin = new System.Windows.Forms.Padding(2);
            this.uc_processor_scenario_merch_chip_value1.Name = "uc_processor_scenario_merch_chip_value1";
            this.uc_processor_scenario_merch_chip_value1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_merch_chip_value1.ScenarioId = "";
            this.uc_processor_scenario_merch_chip_value1.Size = new System.Drawing.Size(192, 70);
            this.uc_processor_scenario_merch_chip_value1.TabIndex = 0;
            // 
            // tbEscalators
            // 
            this.tbEscalators.Controls.Add(this.uc_processor_scenario_escalators1);
            this.tbEscalators.Location = new System.Drawing.Point(4, 26);
            this.tbEscalators.Margin = new System.Windows.Forms.Padding(2);
            this.tbEscalators.Name = "tbEscalators";
            this.tbEscalators.Size = new System.Drawing.Size(192, 70);
            this.tbEscalators.TabIndex = 5;
            this.tbEscalators.Text = "Escalators";
            this.tbEscalators.UseVisualStyleBackColor = true;
            this.tbEscalators.Visible = false;
            // 
            // uc_processor_scenario_escalators1
            // 
            this.uc_processor_scenario_escalators1.BackColor = System.Drawing.SystemColors.Control;
            this.uc_processor_scenario_escalators1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_processor_scenario_escalators1.Location = new System.Drawing.Point(0, 0);
            this.uc_processor_scenario_escalators1.Margin = new System.Windows.Forms.Padding(2);
            this.uc_processor_scenario_escalators1.Name = "uc_processor_scenario_escalators1";
            this.uc_processor_scenario_escalators1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_escalators1.ScenarioId = "";
            this.uc_processor_scenario_escalators1.Size = new System.Drawing.Size(192, 70);
            this.uc_processor_scenario_escalators1.TabIndex = 0;
            // 
            // tbAddHarvestCosts
            // 
            this.tbAddHarvestCosts.Controls.Add(this.uc_processor_scenario_additional_harvest_cost_columns1);
            this.tbAddHarvestCosts.Location = new System.Drawing.Point(4, 26);
            this.tbAddHarvestCosts.Margin = new System.Windows.Forms.Padding(2);
            this.tbAddHarvestCosts.Name = "tbAddHarvestCosts";
            this.tbAddHarvestCosts.Size = new System.Drawing.Size(192, 70);
            this.tbAddHarvestCosts.TabIndex = 8;
            this.tbAddHarvestCosts.Text = "Supplemental Harvest Costs";
            this.tbAddHarvestCosts.UseVisualStyleBackColor = true;
            this.tbAddHarvestCosts.Visible = false;
            // 
            // uc_processor_scenario_additional_harvest_cost_columns1
            // 
            this.uc_processor_scenario_additional_harvest_cost_columns1.BackColor = System.Drawing.SystemColors.Control;
            this.uc_processor_scenario_additional_harvest_cost_columns1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_processor_scenario_additional_harvest_cost_columns1.Location = new System.Drawing.Point(0, 0);
            this.uc_processor_scenario_additional_harvest_cost_columns1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.uc_processor_scenario_additional_harvest_cost_columns1.Name = "uc_processor_scenario_additional_harvest_cost_columns1";
            this.uc_processor_scenario_additional_harvest_cost_columns1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_additional_harvest_cost_columns1.ScenarioId = "";
            this.uc_processor_scenario_additional_harvest_cost_columns1.Size = new System.Drawing.Size(192, 70);
            this.uc_processor_scenario_additional_harvest_cost_columns1.TabIndex = 0;
            // 
            // tbRun
            // 
            this.tbRun.BackColor = System.Drawing.SystemColors.Control;
            this.tbRun.Controls.Add(this.uc_processor_scenario_run1);
            this.tbRun.ForeColor = System.Drawing.Color.Black;
            this.tbRun.Location = new System.Drawing.Point(4, 26);
            this.tbRun.Margin = new System.Windows.Forms.Padding(2);
            this.tbRun.Name = "tbRun";
            this.tbRun.Size = new System.Drawing.Size(192, 70);
            this.tbRun.TabIndex = 6;
            this.tbRun.Text = "Run";
            this.tbRun.Visible = false;
            // 
            // uc_processor_scenario_run1
            // 
            this.uc_processor_scenario_run1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.uc_processor_scenario_run1.Location = new System.Drawing.Point(0, 0);
            this.uc_processor_scenario_run1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.uc_processor_scenario_run1.Name = "uc_processor_scenario_run1";
            this.uc_processor_scenario_run1.ReferenceProcessorScenarioForm = null;
            this.uc_processor_scenario_run1.ScenarioId = "";
            this.uc_processor_scenario_run1.Size = new System.Drawing.Size(192, 70);
            this.uc_processor_scenario_run1.TabIndex = 0;
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(5, 662);
            this.btnHelp.Margin = new System.Windows.Forms.Padding(2);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(128, 39);
            this.btnHelp.TabIndex = 47;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // frmProcessorScenario
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1052, 703);
            this.Controls.Add(this.tabControlScenario);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.tlbScenario);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmProcessorScenario";
            this.Text = "Processor Scenario";
            this.Activated += new System.EventHandler(this.frmProcessorScenario_Activated);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmProcessorScenario_FormClosing);
            this.Resize += new System.EventHandler(this.frmProcessorScenario_Resize);
            this.tabControlScenario.ResumeLayout(false);
            this.tbTreeGroupings.ResumeLayout(false);
            this.tbRules.ResumeLayout(false);
            this.tabControlRules.ResumeLayout(false);
            this.tbHarvestMethod.ResumeLayout(false);
            this.tbMoveInCosts.ResumeLayout(false);
            this.tbWoodValue.ResumeLayout(false);
            this.tbEscalators.ResumeLayout(false);
            this.tbAddHarvestCosts.ResumeLayout(false);
            this.tbRun.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		
		private void tlbScenario_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.Text.Trim().ToUpper())
			{
					//
					//open scenario
					//
				case "OPEN":
					frmMain.g_oFrmMain.OpenProcessorScenario("Open",this);
					break;
					//
					//new scenario
					//
				case "NEW":
					frmMain.g_oFrmMain.OpenProcessorScenario("New",this);
					break;
				case "SAVE":
					this.SaveRuleDefinitions();
					break;
				case "DELETE":
                    bool bDeleted = false;
                    if (this.uc_scenario_open1 != null)
					{
                        if (!m_bUsingSqlite)
                        {
                            bDeleted = uc_scenario1.DeleteScenario(uc_scenario_open1.txtScenarioId.Text.Trim());
                        }
                        else
                        {
                            bDeleted = uc_scenario1.DeleteScenarioSqlite(uc_scenario_open1.txtScenarioId.Text.Trim());
                        }
                        if (bDeleted)
						    uc_scenario_open1.lstScenario.Items.Remove(uc_scenario_open1.lstScenario.SelectedItems[0]);
					}
					else
					{
                        if (!m_bUsingSqlite)
                        {
                            bDeleted = uc_scenario1.DeleteScenario(uc_scenario1.txtScenarioId.Text.Trim());
                        }
                        else
                        {
                            bDeleted = uc_scenario1.DeleteScenarioSqlite(uc_scenario1.txtScenarioId.Text.Trim());
                        }
                        if (bDeleted)
                            this.Close();
					}
					break;
                case "PROPERTIES":
                    frmMain.g_oFrmMain.ActivateStandByAnimation(
                        this.WindowState, this.Left, this.Height, this.Width, this.Top);
                    if (m_bRulesFirstTime == true)
                    {
                        if (m_bTreeGroupsFirstTime == true &&
                            m_bTreeGroupsCopied == false)
                        {
                            // Load tree groupings so they are there for the wood values tab
                            LoadTreeGroupings();
                        }
                        LoadRuleDefinitions();
                    }
                    else
                        SaveRuleDefinitions();
                    FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
			        frmTemp.Text = "FIA Biosum";
                    frmTemp.AutoScroll = false;
                    uc_textbox uc_textbox1 = new uc_textbox();
                    frmTemp.Controls.Add(uc_textbox1);
                    uc_textbox1.Dock = DockStyle.Fill;
                    uc_textbox1.lblTitle.Text = "Properties";
                    uc_textbox1.TextValue=m_oProcessorScenarioTools.ScenarioProperties(m_oProcessorScenarioItem);
                    frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    frmTemp.Show();
                    break;
                case "COPY":
                    CopyScenario();
                    break;
			}
		}
		public void InitializeNewScenario()
		{
			this.uc_scenario1 = new uc_scenario();

			this.Controls.Add(uc_scenario1);

			uc_scenario1.Dock = System.Windows.Forms.DockStyle.Fill;

			this.tlbScenario.Hide();

			this.tabControlScenario.Hide();

			this.btnClose.Hide();

			this.uc_scenario1.ReferenceProcessorScenarioForm=this;

            this.uc_scenario1.ScenarioType = "processor";

			this.uc_scenario1.NewScenario();

			


		}
		public void InitializeOpenScenario()
		{
			this.uc_scenario_open1 = new uc_scenario_open(); 

			this.Controls.Add(uc_scenario_open1);


			uc_scenario_open1.Left = 0;
			uc_scenario_open1.Width = this.ClientSize.Width;
			uc_scenario_open1.Top = this.tlbScenario.Top + this.tlbScenario.Height;
    		
			this.btnScenarioSave.Enabled=false;
			this.btnScenarioOpen.Enabled=false;
            this.btnScenarioProperties.Enabled = false;
            this.btnScenarioProperties.Visible = false;
            this.btnScenarioCopy.Enabled = false;
            this.btnScenarioCopy.Visible = false;

			this.tabControlScenario.Hide();

			this.btnClose.Hide();

			this.uc_scenario_open1.ReferenceProcessorScenarioForm=this;
			this.uc_scenario_open1.ScenarioType="processor";
            this.uc_scenario_open1.OpenScenario();

			this.Height = 200;
			int intHt = this.Height;
			int intHt2=this.uc_scenario_open1.btnOpen.Height;
			int intTop=this.uc_scenario_open1.btnOpen.Top;
			while (intTop + intHt2 + 20
				>=  intHt)
			{
				intHt += 10;

			}
			this.Height = intHt;

			


		}
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			CheckToSave();
			this.Close();
		}
		private void CheckToSave()
		{
			if (m_bSave)
			{
				DialogResult result = MessageBox.Show("Save Changes Y/N","Scenario",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
				if (result == System.Windows.Forms.DialogResult.Yes)
				{
					this.SaveRuleDefinitions();
				}				
			}
		}

		private void tabControlScenario_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.tabControlRules.Enabled)
			{
				if (tabControlScenario.SelectedTab.Text.Trim().ToUpper()=="RULE DEFINITIONS")
				{
                    btnHelp.Show();
                    SetHelpChapterForRulesTab();
                    // Loading the rule definitions for the first time
                    if (m_bRulesFirstTime == true)
                    {
                        frmMain.g_oFrmMain.ActivateStandByAnimation(
                            frmMain.g_oFrmMain.WindowState,
                            frmMain.g_oFrmMain.Left,
                            frmMain.g_oFrmMain.Height,
                            frmMain.g_oFrmMain.Width,
                            frmMain.g_oFrmMain.Top);
                        if (m_bTreeGroupsFirstTime == true && 
                            m_bTreeGroupsCopied == false)
                        {
                            // Load tree groupings so they are there for the wood values tab
                            LoadTreeGroupings();
                        }
                        LoadRuleDefinitions();
                        frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    }
                    // The tree groups have changed so we need to reload the wood values tab
                    else if (m_bTreeGroupsFirstTime == true &&
                             m_bTreeGroupsCopied == false)
                    {
                        frmMain.g_oFrmMain.ActivateStandByAnimation(
                            frmMain.g_oFrmMain.WindowState,
                            frmMain.g_oFrmMain.Left,
                            frmMain.g_oFrmMain.Height,
                            frmMain.g_oFrmMain.Width,
                            frmMain.g_oFrmMain.Top);

                        // Load tree groupings so they are there for the wood values tab
                        LoadTreeGroupings();
                        frmMain.g_sbpInfo.Text = "Loading Scenario Merch and Chip Market Value Rule Definitions...Stand By";
                        this.uc_processor_scenario_merch_chip_value1.loadvalues();
                        frmMain.g_oFrmMain.DeactivateStandByAnimation();
                    }
				}
                if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "TREE GROUPINGS")
                {
                    btnHelp.Hide();
                    // We load the tree groupings (maybe) when user clicks a button on uc_scenario_tree_groupings.cs
                    // This way we reload from the database if user saved a change and went right back in
                    // without returning to the main Processor scenario page.
                }
				if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "NOTES")
				{
                    btnHelp.Hide();
                    if (((Control)this.tbNotes).Enabled)
					{
						this.uc_scenario_notes1.lblTitle.Text = "Notes";
					}
					else
					{
						this.uc_scenario_notes1.lblTitle.Text = "Notes (Read Only)";
					}
				}
				else if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "DESCRIPTION")
				{
                    btnHelp.Show();
                    m_helpChapter = "PROCESSOR_DESCRIPTION";
                    if (((Control)this.tbDesc).Enabled)
					{
						this.uc_scenario1.lblTitle.Text = "Description";
					}
					else
					{
						this.uc_scenario1.lblTitle.Text = "Description (Read Only)";
					}
				}
				else if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "DATA SOURCES")
				{
                    btnHelp.Hide();
                    if (((Control)this.tbDataSources).Enabled)
					{
						this.uc_datasource1.lblTitle.Text = "Scenario Datasource";
					}
					else
					{
						this.uc_datasource1.lblTitle.Text = "Scenario Datasource (Read Only)";
					}
				}


			}
		}
		public void LoadRuleDefinitions()
		{
            // Initialize Queries object for use by child forms
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oReference.LoadDatasource = true;
            string ScenarioId = this.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            m_oQueries.LoadDatasources(true, m_bUsingSqlite, "processor", ScenarioId);
            
            this.m_oProcessorScenarioItem.ScenarioId = uc_scenario1.txtScenarioId.Text.Trim();
 			this.uc_processor_scenario_harvest_method1.ReferenceProcessorScenarioForm=this;
            frmMain.g_sbpInfo.Text = "Loading Scenario Harvest Method Rule Definitions...Stand By";
            this.uc_processor_scenario_harvest_method1.loadvalues();
            this.uc_processor_scenario_movein_costs1.ReferenceProcessorScenarioForm = this;
            frmMain.g_sbpInfo.Text = "Loading Scenario Move-in Cost Rule Definitions...Stand By";
            this.uc_processor_scenario_movein_costs1.loadvalues();
            frmMain.g_sbpInfo.Text = "Loading Scenario Merch and Chip Market Value Rule Definitions...Stand By";
			this.uc_processor_scenario_merch_chip_value1.loadvalues();
            frmMain.g_sbpInfo.Text = "Loading Scenario Revenue And Cost Escalator Rule Definitions...Stand By";
			this.uc_processor_scenario_escalators1.loadvalues();
            frmMain.g_sbpInfo.Text = "Loading Scenario Supplemental Harvest Component Rule Definitions...Stand By";
            if (!m_bUsingSqlite)
            {
                this.uc_processor_scenario_additional_harvest_cost_columns1.loadvalues();
            }
            else
            {
                this.uc_processor_scenario_additional_harvest_cost_columns1.loadvaluesSqlite();
            }
           
            frmMain.g_sbpInfo.Text = "Loading Scenario Run Data...Stand By";
            this.uc_processor_scenario_run1.loadvalues();
            frmMain.g_sbpInfo.Text = "Ready";
			this.m_bRulesFirstTime=false;
		}
	
		public void SaveRuleDefinitions()
		{
            m_intError = 0; m_strError = "";
			frmMain.g_sbpInfo.Text = "Saving scenario rule definitions...Stand by";
			if (m_bRulesFirstTime==false)
			{
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                      this.WindowState,
                      this.Left,
                      this.Height,
                      this.Width,
                      this.Top);

                this.uc_processor_scenario_harvest_method1.savevalues();
                this.m_intError = uc_processor_scenario_harvest_method1.m_intError;
                this.uc_processor_scenario_movein_costs1.savevalues();
                this.m_intError = uc_processor_scenario_movein_costs1.m_intError;
                this.uc_processor_scenario_merch_chip_value1.savevalues();
                if (m_intError == 0) m_intError = uc_processor_scenario_merch_chip_value1.m_intError;
                this.uc_processor_scenario_escalators1.savevalues();
                if (m_intError == 0) m_intError = uc_processor_scenario_escalators1.m_intError;
                if (!m_bUsingSqlite)
                {
                    this.uc_processor_scenario_additional_harvest_cost_columns1.savevalues();
                }
                else
                {
                    this.uc_processor_scenario_additional_harvest_cost_columns1.savevaluesSqlite();
                }
                if (m_intError == 0) m_intError = uc_processor_scenario_additional_harvest_cost_columns1.m_intError;
                this.uc_scenario_notes1.SaveScenarioNotes();
                this.uc_scenario1.UpdateDescription();

                frmMain.g_oFrmMain.DeactivateStandByAnimation();
			}
            if (m_bTreeGroupsCopied == true)
            {
                this.uc_scenario_tree_groupings1.saveTreeGroupings_FromProperties();
            }
			this.m_bSave=false;
			frmMain.g_sbpInfo.Text = "Ready";
		}

        public void ValidateRuleDefinitions(System.Collections.Generic.IList<string> lstRx)
        {
            m_intError = 0; m_strError = "";
            frmMain.g_sbpInfo.Text = "Validating scenario rule definitions...Stand by";

            frmMain.g_oFrmMain.ActivateStandByAnimation(
                this.WindowState,
                this.Left,
                this.Height,
                this.Width,
                this.Top);

            this.uc_processor_scenario_harvest_method1.val_data(lstRx);
            this.m_intError = uc_processor_scenario_harvest_method1.m_intError;
            //Placeholder for future validations
            //if (m_intError == 0) m_intError = uc_processor_scenario_merch_chip_value1.m_intError;

            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            frmMain.g_sbpInfo.Text = "Ready";
        }

        public void LoadTreeGroupings()
        {
            string strScenario = uc_scenario1.txtScenarioId.Text.Trim();
            this.m_oProcessorScenarioItem.ScenarioId = strScenario;
            if (!m_bUsingSqlite)
            {
                string strScenarioMDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor" + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsDbFile;
                this.m_oProcessorScenarioTools.LoadTreeDiameterGroupValues(strScenarioMDB,
                    strScenario, this.m_oProcessorScenarioItem);
                this.m_oProcessorScenarioTools.LoadTreeSpeciesGroupValues(strScenarioMDB, strScenario,
                    this.m_oProcessorScenarioItem);
            }
            else
            {
                string strScenarioMDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                this.m_oProcessorScenarioTools.LoadTreeDiameterGroupValuesSqlite(strScenarioMDB,
                    strScenario, this.m_oProcessorScenarioItem);
                this.m_oProcessorScenarioTools.LoadTreeSpeciesGroupValuesSqlite(strScenarioMDB, strScenario,
                    this.m_oProcessorScenarioItem);
            }
            this.m_bTreeGroupsFirstTime = false;
        }
		public void EnableTabPage(System.Windows.Forms.TabPage p_oTabPage,bool p_bEnable)
		{
			frmMain.g_oDelegate.SetControlPropertyValue((Control)p_oTabPage,"Enabled",p_bEnable);	
		}
		public void EnableTabPage(System.Windows.Forms.TabControl p_oTabControl,string p_strTabPageNameList,bool p_bEnable)
		{
			int y,z;
			string[] strTabPageNameArray = frmMain.g_oUtils.ConvertListToArray(p_strTabPageNameList,",");
			for (y=0;y<=p_oTabControl.TabPages.Count-1;y++)
			{
				for (z=0;z<=strTabPageNameArray.Length-1;z++)
				{
					if (p_oTabControl.TabPages[y].Name.Trim().ToUpper()==strTabPageNameArray[z].Trim().ToUpper())
					{
						EnableTabPage(p_oTabControl.TabPages[y],p_bEnable);
					}
				}
			}
		}

		public void EnableControl(string p_strControlNameList,bool p_bEnable)
		{
			int x,y;
			string[] strControlNameArray = frmMain.g_oUtils.ConvertListToArray(p_strControlNameList,",");

			foreach (System.Windows.Forms.Control c in this.Controls)
			{
				if (c.Name.Trim().Length > 0)
				{
					for (x=0;x<=strControlNameArray.Length -1;x++)
					{
						if (c.Name.Trim().ToUpper()==strControlNameArray[x].Trim().ToUpper())
						{
							c.Enabled=p_bEnable;
							break;
						}
					}
				}
			}
		}

		private void tabControlScenario_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.Green,System.Drawing.Brushes.White);
		}
		public void TabControl_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e,System.Drawing.Color p_oSelectedBackgroundColor, System.Drawing.Brush p_oSelectedForegroundColor)
		{
			try
			{
				//This line of code will help you to change the apperance like size,name,style.
				System.Windows.Forms.TabControl tc = (System.Windows.Forms.TabControl)sender;
				
				Font f;
				//For background color
				Brush backBrush;
				//For forground color
				Brush foreBrush;
			
				//This construct will hell you to deside which tab page have current focus
				//to change the style.
				if(e.Index == tc.SelectedIndex) //tabControlScenario.SelectedIndex)
				{
					//This line of code will help you to change the apperance like size,name,style.
					f = new Font(e.Font, FontStyle.Regular | FontStyle.Regular);
					f = new Font(e.Font,FontStyle.Regular);

					backBrush = new System.Drawing.SolidBrush(p_oSelectedBackgroundColor);

					foreBrush = p_oSelectedForegroundColor;
				}
				else
				{
					f = e.Font;
					backBrush = new SolidBrush(SystemColors.Control); 
					foreBrush = new SolidBrush(e.ForeColor);
				}
                
				//To set the alignment of the caption.
				string tabName = tc.TabPages[e.Index].Text; //tc.Indexe.IndextabControlScenario.TabPages[e.Index].Text;
				StringFormat sf = new StringFormat();
				sf.Alignment = StringAlignment.Center;
	
				//This will help you to fill the interior portion of
				//selected tabpage.
				e.Graphics.FillRectangle(backBrush, e.Bounds);
				Rectangle r = e.Bounds;
				r = new Rectangle(r.X, r.Y + 3, r.Width, r.Height - 3);

				e.Graphics.DrawString(tabName, f, foreBrush, r, sf);

				sf.Dispose();
				if(e.Index == tc.SelectedIndex)
				{
					f.Dispose();
					backBrush.Dispose();
				}
				else
				{
					backBrush.Dispose();
					foreBrush.Dispose();
				}
			}
			catch(Exception Ex)
			{
				MessageBox.Show(Ex.Message.ToString(),"Error Occured",MessageBoxButtons.OK,MessageBoxIcon.Information);
					
			}

		}
		public void resize_frmProcessorScenario()
		{
			try
			{
				if (this.uc_scenario_open1 !=null)
				{
					this.uc_scenario_open1.Width = this.ClientSize.Width - 2;
					this.uc_scenario_open1.Height = this.ClientSize.Height - this.uc_scenario_open1.Top - 2;
				}

				
				this.btnClose.Top = this.ClientSize.Height - this.btnClose.Height - 2;
				this.btnClose.Left = this.ClientSize.Width - this.btnClose.Width - 2;
				this.tabControlScenario.Top = this.tlbScenario.Top + this.tlbScenario.Height + 2;
				this.tabControlScenario.Height = this.btnClose.Top - this.tabControlScenario.Top - 2;
				this.tabControlScenario.Width = this.ClientSize.Width;
                this.btnHelp.Top = this.ClientSize.Height - this.btnHelp.Height - 2;

				
			}
			catch
			{
			}
			finally 
			{
			}

		}

		private void frmProcessorScenario_Resize(object sender, System.EventArgs e)
		{
			resize_frmProcessorScenario();
		}

		private void tabControlRules_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		  // MessageBox.Show(tabControlRules.SelectedTab.Text);
            SetHelpChapterForRulesTab();
		}

        private void SetHelpChapterForRulesTab()
        {
            if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "MOVE-IN COSTS")
            {
                m_helpChapter = "MOVEIN_COSTS";
            }
            else if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "HARVEST METHOD")
            {
                m_helpChapter = "HARVEST_METHOD";
            }
            else if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "WOOD VALUE")
            {
                m_helpChapter = "WOOD_VALUE";
            }
            else if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "ESCALATORS")
            {
                m_helpChapter = "ESCALATORS";
            }
            else if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "SUPPLEMENTAL HARVEST COSTS")
            {
                m_helpChapter = "SUPPLEMENTAL_HARVEST_COSTS";
            }
            else if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "RUN")
            {
                m_helpChapter = "RUN_PROCESSOR";
            }
        }

		private void tabControlRules_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.DarkBlue,System.Drawing.Brushes.White);
		}
        //
        //HARVEST METHOD PROPERTIES
        //
        public string SelectedHarvestMethod
        {
            get { return this.uc_processor_scenario_harvest_method1.SelectedHarvestMethod; }
        }
        public string HarvestMethodSteepSlope
        {
            get { return this.uc_processor_scenario_harvest_method1.HarvestMethodSteepSlope; }
        }
        public string HarvestMethodLowSlope
        {
            get { return this.uc_processor_scenario_harvest_method1.HarvestMethodLowSlope; }
        }
        public string MinDiaForChips
        {
            get { return uc_processor_scenario_harvest_method1.MinDiaForChips; }
        }
        public string MinDiaForSmallLogs
        {
            get { return uc_processor_scenario_harvest_method1.MinDiaForSmallLogs; }
        }
        public string MinDiaForLargeLogs
        {
            get { return uc_processor_scenario_harvest_method1.MinDiaForLargeLogs; }
        }
        public string SteepSlopePercent
        {
            get { return uc_processor_scenario_harvest_method1.SteepSlopePercent; }
        }
        public string MinDiaForAllTreesSteepSlope
        {
            get { return uc_processor_scenario_harvest_method1.MinDiaForAllTreesSteepSlope; }
        }
        //
        //WOOD MARKET VALUE PROPERTIES
        //
        public uc_processor_scenario_spc_dbh_group_value_collection ReferenceUserControlMarketValueSpeciesDbhGroupCollection
        {
            get { return uc_processor_scenario_merch_chip_value1.ReferenceUserControlMarketValueSpeciesDbhGroupCollection; }
        }
        public string MarketValueChips
        {
            get { return uc_processor_scenario_merch_chip_value1.MarketValueChips; }
        }
        //
        //REVENUE AND COST ESCALATOR PROPERTIES
        //
        public uc_processor_scenario_escalators_value ReferenceUserControlOperatingEscalator
        {
            get { return uc_processor_scenario_escalators1.ReferenceUserControlOperatingEscalator; }
        }
        public uc_processor_scenario_escalators_value ReferenceUserControlMerchEscalator
        {
            get { return uc_processor_scenario_escalators1.ReferenceUserControlMerchEscalator; }
        }
        public uc_processor_scenario_escalators_value ReferenceUserControlChipsEscalator
        {
            get { return uc_processor_scenario_escalators1.ReferenceUserControlChipsEscalator; }
        }
        //
        //ADDITIONAL HARVEST COST COLUMNS
        //
        public uc_processor_scenario_additional_harvest_cost_column_collection ReferenceUserControlAdditionalHarvestCostColumnsCollection
        {
            get { return uc_processor_scenario_additional_harvest_cost_columns1.ReferenceUserControlAdditionalHarvestCostColumnsCollection; }
        }
        //
        //HANDLE TO QUERIES OBJECT SO IT CAN BE RE-USED BY CHILD FORMS WHEN LOADONG
        //
        public Queries LoadedQueries
        {
            get { return m_oQueries; }
        }

        private void frmProcessorScenario_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (frmMain.g_oDelegate.CurrentThreadProcessIdle == false)
            {
                e.Cancel = true;
                return;
            }
            CheckToSave();
            // Delete additional harvest costs worktable; Only does something if using sqlite
            this.uc_processor_scenario_additional_harvest_cost_columns1.DeleteWorkTable();
        }

        private void frmProcessorScenario_Activated(object sender, EventArgs e)
        {
            if (uc_processor_scenario_run1 != null)
            {
                uc_processor_scenario_run1.panel1_Resize();
            }
        }

        private void CopyScenario()
        {
            frmDialog frmTemp = new frmDialog();
            frmTemp.Initialize_Processor_Scenario_Copy(this);
            frmTemp.Text = "FIA Biosum";

            if (m_bRulesFirstTime == true)
            {
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
                LoadRuleDefinitions();
                frmMain.g_oFrmMain.DeactivateStandByAnimation();
            }

            frmTemp.uc_processor_scenario_copy1.ReferenceCurrentScenarioItem = m_oProcessorScenarioItem;
            frmTemp.uc_processor_scenario_copy1.loadvalues();
            DialogResult result = frmTemp.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                frmMain.g_sbpInfo.Text = "Copying scenario rule definitions...Stand by";

                // For tree groupings child forms, set flag that this is copied scenario; The form will call
                // the appropriate load method
                m_bTreeGroupsCopied = true;
                this.uc_scenario1.txtDescription.Text = m_oProcessorScenarioItem.Description;
                this.uc_scenario_notes1.ReferenceProcessorScenarioForm = this;
                frmMain.g_sbpInfo.Text = "Loading Scenario Notes...Stand By";
                this.uc_scenario_notes1.loadvalues_FromProperties();
                this.uc_processor_scenario_harvest_method1.ReferenceProcessorScenarioForm = this;
                frmMain.g_sbpInfo.Text = "Loading Scenario Harvest Method Rule Definitions...Stand By";
                this.uc_processor_scenario_harvest_method1.loadvalues_FromProperties();
                this.uc_processor_scenario_movein_costs1.ReferenceProcessorScenarioForm = this;
                frmMain.g_sbpInfo.Text = "Loading Scenario Move-in Cost Rule Definitions...Stand By";
                this.uc_processor_scenario_movein_costs1.loadvalues_FromProperties();
                frmMain.g_sbpInfo.Text = "Loading Scenario Merch and Chip Market Value Rule Definitions...Stand By";
                this.uc_processor_scenario_merch_chip_value1.loadvalues_FromProperties();
                frmMain.g_sbpInfo.Text = "Loading Scenario Revenue And Cost Escalator Rule Definitions...Stand By";
                this.uc_processor_scenario_escalators1.loadvalues_FromProperties();
                frmMain.g_sbpInfo.Text = "Loading Scenario Supplemental Harvest Component Rule Definitions...Stand By";
                if (!this.m_bUsingSqlite)
                {
                    this.uc_processor_scenario_additional_harvest_cost_columns1.loadvaluesFromProperties();
                }
                else
                {
                    this.uc_processor_scenario_additional_harvest_cost_columns1.loadvaluesFromPropertiesSqlite();
                }

                frmMain.g_sbpInfo.Text = "Ready";
                m_bSave = true;
                MessageBox.Show("Done");
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            //ProcessorScenarioTools oTools = new ProcessorScenarioTools();
            //oTools.migrate_access_data();
            if (!String.IsNullOrEmpty(m_helpChapter))
            {
                if (m_oHelp == null)
                {
                    m_oHelp = new Help(m_xpsFile, m_oEnv);
                }
                m_oHelp.ShowHelp(new string[] { "PROCESSOR", m_helpChapter });
            }
        }

	}
    public class ProcessorScenarioItem
    {
        public HarvestMethod m_oHarvestMethod = new HarvestMethod();
        public MoveInCosts m_oMoveInCosts = new MoveInCosts();
        public Escalators m_oEscalators = new Escalators();
        public HarvestCostItem_Collection m_oHarvestCostItem_Collection = new HarvestCostItem_Collection();
        public TreeSpeciesAndDbhDollarValuesItem_Collection m_oTreeSpeciesAndDbhDollarValuesItem_Collection
               = new TreeSpeciesAndDbhDollarValuesItem_Collection();
        public TreeDiamGroupsItem_Collection m_oTreeDiamGroupsItem_Collection
               = new TreeDiamGroupsItem_Collection();
        public SpcGroupItemCollection m_oSpcGroupItem_Collection
            = new SpcGroupItemCollection();
        public SpcGroupListItemCollection m_oSpcGroupListItem_Collection
             = new SpcGroupListItemCollection();

        public ProcessorScenarioItem()
        {
        }
        private string _strScenarioId = "";
        public string ScenarioId
        {
            get { return _strScenarioId; }
            set { _strScenarioId = value; }
        }
        private string _strDesc = "";
        public string Description
        {
            get { return _strDesc; }
            set { _strDesc = value; }
        }
        private string _strDbPath = "";
        public string DbPath
        {
            get { return _strDbPath; }
            set { _strDbPath = value; }
        }
        private string _strDbFileName = "";
        public string DbFileName
        {
            get { return _strDbFileName; }
            set { _strDbFileName = value; }
        }
        private string _strFullDetailsYN = "N";
        public string DisplayFullDetailsYN
        {
            get { return _strFullDetailsYN; }
            set { _strFullDetailsYN = value; }
        }

        private bool _bSelected = false;
        public bool Selected
        {
            get { return _bSelected; }
            set { _bSelected = value; }
        }

        private string _strNotes = "";
        public string Notes
        {
            get { return _strNotes; }
            set { _strNotes = value; }
        }
        private string _strSourceScenarioId = "";
        public string SourceScenarioId
        {
            get { return _strSourceScenarioId; }
            set { _strSourceScenarioId = value; }
        }
        public void Copy(ProcessorScenarioItem p_oSource,
                         ProcessorScenarioItem p_oDest)
        {
            p_oDest.DbFileName = p_oSource.DbFileName;
            p_oDest.DbPath = p_oSource.DbPath;
            p_oDest.ScenarioId = p_oSource.ScenarioId;
            p_oDest.Description = p_oSource.Description;
            p_oDest.DisplayFullDetailsYN = p_oSource.DisplayFullDetailsYN;
            p_oDest.Selected = p_oSource.Selected;
            p_oDest.Notes = p_oSource.Notes;
            p_oDest.m_oTreeDiamGroupsItem_Collection.Copy(p_oSource.m_oTreeDiamGroupsItem_Collection,
            ref p_oDest.m_oTreeDiamGroupsItem_Collection, true);
            p_oDest.m_oEscalators.Copy(p_oSource.m_oEscalators, p_oDest.m_oEscalators);
            p_oDest.m_oHarvestCostItem_Collection.Copy(
                p_oSource.m_oHarvestCostItem_Collection,
            ref p_oDest.m_oHarvestCostItem_Collection,true);
            p_oDest.m_oHarvestMethod.Copy(p_oSource.m_oHarvestMethod, p_oDest.m_oHarvestMethod);
            p_oDest.m_oMoveInCosts.Copy(p_oSource.m_oMoveInCosts, p_oDest.m_oMoveInCosts);
            p_oDest.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Copy(
                p_oSource.m_oTreeSpeciesAndDbhDollarValuesItem_Collection,
            ref p_oDest.m_oTreeSpeciesAndDbhDollarValuesItem_Collection, true);
        }
       
        public class HarvestMethod
        {
            private HarvestMethodSelection _objHarvestMethodSelection = HarvestMethodSelection.RX;
            public HarvestMethodSelection SelectedHarvestMethod
            {
                get { return _objHarvestMethodSelection; }
                set { _objHarvestMethodSelection = value; }
            }
            private string _strHarvestMethodSteepSlope="";
            public string HarvestMethodSteepSlope
            {
                get { return _strHarvestMethodSteepSlope; }
                set { _strHarvestMethodSteepSlope = value; }
            }
            private string _strHarvestMethodLowSlope = "";
            public string HarvestMethodLowSlope
            {
                get { return _strHarvestMethodLowSlope; }
                set { _strHarvestMethodLowSlope = value; }
            }
            private string _strMinDiaForChips = "3";
            public string MinDiaForChips
            {
                get { return _strMinDiaForChips; }
                set { _strMinDiaForChips = value; }
            }
            private string _strMinDiaForSmallLogs = "7";
            public string MinDiaForSmallLogs
            {
                get { return _strMinDiaForSmallLogs; }
                set { _strMinDiaForSmallLogs = value; }
            }
            private string _strMinDiaForLargeLogs = "20";
            public string MinDiaForLargeLogs
            {
                get { return _strMinDiaForLargeLogs; }
                set { _strMinDiaForLargeLogs = value; }
            }
            private string _strSteepSlopePercent = "40";
            public string SteepSlopePercent
            {
                get { return _strSteepSlopePercent;}
                set { _strSteepSlopePercent = value; }
            }
            private string _strMinDiaForAllTreesSteepSlope = "5";
            public string MinDiaForAllTreesSteepSlope
            {
                get { return _strMinDiaForAllTreesSteepSlope; }
                set { _strMinDiaForAllTreesSteepSlope = value; }
            }
            private bool _bProcessSteepSlope = true;
            public bool ProcessSteepSlope
            {
                get { return _bProcessSteepSlope; }
                set { _bProcessSteepSlope = value; }
            }
            private bool _bProcessLowSlope = true;
            public bool ProcessLowSlope
            {
                get { return _bProcessLowSlope; }
                set { _bProcessLowSlope = value; }
            }
            private string _strWoodlandMerchAsPctOfTotalVol = "60";
            public string WoodlandMerchAsPctOfTotalVol
            {
                get { return _strWoodlandMerchAsPctOfTotalVol; }
                set { _strWoodlandMerchAsPctOfTotalVol = value; }
            }
            private string _strSaplingMerchAsPctOfTotalVol = "80";
            public string SaplingMerchAsPctOfTotalVol
            {
                get { return _strSaplingMerchAsPctOfTotalVol; }
                set { _strSaplingMerchAsPctOfTotalVol = value; }
            }
            private string _strCullPctThreshold = "50";
            public string CullPctThreshold
            {
                get { return _strCullPctThreshold; }
                set { _strCullPctThreshold = value; }
            }
            public void Copy(HarvestMethod p_oSource, HarvestMethod p_oDest)
            {
                p_oDest.HarvestMethodLowSlope = p_oSource.HarvestMethodLowSlope;
                p_oDest.HarvestMethodSteepSlope = p_oSource.HarvestMethodSteepSlope;
                p_oDest.MinDiaForAllTreesSteepSlope = p_oSource.MinDiaForAllTreesSteepSlope;
                p_oDest.MinDiaForChips = p_oSource.MinDiaForChips;
                p_oDest.MinDiaForLargeLogs = p_oSource.MinDiaForLargeLogs;
                p_oDest.MinDiaForSmallLogs = p_oSource.MinDiaForSmallLogs;
                p_oDest.ProcessLowSlope = p_oSource.ProcessLowSlope;
                p_oDest.ProcessSteepSlope = p_oSource.ProcessSteepSlope;
                p_oDest.SteepSlopePercent = p_oSource.SteepSlopePercent;
                p_oDest.SelectedHarvestMethod = p_oSource.SelectedHarvestMethod;
                p_oDest.WoodlandMerchAsPctOfTotalVol = p_oSource.WoodlandMerchAsPctOfTotalVol;
                p_oDest.SaplingMerchAsPctOfTotalVol = p_oSource.SaplingMerchAsPctOfTotalVol;
                p_oDest.CullPctThreshold = p_oSource.CullPctThreshold;
            }
            
        }
        public class MoveInCosts
        {
            private string _strYardDistThreshold = "150";
            public string YardDistThreshold
            {
                get { return _strYardDistThreshold; }
                set { _strYardDistThreshold = value; }
            }
            private string _strAssumedHarvestAreaAc = "80";
            public string AssumedHarvestAreaAc
            {
                get { return _strAssumedHarvestAreaAc; }
                set { _strAssumedHarvestAreaAc = value; }
            }
            private string _strMoveInTimeMultiplier = "0";
            public string MoveInTimeMultiplier
            {
                get { return _strMoveInTimeMultiplier; }
                set { _strMoveInTimeMultiplier = value; }
            }
            private string _strMoveInHoursAddend = "1";
            public string MoveInHoursAddend
            {
                get { return _strMoveInHoursAddend; }
                set { _strMoveInHoursAddend = value; }
            }

            public void Copy(MoveInCosts p_oSource, MoveInCosts p_oDest)
            {
                p_oDest.YardDistThreshold = p_oSource.YardDistThreshold;
                p_oDest.AssumedHarvestAreaAc = p_oSource.AssumedHarvestAreaAc;
                p_oDest.MoveInTimeMultiplier = p_oSource.MoveInTimeMultiplier;
                p_oDest.MoveInHoursAddend = p_oSource.MoveInHoursAddend;
            }

        }
        public class Escalators
        {
            public Escalators()
            {
            }
            private string _strOperatingCostsCycle2="1.00";
            public string OperatingCostsCycle2
            {
                get {return _strOperatingCostsCycle2;}
                set {_strOperatingCostsCycle2=value;}
            }
            private string _strOperatingCostsCycle3="1.00";
            public string OperatingCostsCycle3
            {
                get {return _strOperatingCostsCycle3;}
                set {_strOperatingCostsCycle3=value;}
            }
            private string _strOperatingCostsCycle4="1.00";
            public string OperatingCostsCycle4
            {
                get {return _strOperatingCostsCycle4;}
                set {_strOperatingCostsCycle4=value;}
            }
            private string _strMerchWoodRevenueCycle2="1.00";
            public string MerchWoodRevenueCycle2
            {
                get {return _strMerchWoodRevenueCycle2;}
                set {_strMerchWoodRevenueCycle2=value;}
            }
            private string _strMerchWoodRevenueCycle3="1.00";
            public string MerchWoodRevenueCycle3
            {
                get {return _strMerchWoodRevenueCycle3;}
                set {_strMerchWoodRevenueCycle3=value;}
            }
            private string _strMerchWoodRevenueCycle4="1.00";
            public string MerchWoodRevenueCycle4
            {
                get {return _strMerchWoodRevenueCycle4;}
                set {_strMerchWoodRevenueCycle4=value;}
            }
            private string _strEnergyWoodRevenueCycle2="1.00";
            public string EnergyWoodRevenueCycle2
            {
                get {return _strEnergyWoodRevenueCycle2;}
                set {_strEnergyWoodRevenueCycle2=value;}
            }
            private string _strEnergyWoodRevenueCycle3="1.00";
            public string EnergyWoodRevenueCycle3
            {
                get {return _strEnergyWoodRevenueCycle3;}
                set {_strEnergyWoodRevenueCycle3=value;}
            }
            private string _strEnergyWoodRevenueCycle4="1.00";
            public string EnergyWoodRevenueCycle4
            {
                get {return _strEnergyWoodRevenueCycle4;}
                set {_strEnergyWoodRevenueCycle4=value;}
            }
            private int _intCycleLength=10;
            public int CycleLength
            {
                get {return _intCycleLength;}
                set {_intCycleLength=value;}
            }
            public void Copy(Escalators p_oSource, Escalators p_oDest)
            {
                p_oDest.CycleLength = p_oSource.CycleLength;
                p_oDest.EnergyWoodRevenueCycle2 = p_oSource.EnergyWoodRevenueCycle2;
                p_oDest.EnergyWoodRevenueCycle3 = p_oSource.EnergyWoodRevenueCycle3;
                p_oDest.EnergyWoodRevenueCycle4 = p_oSource.EnergyWoodRevenueCycle4;
                p_oDest.MerchWoodRevenueCycle2 = p_oSource.MerchWoodRevenueCycle2;
                p_oDest.MerchWoodRevenueCycle3 = p_oSource.MerchWoodRevenueCycle3;
                p_oDest.MerchWoodRevenueCycle4 = p_oSource.MerchWoodRevenueCycle4;
                p_oDest.OperatingCostsCycle2 = p_oSource.OperatingCostsCycle2;
                p_oDest.OperatingCostsCycle3 = p_oSource.OperatingCostsCycle3;
                p_oDest.OperatingCostsCycle4 = p_oSource.OperatingCostsCycle4;
            }
            
        }
        public class HarvestCostItem
        {
            public HarvestCostItem()
            {
        
            }
            private string _strColumnName = "";
            public string ColumnName
            {
                get { return _strColumnName; }
                set { _strColumnName = value; }
            }
            private string _strDesc = "";
            public string Description
            {
                get { return _strDesc; }
                set { _strDesc = value; }
            }
            private string _strCPA = "";
            public string DefaultCostPerAcre
            {
                get { return _strCPA; }
                set { _strCPA = value; }
            }
            private string _strType = "";
            /// <summary>
            /// S=SCENARIO ONLY, R=RX
            /// </summary>
            public string Type
            {
                get { return _strType; }
                set { _strType = value; }
            }
            private string _strRx = "";
            public string Rx
            {
                get { return _strRx; }
                set { _strRx = value; }
            }
            public void Copy(HarvestCostItem p_oSource, HarvestCostItem p_oDest)
            {
                p_oDest.ColumnName = p_oSource.ColumnName;
                p_oDest.DefaultCostPerAcre = p_oSource.DefaultCostPerAcre;
                p_oDest.Description = p_oSource.Description;
                p_oDest.Rx = p_oSource.Rx;
                p_oDest.Type = p_oSource.Type;
            }
        }
        public class HarvestCostItem_Collection : System.Collections.CollectionBase
        {
            public HarvestCostItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(FIA_Biosum_Manager.ProcessorScenarioItem.HarvestCostItem m_HarvestCostItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_HarvestCostItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_HarvestCostItem);

                // return collection
                //return this;
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
            public FIA_Biosum_Manager.ProcessorScenarioItem.HarvestCostItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (FIA_Biosum_Manager.ProcessorScenarioItem.HarvestCostItem)List[Index];
            }
            public void Copy(HarvestCostItem_Collection p_oSource,
                         ref HarvestCostItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    HarvestCostItem oItem = new HarvestCostItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);

                }
            }

        }
        public class TreeSpeciesAndDbhDollarValuesItem
        {
            public TreeSpeciesAndDbhDollarValuesItem()
            {
            }
            private string _strSpeciesGroup = "";
            public string SpeciesGroup
            {
                get { return _strSpeciesGroup; }
                set { _strSpeciesGroup = value; }
            }
            private string _strDbhGroup = "";
            public string DbhGroup
            {
                get { return _strDbhGroup; }
                set { _strDbhGroup = value; }
            }
            private bool _bUseAsEnergyWood = false;
            public bool UseAsEnergyWood
            {
                get { return _bUseAsEnergyWood; }
                set { _bUseAsEnergyWood = value; }
            }
            private string _strMerchDollarPerCubicFootValue = "0.00";
            public string MerchDollarPerCubicFootValue
            {
                get { return _strMerchDollarPerCubicFootValue; }
                set { _strMerchDollarPerCubicFootValue = value; }
            }
            private string _strChipsDollarPerCubicFootValue = "0.00";
            public string ChipsDollarPerCubicFootValue
            {
                get { return _strChipsDollarPerCubicFootValue; }
                set { _strChipsDollarPerCubicFootValue = value; }
            }
            private int _intSpeciesGroupId;
            public int SpeciesGroupId
            {
                get { return _intSpeciesGroupId; }
                set { _intSpeciesGroupId = value; }
            }
            private int _intDiameterGroupId;
            public int DiameterGroupId
            {
                get { return _intDiameterGroupId; }
                set { _intDiameterGroupId = value; }
            }
            public void Copy(TreeSpeciesAndDbhDollarValuesItem p_oSource,
                             TreeSpeciesAndDbhDollarValuesItem p_oDest)
            {
                p_oDest.ChipsDollarPerCubicFootValue = p_oSource.ChipsDollarPerCubicFootValue;
                p_oDest.DbhGroup = p_oSource.DbhGroup;
                p_oDest.MerchDollarPerCubicFootValue = p_oSource.MerchDollarPerCubicFootValue;
                p_oDest.SpeciesGroup = p_oSource.SpeciesGroup;
                p_oDest.DiameterGroupId = p_oSource.DiameterGroupId;
                p_oDest.SpeciesGroupId = p_oSource.SpeciesGroupId;
                p_oDest.UseAsEnergyWood = p_oSource.UseAsEnergyWood;
   
            }

        }
        public class TreeSpeciesAndDbhDollarValuesItem_Collection : System.Collections.CollectionBase
        {
            public TreeSpeciesAndDbhDollarValuesItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(FIA_Biosum_Manager.ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem m_TreeSpeciesAndDbhDollarValuesItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_TreeSpeciesAndDbhDollarValuesItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_TreeSpeciesAndDbhDollarValuesItem);

                // return collection
                //return this;
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
            public FIA_Biosum_Manager.ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (FIA_Biosum_Manager.ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem)List[Index];
            }
            public void Copy(TreeSpeciesAndDbhDollarValuesItem_Collection p_oSource,
                         ref TreeSpeciesAndDbhDollarValuesItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    TreeSpeciesAndDbhDollarValuesItem oItem = new TreeSpeciesAndDbhDollarValuesItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);

                }
            }
        }

        public class TreeDiamGroupsItem
        {
            public TreeDiamGroupsItem()
            {
            }
            private string _strDiamGroup = "";
            public string DiamGroup
            {
                get { return _strDiamGroup; }
                set { _strDiamGroup = value; }
            }
            private string _strMinDiam = "";
            public string MinDiam
            {
                get { return _strMinDiam; }
                set { _strMinDiam = value; }
            }
            private string _strMaxDiam = "";
            public string MaxDiam
            {
                get { return _strMaxDiam; }
                set { _strMaxDiam = value; }
            }
            private string _strDiamClass = "";
            public string DiamClass
            {
                get { return _strDiamClass; }
                set { _strDiamClass = value; }
            }
            public void Copy(TreeDiamGroupsItem p_oSource,
                             TreeDiamGroupsItem p_oDest)
            {
                p_oDest.DiamGroup = p_oSource.DiamGroup;
                p_oDest.MinDiam = p_oSource.MinDiam;
                p_oDest.MaxDiam = p_oSource.MaxDiam;
                p_oDest.DiamClass = p_oSource.DiamClass;
            }
        }

        public class TreeDiamGroupsItem_Collection : System.Collections.CollectionBase
        {
            public TreeDiamGroupsItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(FIA_Biosum_Manager.ProcessorScenarioItem.TreeDiamGroupsItem m_TreeDiamGroupsItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_TreeDiamGroupsItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_TreeDiamGroupsItem);

                // return collection
                //return this;
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
            public FIA_Biosum_Manager.ProcessorScenarioItem.TreeDiamGroupsItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (FIA_Biosum_Manager.ProcessorScenarioItem.TreeDiamGroupsItem)List[Index];
            }
            public void Copy(TreeDiamGroupsItem_Collection p_oSource,
                         ref TreeDiamGroupsItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    TreeDiamGroupsItem oItem = new TreeDiamGroupsItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);

                }
            }

        }
        public class SpcGroupItem
        {

            private string _strSpeciesGroupLabel = "";
            private int _intSpeciesGrp = -1;

            public SpcGroupItem()
            {
            }
            public string SpeciesGroupLabel
            {
                get { return _strSpeciesGroupLabel; }
                set { _strSpeciesGroupLabel = value; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGrp; }
                set { _intSpeciesGrp = value; }
            }
            public void Copy(SpcGroupItem p_oSource,
                             SpcGroupItem p_oDest)
            {
                p_oDest.SpeciesGroupLabel = p_oSource.SpeciesGroupLabel;
                p_oDest.SpeciesGroup = p_oSource.SpeciesGroup;
            }
        }
        public class SpcGroupItemCollection : System.Collections.CollectionBase
        {
            public SpcGroupItemCollection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(SpcGroupItem spc_group_item1)
            {
                // vérify if object is not already in
                if (this.List.Contains(spc_group_item1))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(spc_group_item1);

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
            public SpcGroupItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (SpcGroupItem)List[Index];
            }
            public void Copy(SpcGroupItemCollection p_oSource,
             ref SpcGroupItemCollection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    SpcGroupItem oItem = new SpcGroupItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);
                }
            }
        }
        public class SpcGroupListItem
        {
            private int _intSpeciesGrp = -1;
            private string _strCommonName = "";
            private int _intSpCd = -1;
            

            public SpcGroupListItem()
            {
            }
            public string CommonName
            {
                get { return _strCommonName; }
                set { _strCommonName = value; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGrp; }
                set { _intSpeciesGrp = value; }
            }
            public int SpeciesCode
            {
                get { return _intSpCd; }
                set { _intSpCd = value; }
            }
            public void Copy(SpcGroupListItem p_oSource,
                             SpcGroupListItem p_oDest)
            {
                p_oDest.CommonName = p_oSource.CommonName;
                p_oDest.SpeciesGroup = p_oSource.SpeciesGroup;
                p_oDest.SpeciesCode = p_oSource.SpeciesCode;
            }
        }
        public class SpcGroupListItemCollection : System.Collections.CollectionBase
        {
            public SpcGroupListItemCollection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(SpcGroupListItem spc_group_item1)
            {
                // vérify if object is not already in
                if (this.List.Contains(spc_group_item1))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(spc_group_item1);

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
            public SpcGroupListItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (SpcGroupListItem)List[Index];
            }
            public void Copy(SpcGroupListItemCollection p_oSource,
                                ref SpcGroupListItemCollection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    SpcGroupListItem oItem = new SpcGroupListItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);
                }
            }
        }
    }
    public class ProcessorScenarioItem_Collection : System.Collections.CollectionBase
    {
        public ProcessorScenarioItem_Collection()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public void Add(FIA_Biosum_Manager.ProcessorScenarioItem m_ProcessorScenarioItem)
        {
            // vérify if object is not already in
            if (this.List.Contains(m_ProcessorScenarioItem))
                throw new InvalidOperationException();

            // adding it
            this.List.Add(m_ProcessorScenarioItem);

            // return collection
            //return this;
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
        public FIA_Biosum_Manager.ProcessorScenarioItem Item(int Index)
        {
            // The appropriate item is retrieved from the List object and
            // explicitly cast to the Widget type, then returned to the 
            // caller.
            return (FIA_Biosum_Manager.ProcessorScenarioItem)List[Index];
        }
        public void Copy(ProcessorScenarioItem_Collection p_oSource,
                     ref ProcessorScenarioItem_Collection p_oDest, bool p_bInitializeDest)
        {
            int x;
            if (p_bInitializeDest) p_oDest.Clear();
            for (x = 0; x <= p_oSource.Count - 1; x++)
            {
                ProcessorScenarioItem oItem = new ProcessorScenarioItem();
                oItem.Copy(p_oSource.Item(x), oItem);
                p_oDest.Add(oItem);

            }
        }

    }
   

    public class ProcessorScenarioTools
    {
        public string m_strError = "";
        public int m_intError = 0;
        public ProcessorScenarioTools()
        {
        }

        public void LoadScenario(string p_strScenarioId, Queries p_oQueries, bool bUsingSqlite,
            ProcessorScenarioItem_Collection p_oProcessorScenarioItem_Collection)
        {

            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            p_oQueries.m_oFvs.LoadDatasource = true;
            p_oQueries.m_oFIAPlot.LoadDatasource = true;
            p_oQueries.m_oProcessor.LoadDatasource = true;
            p_oQueries.m_oReference.LoadDatasource = true;
            p_oQueries.LoadDatasources(true, bUsingSqlite, "processor", p_strScenarioId);
            if (!bUsingSqlite)
            {
                p_oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                    p_oQueries.m_strTempDbFile,
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                    "P");
                LoadAll(p_oQueries.m_strTempDbFile, p_oQueries, p_strScenarioId, p_oProcessorScenarioItem_Collection);
            }
            else
            {
                LoadAllSqlite(p_oQueries, p_strScenarioId, p_oProcessorScenarioItem_Collection);
            }
        }

        public void LoadAll(string p_strDbFile,Queries p_oQueries, string p_strScenarioId, FIA_Biosum_Manager.ProcessorScenarioItem_Collection p_oProcessorScenarioItem_Collection)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                
                    ProcessorScenarioItem oItem = new ProcessorScenarioItem();
                    this.LoadGeneral(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                    this.LoadTreeDiameterGroupValues(p_oQueries.m_strTempDbFile,p_strScenarioId, oItem);
                    this.LoadTreeSpeciesGroupValues(p_oQueries.m_strTempDbFile, p_strScenarioId, oItem);
                    this.LoadHarvestMethod(oAdo, oAdo.m_OleDbConnection, oItem);
                    this.LoadMoveInCosts(p_oQueries.m_strTempDbFile, oItem);
                    this.LoadSpeciesAndDiameterGroupDollarValues(
                        oAdo, oAdo.m_OleDbConnection, oItem);
                    this.LoadHarvestCostComponents(oAdo,
                        oAdo.m_OleDbConnection, oItem);
                    this.LoadEscalators(oAdo, oAdo.m_OleDbConnection,p_oQueries, oItem);
                    p_oProcessorScenarioItem_Collection.Add(oItem);
                
            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }

        public void LoadAllSqlite(Queries p_oQueries, string p_strScenarioId, 
            FIA_Biosum_Manager.ProcessorScenarioItem_Collection p_oProcessorScenarioItem_Collection)
        {
            // Access version used a temp file with links; Trying to skip that
            string strScenarioDB = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            //ado_data_access oAdo = new ado_data_access();
            //oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            //if (oAdo.m_intError == 0)
            //{

            ProcessorScenarioItem oItem = new ProcessorScenarioItem();
            this.LoadGeneralSqlite(strScenarioDB, p_strScenarioId, oItem);
            this.LoadTreeDiameterGroupValuesSqlite(strScenarioDB, p_strScenarioId, oItem);
            this.LoadTreeSpeciesGroupValuesSqlite(strScenarioDB, p_strScenarioId, oItem);
            this.LoadHarvestMethodSqlite(strScenarioDB, oItem);
            this.LoadMoveInCostsSqlite(strScenarioDB, oItem);
            this.LoadSpeciesAndDiameterGroupDollarValuesSqlite(strScenarioDB, oItem);
            this.LoadHarvestCostComponentsSqlite(strScenarioDB, oItem);

            this.LoadEscalatorsSqlite(p_oQueries, strScenarioDB, oItem);
            p_oProcessorScenarioItem_Collection.Add(oItem);

            //}
            //m_intError = oAdo.m_intError;
            //oAdo.CloseConnection(oAdo.m_OleDbConnection);
            //oAdo = null;
        }

        public void LoadGeneral(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            p_oAdo.SqlQueryReader(p_oConn, "SELECT * FROM scenario " + " WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'");

            if (p_oAdo.m_intError == 0)
            {
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        //
                        //SCENARIO ID
                        //
                        if (p_oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                        {
                                p_oProcessorScenarioItem.ScenarioId = p_oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                        }
                        //
                        //DESCRIPTION
                        //
                        if (p_oAdo.m_OleDbDataReader["description"] != System.DBNull.Value)
                        {
                                p_oProcessorScenarioItem.Description = p_oAdo.m_OleDbDataReader["description"].ToString().Trim();
                        }
                        //
                        //PATH
                        //
                        if (p_oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        {
                                p_oProcessorScenarioItem.DbPath = p_oAdo.m_OleDbDataReader["path"].ToString().Trim();
                        }
                        //
                        //FILE
                        //
                        if (p_oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                        {
                            p_oProcessorScenarioItem.DbFileName = p_oAdo.m_OleDbDataReader["file"].ToString().Trim();
                        }
                        //
                        //NOTES
                        //
                        if (p_oAdo.m_OleDbDataReader["notes"] != System.DBNull.Value)
                        {
                            p_oProcessorScenarioItem.Notes = p_oAdo.m_OleDbDataReader["notes"].ToString().Trim();
                        }


                    }
                    p_oAdo.m_OleDbDataReader.Close();
                }
            }
        }
        public void LoadGeneralSqlite(string p_strDbFile, string p_strScenarioId, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.SqlQueryReader(oConn, "SELECT * FROM scenario " + " WHERE TRIM(UPPER(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'");
                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            //
                            //SCENARIO ID
                            //
                            if (dataMgr.m_DataReader["scenario_id"] != System.DBNull.Value)
                            {
                                p_oProcessorScenarioItem.ScenarioId = dataMgr.m_DataReader["scenario_id"].ToString().Trim();
                            }
                            //
                            //DESCRIPTION
                            //
                            if (dataMgr.m_DataReader["description"] != System.DBNull.Value)
                            {
                                p_oProcessorScenarioItem.Description = dataMgr.m_DataReader["description"].ToString().Trim();
                            }
                            //
                            //PATH
                            //
                            if (dataMgr.m_DataReader["path"] != System.DBNull.Value)
                            {
                                p_oProcessorScenarioItem.DbPath = dataMgr.m_DataReader["path"].ToString().Trim();
                            }
                            //
                            //FILE
                            //
                            if (dataMgr.m_DataReader["file"] != System.DBNull.Value)
                            {
                                p_oProcessorScenarioItem.DbFileName = dataMgr.m_DataReader["file"].ToString().Trim();
                            }
                            //
                            //NOTES
                            //
                            if (dataMgr.m_DataReader["notes"] != System.DBNull.Value)
                            {
                                p_oProcessorScenarioItem.Notes = dataMgr.m_DataReader["notes"].ToString().Trim();
                            }
                        }
                        dataMgr.m_DataReader.Close();
                    }
                }
            }


        }
        public void LoadHarvestMethod(string p_strDbFile, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                this.LoadHarvestMethod(oAdo,oAdo.m_OleDbConnection, p_oProcessorScenarioItem);
            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        public void LoadHarvestMethod(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            p_oAdo.SqlQueryReader(p_oConn, "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " WHERE TRIM(UCASE(scenario_id))='" + p_oProcessorScenarioItem.ScenarioId.Trim().ToUpper() + "'");

            if (p_oAdo.m_intError == 0)
            {
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        //
                        //HARVEST METHOD SELECTION
                        //
                        string strHarvestMethodSelection = p_oAdo.m_OleDbDataReader["HarvestMethodSelection"].ToString().Trim();
                        if (strHarvestMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST.Value))
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod = HarvestMethodSelection.LOWEST_COST;
                        }
                        else if (strHarvestMethodSelection.Equals(HarvestMethodSelection.SELECTED.Value))
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod = HarvestMethodSelection.SELECTED;
                        }
                        else
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod = HarvestMethodSelection.RX;
                        }
                        //
                        //HARVEST METHOD LOW SLOPE
                        //
                        if (p_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodLowSlope = p_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"].ToString().Trim();
                            }
                        }
                        //
                        //HARVEST METHOD STEEP SLOPE
                        //
                        if (p_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodSteepSlope = p_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"].ToString().Trim();
                            }
                        }
                        //
                        //MINIMUM CHIPS DBH
                        //
                        if (p_oAdo.m_OleDbDataReader["min_chip_dbh"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["min_chip_dbh"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForChips= p_oAdo.m_OleDbDataReader["min_chip_dbh"].ToString().Trim();
                            }
                        }
                        //
                        //MINIMUM DBH FOR SMALL LOGS
                        //
                        if (p_oAdo.m_OleDbDataReader["min_sm_log_dbh"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["min_sm_log_dbh"].ToString().Trim().Length > 0)
                            {
                               p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForSmallLogs = p_oAdo.m_OleDbDataReader["min_sm_log_dbh"].ToString().Trim();
                            }
                        }
                        //
                        //MINIMUM DBH FOR LARGE LOGS
                        //
                        if (p_oAdo.m_OleDbDataReader["min_lg_log_dbh"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["min_lg_log_dbh"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForLargeLogs = p_oAdo.m_OleDbDataReader["min_lg_log_dbh"].ToString().Trim();
                            }
                        }
                        //
                        //STEEP SLOPE PERCENT
                        //
                        if (p_oAdo.m_OleDbDataReader["SteepSlope"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["SteepSlope"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent = p_oAdo.m_OleDbDataReader["SteepSlope"].ToString().Trim();
                            }
                        }
                        //
                        //MINIMUM DBH FOR STEEP SLOPE
                        //
                        if (p_oAdo.m_OleDbDataReader["min_dbh_steep_slope"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["min_dbh_steep_slope"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope = p_oAdo.m_OleDbDataReader["min_dbh_steep_slope"].ToString().Trim();
                            }
                        }
                        //
                        //PROCESS LOW AND STEEP SLOPE DURING RUN
                        //
                        if (p_oAdo.m_OleDbDataReader["ProcessLowSlopeYN"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["ProcessLowSlopeYN"].ToString().Trim() == "Y")
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.ProcessLowSlope = true;
                        }
                        else
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.ProcessLowSlope = false;
                        }
                        if (p_oAdo.m_OleDbDataReader["ProcessSteepSlopeYN"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["ProcessSteepSlopeYN"].ToString().Trim() == "Y")
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.ProcessSteepSlope = true;
                        }
                        else
                        {
                            p_oProcessorScenarioItem.m_oHarvestMethod.ProcessSteepSlope = false;
                        }
                        //
                        //WOODLAND MERCH AS PCT OF TOTAL VOLUME
                        //
                        if (p_oAdo.m_OleDbDataReader["WoodlandMerchAsPercentOfTotalVol"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["WoodlandMerchAsPercentOfTotalVol"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.WoodlandMerchAsPctOfTotalVol = p_oAdo.m_OleDbDataReader["WoodlandMerchAsPercentOfTotalVol"].ToString().Trim();
                            }
                        }
                        //
                        //SAPLING MERCH AS PCT OF TOTAL VOLUME
                        //
                        if (p_oAdo.m_OleDbDataReader["SaplingMerchAsPercentOfTotalVol"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["SaplingMerchAsPercentOfTotalVol"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.SaplingMerchAsPctOfTotalVol = p_oAdo.m_OleDbDataReader["SaplingMerchAsPercentOfTotalVol"].ToString().Trim();
                            }
                        }
                        //
                        //CULL PCT THRESHOLD
                        //
                        if (p_oAdo.m_OleDbDataReader["CullPctThreshold"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["CullPctThreshold"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.CullPctThreshold = p_oAdo.m_OleDbDataReader["CullPctThreshold"].ToString().Trim();
                            }
                        }
                    }
                    p_oAdo.m_OleDbDataReader.Close();
                }
            }
        }
        public void LoadHarvestMethodSqlite(string p_strDbFile, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.SqlQueryReader(oConn, "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " WHERE TRIM(UPPER(scenario_id))='" + p_oProcessorScenarioItem.ScenarioId.Trim().ToUpper() + "'");

                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            //
                            //HARVEST METHOD SELECTION
                            //
                            string strHarvestMethodSelection = dataMgr.m_DataReader["HarvestMethodSelection"].ToString().Trim();
                            if (strHarvestMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST.Value))
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod = HarvestMethodSelection.LOWEST_COST;
                            }
                            else if (strHarvestMethodSelection.Equals(HarvestMethodSelection.SELECTED.Value))
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod = HarvestMethodSelection.SELECTED;
                            }
                            else
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod = HarvestMethodSelection.RX;
                            }
                            //
                            //HARVEST METHOD LOW SLOPE
                            //
                            if (dataMgr.m_DataReader["HarvestMethodLowSlope"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["HarvestMethodLowSlope"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodLowSlope = dataMgr.m_DataReader["HarvestMethodLowSlope"].ToString().Trim();
                                }
                            }
                            //
                            //HARVEST METHOD STEEP SLOPE
                            //
                            if (dataMgr.m_DataReader["HarvestMethodSteepSlope"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["HarvestMethodSteepSlope"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodSteepSlope = dataMgr.m_DataReader["HarvestMethodSteepSlope"].ToString().Trim();
                                }
                            }
                            //
                            //MINIMUM CHIPS DBH
                            //
                            if (dataMgr.m_DataReader["min_chip_dbh"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["min_chip_dbh"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForChips = dataMgr.m_DataReader["min_chip_dbh"].ToString().Trim();
                                }
                            }
                            //
                            //MINIMUM DBH FOR SMALL LOGS
                            //
                            if (dataMgr.m_DataReader["min_sm_log_dbh"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["min_sm_log_dbh"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForSmallLogs = dataMgr.m_DataReader["min_sm_log_dbh"].ToString().Trim();
                                }
                            }
                            //
                            //MINIMUM DBH FOR LARGE LOGS
                            //
                            if (dataMgr.m_DataReader["min_lg_log_dbh"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["min_lg_log_dbh"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForLargeLogs = dataMgr.m_DataReader["min_lg_log_dbh"].ToString().Trim();
                                }
                            }
                            //
                            //STEEP SLOPE PERCENT
                            //
                            if (dataMgr.m_DataReader["SteepSlope"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["SteepSlope"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent = dataMgr.m_DataReader["SteepSlope"].ToString().Trim();
                                }
                            }
                            //
                            //MINIMUM DBH FOR STEEP SLOPE
                            //
                            if (dataMgr.m_DataReader["min_dbh_steep_slope"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["min_dbh_steep_slope"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope = dataMgr.m_DataReader["min_dbh_steep_slope"].ToString().Trim();
                                }
                            }
                            //
                            //PROCESS LOW AND STEEP SLOPE DURING RUN
                            //
                            if (dataMgr.m_DataReader["ProcessLowSlopeYN"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["ProcessLowSlopeYN"].ToString().Trim() == "Y")
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.ProcessLowSlope = true;
                            }
                            else
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.ProcessLowSlope = false;
                            }
                            if (dataMgr.m_DataReader["ProcessSteepSlopeYN"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["ProcessSteepSlopeYN"].ToString().Trim() == "Y")
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.ProcessSteepSlope = true;
                            }
                            else
                            {
                                p_oProcessorScenarioItem.m_oHarvestMethod.ProcessSteepSlope = false;
                            }
                            //
                            //WOODLAND MERCH AS PCT OF TOTAL VOLUME
                            //
                            if (dataMgr.m_DataReader["WoodlandMerchAsPercentOfTotalVol"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["WoodlandMerchAsPercentOfTotalVol"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.WoodlandMerchAsPctOfTotalVol = dataMgr.m_DataReader["WoodlandMerchAsPercentOfTotalVol"].ToString().Trim();
                                }
                            }
                            //
                            //SAPLING MERCH AS PCT OF TOTAL VOLUME
                            //
                            if (dataMgr.m_DataReader["SaplingMerchAsPercentOfTotalVol"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["SaplingMerchAsPercentOfTotalVol"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.SaplingMerchAsPctOfTotalVol = dataMgr.m_DataReader["SaplingMerchAsPercentOfTotalVol"].ToString().Trim();
                                }
                            }
                            //
                            //CULL PCT THRESHOLD
                            //
                            if (dataMgr.m_DataReader["CullPctThreshold"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["CullPctThreshold"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oHarvestMethod.CullPctThreshold = dataMgr.m_DataReader["CullPctThreshold"].ToString().Trim();
                                }
                            }
                        }
                        dataMgr.m_DataReader.Close();
                    }
                }
            }
            m_intError = dataMgr.m_intError;
        }
        public void LoadMoveInCosts(string p_strDbFile, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                this.LoadMoveInCosts(oAdo, oAdo.m_OleDbConnection, p_oProcessorScenarioItem);
            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        public void LoadMoveInCosts(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            p_oAdo.SqlQueryReader(p_oConn, "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName + " WHERE TRIM(UCASE(scenario_id))='" + p_oProcessorScenarioItem.ScenarioId.Trim().ToUpper() + "'");

            if (p_oAdo.m_intError == 0)
            {
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        //
                        //YARDING DISTANCE THRESHOLD
                        //
                        if (p_oAdo.m_OleDbDataReader["yard_dist_threshold"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["yard_dist_threshold"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oMoveInCosts.YardDistThreshold = p_oAdo.m_OleDbDataReader["yard_dist_threshold"].ToString().Trim();
                            }
                        }
                        //
                        //ASSUMED HARVEST AREA
                        //
                        if (p_oAdo.m_OleDbDataReader["assumed_harvest_area_ac"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["assumed_harvest_area_ac"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oMoveInCosts.AssumedHarvestAreaAc = p_oAdo.m_OleDbDataReader["assumed_harvest_area_ac"].ToString().Trim();
                            }
                        }
                        //
                        //MOVE-IN TIME MULTIPLIER
                        //
                        if (p_oAdo.m_OleDbDataReader["move_in_time_multiplier"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["move_in_time_multiplier"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oMoveInCosts.MoveInTimeMultiplier = p_oAdo.m_OleDbDataReader["move_in_time_multiplier"].ToString().Trim();
                            }
                        }
                        //
                        //MOVE-IN HOURS ADDEND
                        //
                        if (p_oAdo.m_OleDbDataReader["move_in_hours_addend"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["move_in_hours_addend"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oMoveInCosts.MoveInHoursAddend = p_oAdo.m_OleDbDataReader["move_in_hours_addend"].ToString().Trim();
                            }
                        }
                    }
                    p_oAdo.m_OleDbDataReader.Close();
                }
            }
        }
        public void LoadMoveInCostsSqlite(string p_strDbFile, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.SqlQueryReader(oConn, "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName + 
                    " WHERE TRIM(UPPER(scenario_id)) = '" + p_oProcessorScenarioItem.ScenarioId.Trim().ToUpper() + "'");

                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            //
                            //YARDING DISTANCE THRESHOLD
                            //
                            if (dataMgr.m_DataReader["yard_dist_threshold"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["yard_dist_threshold"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oMoveInCosts.YardDistThreshold = dataMgr.m_DataReader["yard_dist_threshold"].ToString().Trim();
                                }
                            }
                            //
                            //ASSUMED HARVEST AREA
                            //
                            if (dataMgr.m_DataReader["assumed_harvest_area_ac"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["assumed_harvest_area_ac"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oMoveInCosts.AssumedHarvestAreaAc = dataMgr.m_DataReader["assumed_harvest_area_ac"].ToString().Trim();
                                }
                            }
                            //
                            //MOVE-IN TIME MULTIPLIER
                            //
                            if (dataMgr.m_DataReader["move_in_time_multiplier"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["move_in_time_multiplier"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oMoveInCosts.MoveInTimeMultiplier = dataMgr.m_DataReader["move_in_time_multiplier"].ToString().Trim();
                                }
                            }
                            //
                            //MOVE-IN HOURS ADDEND
                            //
                            if (dataMgr.m_DataReader["move_in_hours_addend"] != System.DBNull.Value)
                            {
                                if (dataMgr.m_DataReader["move_in_hours_addend"].ToString().Trim().Length > 0)
                                {
                                    p_oProcessorScenarioItem.m_oMoveInCosts.MoveInHoursAddend = dataMgr.m_DataReader["move_in_hours_addend"].ToString().Trim();
                                }
                            }
                        }
                        dataMgr.m_DataReader.Close();
                    }
                }
            }
        }

        private void CreateSpeciesDiameterGroupCollection(ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            int x;
            for (x = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x >= 0; x--)
            {
                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Remove(x);
            }
            //
            //CREATE A COLLECTION CONTAINING EACH SPECIES AND DBH GROUP COMBINATION
            //
            foreach (ProcessorScenarioItem.SpcGroupItem objSpcGroup in p_oProcessorScenarioItem.m_oSpcGroupItem_Collection)
            {
                foreach (ProcessorScenarioItem.TreeDiamGroupsItem objDiamGroup in p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection)
                {
                    ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem oItem = new ProcessorScenarioItem.TreeSpeciesAndDbhDollarValuesItem();
                    oItem.SpeciesGroup = objSpcGroup.SpeciesGroupLabel;
                    oItem.DbhGroup = objDiamGroup.DiamClass;
                    // Added to support Optimizer context table DIAMETER_SPP_GRP_REF_C
                    oItem.SpeciesGroupId = objSpcGroup.SpeciesGroup;
                    oItem.DiameterGroupId = Convert.ToInt16(objDiamGroup.DiamGroup);
                    p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Add(oItem);
                }
            }
        }

        public void LoadSpeciesAndDiameterGroupDollarValues(ado_data_access p_oAdo,
                                                            System.Data.OleDb.OleDbConnection p_oConn,
                                                            ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            CreateSpeciesDiameterGroupCollection(p_oProcessorScenarioItem);
            //
            //UPDATE MERCH AND CHIP DOLLAR VALUE
            //
            p_oAdo.m_strSQL = "SELECT wood_bin,merch_value,chip_value,species_group,diam_group " +
                "FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName +
                " WHERE TRIM(scenario_id)='" + p_oProcessorScenarioItem.ScenarioId.Trim() + "'"; 
            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    string strSpcGrp = "";
                    string strDbhGrp = "";
                    string strMerchValue = "0.00";
                    string strChipValue = "0.00";
                    string strWoodBin = "";

                    if (p_oAdo.m_OleDbDataReader["species_group"] != System.DBNull.Value)
                    {
                        int intSpcGrp = (int) p_oAdo.m_OleDbDataReader["species_group"];
                        foreach (ProcessorScenarioItem.SpcGroupItem objSpcGroup in p_oProcessorScenarioItem.m_oSpcGroupItem_Collection)
                        {

                            if (objSpcGroup.SpeciesGroup == intSpcGrp)
                            {
                                strSpcGrp = objSpcGroup.SpeciesGroupLabel.Trim();
                                break; // Exit for loop when we've found the species group
                            }
                        }
                    }
                    if (p_oAdo.m_OleDbDataReader["diam_group"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["diam_group"].ToString().Trim().Length > 0)
                    {
                        string strDiamGroup = p_oAdo.m_OleDbDataReader["diam_group"].ToString().Trim();
                        foreach (ProcessorScenarioItem.TreeDiamGroupsItem objTreeDiam in p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection)
                        {

                            if (objTreeDiam.DiamGroup.Equals(strDiamGroup))
                            {
                                strDbhGrp = objTreeDiam.DiamClass.Trim();
                                break; // Exit for loop when we've found the diameter group
                            }
                        }
                    }
                    if (p_oAdo.m_OleDbDataReader["merch_value"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["merch_value"].ToString().Trim().Length > 0)
                    {
                        strMerchValue = p_oAdo.m_OleDbDataReader["merch_value"].ToString().Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["chip_value"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["chip_value"].ToString().Trim().Length > 0)
                    {
                        strChipValue = p_oAdo.m_OleDbDataReader["chip_value"].ToString().Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["wood_bin"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["wood_bin"].ToString().Trim().Length > 0)
                    {
                        strWoodBin = p_oAdo.m_OleDbDataReader["wood_bin"].ToString().Trim();
                    }
                    if (strSpcGrp.Length > 0 && strDbhGrp.Length > 0)
                    {
                        for (int x = 0; x <= p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                        {
                            string tempSpGrp = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup;
                            string tempDbhGrp = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup;
                            //merch value is associated with each dbh and spc group
                            if (strSpcGrp.ToUpper() ==
                                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().ToUpper() &&
                                strDbhGrp.ToUpper() ==
                                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim().ToUpper())
                            {
                                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue = strMerchValue;
                                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue = strChipValue;
                                if (strWoodBin == "M")
                                {
                                    p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood = false;
                                }
                                else
                                {
                                    p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood = true;
                                }
                            }
                        }
                    }
                }
            }
            p_oAdo.m_OleDbDataReader.Close();
        }
        public void LoadSpeciesAndDiameterGroupDollarValuesSqlite(string p_strDbFile, ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            CreateSpeciesDiameterGroupCollection(p_oProcessorScenarioItem);

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.SqlQueryReader(oConn, "SELECT wood_bin,merch_value,chip_value,species_group,diam_group " +
                    "FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName +
                    " WHERE TRIM(scenario_id)='" + p_oProcessorScenarioItem.ScenarioId.Trim() + "'");

                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            string strSpcGrp = "";
                            string strDbhGrp = "";
                            string strMerchValue = "0.00";
                            string strChipValue = "0.00";
                            string strWoodBin = "";

                            if (dataMgr.m_DataReader["species_group"] != System.DBNull.Value)
                            {
                                int intSpcGrp = Convert.ToInt16(dataMgr.m_DataReader["species_group"]);
                                foreach (ProcessorScenarioItem.SpcGroupItem objSpcGroup in p_oProcessorScenarioItem.m_oSpcGroupItem_Collection)
                                {

                                    if (objSpcGroup.SpeciesGroup == intSpcGrp)
                                    {
                                        strSpcGrp = objSpcGroup.SpeciesGroupLabel.Trim();
                                        break; // Exit for loop when we've found the species group
                                    }
                                }
                            }
                            if (dataMgr.m_DataReader["diam_group"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["diam_group"].ToString().Trim().Length > 0)
                            {
                                string strDiamGroup = dataMgr.m_DataReader["diam_group"].ToString().Trim();
                                foreach (ProcessorScenarioItem.TreeDiamGroupsItem objTreeDiam in p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection)
                                {

                                    if (objTreeDiam.DiamGroup.Equals(strDiamGroup))
                                    {
                                        strDbhGrp = objTreeDiam.DiamClass.Trim();
                                        break; // Exit for loop when we've found the diameter group
                                    }
                                }
                            }
                            if (dataMgr.m_DataReader["merch_value"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["merch_value"].ToString().Trim().Length > 0)
                            {
                                strMerchValue = String.Format("{0:0.00}", Convert.ToDouble(dataMgr.m_DataReader["merch_value"]));
                            }
                            if (dataMgr.m_DataReader["chip_value"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["chip_value"].ToString().Trim().Length > 0)
                            {
                                strChipValue = String.Format("{0:0.00}", Convert.ToDouble(dataMgr.m_DataReader["chip_value"]));
                            }
                            if (dataMgr.m_DataReader["wood_bin"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["wood_bin"].ToString().Trim().Length > 0)
                            {
                                strWoodBin = dataMgr.m_DataReader["wood_bin"].ToString().Trim();
                            }
                            if (strSpcGrp.Length > 0 && strDbhGrp.Length > 0)
                            {
                                for (int x = 0; x <= p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
                                {
                                    string tempSpGrp = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup;
                                    string tempDbhGrp = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup;
                                    //merch value is associated with each dbh and spc group
                                    if (strSpcGrp.ToUpper() ==
                                        p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup.Trim().ToUpper() &&
                                        strDbhGrp.ToUpper() ==
                                        p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup.Trim().ToUpper())
                                    {
                                        p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue = strMerchValue;
                                        p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue = strChipValue;
                                        if (strWoodBin == "M")
                                        {
                                            p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood = false;
                                        }
                                        else
                                        {
                                            p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood = true;
                                        }
                                    }
                                }
                            }
                        }
                        dataMgr.m_DataReader.Close();
                    }
                }
            }
        }
        public void LoadEscalators(string p_strDbFile, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                this.LoadHarvestMethod(oAdo, oAdo.m_OleDbConnection, p_oProcessorScenarioItem);
            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        public void LoadEscalators(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn,Queries p_oQueries, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
                //
				//UPDATE CYCLE ESCALATOR TEXT BOXES
				//
				//update the text boxes that were previously saved
				//scenario_cost_revenue_escalators
				p_oAdo.m_strSQL = "SELECT EscalatorOperatingCosts_Cycle2," + 
					                     "EscalatorOperatingCosts_Cycle3," +
                                         "EscalatorOperatingCosts_Cycle4," +
                                         "EscalatorMerchWoodRevenue_Cycle2," +
                                         "EscalatorMerchWoodRevenue_Cycle3," +
                                         "EscalatorMerchWoodRevenue_Cycle4," +
                                         "EscalatorEnergyWoodRevenue_Cycle2," +
                                         "EscalatorEnergyWoodRevenue_Cycle3," +
                                         "EscalatorEnergyWoodRevenue_Cycle4 " + 
					              "FROM scenario_cost_revenue_escalators " + 
					              "WHERE TRIM(scenario_id) = '" + p_oProcessorScenarioItem.ScenarioId.Trim() + "'";
				
				p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection,p_oAdo.m_strSQL);
				if (p_oAdo.m_intError==0)
				{
					if (p_oAdo.m_OleDbDataReader.HasRows)
					{
						while (p_oAdo.m_OleDbDataReader.Read())
						{
                            //
                            //OPERATING COSTS CYCLE2
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle2"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle2"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 = p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle2"].ToString().Trim();
                            }
                            //
                            //OPERATING COSTS CYCLE3
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle3"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle3"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 = p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle3"].ToString().Trim();
                            }
                            //
                            //OPERATING COSTS CYCLE4
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle4"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle4"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 = p_oAdo.m_OleDbDataReader["EscalatorOperatingCosts_Cycle4"].ToString().Trim();
                            }
                            //														//
                            //MERCH WOOD REVENUE CYCLE2
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle2"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle2"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle2 = p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle2"].ToString().Trim();
                            }
                            //
                            //MERCH WOOD REVENUE CYCLE3
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle3"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle3"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle3 = p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle3"].ToString().Trim();
                            }
                            //
                            //MERCH WOOD REVENUE CYCLE4
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle4"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle4"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle4 = p_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle4"].ToString().Trim();
                            }
                            //														//
                            //ENERGY WOOD REVENUE CYCLE2
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle2"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle2"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2 = p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle2"].ToString().Trim();
                            }
                            //
                            //ENERGY WOOD REVENUE CYCLE3
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle3"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle3"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3 = p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle3"].ToString().Trim();
                            }
                            //
                            //ENERGY WOOD REVENUE CYCLE4
                            //
							if (p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle4"] != System.DBNull.Value &&
                                p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle4"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4 = p_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle4"].ToString().Trim();
                            }
						}
					}
					p_oAdo.m_OleDbDataReader.Close();
					
				}



            if (p_oQueries != null)
		    	p_oProcessorScenarioItem.m_oEscalators.CycleLength=Convert.ToInt32(p_oAdo.getSingleDoubleValueFromSQLQuery(p_oAdo.m_OleDbConnection,"SELECT TOP 1 rxcycle_length FROM " + p_oQueries.m_oFvs.m_strRxPackageTable,"temp"));            

        }
        public void LoadEscalatorsSqlite(Queries p_oQueries, string p_strDbFile, FIA_Biosum_Manager.ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.SqlQueryReader(oConn, "SELECT EscalatorOperatingCosts_Cycle2, " + 
                                         "EscalatorOperatingCosts_Cycle3," +
                                         "EscalatorOperatingCosts_Cycle4," +
                                         "EscalatorMerchWoodRevenue_Cycle2," +
                                         "EscalatorMerchWoodRevenue_Cycle3," +
                                         "EscalatorMerchWoodRevenue_Cycle4," +
                                         "EscalatorEnergyWoodRevenue_Cycle2," +
                                         "EscalatorEnergyWoodRevenue_Cycle3," +
                                         "EscalatorEnergyWoodRevenue_Cycle4 " +
                                  "FROM scenario_cost_revenue_escalators " +
                                  "WHERE TRIM(scenario_id) = '" + p_oProcessorScenarioItem.ScenarioId.Trim() + "'");
                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            //
                            //OPERATING COSTS CYCLE2
                            //
                            if (dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle2"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle2"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle2 = dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle2"].ToString().Trim();
                            }
                            //
                            //OPERATING COSTS CYCLE3
                            //
                            if (dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle3"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle3"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle3 = dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle3"].ToString().Trim();
                            }
                            //
                            //OPERATING COSTS CYCLE4
                            //
                            if (dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle4"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle4"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.OperatingCostsCycle4 = dataMgr.m_DataReader["EscalatorOperatingCosts_Cycle4"].ToString().Trim();
                            }
                            //														//
                            //MERCH WOOD REVENUE CYCLE2
                            //
                            if (dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle2"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle2"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle2 = dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle2"].ToString().Trim();
                            }
                            //
                            //MERCH WOOD REVENUE CYCLE3
                            //
                            if (dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle3"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle3"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle3 = dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle3"].ToString().Trim();
                            }
                            //
                            //MERCH WOOD REVENUE CYCLE4
                            //
                            if (dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle4"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle4"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle4 = dataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle4"].ToString().Trim();
                            }
                            //														//
                            //ENERGY WOOD REVENUE CYCLE2
                            //
                            if (dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle2"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle2"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2 = dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle2"].ToString().Trim();
                            }
                            //
                            //ENERGY WOOD REVENUE CYCLE3
                            //
                            if (dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle3"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle3"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3 = dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle3"].ToString().Trim();
                            }
                            //
                            //ENERGY WOOD REVENUE CYCLE4
                            //
                            if (dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle4"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle4"].ToString().Trim().Length > 0)
                            {
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4 = dataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle4"].ToString().Trim();
                            }
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
            }

            // Retrieve cycle length from any rx package item
            RxTools oRxTools = new RxTools();
            RxPackageItem_Collection p_oRxPackageItemCollection = new RxPackageItem_Collection();
            string rxPackageDb = p_oQueries.m_oDataSource.getFullPathAndFile("TREATMENT PACKAGES");
            oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(rxPackageDb, p_oQueries, p_oRxPackageItemCollection);
            if (p_oRxPackageItemCollection != null && p_oRxPackageItemCollection.Count > 0)
            {
                RxPackageItem rxPackageItem = p_oRxPackageItemCollection.Item(0);
                p_oProcessorScenarioItem.m_oEscalators.CycleLength = rxPackageItem.RxCycleLength;
            }
        }
        public void LoadHarvestCostComponents(ado_data_access p_oAdo,
                                              System.Data.OleDb.OleDbConnection p_oConn,
                                              ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            for (int x = p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count-1;
                     x >= 0;
                     x--)
            {
                p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Remove(x);
            }
            //
            //load up any scenario columns and the default values
            //
            p_oAdo.m_strSQL = "SELECT rx,[ColumnName],[Description],Default_CPA FROM scenario_harvest_cost_columns WHERE TRIM(scenario_id)='" + p_oProcessorScenarioItem.ScenarioId.Trim() + "'";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    ProcessorScenarioItem.HarvestCostItem oItem = new ProcessorScenarioItem.HarvestCostItem();
                    if (p_oAdo.m_OleDbDataReader["rx"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["rx"].ToString().Trim().Length > 0)
                    {
                        oItem.Rx = p_oAdo.m_OleDbDataReader["rx"].ToString().Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["ColumnName"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["ColumnName"].ToString().Trim().Length > 0)
                    {
                        oItem.ColumnName = p_oAdo.m_OleDbDataReader["ColumnName"].ToString().Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["Description"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["Description"].ToString().Trim().Length > 0)
                    {
                        oItem.Description = p_oAdo.m_OleDbDataReader["Description"].ToString().Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["Default_CPA"] != System.DBNull.Value &&
                        p_oAdo.m_OleDbDataReader["Default_CPA"].ToString().Trim().Length > 0)
                    {
                        oItem.DefaultCostPerAcre = p_oAdo.m_OleDbDataReader["Default_CPA"].ToString().Trim();
                    }
                    if (oItem.Rx.Trim().Length == 0)
                    {
                        oItem.Type = "S";
                    }
                    else
                    {
                        oItem.Type = "R";
                    }
                    p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Add(oItem);

                }
            }
            p_oAdo.m_OleDbDataReader.Close();
        }

        public void LoadHarvestCostComponentsSqlite(string p_strDbFile, ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            for (int x = p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count - 1;
                     x >= 0;
                     x--)
            {
                p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Remove(x);
            }
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                //
                //load up any scenario columns and the default values
                //
                dataMgr.SqlQueryReader(oConn, "SELECT rx,[ColumnName],[Description],Default_CPA FROM scenario_harvest_cost_columns WHERE TRIM(scenario_id)='"
                    + p_oProcessorScenarioItem.ScenarioId.Trim() + "'");
                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            ProcessorScenarioItem.HarvestCostItem oItem = new ProcessorScenarioItem.HarvestCostItem();
                            if (dataMgr.m_DataReader["rx"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["rx"].ToString().Trim().Length > 0)
                            {
                                oItem.Rx = dataMgr.m_DataReader["rx"].ToString().Trim();
                            }
                            if (dataMgr.m_DataReader["ColumnName"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["ColumnName"].ToString().Trim().Length > 0)
                            {
                                oItem.ColumnName = dataMgr.m_DataReader["ColumnName"].ToString().Trim();
                            }
                            if (dataMgr.m_DataReader["Description"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["Description"].ToString().Trim().Length > 0)
                            {
                                oItem.Description = dataMgr.m_DataReader["Description"].ToString().Trim();
                            }
                            if (dataMgr.m_DataReader["Default_CPA"] != System.DBNull.Value &&
                                dataMgr.m_DataReader["Default_CPA"].ToString().Trim().Length > 0)
                            {
                                string strDefaultCPA = String.Format("{0:0.00}", Convert.ToDouble(dataMgr.m_DataReader["Default_CPA"]));
                                oItem.DefaultCostPerAcre = strDefaultCPA;
                            }
                            if (oItem.Rx.Trim().Length == 0)
                            {
                                oItem.Type = "S";
                            }
                            else
                            {
                                oItem.Type = "R";
                            }
                            p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Add(oItem);
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
            }            
        }

        public void LoadTreeDiameterGroupValues(string p_strDbFile, string p_strScenarioId,
                                                ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            int x;
            //REMOVE ANY EXISTING ITEMS
            for (x = p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x >= 0; x--)
            {
                p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Remove(x);
            }
            //
            //QUERY ALL DIAMETER GROUPS
            //
            ado_data_access oAdo = new ado_data_access();
            string strConn = oAdo.getMDBConnString(p_strDbFile, "", "");
            using (var oConn = new System.Data.OleDb.OleDbConnection(strConn))
            {
                oConn.Open();
                oAdo.SqlQueryReader(oConn, "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                    " WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' ORDER BY DIAM_GROUP");

                //
                //CREATE A COLLECTION CONTAINING EACH DIAMETER GROUP
                //
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        ProcessorScenarioItem.TreeDiamGroupsItem oItem = new ProcessorScenarioItem.TreeDiamGroupsItem();
                        oItem.DiamGroup = oAdo.m_OleDbDataReader["diam_group"].ToString().Trim();
                        oItem.DiamClass = oAdo.m_OleDbDataReader["diam_class"].ToString().Trim();
                        oItem.MaxDiam = oAdo.m_OleDbDataReader["max_diam"].ToString().Trim();
                        oItem.MinDiam = oAdo.m_OleDbDataReader["min_diam"].ToString().Trim();
                        p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Add(oItem);
                    }
                }
                oAdo.m_OleDbDataReader.Close();
            }
        }
        public void LoadTreeDiameterGroupValuesSqlite(string p_strDbFile, string p_strScenarioId,
                                                      ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            int x;
            //REMOVE ANY EXISTING ITEMS
            for (x = p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x >= 0; x--)
            {
                p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Remove(x);
            }
            //
            //QUERY ALL DIAMETER GROUPS
            //
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                dataMgr.SqlQueryReader(oConn, "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                    " WHERE TRIM(UPPER(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' ORDER BY DIAM_GROUP");

                //
                //CREATE A COLLECTION CONTAINING EACH DIAMETER GROUP
                //
                if (dataMgr.m_DataReader.HasRows)
                {
                    while (dataMgr.m_DataReader.Read())
                    {
                        ProcessorScenarioItem.TreeDiamGroupsItem oItem = new ProcessorScenarioItem.TreeDiamGroupsItem();
                        oItem.DiamGroup = dataMgr.m_DataReader["diam_group"].ToString().Trim();
                        oItem.DiamClass = dataMgr.m_DataReader["diam_class"].ToString().Trim();
                        oItem.MaxDiam = dataMgr.m_DataReader["max_diam"].ToString().Trim();
                        oItem.MinDiam = dataMgr.m_DataReader["min_diam"].ToString().Trim();
                        p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Add(oItem);
                    }
                }
                dataMgr.m_DataReader.Close();
            }
        }

        public void LoadTreeSpeciesGroupValues(string p_strDbFile,
                                        string p_strScenarioId,
                                        ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            int x;
            //REMOVE ANY EXISTING ITEMS
            for (x = p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count - 1; x >= 0; x--)
            {
                p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Remove(x);
            }

            ado_data_access oAdo = new ado_data_access();
            string strConn = oAdo.getMDBConnString(p_strDbFile, "", "");
            using (var oConn = new System.Data.OleDb.OleDbConnection(strConn))
            {
                oConn.Open();
                //LOAD USER SPECIES COMMON NAME GROUPING ASSIGNMENTS
                /****************************************************************************************
                 **load any previous group assignments 
                 ****************************************************************************************/
                oAdo.m_strSQL = "select * from " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName +
                " WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' order by species_group";
                oAdo.SqlQueryReader(oConn, oAdo.m_strSQL);
                if (oAdo.m_intError == 0)
                {
                    /**************************************************************************************
                     **go through the species group table and assign values to the group label text box
                     **************************************************************************************/
                    int intSpcGrp;
                    string strSpeciesLabel;
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            if (oAdo.m_OleDbDataReader["species_group"] != DBNull.Value)
                            {
                                intSpcGrp = Convert.ToInt32(oAdo.m_OleDbDataReader["species_group"]);
                                strSpeciesLabel = Convert.ToString(oAdo.m_OleDbDataReader["species_label"]).Trim();
                                ProcessorScenarioItem.SpcGroupItem oItem = new ProcessorScenarioItem.SpcGroupItem();
                                oItem.SpeciesGroup = intSpcGrp;
                                oItem.SpeciesGroupLabel = strSpeciesLabel;
                                p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Add(oItem);
                            }
                        }
                    }
                    oAdo.m_OleDbDataReader.Close();
                }
                //REMOVE ANY EXISTING ITEMS
                for (x = p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Count - 1; x >= 0; x--)
                {
                    p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Remove(x);
                }

                //go through the species table and load its 
                //group list box with the species common name
                oAdo.m_strSQL = "SELECT DISTINCT common_name,species_group, spcd " +
                    "FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                    " WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + " '";
                oAdo.SqlQueryReader(oConn, oAdo.m_strSQL);
                if (oAdo.m_intError == 0)
                {
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            if (oAdo.m_OleDbDataReader["common_name"] != DBNull.Value)
                            {
                                ProcessorScenarioItem.SpcGroupListItem oItem = new ProcessorScenarioItem.SpcGroupListItem();
                                oItem.SpeciesGroup = Convert.ToInt32(oAdo.m_OleDbDataReader["species_group"]);
                                oItem.CommonName = Convert.ToString(oAdo.m_OleDbDataReader["common_name"]).Trim();
                                oItem.SpeciesCode = Convert.ToInt32(oAdo.m_OleDbDataReader["spcd"]);
                                p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Add(oItem);
                            }
                        }
                    }
                    oAdo.m_OleDbDataReader.Close();
                }
            }
        }

        public void LoadTreeSpeciesGroupValuesSqlite(string p_strDbFile,
                                                     string p_strScenarioId,
                                                     ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            int x;
            //REMOVE ANY EXISTING ITEMS
            for (x = p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count - 1; x >= 0; x--)
            {
                p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Remove(x);
            }

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(p_strDbFile);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                //LOAD USER SPECIES COMMON NAME GROUPING ASSIGNMENTS
                /****************************************************************************************
                 **load any previous group assignments 
                 ****************************************************************************************/
                dataMgr.m_strSQL = "select * from " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName +
                " WHERE TRIM(UPPER(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' order by species_group";
                dataMgr.SqlQueryReader(oConn, dataMgr.m_strSQL);
                if (dataMgr.m_intError == 0)
                {
                    /**************************************************************************************
                     **go through the species group table and assign values to the group label text box
                     **************************************************************************************/
                    int intSpcGrp;
                    string strSpeciesLabel;
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            if (dataMgr.m_DataReader["species_group"] != DBNull.Value)
                            {
                                intSpcGrp = Convert.ToInt32(dataMgr.m_DataReader["species_group"]);
                                strSpeciesLabel = Convert.ToString(dataMgr.m_DataReader["species_label"]).Trim();
                                ProcessorScenarioItem.SpcGroupItem oItem = new ProcessorScenarioItem.SpcGroupItem();
                                oItem.SpeciesGroup = intSpcGrp;
                                oItem.SpeciesGroupLabel = strSpeciesLabel;
                                p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Add(oItem);
                            }
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
                //REMOVE ANY EXISTING ITEMS
                for (x = p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Count - 1; x >= 0; x--)
                {
                    p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Remove(x);
                }

                //go through the species table and load its 
                //group list box with the species common name
                dataMgr.m_strSQL = "SELECT DISTINCT common_name,species_group, spcd " +
                    "FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName +
                    " WHERE TRIM(UPPER(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'" +
                    " ORDER BY COMMON_NAME COLLATE NOCASE";
                dataMgr.SqlQueryReader(oConn, dataMgr.m_strSQL);
                if (dataMgr.m_intError == 0)
                {
                    if (dataMgr.m_DataReader.HasRows)
                    {
                        while (dataMgr.m_DataReader.Read())
                        {
                            if (dataMgr.m_DataReader["common_name"] != DBNull.Value)
                            {
                                ProcessorScenarioItem.SpcGroupListItem oItem = new ProcessorScenarioItem.SpcGroupListItem();
                                oItem.SpeciesGroup = Convert.ToInt32(dataMgr.m_DataReader["species_group"]);
                                oItem.CommonName = Convert.ToString(dataMgr.m_DataReader["common_name"]).Trim();
                                oItem.SpeciesCode = Convert.ToInt32(dataMgr.m_DataReader["spcd"]);
                                p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Add(oItem);
                            }
                        }
                    }
                    dataMgr.m_DataReader.Close();
                }
            }
        }

        public string ScenarioProperties(ProcessorScenarioItem p_oProcessorScenarioItem)
        {
            int x;
            string strLine = "";
            strLine = "***PROCESSOR SCENARIO***\r\n\r\n";
            strLine = strLine + "ID\r\n";
            strLine = strLine + "-----\r\n";
            strLine = strLine + p_oProcessorScenarioItem.ScenarioId + "\r\n\r\n";
            strLine = strLine + "Description\r\n";
            strLine = strLine + "---------------------\r\n";
            strLine = strLine + p_oProcessorScenarioItem.Description + "\r\n\r\n";
            strLine = strLine + "Tree Diameter Groups\r\n";
            strLine = strLine + "---------------------\r\n";
            strLine = strLine + "--Group ID--    --Minimum Diameter--   --Maximum Diameter--    --Definition--   \r\n";
            for (x = 0; x <= p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Count - 1; x++)
            {
                strLine = strLine + String.Format("{0,7}{1,19}{2, 24}{3,25}",
                    p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x).DiamGroup,
                    p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x).MinDiam,
                    p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x).MaxDiam,
                    p_oProcessorScenarioItem.m_oTreeDiamGroupsItem_Collection.Item(x).DiamClass);
                strLine = strLine + "\r\n";
            }
            strLine = strLine + "\r\n";
            strLine = strLine + "Tree Species Groups\r\n";
            strLine = strLine + "---------------------\r\n";
            strLine = strLine + "--Group ID--     --Label--                 --Common Name--      --FIA spcd-- \r\n";
            // Cache species groups items in dictionary so we can get the label for each species
            IDictionary<String, String> _dictSpcGroupLabels =
                new Dictionary<String, String>();
            for (x = 0; x <= p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Count - 1; x++)
            {
                ProcessorScenarioItem.SpcGroupItem _oItem = p_oProcessorScenarioItem.m_oSpcGroupItem_Collection.Item(x);
                if (!_dictSpcGroupLabels.ContainsKey(Convert.ToString(_oItem.SpeciesGroup)))
                    _dictSpcGroupLabels.Add(Convert.ToString(_oItem.SpeciesGroup), _oItem.SpeciesGroupLabel);
            }
            for (x = 0; x <= p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Count - 1; x++)
            {
                string _strLabel = "Missing";
                ProcessorScenarioItem.SpcGroupListItem _oItem = p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(x);
                if (_dictSpcGroupLabels.ContainsKey(Convert.ToString(_oItem.SpeciesGroup)))
                {
                    _strLabel = _dictSpcGroupLabels[Convert.ToString(_oItem.SpeciesGroup)];
                }
                    
                strLine = strLine + String.Format("{0,7}{1,19}{2, 32}{3,16}",
                        p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(x).SpeciesGroup,
                        _strLabel,
                        p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(x).CommonName,
                        p_oProcessorScenarioItem.m_oSpcGroupListItem_Collection.Item(x).SpeciesCode);
                        strLine = strLine + "\r\n";
            }
            strLine = strLine + "\r\n";
            strLine = strLine + "Harvest Method Low Slope\r\n";
            strLine = strLine + "-------------------------------\r\n";
            strLine = strLine + "Process Low Slope: ";
            if (p_oProcessorScenarioItem.m_oHarvestMethod.ProcessLowSlope == true)
            {
                strLine = strLine + "Yes\r\n";
            }
            else
            {
                strLine = strLine + "No\r\n";
            }
            strLine = strLine + "Harvest method selection: ";
            strLine = strLine + p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod.Value + "\r\n";
            if (p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod == HarvestMethodSelection.SELECTED)
            {
                strLine = strLine + "Processor selected harvest method\r\n";
                strLine = strLine + "Harvest Method Low Slope: " + p_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodLowSlope + "\r\n";
            }
            strLine = strLine + "Chips Minimum Diameter: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForChips + "\r\n";
            strLine = strLine + "Small Logs Minimum Diameter: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForSmallLogs + "\r\n";
            strLine = strLine + "Large Logs Minimum Diameter: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForLargeLogs + "\r\n";
            strLine = strLine + "Slope Percent: < " + p_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + "\r\n";
            strLine = strLine + "Woodland Merch as Pct of Total Vol: = " + p_oProcessorScenarioItem.m_oHarvestMethod.WoodlandMerchAsPctOfTotalVol + "\r\n";
            strLine = strLine + "Sapling Merch as Pct of Total Vol: = " + p_oProcessorScenarioItem.m_oHarvestMethod.SaplingMerchAsPctOfTotalVol + "\r\n";
            strLine = strLine + "Cull Pct Threshold: = " + p_oProcessorScenarioItem.m_oHarvestMethod.CullPctThreshold + "\r\n";

            strLine = strLine + "\r\nHarvest Method Steep Slope\r\n";
            strLine = strLine + "-------------------------------\r\n";
            strLine = strLine + "Process Steep Slope: ";
            if (p_oProcessorScenarioItem.m_oHarvestMethod.ProcessSteepSlope == true)
            {
                strLine = strLine + "Yes\r\n";
            }
            else
            {
                strLine = strLine + "No\r\n";
            }
            strLine = strLine + "Harvest method selection: ";
            strLine = strLine + p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod.Value + "\r\n";
            if (p_oProcessorScenarioItem.m_oHarvestMethod.SelectedHarvestMethod == HarvestMethodSelection.SELECTED)
            {
                strLine = strLine + "Processor selected harvest method\r\n";
                strLine = strLine + "Harvest Method Steep Slope: " + p_oProcessorScenarioItem.m_oHarvestMethod.HarvestMethodSteepSlope + "\r\n";
            }
            strLine = strLine + "All Trees Minimum Diameter: > " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForAllTreesSteepSlope + "\r\n";
            strLine = strLine + "Chips Minimum Diameter: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForChips + "\r\n";
            strLine = strLine + "Small Logs Minimum Diameter: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForSmallLogs + "\r\n";
            strLine = strLine + "Large Logs Minimum Diameter: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.MinDiaForLargeLogs + "\r\n";
            strLine = strLine + "Slope Percent: >= " + p_oProcessorScenarioItem.m_oHarvestMethod.SteepSlopePercent + "\r\n";
            strLine = strLine + "Woodland Merch as Pct of Total Vol: = " + p_oProcessorScenarioItem.m_oHarvestMethod.WoodlandMerchAsPctOfTotalVol + "\r\n";
            strLine = strLine + "Sapling Merch as Pct of Total Vol: = " + p_oProcessorScenarioItem.m_oHarvestMethod.SaplingMerchAsPctOfTotalVol + "\r\n";
            strLine = strLine + "Cull Pct Threshold: = " + p_oProcessorScenarioItem.m_oHarvestMethod.CullPctThreshold + "\r\n";


            strLine = strLine + "\r\nAdditional Harvest Cost Components \r\n";
            strLine = strLine + "----------------------------------------------------------------------------\r\n";
            for (x = 0; x <= p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Count - 1; x++)
            {
                if (p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x).ColumnName.Trim().Length > 0)
                {
                    strLine = strLine + "Cost Component: " +
                           p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x).ColumnName + "\r\n";
                    strLine = strLine + "Description: " +
                          p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x).Description + "\r\n";
                    strLine = strLine + "Association: ";
                    if (p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x).Type == "S")
                    {
                        strLine = strLine + "Scenario Only\r\n";
                        strLine = strLine + "Treatment: NA\r\n";
                    }
                    else
                    {
                        strLine = strLine + "Treatment (RX)\r\n";
                        strLine = strLine + "Treatment: " + p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x).Rx + "\r\n";
                    }
                    strLine = strLine + "Default Cost Per Acre: " + p_oProcessorScenarioItem.m_oHarvestCostItem_Collection.Item(x).DefaultCostPerAcre + "\r\n\r\n";


                }
            }

            //Check to make sure we have items in m_oTreeSpeciesAndDbhDollarValuesItem_Collection before trying to access the first item
            string strChipMarketValue = "NA";
            if (p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count > 0)
                strChipMarketValue = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(0).ChipsDollarPerCubicFootValue;
            strLine = strLine + "\r\nTree Species And Diameter Group Market Value Assignment\r\n";
            strLine = strLine + "------------------------------------------------------------\r\n";
            strLine = strLine + "Chip Market Value (Dollars Per Green Ton): " + strChipMarketValue + "\r\n\r\n";
            strLine = strLine + "--Species--              --Dia--         --EnergyWood-- --*MerchValue-- --ChipValue--\r\n";
            for (x = 0; x <= p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Count - 1; x++)
            {
                string strEnergyWood = "";
                string strMerchValue = "";
                string strChipValue = "";
                if (p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).UseAsEnergyWood)
                {
                    strEnergyWood = "Yes";
                    strMerchValue = "NA";
                    strChipValue = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).ChipsDollarPerCubicFootValue;
                }
                else
                {
                    strEnergyWood = "No";
                    strChipValue = "NA";
                    strMerchValue = p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).MerchDollarPerCubicFootValue;
                }
                strLine = strLine + String.Format("{0,-25}{1,-11}{2, 15}{3,13}{4,14}",
                                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).SpeciesGroup,
                                p_oProcessorScenarioItem.m_oTreeSpeciesAndDbhDollarValuesItem_Collection.Item(x).DbhGroup,
                                strEnergyWood,
                                strMerchValue,
                                strChipValue);
                strLine = strLine + "\r\n";

            }
            strLine = strLine + "*Dollars Per Cubic Foot\r\n";

            strLine = strLine + "\r\nCost And Revenue Escalators\r\n";
            strLine = strLine + "-------------------------------------------------------\r\n";
            strLine = strLine + "--Type--               --Cycle2--  --Cycle3--  --Cycle4--\r\n";
            strLine = strLine + String.Format("{0,-19}{1,12}{2, 12}{3,12}",
                                "Operating Costs",
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2,
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3,
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4);
            strLine = strLine + "\r\n";
            strLine = strLine + String.Format("{0,-19}{1,12}{2, 12}{3,12}",
                                "Merch Wood Revenue",
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle2,
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle3,
                                p_oProcessorScenarioItem.m_oEscalators.MerchWoodRevenueCycle4);
            strLine = strLine + "\r\n";
            strLine = strLine + String.Format("{0,-19}{1,12}{2, 12}{3,12}",
                                "Energy Wood Revenue",
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle2,
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle3,
                                p_oProcessorScenarioItem.m_oEscalators.EnergyWoodRevenueCycle4);

            strLine = strLine + "\r\n\r\nEOF";
            return strLine;

        }

        public void migrate_access_data()
        {
            string targetDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                @"\processor\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            string sourceDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                @"\processor\" + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodDbFile;
            if (System.IO.File.Exists(targetDbFile) == false)
            {
                frmMain.g_oFrmMain.frmProject.uc_project1.CreateProcessorScenarioRuleDefinitionSqliteDbAndTables(targetDbFile);
            }
            else
            {
                MessageBox.Show(targetDbFile + " already exists. New tables cannot be created!!", "FIA Biosum");
                return;
            }

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            ODBCMgr odbcmgr = new ODBCMgr();
            try
            {

                string[] arrTargetTables = { };
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dataMgr.GetConnectionString(targetDbFile)))
                {
                    conn.Open();
                    arrTargetTables = dataMgr.getTableNames(conn);
                    if (arrTargetTables.Length < 1)
                    {
                        MessageBox.Show("Target SQLite tables could not be created. Migration stopped!!", "FIA Biosum");
                        return;
                    }

                    // custom processing for scenario_additional_harvest_costs
                    string[] strSourceColumnsArray = new string[0];
                    string strTableName = Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName;
                    oDao.getFieldNames(sourceDbFile, strTableName, ref strSourceColumnsArray);
                    foreach (string strColumn in strSourceColumnsArray)
                    {
                        if (!dataMgr.ColumnExists(conn, strTableName, strColumn))
                        {
                            dataMgr.AddColumn(conn, strTableName, strColumn, "REAL", "");
                        }
                    }
                }

                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, targetDbFile);

                // Create temporary database
                utils oUtils = new utils();
                env oEnv = new env();
                string strTempAccdb = oUtils.getRandomFile(oEnv.strTempDir, "accdb");
                oDao.CreateMDB(strTempAccdb);

                // Link all the target tables to the database
                for (int i = 0; i < arrTargetTables.Length; i++)
                {
                    oDao.CreateSQLiteTableLink(strTempAccdb, arrTargetTables[i], arrTargetTables[i] + "_1",
                        ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, targetDbFile);
                }
                oDao.CreateTableLinks(strTempAccdb, sourceDbFile);  // Link all the source tables to the database

                string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                using (var copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    foreach (var strTable in arrTargetTables)
                    {
                        oAdo.m_strSQL = "INSERT INTO " + strTable + "_1" +
                        " SELECT * FROM " + strTable;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }

                    if (oAdo.m_intError == 0)
                    {
                        // Primary key required to update
                        oAdo.AddPrimaryKey(copyConn, "scenario_1", "scenario_id_pk", "scenario_id");
                        // Set file (database) field to new Sqlite DB
                        string newDbFile = System.IO.Path.GetFileName(Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile);
                        oAdo.m_strSQL = "UPDATE scenario_1 set file = '" +
                            newDbFile + "'";
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                // Clean-up
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                if (oAdo != null)
                {
                    oAdo = null;
                }
                if (oDao != null)
                {
                    oDao.m_DaoWorkspace.Close();
                    oDao = null;
                }
            }
        }

    }

}

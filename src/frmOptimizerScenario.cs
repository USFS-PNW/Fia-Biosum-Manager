using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.OleDb;

namespace FIA_Biosum_Manager
{
	/// <summary>
    /// Summary description for frmOptimizerScenario.
	/// </summary>
	public class frmOptimizerScenario : System.Windows.Forms.Form
	{
		public System.Windows.Forms.TextBox txtDropDown;
		private System.Data.DataView dataView1;
		private System.Windows.Forms.ImageList imgSize;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ImageList imageList2;
		public bool m_bScenarioOpen = false;
		public bool m_ldatasourcefirsttime;
		public bool m_lrulesfirsttime=true;
		private System.Windows.Forms.Button btnClose;
		public FIA_Biosum_Manager.uc_scenario uc_scenario1;
		public FIA_Biosum_Manager.uc_datasource uc_datasource1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_effective uc_scenario_fvs_prepost_variables_effective1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_optimization uc_scenario_fvs_prepost_optimization1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker uc_scenario_fvs_prepost_variables_tiebreaker1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_costs uc_scenario_costs1;
		public FIA_Biosum_Manager.uc_select_list_item uc_select_list_item1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_filter uc_scenario_filter1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_filter uc_scenario_cond_filter1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_psite uc_scenario_psite1;
		public FIA_Biosum_Manager.uc_scenario_open uc_scenario_open1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_run uc_scenario_run1=null;
		public FIA_Biosum_Manager.uc_scenario_notes uc_scenario_notes1;
        public FIA_Biosum_Manager.uc_optimizer_scenario_processor_scenario_select uc_scenario_processor_scenario_select1;
		public FIA_Biosum_Manager.uc_optimizer_scenario_owner_groups uc_scenario_owner_groups1;
        public FIA_Biosum_Manager.uc_optimizer_scenario_select_packages uc_optimizer_scenario_select_packages1;
		public FIA_Biosum_Manager.frmGridView frmGridView1;
		private FIA_Biosum_Manager.frmRunOptimizerScenario frmRunOptimizerScenario1;
		
		public System.Windows.Forms.VScrollBar m_vScrollBar;
		public System.Windows.Forms.HScrollBar m_hScrollBar;
		protected int m_intIntensityTop=0;
		protected int m_intLblTitleTop = 0;
		private int m_intHScrollValue=0;
		private double m_dblHScrollOldPerc=0;
		private double m_dblHScrollNewPerc=0;
		private int m_intHScrollMaxSize=0;
		public System.Windows.Forms.TabControl tabControlScenario;
		private System.Windows.Forms.TabPage tbDesc;
		private System.Windows.Forms.TabPage tbNotes;
		private System.Windows.Forms.TabPage tbDataSources;
		private System.Windows.Forms.TabPage tbRules;
		public System.Windows.Forms.ToolBar tlbScenario;
		public System.Windows.Forms.TabControl tabControlRules;
		private System.Windows.Forms.TabPage tbOwners;
		private System.Windows.Forms.TabPage tbPSites;
		private System.Windows.Forms.TabPage tbFilterPlots;
		private System.Windows.Forms.ToolBarButton btnScenarioNew;
		private System.Windows.Forms.ToolBarButton btnScenarioOpen;
		private System.Windows.Forms.ToolBarButton btnScenarioSave;
		private System.Windows.Forms.ToolBarButton toolBarButton1;
		private System.Windows.Forms.ToolBarButton btnScenarioDelete;
		private System.Windows.Forms.TabPage tbRun;
        public bool m_bSave = false;
		private System.Windows.Forms.TabPage tbFVSVariables;
		public System.Windows.Forms.TabPage tbOptimization;
		public bool m_bOptimizeTabPageEnabled=true;

		private bool m_oTabPageLast;


		public string m_strAllowLeaveTabPageMsg="";
		public bool m_bEnableSelectedTabPage=true;
		public int  m_intEditTabPageIndex=0;
		public string m_strCurrentEditTabControlName="";
		public string m_strCurrentEditTabPageText="";
		public System.Windows.Forms.TabControl tabControlFVSVariables;
		private System.Windows.Forms.TabPage tbTieBreaker;
		public System.Windows.Forms.TabPage tbEffective;
        private System.Windows.Forms.TabPage tbFilterCond;
        private TabPage tbCostsAndRevenue;
        public TabControl tabControlCosts;
        private TabPage tbCosts;
        private TabPage tbProcessorScenario;

        public const int TOTALCYCLES = 4;
		
		
		public bool m_bPopup = false;

        public FIA_Biosum_Manager.OptimizerScenarioItem m_oOptimizerScenarioItem = new OptimizerScenarioItem();
        public FIA_Biosum_Manager.OptimizerScenarioItem_Collection m_oOptimizerScenarioItem_Collection = new OptimizerScenarioItem_Collection();
        public FIA_Biosum_Manager.OptimizerScenarioTools m_oOptimizerScenarioTools = new OptimizerScenarioTools();
        private ToolBarButton btnScenarioProperties;
        private ToolBarButton tlbSeparator;
        private ToolBarButton btnScenarioCopy;
        public FIA_Biosum_Manager.OptimizerScenarioItem m_oSavOptimizerScenarioItem = new OptimizerScenarioItem();
        private Button btnHelp;
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultTreatmentOptimizerFile;
        private string m_helpChapter = "OPEN_SCENARIO";
        public FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.Variable_Collection m_oWeightedVariableCollection =
            new FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.Variable_Collection();
        private TabPage tbFilterPkg;
        private string m_strOutputTablePrefix;
        public bool m_bProcessorUsingSqlite = false;

        public frmOptimizerScenario(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			//
			// Required for Windows Form Designer support
			//

			FIA_Biosum_Manager.frmMain.g_oFrmMain=p_frmMain;

			InitializeComponent();
			
			if (this.Width > p_frmMain.Width - p_frmMain.grpboxLeft.Width ) 
			{
				this.Width = p_frmMain.Width - p_frmMain.grpboxLeft.Width  - 10;
			}
			try
			{
			
				this.uc_select_list_item1 = new uc_select_list_item();
				this.frmGridView1 = new frmGridView();
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
				uc_scenario1.ReferenceOptimizerScenarioForm=this;
				uc_scenario1.ScenarioType="optimizer";
				//
				//scenario datasource
				//
				this.uc_datasource1 = new uc_datasource();
				this.tbDataSources.Controls.Add(uc_datasource1);
				uc_datasource1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_datasource1.ReferenceOptimizerScenarioForm=this;
                uc_datasource1.ScenarioType = "optimizer";
				//
				//scenario notes
				//
				this.uc_scenario_notes1 = new uc_scenario_notes();
				this.tbNotes.Controls.Add(uc_scenario_notes1);
				uc_scenario_notes1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_scenario_notes1.ReferenceOptimizerScenarioForm=this;
				uc_scenario_notes1.ScenarioType="optimizer";
				//
				//rule definitions owner groups
				//
				this.uc_scenario_owner_groups1 = new uc_optimizer_scenario_owner_groups();
				this.tbOwners.Controls.Add(uc_scenario_owner_groups1);
				this.uc_scenario_owner_groups1.Dock = System.Windows.Forms.DockStyle.Fill;
				uc_scenario_owner_groups1.ReferenceOptimizerScenarioForm=this;
				//
				//rule definitions costs
				//
				this.uc_scenario_costs1 = new uc_optimizer_scenario_costs();
				this.tbCosts.Controls.Add(uc_scenario_costs1);
              
				this.uc_scenario_costs1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_costs1.ReferenceOptimizerScenarioForm=this;
				//
				//rule definitions processing sites
				//
				this.uc_scenario_psite1 = new uc_optimizer_scenario_psite();
				this.tbPSites.Controls.Add(uc_scenario_psite1);
				this.uc_scenario_psite1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_psite1.ReferenceOptimizerScenarioForm=this;
				//
				//rule custom plot filter
				//
				this.uc_scenario_filter1 = new uc_optimizer_scenario_filter();
				this.tbFilterPlots.Controls.Add(uc_scenario_filter1);
				this.uc_scenario_filter1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_filter1.ReferenceOptimizerScenarioForm=this;
				this.uc_scenario_filter1.FilterType="PLOT";
				//
				//rule custom condition filter
				//
				this.uc_scenario_cond_filter1 = new uc_optimizer_scenario_filter();
				this.tbFilterCond.Controls.Add(uc_scenario_cond_filter1);
				this.uc_scenario_cond_filter1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_cond_filter1.ReferenceOptimizerScenarioForm=this;
				this.uc_scenario_cond_filter1.lblTitle.Text = "Condition Filter";
				this.uc_scenario_cond_filter1.FilterType="COND";
				//
				//rule definitions fvs pre-post variables
				//
				this.uc_scenario_fvs_prepost_variables_effective1 = new uc_optimizer_scenario_fvs_prepost_variables_effective();
				this.tbEffective.Controls.Add(this.uc_scenario_fvs_prepost_variables_effective1);
				this.uc_scenario_fvs_prepost_variables_effective1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_fvs_prepost_variables_effective1.ReferenceOptimizerScenarioForm=this;
				//
				//rule definitions fvs optimization variables
				//
				this.uc_scenario_fvs_prepost_optimization1 = new uc_optimizer_scenario_fvs_prepost_optimization();
				this.tbOptimization.Controls.Add(this.uc_scenario_fvs_prepost_optimization1);
				this.uc_scenario_fvs_prepost_optimization1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_fvs_prepost_optimization1.ReferenceOptimizerScenarioForm = this;
				this.uc_scenario_fvs_prepost_variables_effective1.ReferenceOptimizationUserControl=this.uc_scenario_fvs_prepost_optimization1;

				//
				//rule definitions fvs tie breaker
				//
				this.uc_scenario_fvs_prepost_variables_tiebreaker1 = new FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker();
				this.tbTieBreaker.Controls.Add(this.uc_scenario_fvs_prepost_variables_tiebreaker1);
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.ReferenceOptimizerScenarioForm = this;
				this.uc_scenario_fvs_prepost_variables_effective1.ReferenceTieBreakerUserControl=this.uc_scenario_fvs_prepost_variables_tiebreaker1;
                this.uc_scenario_fvs_prepost_variables_tiebreaker1.uc_scenario_last_tiebreak_rank1.ReferenceOptimizerScenarioForm = this;
				this.uc_scenario_fvs_prepost_variables_tiebreaker1.ReferenceOptimizationUserControl=this.uc_scenario_fvs_prepost_optimization1;
				this.uc_scenario_fvs_prepost_optimization1.ReferenceTieBreaker=this.uc_scenario_fvs_prepost_variables_tiebreaker1;
                //
                //processor scenario select
                //
                this.uc_scenario_processor_scenario_select1 = new uc_optimizer_scenario_processor_scenario_select();
                this.tbProcessorScenario.Controls.Add(this.uc_scenario_processor_scenario_select1);
                this.uc_scenario_processor_scenario_select1.Dock = System.Windows.Forms.DockStyle.Fill;
                this.uc_scenario_processor_scenario_select1.ReferenceOptimizerScenarioForm = this;

                // Check to see if Processor is using Sqlite
                if (System.IO.File.Exists(
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()
                    + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile) )
                {
                    this.uc_scenario_processor_scenario_select1.ReferenceOptimizerScenarioForm.m_bProcessorUsingSqlite = true;
                }

                this.uc_optimizer_scenario_select_packages1 = new uc_optimizer_scenario_select_packages();
                this.tbFilterPkg.Controls.Add(uc_optimizer_scenario_select_packages1);
                this.uc_optimizer_scenario_select_packages1.Dock = System.Windows.Forms.DockStyle.Fill;
                this.uc_optimizer_scenario_select_packages1.ReferenceOptimizerScenarioForm = this;

				//rule execute run
				this.uc_scenario_run1 = new uc_optimizer_scenario_run();
				this.tbRun.Controls.Add(uc_scenario_run1);
				this.uc_scenario_run1.Dock = System.Windows.Forms.DockStyle.Fill;
				this.uc_scenario_run1.ReferenceOptimizerScenarioForm=this;
                this.uc_scenario_run1.ReferenceOptimizerScenarioForm = this;
                this.btnClose.Enabled=true;
                this.resize_frmOptimizerScenario();

                this.m_oEnv = new env();

                //load weighted variable definitions
                ado_data_access oAdo = new ado_data_access();
                m_oOptimizerScenarioTools.LoadWeightedVariables(oAdo, m_oWeightedVariableCollection);
                oAdo.m_OleDbDataReader.Close();
                oAdo.CloseConnection(oAdo.m_OleDbConnection);

			}
			catch (Exception p_msg)
			{
				MessageBox.Show(p_msg.Message);
			}


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

		public frmOptimizerScenario()
		{
			this.InitializeComponent();

		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmOptimizerScenario));
            this.tlbScenario = new System.Windows.Forms.ToolBar();
            this.btnScenarioNew = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioOpen = new System.Windows.Forms.ToolBarButton();
            this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioSave = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioDelete = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioProperties = new System.Windows.Forms.ToolBarButton();
            this.tlbSeparator = new System.Windows.Forms.ToolBarButton();
            this.btnScenarioCopy = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.txtDropDown = new System.Windows.Forms.TextBox();
            this.dataView1 = new System.Data.DataView();
            this.imgSize = new System.Windows.Forms.ImageList(this.components);
            this.imageList2 = new System.Windows.Forms.ImageList(this.components);
            this.btnClose = new System.Windows.Forms.Button();
            this.tabControlScenario = new System.Windows.Forms.TabControl();
            this.tbDesc = new System.Windows.Forms.TabPage();
            this.tbNotes = new System.Windows.Forms.TabPage();
            this.tbDataSources = new System.Windows.Forms.TabPage();
            this.tbRules = new System.Windows.Forms.TabPage();
            this.tabControlRules = new System.Windows.Forms.TabControl();
            this.tbOwners = new System.Windows.Forms.TabPage();
            this.tbCostsAndRevenue = new System.Windows.Forms.TabPage();
            this.tabControlCosts = new System.Windows.Forms.TabControl();
            this.tbCosts = new System.Windows.Forms.TabPage();
            this.tbProcessorScenario = new System.Windows.Forms.TabPage();
            this.tbPSites = new System.Windows.Forms.TabPage();
            this.tbFilterPlots = new System.Windows.Forms.TabPage();
            this.tbFilterCond = new System.Windows.Forms.TabPage();
            this.tbFVSVariables = new System.Windows.Forms.TabPage();
            this.tabControlFVSVariables = new System.Windows.Forms.TabControl();
            this.tbEffective = new System.Windows.Forms.TabPage();
            this.tbOptimization = new System.Windows.Forms.TabPage();
            this.tbTieBreaker = new System.Windows.Forms.TabPage();
            this.tbFilterPkg = new System.Windows.Forms.TabPage();
            this.tbRun = new System.Windows.Forms.TabPage();
            this.btnHelp = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataView1)).BeginInit();
            this.tabControlScenario.SuspendLayout();
            this.tbRules.SuspendLayout();
            this.tabControlRules.SuspendLayout();
            this.tbCostsAndRevenue.SuspendLayout();
            this.tabControlCosts.SuspendLayout();
            this.tbFVSVariables.SuspendLayout();
            this.tabControlFVSVariables.SuspendLayout();
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
            this.tlbSeparator,
            this.btnScenarioCopy});
            this.tlbScenario.DropDownArrows = true;
            this.tlbScenario.ImageList = this.imageList1;
            this.tlbScenario.Location = new System.Drawing.Point(0, 0);
            this.tlbScenario.Margin = new System.Windows.Forms.Padding(2);
            this.tlbScenario.Name = "tlbScenario";
            this.tlbScenario.ShowToolTips = true;
            this.tlbScenario.Size = new System.Drawing.Size(782, 47);
            this.tlbScenario.TabIndex = 42;
            this.tlbScenario.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.tlbScenario_ButtonClick);
            // 
            // btnScenarioNew
            // 
            this.btnScenarioNew.ImageIndex = 1;
            this.btnScenarioNew.Name = "btnScenarioNew";
            this.btnScenarioNew.Text = "New";
            this.btnScenarioNew.ToolTipText = "New";
            // 
            // btnScenarioOpen
            // 
            this.btnScenarioOpen.ImageIndex = 0;
            this.btnScenarioOpen.Name = "btnScenarioOpen";
            this.btnScenarioOpen.Text = "Open";
            this.btnScenarioOpen.ToolTipText = "Open";
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
            this.btnScenarioSave.ToolTipText = "Save";
            // 
            // btnScenarioDelete
            // 
            this.btnScenarioDelete.ImageIndex = 3;
            this.btnScenarioDelete.Name = "btnScenarioDelete";
            this.btnScenarioDelete.Text = "Delete";
            this.btnScenarioDelete.ToolTipText = "Delete";
            // 
            // btnScenarioProperties
            // 
            this.btnScenarioProperties.ImageIndex = 4;
            this.btnScenarioProperties.Name = "btnScenarioProperties";
            this.btnScenarioProperties.Text = "Properties";
            this.btnScenarioProperties.Visible = false;
            // 
            // tlbSeparator
            // 
            this.tlbSeparator.Name = "tlbSeparator";
            this.tlbSeparator.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // btnScenarioCopy
            // 
            this.btnScenarioCopy.ImageIndex = 5;
            this.btnScenarioCopy.Name = "btnScenarioCopy";
            this.btnScenarioCopy.Text = "Copy";
            this.btnScenarioCopy.ToolTipText = "Copy scenario values to this scenario";
            this.btnScenarioCopy.Visible = false;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "open-file-icon.png");
            this.imageList1.Images.SetKeyName(1, "");
            this.imageList1.Images.SetKeyName(2, "");
            this.imageList1.Images.SetKeyName(3, "deletered.png");
            this.imageList1.Images.SetKeyName(4, "properties.png");
            this.imageList1.Images.SetKeyName(5, "copy.bmp");
            // 
            // txtDropDown
            // 
            this.txtDropDown.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.txtDropDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDropDown.ForeColor = System.Drawing.SystemColors.ControlText;
            this.txtDropDown.HideSelection = false;
            this.txtDropDown.Location = new System.Drawing.Point(239, 405);
            this.txtDropDown.Margin = new System.Windows.Forms.Padding(2);
            this.txtDropDown.Multiline = true;
            this.txtDropDown.Name = "txtDropDown";
            this.txtDropDown.ReadOnly = true;
            this.txtDropDown.Size = new System.Drawing.Size(27, 25);
            this.txtDropDown.TabIndex = 12;
            this.txtDropDown.Visible = false;
            // 
            // imgSize
            // 
            this.imgSize.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgSize.ImageStream")));
            this.imgSize.TransparentColor = System.Drawing.Color.Transparent;
            this.imgSize.Images.SetKeyName(0, "");
            this.imgSize.Images.SetKeyName(1, "");
            // 
            // imageList2
            // 
            this.imageList2.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList2.ImageStream")));
            this.imageList2.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList2.Images.SetKeyName(0, "");
            // 
            // btnClose
            // 
            this.btnClose.Enabled = false;
            this.btnClose.Location = new System.Drawing.Point(597, 405);
            this.btnClose.Margin = new System.Windows.Forms.Padding(2);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(102, 31);
            this.btnClose.TabIndex = 40;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // tabControlScenario
            // 
            this.tabControlScenario.Controls.Add(this.tbDesc);
            this.tabControlScenario.Controls.Add(this.tbNotes);
            this.tabControlScenario.Controls.Add(this.tbDataSources);
            this.tabControlScenario.Controls.Add(this.tbRules);
            this.tabControlScenario.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlScenario.ItemSize = new System.Drawing.Size(100, 18);
            this.tabControlScenario.Location = new System.Drawing.Point(0, 31);
            this.tabControlScenario.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlScenario.Name = "tabControlScenario";
            this.tabControlScenario.SelectedIndex = 0;
            this.tabControlScenario.Size = new System.Drawing.Size(699, 351);
            this.tabControlScenario.TabIndex = 41;
            this.tabControlScenario.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlScenario_DrawItem);
            this.tabControlScenario.SelectedIndexChanged += new System.EventHandler(this.tabControlScenario_SelectedIndexChanged);
            this.tabControlScenario.Enter += new System.EventHandler(this.tabControlScenario_Enter);
            this.tabControlScenario.Leave += new System.EventHandler(this.tabControlScenario_Leave);
            this.tabControlScenario.ChangeUICues += new System.Windows.Forms.UICuesEventHandler(this.tabControlScenario_ChangeUICues);
            // 
            // tbDesc
            // 
            this.tbDesc.AutoScroll = true;
            this.tbDesc.Location = new System.Drawing.Point(4, 22);
            this.tbDesc.Margin = new System.Windows.Forms.Padding(2);
            this.tbDesc.Name = "tbDesc";
            this.tbDesc.Size = new System.Drawing.Size(691, 325);
            this.tbDesc.TabIndex = 0;
            this.tbDesc.Text = "Description";
            // 
            // tbNotes
            // 
            this.tbNotes.AutoScroll = true;
            this.tbNotes.Location = new System.Drawing.Point(4, 22);
            this.tbNotes.Margin = new System.Windows.Forms.Padding(2);
            this.tbNotes.Name = "tbNotes";
            this.tbNotes.Size = new System.Drawing.Size(683, 325);
            this.tbNotes.TabIndex = 1;
            this.tbNotes.Text = "Notes";
            // 
            // tbDataSources
            // 
            this.tbDataSources.AutoScroll = true;
            this.tbDataSources.Location = new System.Drawing.Point(4, 22);
            this.tbDataSources.Margin = new System.Windows.Forms.Padding(2);
            this.tbDataSources.Name = "tbDataSources";
            this.tbDataSources.Size = new System.Drawing.Size(683, 325);
            this.tbDataSources.TabIndex = 2;
            this.tbDataSources.Text = "Data Sources";
            // 
            // tbRules
            // 
            this.tbRules.Controls.Add(this.tabControlRules);
            this.tbRules.ForeColor = System.Drawing.Color.Red;
            this.tbRules.Location = new System.Drawing.Point(4, 22);
            this.tbRules.Margin = new System.Windows.Forms.Padding(2);
            this.tbRules.Name = "tbRules";
            this.tbRules.Size = new System.Drawing.Size(683, 325);
            this.tbRules.TabIndex = 3;
            this.tbRules.Text = "Rule Definitions";
            this.tbRules.Click += new System.EventHandler(this.tbRules_Click);
            // 
            // tabControlRules
            // 
            this.tabControlRules.Controls.Add(this.tbOwners);
            this.tabControlRules.Controls.Add(this.tbCostsAndRevenue);
            this.tabControlRules.Controls.Add(this.tbPSites);
            this.tabControlRules.Controls.Add(this.tbFilterPlots);
            this.tabControlRules.Controls.Add(this.tbFilterCond);
            this.tabControlRules.Controls.Add(this.tbFVSVariables);
            this.tabControlRules.Controls.Add(this.tbFilterPkg);
            this.tabControlRules.Controls.Add(this.tbRun);
            this.tabControlRules.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlRules.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlRules.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlRules.ItemSize = new System.Drawing.Size(125, 18);
            this.tabControlRules.Location = new System.Drawing.Point(0, 0);
            this.tabControlRules.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlRules.Name = "tabControlRules";
            this.tabControlRules.SelectedIndex = 0;
            this.tabControlRules.Size = new System.Drawing.Size(683, 325);
            this.tabControlRules.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControlRules.TabIndex = 0;
            this.tabControlRules.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlRules_DrawItem);
            this.tabControlRules.SelectedIndexChanged += new System.EventHandler(this.tabControlRules_SelectedIndexChanged);
            // 
            // tbOwners
            // 
            this.tbOwners.Location = new System.Drawing.Point(4, 22);
            this.tbOwners.Margin = new System.Windows.Forms.Padding(2);
            this.tbOwners.Name = "tbOwners";
            this.tbOwners.Size = new System.Drawing.Size(675, 299);
            this.tbOwners.TabIndex = 1;
            this.tbOwners.Text = "Land Ownership Groups";
            this.tbOwners.Click += new System.EventHandler(this.tbOwners_Click);
            // 
            // tbCostsAndRevenue
            // 
            this.tbCostsAndRevenue.Controls.Add(this.tabControlCosts);
            this.tbCostsAndRevenue.Location = new System.Drawing.Point(4, 22);
            this.tbCostsAndRevenue.Margin = new System.Windows.Forms.Padding(2);
            this.tbCostsAndRevenue.Name = "tbCostsAndRevenue";
            this.tbCostsAndRevenue.Size = new System.Drawing.Size(675, 299);
            this.tbCostsAndRevenue.TabIndex = 10;
            this.tbCostsAndRevenue.Text = "Cost And Revenue";
            this.tbCostsAndRevenue.UseVisualStyleBackColor = true;
            // 
            // tabControlCosts
            // 
            this.tabControlCosts.Controls.Add(this.tbCosts);
            this.tabControlCosts.Controls.Add(this.tbProcessorScenario);
            this.tabControlCosts.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlCosts.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlCosts.Location = new System.Drawing.Point(0, 0);
            this.tabControlCosts.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlCosts.Name = "tabControlCosts";
            this.tabControlCosts.SelectedIndex = 0;
            this.tabControlCosts.Size = new System.Drawing.Size(675, 299);
            this.tabControlCosts.TabIndex = 1;
            this.tabControlCosts.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControlCosts_DrawItem);
            this.tabControlCosts.SelectedIndexChanged += new System.EventHandler(this.tabControlCosts_SelectedIndexChanged);
            // 
            // tbCosts
            // 
            this.tbCosts.Location = new System.Drawing.Point(4, 26);
            this.tbCosts.Margin = new System.Windows.Forms.Padding(2);
            this.tbCosts.Name = "tbCosts";
            this.tbCosts.Padding = new System.Windows.Forms.Padding(2);
            this.tbCosts.Size = new System.Drawing.Size(667, 269);
            this.tbCosts.TabIndex = 0;
            this.tbCosts.Text = "Haul Costs";
            this.tbCosts.UseVisualStyleBackColor = true;
            // 
            // tbProcessorScenario
            // 
            this.tbProcessorScenario.Location = new System.Drawing.Point(4, 25);
            this.tbProcessorScenario.Margin = new System.Windows.Forms.Padding(2);
            this.tbProcessorScenario.Name = "tbProcessorScenario";
            this.tbProcessorScenario.Padding = new System.Windows.Forms.Padding(2);
            this.tbProcessorScenario.Size = new System.Drawing.Size(667, 270);
            this.tbProcessorScenario.TabIndex = 1;
            this.tbProcessorScenario.Text = "Processor Scenario";
            this.tbProcessorScenario.UseVisualStyleBackColor = true;
            // 
            // tbPSites
            // 
            this.tbPSites.Location = new System.Drawing.Point(4, 22);
            this.tbPSites.Margin = new System.Windows.Forms.Padding(2);
            this.tbPSites.Name = "tbPSites";
            this.tbPSites.Size = new System.Drawing.Size(675, 299);
            this.tbPSites.TabIndex = 4;
            this.tbPSites.Text = "Wood Processing Sites";
            this.tbPSites.Click += new System.EventHandler(this.tbPSites_Click);
            this.tbPSites.Enter += new System.EventHandler(this.tbPSites_Enter);
            // 
            // tbFilterPlots
            // 
            this.tbFilterPlots.Location = new System.Drawing.Point(4, 22);
            this.tbFilterPlots.Margin = new System.Windows.Forms.Padding(2);
            this.tbFilterPlots.Name = "tbFilterPlots";
            this.tbFilterPlots.Size = new System.Drawing.Size(675, 299);
            this.tbFilterPlots.TabIndex = 5;
            this.tbFilterPlots.Text = "Filter Plot Records";
            // 
            // tbFilterCond
            // 
            this.tbFilterCond.Location = new System.Drawing.Point(4, 22);
            this.tbFilterCond.Margin = new System.Windows.Forms.Padding(2);
            this.tbFilterCond.Name = "tbFilterCond";
            this.tbFilterCond.Size = new System.Drawing.Size(675, 299);
            this.tbFilterCond.TabIndex = 9;
            this.tbFilterCond.Text = "Filter Condition Records";
            // 
            // tbFVSVariables
            // 
            this.tbFVSVariables.Controls.Add(this.tabControlFVSVariables);
            this.tbFVSVariables.Location = new System.Drawing.Point(4, 22);
            this.tbFVSVariables.Margin = new System.Windows.Forms.Padding(2);
            this.tbFVSVariables.Name = "tbFVSVariables";
            this.tbFVSVariables.Size = new System.Drawing.Size(675, 299);
            this.tbFVSVariables.TabIndex = 8;
            this.tbFVSVariables.Text = "FVS Variables";
            // 
            // tabControlFVSVariables
            // 
            this.tabControlFVSVariables.Controls.Add(this.tbEffective);
            this.tabControlFVSVariables.Controls.Add(this.tbOptimization);
            this.tabControlFVSVariables.Controls.Add(this.tbTieBreaker);
            this.tabControlFVSVariables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlFVSVariables.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControlFVSVariables.ItemSize = new System.Drawing.Size(75, 18);
            this.tabControlFVSVariables.Location = new System.Drawing.Point(0, 0);
            this.tabControlFVSVariables.Margin = new System.Windows.Forms.Padding(2);
            this.tabControlFVSVariables.Name = "tabControlFVSVariables";
            this.tabControlFVSVariables.SelectedIndex = 0;
            this.tabControlFVSVariables.Size = new System.Drawing.Size(675, 299);
            this.tabControlFVSVariables.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControlFVSVariables.TabIndex = 0;
            this.tabControlFVSVariables.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tbFVSVariablesSelect_DrawItem);
            this.tabControlFVSVariables.SelectedIndexChanged += new System.EventHandler(this.tbFVSVariablesSelect_SelectedIndexChanged);
            // 
            // tbEffective
            // 
            this.tbEffective.Location = new System.Drawing.Point(4, 22);
            this.tbEffective.Margin = new System.Windows.Forms.Padding(2);
            this.tbEffective.Name = "tbEffective";
            this.tbEffective.Size = new System.Drawing.Size(667, 273);
            this.tbEffective.TabIndex = 0;
            this.tbEffective.Text = "Effective";
            // 
            // tbOptimization
            // 
            this.tbOptimization.Location = new System.Drawing.Point(4, 22);
            this.tbOptimization.Margin = new System.Windows.Forms.Padding(2);
            this.tbOptimization.Name = "tbOptimization";
            this.tbOptimization.Size = new System.Drawing.Size(667, 273);
            this.tbOptimization.TabIndex = 1;
            this.tbOptimization.Text = "Optimization";
            // 
            // tbTieBreaker
            // 
            this.tbTieBreaker.Location = new System.Drawing.Point(4, 22);
            this.tbTieBreaker.Margin = new System.Windows.Forms.Padding(2);
            this.tbTieBreaker.Name = "tbTieBreaker";
            this.tbTieBreaker.Size = new System.Drawing.Size(667, 273);
            this.tbTieBreaker.TabIndex = 2;
            this.tbTieBreaker.Text = "Tie Breaker";
            // 
            // tbFilterPkg
            // 
            this.tbFilterPkg.Location = new System.Drawing.Point(4, 22);
            this.tbFilterPkg.Margin = new System.Windows.Forms.Padding(2);
            this.tbFilterPkg.Name = "tbFilterPkg";
            this.tbFilterPkg.Size = new System.Drawing.Size(675, 299);
            this.tbFilterPkg.TabIndex = 11;
            this.tbFilterPkg.Text = "Filter RxPackage";
            this.tbFilterPkg.UseVisualStyleBackColor = true;
            // 
            // tbRun
            // 
            this.tbRun.Location = new System.Drawing.Point(4, 22);
            this.tbRun.Margin = new System.Windows.Forms.Padding(2);
            this.tbRun.Name = "tbRun";
            this.tbRun.Size = new System.Drawing.Size(675, 299);
            this.tbRun.TabIndex = 6;
            this.tbRun.Text = "Run";
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(9, 405);
            this.btnHelp.Margin = new System.Windows.Forms.Padding(2);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(103, 31);
            this.btnHelp.TabIndex = 48;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // frmOptimizerScenario
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(782, 603);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.tlbScenario);
            this.Controls.Add(this.tabControlScenario);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.txtDropDown);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmOptimizerScenario";
            this.Text = "Optimizer Scenario";
            this.Activated += new System.EventHandler(this.frmOptimizerScenario_Activated);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmOptimizerScenario_Closing);
            this.Load += new System.EventHandler(this.frmOptimizerScenario_Load);
            this.Click += new System.EventHandler(this.frmOptimizerScenario_Click);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.frmOptimizerScenario_MouseDown);
            this.Resize += new System.EventHandler(this.frmOptimizerScenario_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataView1)).EndInit();
            this.tabControlScenario.ResumeLayout(false);
            this.tbRules.ResumeLayout(false);
            this.tabControlRules.ResumeLayout(false);
            this.tbCostsAndRevenue.ResumeLayout(false);
            this.tabControlCosts.ResumeLayout(false);
            this.tbFVSVariables.ResumeLayout(false);
            this.tabControlFVSVariables.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

   public string OutputTablePrefix
   {
       get { return m_strOutputTablePrefix; }
       set { m_strOutputTablePrefix = value; }
   }

		private void frmOptimizerScenario_Load(object sender, System.EventArgs e)
		{
			this.resize_frmOptimizerScenario();
		}

		private void btnCurrentscenario_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnCurrentscenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{

		}

		private void btnCurrentscenario_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			this.txtDropDown.Visible=false;
		}

		private void btnSize_Click(object sender, System.EventArgs e)
		{
		}

		private void frmOptimizerScenario_Click(object sender, System.EventArgs e)
		{
			this.Focus();
		}

		private void frmOptimizerScenario_Resize(object sender, System.EventArgs e)
		{
			try
			{
				this.resize_frmOptimizerScenario();
			}
			catch
			{
			}
		}
		public void resize_frmOptimizerScenario()
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
                this.btnHelp.Top = btnClose.Top;
				this.tabControlScenario.Top = this.tlbScenario.Top + this.tlbScenario.Height + 2;
				this.tabControlScenario.Height = this.btnClose.Top - this.tabControlScenario.Top - 2;
				this.tabControlScenario.Width = this.ClientSize.Width;
				
			}
			catch
			{
			}
			finally 
			{
			}

		}

		private void btnDataSources_Click(object sender, System.EventArgs e)
		{
		}

		private void mnuFile_Click(object sender, System.EventArgs e)
		{
		
		}

		

		private void btnDataSources_Click_1(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxLeft_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnFile2_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			
		
		}

		public void InitializeNewScenario()
		{
			this.uc_scenario1 = new uc_scenario();

			this.Controls.Add(uc_scenario1);

			uc_scenario1.Dock = System.Windows.Forms.DockStyle.Fill;

			this.tlbScenario.Hide();

			this.tabControlScenario.Hide();

			this.btnClose.Hide();

			this.uc_scenario1.ReferenceOptimizerScenarioForm=this;

			this.uc_scenario1.ScenarioType="optimizer";

           

			this.uc_scenario1.NewScenario();

			


		}

		public void InitializeOpenScenario(string strDebugFile)
		{
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "=====================   InitializeOpenScenario   =====================\r\n");
            this.uc_scenario_open1 = new uc_scenario_open(); 

			this.Controls.Add(uc_scenario_open1);


			uc_scenario_open1.Left = 0;
			uc_scenario_open1.Width = this.ClientSize.Width;
			uc_scenario_open1.Top = this.tlbScenario.Top + this.tlbScenario.Height;
    		
			this.btnScenarioSave.Enabled=false;
			this.btnScenarioOpen.Enabled=false;

			this.tabControlScenario.Hide();

			this.btnClose.Hide();

			this.uc_scenario_open1.OpenScenario(strDebugFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(strDebugFile, "InitializeOpenScenario: uc_scenario_open.OpenScenario() successful! \r\n");

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

		public void SetMenu(string strType)
		{
			switch (strType)
			{
				case "default" :
					break;
				case "new":
					break;
				case "open":
					break;
				case "scenario":
					break;
			}
		}
		private void btnOpen_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
		
		}

		private void btnCurrentScenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			string strTemp = this.uc_scenario1.txtDescription.Text.Trim();
			int lines = (int) (strTemp.Length / 30);  //40 characters per line
			if (lines == 0) lines = 1;

			this.txtDropDown.Text = strTemp;
			int textWidth = (int)this.CreateGraphics().MeasureString(strTemp, this.txtDropDown.Font).Width;
			int textHeight = (int)this.CreateGraphics().MeasureString(strTemp, this.txtDropDown.Font).Height;
			this.txtDropDown.Width = (int) ((textWidth / lines) * 1.5);
			this.txtDropDown.Height = (int)Math.Round((textHeight * lines) * 1.5,0) ;
			this.txtDropDown.BringToFront();
			this.txtDropDown.Visible = true;

		}

		private void btnCurrentScenario_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			this.txtDropDown.Visible=false;
		}

		private void btnOpen_Click(object sender, System.EventArgs e)
		{
			int intAvailWd = this.ParentForm.ClientSize.Width - ((frmMain)this.ParentForm).grpboxLeft.Left - ((frmMain)this.ParentForm).grpboxLeft.Width - 20;
			int intAvailHt = this.ParentForm.ClientSize.Height - ((frmMain)this.ParentForm).tlbMain.Top - ((frmMain)this.ParentForm).tlbMain.Height - 20;

			
			if (this.m_bScenarioOpen == false) 
			{
				this.uc_datasource1.Visible=false;
				this.uc_scenario_notes1.Visible=false;
				this.uc_scenario_owner_groups1.Visible=false;
				this.uc_scenario_costs1.Visible=false;
				this.uc_scenario_psite1.Visible=false;
				this.uc_scenario_filter1.Visible=false;
				this.m_vScrollBar.Visible=false;
				this.m_hScrollBar.Visible=false;


				this.Height = intAvailHt; 
				this.Width = intAvailWd; 
				this.Left = 0;
				this.Top = 0;
				this.uc_scenario1.lblTitle.Text = "Open Scenario";

				this.uc_scenario1.txtScenarioId.Visible=false;
				this.uc_scenario1.lblNewScenario.Visible=false;
				this.uc_scenario1.resize_uc_scenario();
				this.uc_scenario1.Visible=true;
		
			}
			else 
			{
				frmOptimizerScenario frmTemp = new frmOptimizerScenario((frmMain)this.ParentForm);
				frmTemp.m_vScrollBar.Visible=false;
				frmTemp.m_hScrollBar.Visible=false;
				frmTemp.BackColor = System.Drawing.SystemColors.Control;
				frmTemp.Text = "Treatment Optimizer: Optimization Scenario";
				frmTemp.MdiParent = this.ParentForm;

				frmTemp.SetMenu("new");
				frmTemp.uc_scenario1.lblTitle.Text = "Open Scenario";
				frmTemp.uc_scenario1.lblTitle.Width =
					(int)this.CreateGraphics().MeasureString(frmTemp.uc_scenario1.lblTitle.Text,frmTemp.uc_scenario1.lblTitle.Font).Width;
				frmTemp.uc_scenario1.txtScenarioId.Visible=false;
				frmTemp.uc_scenario1.lblNewScenario.Visible=false;
				frmTemp.uc_scenario1.resize_uc_scenario();
				frmTemp.uc_scenario1.Visible=true;
				frmTemp.Show();
				frmTemp.Height = intAvailHt; 
				frmTemp.Width = intAvailWd; 
				frmTemp.Left = 0;
				frmTemp.Top = 0;
			}

		}

		private void contextMenuView_Popup(object sender, System.EventArgs e)
		{
			m_bPopup=true;
		}

		private void frmOptimizerScenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (m_bPopup == true) m_bPopup =false;
		}

		private void btnSave_Click(object sender, System.EventArgs e)
		{
			SaveRuleDefinitions();


		}
		private void LoadRuleDefinitions()
		{
            int x;
			frmTherm p_frmTherm = new frmTherm();
			p_frmTherm.lblMsg.Text = "";
			p_frmTherm.progressBar1.Minimum = 0;
			p_frmTherm.progressBar1.Maximum = 7;
			p_frmTherm.btnCancel.Visible=false;
			p_frmTherm.lblMsg.Visible=true;
			p_frmTherm.Show();
			p_frmTherm.progressBar1.Value=1;
			p_frmTherm.progressBar1.Value=2;
            Queries oQueries = new Queries();
            
            this.m_oOptimizerScenarioItem.ScenarioId = this.uc_scenario1.txtScenarioId.Text.Trim();

            this.m_oOptimizerScenarioTools.LoadAll(
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + 
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile,
                oQueries, m_bProcessorUsingSqlite, m_oOptimizerScenarioItem.ScenarioId,
                m_oOptimizerScenarioItem_Collection);
            //find the current scenario
            for (x = 0; x <= m_oOptimizerScenarioItem_Collection.Count - 1; x++)
            {
                if (m_oOptimizerScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    m_oOptimizerScenarioItem.ScenarioId.Trim().ToUpper())
                    break;
            }

            this.m_oOptimizerScenarioItem.Copy(m_oOptimizerScenarioItem_Collection.Item(x), m_oSavOptimizerScenarioItem);
            this.m_oOptimizerScenarioItem = m_oOptimizerScenarioItem_Collection.Item(x);

			p_frmTherm.progressBar1.Value=3;
			p_frmTherm.lblMsg.Text = "Rule Definitions: FVS Variables Data";
			p_frmTherm.lblMsg.Refresh();

            this.uc_scenario_fvs_prepost_variables_effective1.loadvalues();
 			p_frmTherm.progressBar1.Value=4;


			p_frmTherm.lblMsg.Text = "Rule Definitions: Owner Group Data";
			p_frmTherm.lblMsg.Refresh();

			this.uc_scenario_owner_groups1.loadvalues();
			p_frmTherm.progressBar1.Value=5;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Cost And Revenue Data";
			p_frmTherm.lblMsg.Refresh();
			this.uc_scenario_costs1.loadvalues();
            this.uc_scenario_processor_scenario_select1.loadvalues(false);
			p_frmTherm.progressBar1.Value=6;
			p_frmTherm.lblMsg.Text = "Rule Definitions: Plot Filter Data";
			p_frmTherm.lblMsg.Refresh();
			this.uc_scenario_filter1.loadvalues(false);
			this.uc_scenario_cond_filter1.loadvalues(false);
			this.uc_scenario_psite1.loadvalues();
            ProcessorScenarioItem_Collection oProcItemCollection = this.m_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection;
            foreach (ProcessorScenarioItem psItem in oProcItemCollection)
            {
                if (psItem.Selected == true)
                {
                    this.uc_optimizer_scenario_select_packages1.loadvalues_FromProperties(psItem);
                }
            }
			p_frmTherm.progressBar1.Value=7;
			p_frmTherm.Close();
			p_frmTherm = null;
			this.m_lrulesfirsttime=false;

		}
		public void SaveRuleDefinitions()
		{
			int savestatus;
            int x;

			frmMain.g_sbpInfo.Text = "Saving scenario rule definitions...Stand by";
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMaximumSteps = 12;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicMinimumSteps = 1;
            FIA_Biosum_Manager.RunOptimizer.g_intCurrentProgressBarBasicCurrentStep = 1;

            if (m_lrulesfirsttime == false)
            {
                frmMain.g_oFrmMain.ActivateStandByAnimation(
                    this.WindowState,
                    this.Left,
                    this.Height,
                    this.Width,
                    this.Top);
                
                savestatus = this.uc_scenario_fvs_prepost_variables_effective1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_fvs_prepost_optimization1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_fvs_prepost_variables_tiebreaker1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_owner_groups1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer .g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_costs1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_psite1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_filter1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                savestatus = this.uc_scenario_cond_filter1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
                this.uc_scenario_processor_scenario_select1.savevalues();
                if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();

                frmMain.g_oFrmMain.DeactivateStandByAnimation();
            }
			this.uc_scenario_notes1.SaveScenarioNotes();
            if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
			this.uc_scenario1.UpdateDescription();
            if (FIA_Biosum_Manager.RunOptimizer.g_bOptimizerRun) FIA_Biosum_Manager.uc_optimizer_scenario_run.UpdateThermPercent();
            
			this.m_bSave=false;
			frmMain.g_sbpInfo.Text = "Ready";
		}
		

		private void btnClose_Click(object sender, System.EventArgs e)
		{
            // CheckToSave() method is called from frmScenario_Closing event
            this.Close();
		}
		private DialogResult CheckToSave()
		{
            DialogResult result=DialogResult.No;
			if (m_bSave)
			{
				result = MessageBox.Show("Save Scenario Changes Y/N","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNoCancel,System.Windows.Forms.MessageBoxIcon.Question);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    this.SaveRuleDefinitions();
                }
			}
            return result;
		}

        public string HelpChapter
        {
            get 
            {
                if (tabControlRules.SelectedTab.Text.Trim().ToUpper() == "FVS VARIABLES")
                {
                    switch (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper())
                    {
                        case "OPTIMIZATION":
                            m_helpChapter = this.uc_scenario_fvs_prepost_optimization1.HelpChapter;
                            break;
                        case "EFFECTIVE":
                            m_helpChapter = this.uc_scenario_fvs_prepost_variables_effective1.HelpChapter;
                            break;
                        case "TIE BREAKER":
                            m_helpChapter = this.uc_scenario_fvs_prepost_variables_tiebreaker1.HelpChapter;
                            break;
                    }
                }
                return m_helpChapter;
            }
            set { this.m_helpChapter = value; }
        }

		private void vScrollBar_Scroll(Object sender, ScrollEventArgs e)
		{
		}

		private void vScrollBar_MouseWheel(Object sender, MouseEventArgs e)
		{
		}

		public void HScrollRepositionControls()
		{
			try
			{
				int intScroll;
				double dblVal = Convert.ToDouble(this.m_hScrollBar.Value);
				double dblMax = Convert.ToDouble(this.m_hScrollBar.Maximum);

				if (this.m_hScrollBar.Value == 0)
				{
					this.m_dblHScrollNewPerc = 0;
				}
				else
				{
					this.m_dblHScrollNewPerc = (dblVal / dblMax);
				}

				if (this.m_dblHScrollNewPerc == 0 && this.m_dblHScrollOldPerc  > 0)
				{
					intScroll =  (int)(this.m_intHScrollMaxSize * this.m_dblHScrollOldPerc);

				}
				else if (this.m_dblHScrollNewPerc > 0 && this.m_dblHScrollOldPerc == 0)
				{
					intScroll = (int)(this.m_intHScrollMaxSize * this.m_dblHScrollNewPerc);
					intScroll = -1 * intScroll;
				}
				else if (this.m_dblHScrollNewPerc > 0 && this.m_dblHScrollOldPerc > 0)
				{
				
					intScroll = (int)(this.m_intHScrollMaxSize * this.m_dblHScrollNewPerc) - (int)(this.m_intHScrollMaxSize * this.m_dblHScrollOldPerc);
					intScroll = -1 * intScroll;
			
				}
				else
				{
					intScroll = 0;
				}
				for (int z=0;z<=this.Controls.Count-1;z++)
				{
					this.Controls[z].Left = this.Controls[z].Left + intScroll;
				}
				this.m_dblHScrollOldPerc = this.m_dblHScrollNewPerc;
			}
			catch
			{
			}
		}
		
		private void button1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("thisHeight=" + this.Height.ToString() + " this.m_vScrollBar.Value=" + this.m_vScrollBar.Value.ToString() + " this.m_vScrollBar.Maximum=" + this.m_vScrollBar.Maximum.ToString());
		}
		private void grpboxMenu_Paint(object sender, PaintEventArgs e) 
		{
		}

		private void btnRun_Click(object sender, System.EventArgs e)
		{
			int intAvailWd = this.ParentForm.ClientSize.Width - ((frmMain)this.ParentForm).grpboxLeft.Left - ((frmMain)this.ParentForm).grpboxLeft.Width - 20;
			int intAvailHt = this.ParentForm.ClientSize.Height - ((frmMain)this.ParentForm).tlbMain.Top - ((frmMain)this.ParentForm).tlbMain.Height - 20;
			FIA_Biosum_Manager.utils p_oUtils = new FIA_Biosum_Manager.utils();
			if (p_oUtils.FindWindowLike((IntPtr)((frmMain)this.ParentForm).Handle, "Treatment Optimizer: Run Scenario (" +  this.uc_scenario1.txtScenarioId.Text.Trim() + ")","*",true,true) > 0)
			{
				this.frmRunOptimizerScenario1.WindowState = System.Windows.Forms.FormWindowState.Normal;
				this.frmRunOptimizerScenario1.Focus();
				this.frmRunOptimizerScenario1.Height = intAvailHt; 
				this.frmRunOptimizerScenario1.Width = intAvailWd; 
				this.frmRunOptimizerScenario1.Left = 0;
				this.frmRunOptimizerScenario1.Top = 0;
				return;

			}
			this.frmRunOptimizerScenario1 = new frmRunOptimizerScenario(this);
			this.frmRunOptimizerScenario1.Text = "Treatment Optimizer: Run Scenario (" + this.uc_scenario1.txtScenarioId.Text.Trim() + ")";
			this.frmRunOptimizerScenario1.MdiParent = this.ParentForm;
			this.frmRunOptimizerScenario1.btnCancel.Text = "Start";
			this.frmRunOptimizerScenario1.WindowState = System.Windows.Forms.FormWindowState.Normal;
			this.frmRunOptimizerScenario1.Enabled=true;
			this.frmRunOptimizerScenario1.Show();
			
			
			this.frmRunOptimizerScenario1.Height = intAvailHt; 
			this.frmRunOptimizerScenario1.Width = intAvailWd; 
			this.frmRunOptimizerScenario1.Left = 0;
			this.frmRunOptimizerScenario1.Top = 0;
		}
		private void hScrollBar_ValueChanged(Object sender, EventArgs e)
		{
			try
			{
				int myValue = ((HScrollBar)sender).Value;
				this.HScrollRepositionControls();
				this.m_intHScrollValue = myValue;
			}
			catch
			{
			}

		}

		private void uc_scenario_notes1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void uc_scenario_notes2_Load(object sender, System.EventArgs e)
		{
		
		}

		private void tabControlRules_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			switch (tabControlRules.SelectedTab.Text.Trim().ToUpper())
			{
				case "WOOD PROCESSING SITES":
                    m_helpChapter = "WOOD_PROCESSING_SITES";
                    btnHelp.Enabled = true;
					if (((Control)this.tbPSites).Enabled)
					    this.uc_scenario_psite1.lblTitle.Text = "Wood Processing Sites";
					else
						this.uc_scenario_psite1.lblTitle.Text = "Wood Processing Sites (Read Only)";

					break;
				case "COST AND REVENUE":
                    if (tabControlCosts.SelectedTab.Text.Trim().ToUpper() == "HAUL COSTS")
                    {
                        m_helpChapter = "HAUL_COSTS";
                    }
                    else
                    {
                        m_helpChapter = "PROCESSOR_SCENARIO";
                    }
                    btnHelp.Enabled = true;
					if (((Control)this.tbCosts).Enabled)
						this.uc_scenario_costs1.lblTitle.Text = "Cost And Revenue";
					else
						this.uc_scenario_costs1.lblTitle.Text = "Cost And Revenue (Read Only)";
					break;
				case "FILTER PLOT RECORDS":
                    m_helpChapter = "FILTER_PLOT";
                    btnHelp.Enabled = true;
					if (((Control)this.tbFilterPlots).Enabled)
						this.uc_scenario_filter1.lblTitle.Text = "Plot Filter";
					else
						this.uc_scenario_filter1.lblTitle.Text = "Plot Filter (Read Only)";
					break;
                case "FILTER CONDITION RECORDS":
                    m_helpChapter = "FILTER_CONDITION";
                    btnHelp.Enabled = true;
                    if (((Control)this.tbFilterPlots).Enabled)
                        this.uc_scenario_filter1.lblTitle.Text = "Condition Filter";
                    else
                        this.uc_scenario_filter1.lblTitle.Text = "Condition Filter (Read Only)";
                    break;
				case "LAND OWNERSHIP GROUPS":
                    m_helpChapter = "LAND_OWNERSHIP_GROUPS";
                    btnHelp.Enabled = true;
					if (((Control)this.tbOwners).Enabled)
						this.uc_scenario_owner_groups1.lblTitle.Text = "Owner Groups";
					else
						this.uc_scenario_owner_groups1.lblTitle.Text = "Owner Groups (Read Only)";
					break;
                case "FVS VARIABLES":
                    // This logic is in the HelpChapter getter
                    btnHelp.Enabled = true;
                    break;
                case "FILTER RXPACKAGE":
                    btnHelp.Enabled = false;
                    if (((Control)this.tbFilterPkg).Enabled)
						this.uc_scenario_run1.lblTitle.Text = "Run";
					else
						this.uc_scenario_run1.lblTitle.Text = "Run (Read Only)";
					break;
				case "RUN":
                    m_helpChapter = "RUN";
                    btnHelp.Enabled = true;
					if (((Control)this.tbRun).Enabled)
						this.uc_scenario_run1.lblTitle.Text = "Run";
					else
						this.uc_scenario_run1.lblTitle.Text = "Run (Read Only)";
					break;



			}
		}

		private void tlbScenario_ButtonClick(object sender, System.Windows.Forms.ToolBarButtonClickEventArgs e)
		{
			switch (e.Button.ImageIndex)
			{
					//
					//open scenario
					//
				case 0:
					frmMain.g_oFrmMain.OpenOptimizerScenario("Open", this);
					break;
					//
					//new scenario
					//
				case 1:
					frmMain.g_oFrmMain.OpenOptimizerScenario("New", this);
					break;
				case 2:
					this.SaveRuleDefinitions();
					break;
				case 3:
                    if (this.uc_scenario_open1 != null)
					{
                        if(uc_scenario1.DeleteScenario(uc_scenario_open1.txtScenarioId.Text.Trim()))
						    uc_scenario_open1.lstScenario.Items.Remove(uc_scenario_open1.lstScenario.SelectedItems[0]);
					}
					else
					{
						if (uc_scenario1.DeleteScenario(uc_scenario1.txtScenarioId.Text.Trim()))
							this.Close();
					}
                    break;
                case 4:
                    if (this.m_lrulesfirsttime == true)
                        LoadRuleDefinitions();
                    else
                    {
                        SaveRuleDefinitions();
                       
                    }
                    FIA_Biosum_Manager.frmDialog frmTemp = new frmDialog();
                    frmTemp.Text = "FIA Biosum";
                    frmTemp.AutoScroll = false;
                    uc_textbox uc_textbox1 = new uc_textbox();
                    frmTemp.Controls.Add(uc_textbox1);
                    uc_textbox1.Dock = DockStyle.Fill;
                    uc_textbox1.lblTitle.Text = "Properties";
                    uc_textbox1.TextValue = m_oOptimizerScenarioTools.ScenarioProperties(m_oOptimizerScenarioItem);
                    frmTemp.Show();
                    break;
                case 5:
                    CopyScenario();
                    break;
			}
		}
        private void CopyScenario()
        {
            frmDialog frmTemp = new frmDialog();
            frmTemp.Initialize_Scenario_Optimizer_Scenario_Copy(this);
            frmTemp.Text = "FIA Biosum";
            if (m_oOptimizerScenarioItem.ScenarioId.Trim().Length == 0) LoadRuleDefinitions();

            frmTemp.uc_scenario_optimizer_scenario_copy1.ReferenceCurrentScenarioItem = m_oOptimizerScenarioItem;
            frmTemp.uc_scenario_optimizer_scenario_copy1.loadvalues();
            DialogResult result = frmTemp.ShowDialog(this);
            if (result == DialogResult.OK)
            {

                frmMain.g_sbpInfo.Text = "Copying scenario rule definitions...Stand by";

                this.uc_scenario_fvs_prepost_variables_effective1.loadvalues_FromProperties();

                this.uc_scenario1.txtDescription.Text = m_oOptimizerScenarioItem.Description;
                frmMain.g_sbpInfo.Text = "Loading Scenario Notes...Stand By";
                this.uc_scenario_notes1.ReferenceOptimizerScenarioForm = this;
                this.uc_scenario_notes1.loadvalues_FromProperties();
                this.uc_scenario_owner_groups1.loadvalues();
                
                
                this.uc_scenario_costs1.loadvalues();
                this.uc_scenario_processor_scenario_select1.loadvalues(true);
                
                
                this.uc_scenario_filter1.loadvalues(true);
                this.uc_scenario_cond_filter1.loadvalues(true);
                this.uc_scenario_psite1.loadvalues_FromProperties();
                ProcessorScenarioItem_Collection oProcItemCollection = this.m_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection;
                //if (oProcItemCollection != null && oProcItemCollection.Count > 0)
                //{
                //    ProcessorScenarioItem oProcItem = oProcItemCollection.Item(0);
                //    this.uc_optimizer_scenario_select_packages1.loadvalues_FromProperties(oProcItem);
                //}
                foreach (ProcessorScenarioItem psItem in oProcItemCollection)
                {
                    if (psItem.Selected == true)
                    {
                        this.uc_optimizer_scenario_select_packages1.loadvalues_FromProperties(psItem);
                    }
                }
                
                frmMain.g_sbpInfo.Text = "Ready";
                m_bSave = true;
            }
           
        }
		private void tbRules_Click(object sender, System.EventArgs e)
		{
		}

		private void tabControlScenario_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.tabControlRules.Enabled)
			{
                btnHelp.Visible = true;
                m_helpChapter = "LAND_OWNERSHIP_GROUPS";
                if (tabControlScenario.SelectedTab.Text.Trim().ToUpper()=="RULE DEFINITIONS")
				{
					if (m_lrulesfirsttime==true)
						LoadRuleDefinitions();

				}
				else if (tabControlScenario.SelectedTab.Text.Trim().ToUpper() == "NOTES")
				{
                    btnHelp.Visible = true;
                    m_helpChapter = "EDIT_SCENARIO";
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
                    btnHelp.Visible = true;
                    m_helpChapter = "EDIT_SCENARIO";
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
                    btnHelp.Visible = false;
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

		private void tbPSites_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void tbPSites_Click(object sender, System.EventArgs e)
		{
		
		}

		private void tbOwners_Click(object sender, System.EventArgs e)
		{
		
		}

		private void tbFVSVariablesSelect_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			try
			{
				if (tabControlFVSVariables.Enabled)
				{
					if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper()=="OPTIMIZATION")
					{
							if (((Control)this.tbOptimization).Enabled)
							{
						
								uc_scenario_fvs_prepost_optimization1.lblTitle.Text = "Optimization Settings";
							}
							else
								uc_scenario_fvs_prepost_optimization1.lblTitle.Text = "Optimization Settings (Read Only)";

					}
					else if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper()=="EFFECTIVE")
					{
						if (((Control)this.tbEffective).Enabled)
						{
						    uc_scenario_fvs_prepost_variables_effective1.lblTitle.Text  = "Effective Settings";
						}
						else
							uc_scenario_fvs_prepost_variables_effective1.lblTitle.Text = "Effective Settings (Read Only)";

					}
					else if (tabControlFVSVariables.SelectedTab.Text.Trim().ToUpper()=="TIE BREAKER")
					{
						if (((Control)this.tbTieBreaker).Enabled)
						{
							uc_scenario_fvs_prepost_variables_tiebreaker1.lblTitle.Text  = "Tie Breaker Settings";
						}
						else
							uc_scenario_fvs_prepost_variables_tiebreaker1.lblTitle.Text = "Tie Breaker Settings (Read Only)";

					}
                    m_helpChapter = this.HelpChapter;
					
				}
				
			}
			catch (Exception err)
			{
				MessageBox.Show(err.Message);
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
			
				//This construct will help you to deside which tab page have current focus
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

		private void tabControlRules_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.DarkBlue,System.Drawing.Brushes.White);
		}

		private void tbFVSVariablesSelect_DrawItem(object sender, System.Windows.Forms.DrawItemEventArgs e)
		{
			TabControl_DrawItem(sender,e,Color.DarkRed,System.Drawing.Brushes.White);
		}

		private void tabControlScenario_Leave(object sender, System.EventArgs e)
		{
			
		}

		private void tabControlScenario_Enter(object sender, System.EventArgs e)
		{
			
		}

		private void tabControlScenario_ChangeUICues(object sender, System.Windows.Forms.UICuesEventArgs e)
		{
		
		}
		public void EnableTabPage(System.Windows.Forms.TabPage p_oTabPage,bool p_bEnable)
		{
				((Control)p_oTabPage).Enabled = p_bEnable;	
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

		private void frmOptimizerScenario_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (frmMain.g_oDelegate.CurrentThreadProcessIdle==false)
			{
				e.Cancel = true;
				return;
			}
            if (frmMain.g_oFrmMain.tlbMain.Enabled == false)
            {
                MessageBox.Show("!!Finish Edit Session Before Closing!!", "FIA Biosum");
                e.Cancel = true;
            }
            else
            {
                DialogResult result = CheckToSave();
                switch (result)
                {
                    case (DialogResult.Cancel):
                        e.Cancel = true;
                        break;
                    default:
                        break;

                }
            }
		}

        private void tabControlCosts_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabControl_DrawItem(sender, e, Color.DarkRed, System.Drawing.Brushes.White);
        }

        private void tabControlCosts_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.tabControlCosts.Enabled)
                {
                    if (tabControlCosts.SelectedTab.Text.Trim().ToUpper() == "HAUL COSTS")
                    {
                        m_helpChapter = "HAUL_COSTS";
                        if (((Control)this.tbCosts).Enabled)
                        {
                            this.uc_scenario_costs1.lblTitle.Text = "Costs And Revenue";
                        }
                        else
                            this.uc_scenario_costs1.lblTitle.Text = "Costs And Revenue (Read Only)";

                    }
                    else if (tabControlCosts.SelectedTab.Text.Trim().ToUpper() == "PROCESSOR SCENARIO")
                    {
                        m_helpChapter = "PROCESSOR_SCENARIO";
                        if (((Control)this.tbProcessorScenario).Enabled)
                        {
                        }
                        else
                        {
                        }

                    }
                    

                }

            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void frmOptimizerScenario_Activated(object sender, EventArgs e)
        {
            if (uc_scenario_run1 != null)
            {
                this.uc_scenario_run1.groupBox1_Resize();
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(m_helpChapter))
            {
                if (m_oHelp == null)
                {
                    m_oHelp = new Help(m_xpsFile, m_oEnv);
                }
                m_oHelp.ShowHelp(new string[] { "TREATMENT_OPTIMIZER", this.HelpChapter });
            }
        }
	
	}
    public class OptimizerScenarioItem
    {
        
        public EffectiveVariablesItem_Collection m_oEffectiveVariablesItem_Collection = new EffectiveVariablesItem_Collection();
        public OptimizationVariableItem_Collection m_oOptimizationVariableItem_Collection = new OptimizationVariableItem_Collection();
        public TieBreaker_Collection m_oTieBreaker_Collection = new TieBreaker_Collection();
        public LastTieBreakRankItem_Collection m_oLastTieBreakRankItem_Collection = new LastTieBreakRankItem_Collection();
        public ProcessingSiteItem_Collection m_oProcessingSiteItem_Collection = new ProcessingSiteItem_Collection();
        public TransportationCosts m_oTranCosts = new TransportationCosts();
        public ProcessorScenarioItem_Collection m_oProcessorScenarioItem_Collection = new ProcessorScenarioItem_Collection();
        public ConditionTableSQLFilter m_oCondTableSQLFilter = new ConditionTableSQLFilter();

        public OptimizerScenarioItem()
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
        private string _strOwnerGroupCdList = "";
        public string OwnerGroupCodeList
        {
            get { return _strOwnerGroupCdList; }
            set { _strOwnerGroupCdList = value; }
        }
               
        private string _strPlotTableSQLFilter = "";
        public string PlotTableSQLFilter
        {
            get {return _strPlotTableSQLFilter;}
            set {_strPlotTableSQLFilter=value;}
        }
        private string _strNotes = "";
        public string Notes
        {
            get { return _strNotes; }
            set { _strNotes = value; }
        }
        public struct ConditionTableSQLFilter
        {
            public string SQL;
            public string LowSlopeMaximumYardingDistanceFeet;
            public string SteepSlopeMaximumYardingDistanceFeet;
        }
        public void Copy(OptimizerScenarioItem p_oSource,
                         OptimizerScenarioItem p_oDest)
        {
            p_oDest.DbFileName = p_oSource.DbFileName;
            p_oDest.DbPath = p_oSource.DbPath;
            p_oDest.OwnerGroupCodeList = p_oSource.OwnerGroupCodeList;
            p_oDest.PlotTableSQLFilter = p_oSource.PlotTableSQLFilter;
            p_oDest.ScenarioId = p_oSource.ScenarioId;
            p_oDest.Description = p_oSource.Description;
            p_oDest.Notes = p_oSource.Notes;
            p_oDest.m_oEffectiveVariablesItem_Collection.Copy(
                p_oSource.m_oEffectiveVariablesItem_Collection,
            ref p_oDest.m_oEffectiveVariablesItem_Collection, true);
            p_oDest.m_oOptimizationVariableItem_Collection.Copy(
                p_oSource.m_oOptimizationVariableItem_Collection,
            ref p_oDest.m_oOptimizationVariableItem_Collection,true);
            p_oDest.m_oProcessingSiteItem_Collection.Copy(
                p_oSource.m_oProcessingSiteItem_Collection,
            ref p_oDest.m_oProcessingSiteItem_Collection, true);
            p_oDest.m_oLastTieBreakRankItem_Collection.Copy(
                p_oSource.m_oLastTieBreakRankItem_Collection,
            ref p_oDest.m_oLastTieBreakRankItem_Collection, true);
            p_oDest.m_oTieBreaker_Collection.Copy(
                p_oSource.m_oTieBreaker_Collection,
            ref p_oDest.m_oTieBreaker_Collection, true);
            p_oDest.m_oTranCosts.Copy(
                p_oSource.m_oTranCosts,
                p_oDest.m_oTranCosts);
            p_oDest.m_oCondTableSQLFilter = p_oSource.m_oCondTableSQLFilter;
            p_oDest.m_oProcessorScenarioItem_Collection.Copy(
                p_oSource.m_oProcessorScenarioItem_Collection,
            ref p_oDest.m_oProcessorScenarioItem_Collection,true);

            
        }



        public class TransportationCosts
        {

            public string RoadHaulCostPerGreenTonPerHour = "";
            public string RailHaulCostPerGreenTonPerMile = "";
            public string RailChipTransferPerGreenTon = "";
            public string RailMerchTransferPerGreenTon = "";

            public void Copy(TransportationCosts p_oSource, TransportationCosts p_oDest)
            {
                p_oDest.RailChipTransferPerGreenTon = p_oSource.RailChipTransferPerGreenTon;
                p_oDest.RailHaulCostPerGreenTonPerMile = p_oSource.RailHaulCostPerGreenTonPerMile;
                p_oDest.RailMerchTransferPerGreenTon = p_oSource.RailMerchTransferPerGreenTon;
                p_oDest.RoadHaulCostPerGreenTonPerHour = p_oSource.RoadHaulCostPerGreenTonPerHour;
            }
            
        }
        
        public class EffectiveVariablesItem
        {
            public static byte NUMBER_OF_VARIABLES = 4;
            public string[] m_strPreVarArray = new string[NUMBER_OF_VARIABLES];
            public string[] m_strPostVarArray = new string[NUMBER_OF_VARIABLES];
            public string[] m_strBetterExpr = new string[NUMBER_OF_VARIABLES];
            public string[] m_strWorseExpr = new string[NUMBER_OF_VARIABLES];
            public string[] m_strEffectiveExpr = new string[NUMBER_OF_VARIABLES];
            public string m_strOverallEffectiveExpr = "";
            private string _strRxCycle = "";
            private string _strType="";

            public EffectiveVariablesItem()
            {
                for (int x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                {
                    m_strPreVarArray[x] = "Not Defined";
                    m_strPostVarArray[x] = "Not Defined";
                    m_strBetterExpr[x] = "";
                    m_strWorseExpr[x] = "";
                    m_strEffectiveExpr[x] = "";
                    _strRxCycle = "";
                    _strType="";
                }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public string Type
            {
                get {return _strType;}
                set {_strType=value;}
            }
            public void Copy(EffectiveVariablesItem p_oSource,
                             EffectiveVariablesItem p_oDest)
            {
                for (int x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                {
                    p_oDest.m_strPreVarArray[x] = p_oSource.m_strPreVarArray[x];
                    p_oDest.m_strPostVarArray[x] = p_oSource.m_strPostVarArray[x];
                    p_oDest.m_strBetterExpr[x] = p_oSource.m_strBetterExpr[x];
                    p_oDest.m_strWorseExpr[x] = p_oSource.m_strWorseExpr[x];
                    p_oDest.m_strEffectiveExpr[x] = p_oSource.m_strEffectiveExpr[x];

                }
                p_oDest.RxCycle = p_oSource.RxCycle;
                p_oDest.m_strOverallEffectiveExpr = p_oSource.m_strOverallEffectiveExpr;                
            }
            /// <summary>
            /// return the table names found in the either m_strPreVarArray or m_strPostVarArray variables
            /// </summary>
            /// <param name="p_oVarArray">m_strPreVarArray or m_strPostVarArray values</param>
            /// <returns></returns>
            public string[] TableNames(string[] p_oVarArray)
            {
                string strTable = "";
                string strTableList = ",";
                string[] strTableArray = null;
                for (int x = 0; x <= p_oVarArray.Length - 1; x++)
                {
                    strTable = this.TableName(p_oVarArray[x]);
                    if (strTable.Trim().Length > 0)
                    {
                        if (strTableList.Trim().ToUpper().IndexOf("," + strTable.ToUpper().Trim().ToUpper() + ",", 0) != 0)
                        {
                            strTableList = strTableList + strTable.Trim() + ",";
                        }

                    }
                }

                if (strTableList.Trim().Length > 1)
                {
                    strTableList = strTableList.Substring(1, strTableList.Length - 2);
                    strTableArray = frmMain.g_oUtils.ConvertListToArray(strTableList, ",");

                }
                else
                {
                    strTableList = "";
                }
                return strTableArray;
            }
            /// <summary>
            /// expecting a value in the format of tablename.columnname
            /// </summary>
            /// <param name="p_strValue"></param>
            /// <returns></returns>
            public string TableName(string p_strValue)
            {
                string strTableName = "";
                if (p_strValue.Trim().ToUpper() != "NOT DEFINED")
                {
                    if (p_strValue.IndexOf(".", 0) > 0)
                    {
                        string[] strArray = frmMain.g_oUtils.ConvertListToArray(p_strValue, ".");
                        if (strArray[0].Trim().Length > 0)
                        {
                            strTableName = strArray[0].Trim();
                        }
                    }
                }
                return strTableName;
            }
            /// <summary>
            /// expecting a value in the format of tablename.columnname
            /// </summary>
            /// <param name="p_strValue"></param>
            /// <returns></returns>
            public string ColumnName(string p_strValue)
            {
                string strColumnName = "";
                if (p_strValue.Trim().ToUpper() != "NOT DEFINED")
                {
                    if (p_strValue.IndexOf(".", 0) > 0)
                    {
                        string[] strArray = frmMain.g_oUtils.ConvertListToArray(p_strValue, ".");
                        if (strArray.Length == 2)
                        {
                            if (strArray[1].Trim().Length > 0)
                            {
                                strColumnName = strArray[1].Trim();
                            }
                        }
                    }
                }
                return strColumnName;
            }


            public bool Modified(EffectiveVariablesItem p_oCompare)
            {
                if (m_strOverallEffectiveExpr.Trim().ToUpper() != p_oCompare.m_strOverallEffectiveExpr.Trim().ToUpper()) return true;

                for (int x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                {
                    if (m_strPreVarArray[x].Trim().ToUpper() != p_oCompare.m_strPreVarArray[x].Trim().ToUpper())
                        return true;

                    if (m_strPostVarArray[x].Trim().ToUpper() != p_oCompare.m_strPostVarArray[x].Trim().ToUpper())
                        return true;

                    if (m_strBetterExpr[x].Trim().ToUpper() != p_oCompare.m_strBetterExpr[x].Trim().ToUpper())
                        return true;

                    if (this.m_strWorseExpr[x].Trim().ToUpper() != p_oCompare.m_strWorseExpr[x].Trim().ToUpper())
                        return true;

                    if (this.m_strEffectiveExpr[x].Trim().ToUpper() != p_oCompare.m_strEffectiveExpr[x].Trim().ToUpper())
                        return true;



                }
                return false;
            }
        }
        public class EffectiveVariablesItem_Collection : System.Collections.CollectionBase
        {
            public EffectiveVariablesItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(EffectiveVariablesItem m_EffectiveVariablesItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_EffectiveVariablesItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_EffectiveVariablesItem);

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
            public EffectiveVariablesItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (EffectiveVariablesItem)List[Index];
            }
            public void Copy(EffectiveVariablesItem_Collection p_oSource,
                         ref EffectiveVariablesItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    EffectiveVariablesItem oItem = new EffectiveVariablesItem();
                    oItem.Copy(p_oSource.Item(x), oItem);
                    p_oDest.Add(oItem);

                }
            }

        }
        public class OptimizationVariableItem
		{
			public bool bSelected=false;
			public string strOptimizedVariable="";
			public string strFVSVariableName="";
			public string strValueSource="";
			public string strMaxYN="N";
			public string strMinYN="N";
			public bool bUseFilter=true;
			public string strFilterOperator="";
			public double dblFilterValue=0;
			public int intListViewIndex=-1;
            private string _strRxCycle = "";
            private string _strType = "";
            public string strRevenueAttribute = "";

            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public string Type
            {
                get { return _strType; }
                set { _strType = value; }
            }
			public void Copy(OptimizationVariableItem p_oSource,
                         ref OptimizationVariableItem p_oDest)
			{
				p_oDest.bSelected = p_oSource.bSelected;
				p_oDest.strOptimizedVariable=p_oSource.strOptimizedVariable;
				p_oDest.strValueSource = p_oSource.strValueSource;
				p_oDest.strMaxYN = p_oSource.strMaxYN;
				p_oDest.strMinYN = p_oSource.strMinYN;
				p_oDest.strFilterOperator = p_oSource.strFilterOperator;
				p_oDest.strFVSVariableName = p_oSource.strFVSVariableName;
				p_oDest.dblFilterValue = p_oSource.dblFilterValue;
				p_oDest.bUseFilter = p_oSource.bUseFilter;
				p_oDest.intListViewIndex = p_oSource.intListViewIndex;
                p_oDest.RxCycle = p_oSource.RxCycle;
                p_oDest.Type = p_oSource.Type;
                p_oDest.strRevenueAttribute = p_oSource.strRevenueAttribute;
			}
		}

		public class OptimizationVariableItem_Collection : System.Collections.CollectionBase
		{
			public OptimizationVariableItem_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(OptimizationVariableItem m_oVariable)
			{
				// vérify if object is not already in
				if (this.List.Contains(m_oVariable))
					throw new InvalidOperationException();
 
				// adding it
				this.List.Add(m_oVariable);
 
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
			public OptimizationVariableItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (OptimizationVariableItem) List[Index];
			}
			public void Copy(OptimizationVariableItem_Collection p_oSource,
			             ref OptimizationVariableItem_Collection p_oDest,bool p_bInitializeDest)
			{
				int x;
				if (p_bInitializeDest) p_oDest.Clear();
				for (x=0;x<=p_oSource.Count-1;x++)
				{
					OptimizationVariableItem oItem = new OptimizationVariableItem();
					oItem.Copy(p_oSource.Item(x),ref oItem);
					p_oDest.Add(oItem);

				}
			}


		}
        public class TieBreakerItem
        {
            public bool bSelected = true;
            public string strMethod = "";
            public string strFVSVariableName = "";
            public string strMaxYN = "N";
            public string strMinYN = "N";
            public string strValueSource = "";
            public int intListViewIndex = -1;
            private string _strRxCycle = "";

            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public void Copy(TieBreakerItem p_oSource, ref TieBreakerItem p_oDest)
            {
                p_oDest.bSelected = p_oSource.bSelected;
                p_oDest.strMethod = p_oSource.strMethod;
                p_oDest.strFVSVariableName = p_oSource.strFVSVariableName;
                p_oDest.strMaxYN = p_oSource.strMaxYN;
                p_oDest.strMinYN = p_oSource.strMinYN;
                p_oDest.strValueSource = p_oSource.strValueSource;
                p_oDest.intListViewIndex = p_oSource.intListViewIndex;
                p_oDest.RxCycle = p_oSource.RxCycle;


            }
        }

        public class TieBreaker_Collection : System.Collections.CollectionBase
        {
            public TieBreaker_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(TieBreakerItem m_oTieBreaker)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oTieBreaker))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oTieBreaker);

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
            public TieBreakerItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (TieBreakerItem)List[Index];
            }
            public void Copy(TieBreaker_Collection p_oSource,
                         ref TieBreaker_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    TieBreakerItem oItem = new TieBreakerItem();
                    oItem.Copy(p_oSource.Item(x), ref oItem);
                    p_oDest.Add(oItem);

                }
            }


        }
        public class LastTieBreakRankItem
        {
            private string _strRxPackage = "";
            public string RxPackage
            {
                get { return _strRxPackage; }
                set { _strRxPackage = value; }
            }
            private int _intLastTieBreakRank = -1;
            public int LastTieBreakRank
            {
                get { return _intLastTieBreakRank; }
                set { _intLastTieBreakRank = value; }
            }
            public void Copy(LastTieBreakRankItem p_oSource, ref LastTieBreakRankItem p_oDest)
            {
                p_oDest.RxPackage = p_oSource.RxPackage;
                p_oDest.LastTieBreakRank = p_oSource.LastTieBreakRank;
                

            }
        }

        public class LastTieBreakRankItem_Collection : System.Collections.CollectionBase
        {
            public LastTieBreakRankItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(LastTieBreakRankItem m_oLastTieBreakRankItem)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oLastTieBreakRankItem))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oLastTieBreakRankItem);

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
            public LastTieBreakRankItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (LastTieBreakRankItem)List[Index];
            }
            public void Copy(LastTieBreakRankItem_Collection p_oSource,
                         ref LastTieBreakRankItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    LastTieBreakRankItem oItem = new LastTieBreakRankItem();
                    oItem.Copy(p_oSource.Item(x), ref oItem);
                    p_oDest.Add(oItem);

                }
            }


        }
        public class ProcessingSiteItem
        {
            public int intListViewIndex = -1;
            string _strPSiteId = "";
            public string ProcessingSiteId
            {
                get { return _strPSiteId; }
                set { _strPSiteId = value; }
            }
            string _strPSiteName = "";
            public string ProcessingSiteName
            {
                get { return _strPSiteName; }
                set { _strPSiteName = value; }
            }
            public string[,] m_strTranCdDescArray = new string[3, 2] {{"1","Regular - PSite With Road Only Access"},
		  															  {"2","Railhead - Road To Rail Wood Transfer Point"},
		  															  {"3","Rail Collector - PSite With Both Road And Rail Access"}
		  															 };
            string _strTranCd = "";
            public string TransportationCode
            {
                get { return _strTranCd; }
                set { _strTranCd = value; }
            }
            public string[,] m_strBioCdDescArray = new string[3, 2] {{"1","Merchantable - Logs Only"},
						  											 {"2","Chips - Chips Only"},
							  										 {"3","Both - Logs And Chips"}
							  										};
            string _strBioCd = "";
            public string BiomassCode
            {
                get { return _strBioCd; }
                set { _strBioCd = value; }
            }

            string _strExistsYN = "N";
            public string ProcessingSiteExistYN
            {
                get { return _strExistsYN; }
                set { _strExistsYN = value; }
            }

            double _dblLat = -1;
            public double GPSLatitude
            {
                get { return _dblLat; }
                set { _dblLat = value; }
            }

            double _dblLon = -1;
            public double GPSLongitude
            {
                get { return _dblLon; }
                set { _dblLon = value; }
            }
            bool _bSelected = false;
            public bool Selected
            {
                get { return _bSelected; }
                set { _bSelected = value; }
            }




            public void Copy(ProcessingSiteItem p_oSource, ref ProcessingSiteItem p_oDest)
            {
                p_oDest.ProcessingSiteId = p_oSource.ProcessingSiteId;
                p_oDest.ProcessingSiteName = p_oSource.ProcessingSiteName;
                p_oDest.TransportationCode = p_oSource.TransportationCode;
                p_oDest.BiomassCode = p_oSource.BiomassCode;
                p_oDest.ProcessingSiteExistYN = p_oSource.ProcessingSiteExistYN;
                p_oDest.GPSLatitude = p_oSource.GPSLatitude;
                p_oDest.GPSLongitude = p_oSource.GPSLongitude;
                p_oDest.Selected = p_oSource.Selected;
                p_oDest.intListViewIndex = p_oSource.intListViewIndex;
            }
        }

        public class ProcessingSiteItem_Collection : System.Collections.CollectionBase
        {
            public ProcessingSiteItem_Collection()
            {
                //
                // TODO: Add constructor logic here
                //
            }

            public void Add(ProcessingSiteItem m_oPSite)
            {
                // vérify if object is not already in
                if (this.List.Contains(m_oPSite))
                    throw new InvalidOperationException();

                // adding it
                this.List.Add(m_oPSite);

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
            public ProcessingSiteItem Item(int Index)
            {
                // The appropriate item is retrieved from the List object and
                // explicitly cast to the Widget type, then returned to the 
                // caller.
                return (ProcessingSiteItem)List[Index];
            }
            public void Copy(ProcessingSiteItem_Collection p_oSource,
                         ref ProcessingSiteItem_Collection p_oDest, bool p_bInitializeDest)
            {
                int x;
                if (p_bInitializeDest) p_oDest.Clear();
                for (x = 0; x <= p_oSource.Count - 1; x++)
                {
                    ProcessingSiteItem oItem = new ProcessingSiteItem();
                    oItem.Copy(p_oSource.Item(x), ref oItem);
                    p_oDest.Add(oItem);

                }
            }


        }
    }
    public class OptimizerScenarioItem_Collection : System.Collections.CollectionBase
    {
        public OptimizerScenarioItem_Collection()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public void Add(FIA_Biosum_Manager.OptimizerScenarioItem m_OptimizerScenarioItem)
        {
            // vérify if object is not already in
            if (this.List.Contains(m_OptimizerScenarioItem))
                throw new InvalidOperationException();

            // adding it
            this.List.Add(m_OptimizerScenarioItem);

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
        public FIA_Biosum_Manager.OptimizerScenarioItem Item(int Index)
        {
            // The appropriate item is retrieved from the List object and
            // explicitly cast to the Widget type, then returned to the 
            // caller.
            return (FIA_Biosum_Manager.OptimizerScenarioItem)List[Index];
        }

    }




    public class OptimizerScenarioTools
    {
        public string m_strError = "";
        public int m_intError = 0;
        public OptimizerScenarioTools()
        {
        }
        public void LoadScenario(string p_strScenarioId, Queries p_oQueries, bool p_bProcessorUsingSqlite,
            OptimizerScenarioItem_Collection p_oOptimizerScenarioItem_Collection)
        {
           
            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            p_oQueries.m_oFvs.LoadDatasource = true;
            p_oQueries.m_oFIAPlot.LoadDatasource = true;
            p_oQueries.m_oProcessor.LoadDatasource = true;
            p_oQueries.m_oReference.LoadDatasource = true;
            p_oQueries.LoadDatasources(true, false, "optimizer", p_strScenarioId);
            p_oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                p_oQueries.m_strTempDbFile,
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                "C");
            LoadAll(p_oQueries.m_strTempDbFile, p_oQueries, p_bProcessorUsingSqlite,
                p_strScenarioId, p_oOptimizerScenarioItem_Collection);
        }
        public void LoadAll(string p_strDbFile, Queries p_oQueries, bool p_bProcessorUsingSqlite, 
            string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem_Collection p_oOptimizerScenarioItem_Collection)
        {
            int x;
            for (x = p_oOptimizerScenarioItem_Collection.Count - 1; x >= 0; x--)
            {
                if (p_oOptimizerScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                    p_strScenarioId.Trim().ToUpper())
                {
                    p_oOptimizerScenarioItem_Collection.Remove(x);
                }
            }
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {

                OptimizerScenarioItem oItem = new OptimizerScenarioItem();
                this.LoadGeneral(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadEffectiveVariables(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadOptimizationVariable(oAdo, oAdo.m_OleDbConnection, p_strScenarioId,oItem);
                this.LoadTieBreakerVariables(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadLastTieBreakRank(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                if (!p_bProcessorUsingSqlite)
                {
                    this.LoadProcessorScenarioItems(oAdo, oAdo.m_OleDbConnection, p_bProcessorUsingSqlite, p_strScenarioId, oItem);
                }
                else
                {
                    this.LoadProcessorScenarioItemsSqlite(oAdo, oAdo.m_OleDbConnection, p_bProcessorUsingSqlite, p_strScenarioId, oItem);
                }
                this.LoadPlotFilter(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadCondFilter(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadProcessingSites(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadLandOwnerGroupFilter(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                this.LoadTransportationCosts(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, oItem);
                p_oOptimizerScenarioItem_Collection.Add(oItem);
                




            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }

        public void LoadGeneral(string p_strDbFile, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerAnalysisScenarioItem)
        {
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(p_strDbFile, "", ""));
            if (oAdo.m_intError == 0)
            {
                this.LoadGeneral(oAdo, oAdo.m_OleDbConnection, p_strScenarioId, p_oOptimizerAnalysisScenarioItem);
            }
            m_intError = oAdo.m_intError;
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            oAdo = null;
        }
        public void LoadGeneral(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
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
                            p_oOptimizerScenarioItem.ScenarioId = p_oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                        }
                        //
                        //DESCRIPTION
                        //
                        if (p_oAdo.m_OleDbDataReader["description"] != System.DBNull.Value)
                        {
                            p_oOptimizerScenarioItem.Description = p_oAdo.m_OleDbDataReader["description"].ToString().Trim();
                        }
                        //
                        //PATH
                        //
                        if (p_oAdo.m_OleDbDataReader["path"] != System.DBNull.Value)
                        {
                            p_oOptimizerScenarioItem.DbPath = p_oAdo.m_OleDbDataReader["path"].ToString().Trim();
                        }
                        //
                        //FILE
                        //
                        if (p_oAdo.m_OleDbDataReader["file"] != System.DBNull.Value)
                        {
                            p_oOptimizerScenarioItem.DbFileName = p_oAdo.m_OleDbDataReader["file"].ToString().Trim();
                        }
                        //
                        //NOTES
                        //
                        if (p_oAdo.m_OleDbDataReader["notes"] != System.DBNull.Value)
                        {
                            p_oOptimizerScenarioItem.Notes = p_oAdo.m_OleDbDataReader["notes"].ToString().Trim();
                        }

                    }
                    p_oAdo.m_OleDbDataReader.Close();
                }
            }
        }
        public void LoadEffectiveVariables(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            int x = 0, y = 0;
            string strRxCycle = "";
            //
            //INITIALIZE EFFECTIVE VARIABLES COLLECTION
            //
            for (x = p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1;x>=0 ; x--)
            {
                p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Remove(x);
            }
            //
            //LOAD EFFECTIVE VARIABLES
            //
            p_oAdo.m_strSQL = "SELECT a.* " +
                            "FROM scenario_fvs_variables a " +
                            "WHERE TRIM(a.scenario_id)='" + p_strScenarioId.Trim() + "' AND " +
                            "a.current_yn='Y' ORDER BY a.rxcycle,a.variable_number";

            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                    
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    int intVarNum = Convert.ToInt32(p_oAdo.m_OleDbDataReader["variable_number"]) - 1;
                    strRxCycle = Convert.ToString(p_oAdo.m_OleDbDataReader["rxcycle"]).Trim();
                    //find the current rxcycle in the collection
                    for (x = 0; x <= p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1; x++)
                    {
                        if (p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle == strRxCycle)
                            break;
                    }

                    if (x > p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1)
                    {
                        OptimizerScenarioItem.EffectiveVariablesItem oEffItem = new OptimizerScenarioItem.EffectiveVariablesItem();
                        oEffItem.RxCycle = strRxCycle;
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Add(oEffItem);
                    }
                    p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPreVarArray[intVarNum] =
                        Convert.ToString(p_oAdo.m_OleDbDataReader["pre_fvs_variable"]).Trim();

                    p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPostVarArray[intVarNum] =
                        Convert.ToString(p_oAdo.m_OleDbDataReader["post_fvs_variable"]).Trim();


                    if (p_oAdo.m_OleDbDataReader["better_expression"] != System.DBNull.Value)
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strBetterExpr[intVarNum] =
                            Convert.ToString(p_oAdo.m_OleDbDataReader["better_expression"]).Trim();

                    if (p_oAdo.m_OleDbDataReader["worse_expression"] != System.DBNull.Value)
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strWorseExpr[intVarNum] =
                                Convert.ToString(p_oAdo.m_OleDbDataReader["worse_expression"]).Trim();

                    if (p_oAdo.m_OleDbDataReader["effective_expression"] != System.DBNull.Value)
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strEffectiveExpr[intVarNum] =
                            Convert.ToString(p_oAdo.m_OleDbDataReader["effective_expression"]).Trim();

                    if (p_oAdo.m_OleDbDataReader["rxcycle"] != System.DBNull.Value)
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle = 
                            Convert.ToString(p_oAdo.m_OleDbDataReader["rxcycle"]).Trim();


                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            //
            //LOAD OVERALL EFFECTIVE VARIABLES
            //
            //overall expression
            p_oAdo.m_strSQL = "SELECT b.overall_effective_expression,b.current_yn," +
                            "b.rxcycle " +
                            "FROM scenario_fvs_variables_overall_effective b " +
                            "WHERE TRIM(b.scenario_id)='" + p_strScenarioId.Trim() + "' AND " +
                            "b.current_yn='Y'";



            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    strRxCycle = Convert.ToString(p_oAdo.m_OleDbDataReader["rxcycle"]).Trim();
                    //find the current rxcycle in the collection
                    for (x = 0; x <= p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1; x++)
                    {
                        if (p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle == strRxCycle)
                            break;
                    }
                    if (p_oAdo.m_OleDbDataReader["overall_effective_expression"] != System.DBNull.Value &&
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveExpr.Trim().Length == 0)
                        p_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveExpr =
                                Convert.ToString(p_oAdo.m_OleDbDataReader["overall_effective_expression"]).Trim();

                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            
            

               
        }
        public void LoadOptimizationVariable(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            int x = 0, y = 0;
            int intVarNum = 0;

            //
            //INITIALIZE OPTIMIZATION VARIABLES COLLECTION
            //
            for (x = p_oOptimizerScenarioItem.m_oOptimizationVariableItem_Collection.Count - 1;x>=0 ; x--)
            {
                p_oOptimizerScenarioItem.m_oOptimizationVariableItem_Collection.Remove(x);
            }
            
            p_oAdo.m_strSQL = "SELECT * " +
                "FROM scenario_fvs_variables_optimization " +
                "WHERE TRIM(scenario_id)='" + p_strScenarioId.Trim() + "' AND " +
                "current_yn='Y'";

            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                   

                    if (p_oAdo.m_OleDbDataReader["optimization_variable"] != System.DBNull.Value)
                    {
                        OptimizerScenarioItem.OptimizationVariableItem oItem = new OptimizerScenarioItem.OptimizationVariableItem();
                        oItem.strOptimizedVariable = Convert.ToString(p_oAdo.m_OleDbDataReader["optimization_variable"]);
                        //optimization variable
                        if (p_oAdo.m_OleDbDataReader["fvs_variable_name"] != System.DBNull.Value)
                        {
                            oItem.strFVSVariableName = Convert.ToString(p_oAdo.m_OleDbDataReader["fvs_variable_name"]).Trim();
                        }
                        else
                        {
                            oItem.strFVSVariableName = "NA";
                        }
                        //value source (POST or POST-PRE)
                        if (p_oAdo.m_OleDbDataReader["value_source"] != System.DBNull.Value)
                        {
                            oItem.strValueSource = Convert.ToString(p_oAdo.m_OleDbDataReader["value_source"]).Trim();
                        }
                        else
                        {
                            oItem.strValueSource = "Not Defined";
                        }

                        //max value
                        if (p_oAdo.m_OleDbDataReader["max_yn"] != System.DBNull.Value)
                        {
                            oItem.strMaxYN = Convert.ToString(p_oAdo.m_OleDbDataReader["max_yn"]).Trim();
                        }
                        else
                        {
                            oItem.strMaxYN = "N";
                        }
                        //min value
                        if (p_oAdo.m_OleDbDataReader["min_yn"] != System.DBNull.Value)
                        {
                            oItem.strMinYN = Convert.ToString(p_oAdo.m_OleDbDataReader["min_yn"]).Trim();
                        }
                        else
                        {
                            oItem.strMinYN = "N";
                        }
                        //enable filter
                        if (p_oAdo.m_OleDbDataReader["filter_enabled_yn"] != System.DBNull.Value)
                        {
                            if (Convert.ToString(p_oAdo.m_OleDbDataReader["filter_enabled_yn"]).Trim() == "Y")
                                oItem.bUseFilter = true;
                            else
                                oItem.bUseFilter = false;

                        }
                        else
                        {
                            oItem.bUseFilter = false;
                        }
                        //filter operator
                        if (p_oAdo.m_OleDbDataReader["filter_operator"] != System.DBNull.Value)
                        {

                            oItem.strFilterOperator = Convert.ToString(p_oAdo.m_OleDbDataReader["filter_operator"]).Trim();
                        }
                        else
                        {
                            oItem.strFilterOperator = "";
                        }
                        //filter value
                        if (p_oAdo.m_OleDbDataReader["filter_value"] != System.DBNull.Value)
                        {
                            oItem.dblFilterValue = Convert.ToDouble(p_oAdo.m_OleDbDataReader["filter_value"]);
                        }
                        //revenue filter attribute
                        if (p_oAdo.m_OleDbDataReader["revenue_attribute"] != System.DBNull.Value)
                        {
                            oItem.strRevenueAttribute = Convert.ToString(p_oAdo.m_OleDbDataReader["revenue_attribute"]).Trim();
                        }
                        //filter operator
                        if (p_oAdo.m_OleDbDataReader["checked_yn"] != System.DBNull.Value)
                        {
                            if (Convert.ToString(p_oAdo.m_OleDbDataReader["checked_yn"]).Trim() == "Y")
                            {
                                oItem.bSelected = true;
                                if (oItem.strOptimizedVariable.Trim().ToUpper() == "FVS VARIABLE")
                                {
                                }
                                else
                                {
                                }
                            }
                            else
                                oItem.bSelected = false;
                        }
                        else
                        {
                            oItem.bSelected = false;
                        }
                        oItem.RxCycle = p_oAdo.m_OleDbDataReader["rxcycle"].ToString().Trim();
                        p_oOptimizerScenarioItem.m_oOptimizationVariableItem_Collection.Add(oItem);
                    }
                }
                
            }
            p_oAdo.m_OleDbDataReader.Close();

        }
        public void LoadTieBreakerVariables(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            this.m_intError = 0;
            this.m_strError = "";
            

            int x, y;
            //
            //INITIALIZE TIEBREAKER COLLECTION
            //
            p_oOptimizerScenarioItem.m_oTieBreaker_Collection.Clear();
            //
            //LOAD TIEBREAKER ITEMS
            //
            p_oAdo.m_strSQL = "SELECT * FROM " +
                                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                "WHERE TRIM(scenario_id)='" + p_strScenarioId.Trim() + "'";


            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {


                    if (p_oAdo.m_OleDbDataReader["tiebreaker_method"] != System.DBNull.Value)
                    {
                        OptimizerScenarioItem.TieBreakerItem oItem = new OptimizerScenarioItem.TieBreakerItem();
                        oItem.RxCycle = p_oAdo.m_OleDbDataReader["rxcycle"].ToString().Trim();
                        oItem.strMethod = p_oAdo.m_OleDbDataReader["tiebreaker_method"].ToString().Trim();
                        if (oItem.strMethod.ToUpper().IndexOf("ATTRIBUTE") > -1)
                        {


                            //fvs variable name
                            oItem.strFVSVariableName = p_oAdo.m_OleDbDataReader["fvs_variable_name"].ToString().Trim();
                            oItem.strValueSource = p_oAdo.m_OleDbDataReader["value_source"].ToString().Trim();



                            //MAX or MIN	
                            if (p_oAdo.m_OleDbDataReader["max_yn"].ToString().Trim().ToUpper() == "Y")
                            {
                                oItem.strMaxYN = "Y";
                                oItem.strMinYN = "N";
                            }
                            else
                            {
                                oItem.strMinYN = "Y";
                                oItem.strMaxYN = "N";

                            }
                            if (p_oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper() == "Y")
                            {
                                oItem.bSelected = true;

                            }
                            else
                            {
                                oItem.bSelected = false;
                            }
                        }
                        else if (oItem.strMethod.ToUpper() == "LAST TIE-BREAK RANK")
                        {
                            if (p_oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper() == "Y")
                            {
                                oItem.bSelected = true;

                            }
                            else
                            {
                                oItem.bSelected = false;
                            }
                        }
                        p_oOptimizerScenarioItem.m_oTieBreaker_Collection.Add(oItem);
                    }

                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            

        }
        public void LoadLastTieBreakRank(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            this.m_intError = 0;
            this.m_strError = "";


            int x, y;
            //
            //GET ALL THE CURRENT TREATMENTS
            //
            /*************************************************************************
			 **get the treatment packages mdb file,table, and connection strings
			 *************************************************************************/
            string strRxMDBFile = "", strRxPackageTableName = "", strRxConn = "";
            p_oAdo.getScenarioDataSourceConnStringAndTable(ref strRxMDBFile,
                                            ref strRxPackageTableName, ref strRxConn,
                                            "Treatment Packages",
                                            p_strScenarioId,
                                            p_oConn);
            //
            //GET A LIST OF ALL THE CURRENT TREATMENTS
            //
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            p_oAdo.OpenConnection(strRxConn,ref oConn);
            string strRxList = p_oAdo.CreateCommaDelimitedList(oConn, "SELECT RXPACKAGE FROM " + strRxPackageTableName, "'");
            string[] strRxArray = frmMain.g_oUtils.ConvertListToArray(strRxList, ",");
            p_oAdo.CloseConnection(oConn);
            //
            //INITIALIZE RXINTENSITY COLLECTION
            //
            p_oOptimizerScenarioItem.m_oLastTieBreakRankItem_Collection.Clear();
            //
            //FIRST DELETE THE RXPACKAGES THAT NO LONGER EXIST
            //
            p_oAdo.m_strSQL = "DELETE FROM scenario_last_tiebreak_rank WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' AND rxpackage NOT IN (" + strRxList + ")";
            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            //
            //NEXT INSERT INTO THE SCENARIO_LAST_TIEBREAK_RANK THOSE RX'S THAT DO NOT CURRENTLY EXIST
            //
            if (p_oAdo.TableExist(p_oConn, "rxtemp"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE rxtemp");
            p_oAdo.SqlNonQuery(p_oConn, Tables.OptimizerScenarioRuleDefinitions.CreateScenarioLastTieBreakTableSQL("rxtemp"));
            for (x = 0; x <= strRxArray.Length - 1; x++)
            {
                // SET THE DEFAULT TREATMENT INTENSITY TO BE THE SAME AS RXPACKAGE SO IT IS ALWAYS UNIQUE AND NEVER NULL
                int intIntensity = -1;
                char[] charsToTrim = {' ', '\'' };
                bool bResult = Int32.TryParse(strRxArray[x].Trim(charsToTrim), out intIntensity);
                p_oAdo.SqlNonQuery(p_oConn, "INSERT INTO rxtemp (scenario_id,rxpackage,last_tiebreak_rank) VALUES ('" + p_strScenarioId + "'," + strRxArray[x].Trim() + ", " + intIntensity + " )");
            }
            p_oAdo.m_strSQL = "INSERT INTO scenario_last_tiebreak_rank " +
                             "(scenario_id,rxpackage,last_tiebreak_rank) " +
                             "SELECT a.scenario_id, a.RXPACKAGE, a.last_tiebreak_rank " +
                             "FROM rxtemp a " +
                             "WHERE NOT EXISTS (SELECT b.scenario_id,b.rxpackage " +
                                               "FROM scenario_last_tiebreak_rank b " +
                                               "WHERE a.rxpackage=b.rxpackage AND TRIM(UCASE(a.scenario_id))=TRIM(UCASE(b.scenario_id)))";
            p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
            if (p_oAdo.TableExist(p_oConn, "rxtemp"))
                p_oAdo.SqlNonQuery(p_oConn, "DROP TABLE rxtemp");
            //
            //LOAD RXINTENSITY ITEMS
            //
            p_oAdo.m_strSQL = "SELECT * FROM " +
                              Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName + " " +
                              "WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";

            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows)
            {
                while (p_oAdo.m_OleDbDataReader.Read())
                {


                    if (p_oAdo.m_OleDbDataReader["rxpackage"] != System.DBNull.Value)
                    {
                        OptimizerScenarioItem.LastTieBreakRankItem oItem = new OptimizerScenarioItem.LastTieBreakRankItem();
                        oItem.RxPackage = p_oAdo.m_OleDbDataReader["rxpackage"].ToString().Trim();
                        if (p_oAdo.m_OleDbDataReader["last_tiebreak_rank"] != System.DBNull.Value)
                        {
                            oItem.LastTieBreakRank = Convert.ToInt32(p_oAdo.m_OleDbDataReader["last_tiebreak_rank"]);
                        }
                        else
                        {
                            oItem.LastTieBreakRank = -1;
                        }
                        p_oOptimizerScenarioItem.m_oLastTieBreakRankItem_Collection.Add(oItem);
                    }

                }
            }

            p_oAdo.m_OleDbDataReader.Close();
            

        }
        
        public void LoadProcessorScenarioItems(ado_data_access p_oAdo,
                                         System.Data.OleDb.OleDbConnection p_oConn,
                                         bool bProcessorUsingSqlite, string p_strScenarioId,
                                         OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            int x;
            Queries oQueries = new Queries();
            ProcessorScenarioTools oTools = new ProcessorScenarioTools();

            //scenario mdb connection
            string strProcessorScenarioMDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
              "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            //
            //get a list of all the scenarios
            //
            ado_data_access oAdo = new ado_data_access();
            oAdo.OpenConnection(oAdo.getMDBConnString(strProcessorScenarioMDB, "", ""));
            string[] strScenarioArray =
                frmMain.g_oUtils.ConvertListToArray(
                    oAdo.CreateCommaDelimitedList(
                       oAdo.m_OleDbConnection,
                       "SELECT scenario_id " +
                       "FROM scenario " +
                       "WHERE scenario_id IS NOT NULL AND " +
                                         "LEN(TRIM(scenario_id)) > 0", ""), ",");
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            if (strScenarioArray == null) return;

            for (x = 0; x <= strScenarioArray.Length - 1; x++)
            {
                //
                //LOAD PROJECT DATATASOURCES INFO
                //
                oQueries.m_oFvs.LoadDatasource = true;
                oQueries.m_oFIAPlot.LoadDatasource = true;
                oQueries.m_oProcessor.LoadDatasource = true;
                oQueries.m_oReference.LoadDatasource = true;
                oQueries.LoadDatasources(true, bProcessorUsingSqlite, "processor", strScenarioArray[x]);
                oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                    oQueries.m_strTempDbFile,
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                    "P");
                oTools.LoadAll(oQueries.m_strTempDbFile, oQueries, strScenarioArray[x], p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection);
            }

            if (p_oAdo.TableExist(p_oConn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName))
            {
                p_oAdo.m_strSQL = "SELECT * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " " +
                                "WHERE TRIM(UCASE(scenario_id)) = '" +
                                   p_strScenarioId.Trim().ToUpper() + "';";
                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                string strProcessorScenario = "";
                string strFullDetailsYN = "N";

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["processor_scenario_id"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim().Length > 0)
                        {
                            strProcessorScenario = p_oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim();
                        }
                        if (p_oAdo.m_OleDbDataReader["FullDetailsYN"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim().Length > 0)
                        {
                            strFullDetailsYN = p_oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim();
                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
                //
                //find the current selected scenario in the collection
                //
                for (x = 0; x <= p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Count - 1; x++)
                {
                    if (p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                        strProcessorScenario.Trim().ToUpper())
                    {
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected = true;
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).DisplayFullDetailsYN = strFullDetailsYN;
                    }
                    else
                    {
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected = false;
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).DisplayFullDetailsYN = strFullDetailsYN;

                    }
                }
            }
            else
            {
                frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo, p_oConn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);
            }
        }

        public void LoadProcessorScenarioItemsSqlite(ado_data_access p_oAdo,
                                 System.Data.OleDb.OleDbConnection p_oConn,
                                 bool bProcessorUsingSqlite, string p_strScenarioId,
                                 OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            int x;
            Queries oQueries = new Queries();
            ProcessorScenarioTools oTools = new ProcessorScenarioTools();

            //scenario mdb connection
            string strProcessorScenarioDB =
              frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
              "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            //
            //get a list of all the scenarios
            //
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = dataMgr.GetConnectionString(strProcessorScenarioDB);
            using (System.Data.SQLite.SQLiteConnection oConn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                oConn.Open();
                System.Collections.Generic.IList<string> lstScenarios = dataMgr.getStringList(oConn, 
                       "SELECT scenario_id " +
                       "FROM scenario " +
                       "WHERE scenario_id IS NOT NULL AND " +
                                         "LENGTH(TRIM(scenario_id)) > 0");
                if (lstScenarios.Count == 0) return;
                for (x = 0; x <= lstScenarios.Count - 1; x++)
                {
                    //
                    //LOAD PROJECT DATATASOURCES INFO
                    //
                    oQueries.m_oFvs.LoadDatasource = true;
                    oQueries.m_oFIAPlot.LoadDatasource = true;
                    oQueries.m_oProcessor.LoadDatasource = true;
                    oQueries.m_oReference.LoadDatasource = true;
                    oQueries.LoadDatasources(true, bProcessorUsingSqlite, "processor", lstScenarios[x]);
                    //@ToDo: Do we really need to be creating these table links?
                    //oQueries.m_oDataSource.CreateScenarioRuleDefinitionTableLinks(
                    //    oQueries.m_strTempDbFile,
                    //    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                    //    "P");
                    oTools.LoadAllSqlite(oQueries, lstScenarios[x], p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection);
                }
            }
            ado_data_access oAdo = new ado_data_access();



            if (p_oAdo.TableExist(p_oConn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName))
            {
                p_oAdo.m_strSQL = "SELECT * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " " +
                                "WHERE TRIM(UCASE(scenario_id)) = '" +
                                   p_strScenarioId.Trim().ToUpper() + "';";
                p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);
                string strProcessorScenario = "";
                string strFullDetailsYN = "N";

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["processor_scenario_id"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim().Length > 0)
                        {
                            strProcessorScenario = p_oAdo.m_OleDbDataReader["processor_scenario_id"].ToString().Trim();
                        }
                        if (p_oAdo.m_OleDbDataReader["FullDetailsYN"] != System.DBNull.Value &&
                            p_oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim().Length > 0)
                        {
                            strFullDetailsYN = p_oAdo.m_OleDbDataReader["FullDetailsYN"].ToString().Trim();
                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
                //
                //find the current selected scenario in the collection
                //
                for (x = 0; x <= p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Count - 1; x++)
                {
                    if (p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).ScenarioId.Trim().ToUpper() ==
                        strProcessorScenario.Trim().ToUpper())
                    {
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected = true;
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).DisplayFullDetailsYN = strFullDetailsYN;
                    }
                    else
                    {
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).Selected = false;
                        p_oOptimizerScenarioItem.m_oProcessorScenarioItem_Collection.Item(x).DisplayFullDetailsYN = strFullDetailsYN;

                    }
                }
            }
            else
            {
                frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioProcessorScenarioSelectTable(oAdo, p_oConn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName);
            }




        }
        public void LoadPlotFilter(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {

            p_oAdo.m_strSQL = "SELECT * FROM scenario_plot_filter  WHERE " +
                " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y';";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);



            if (p_oAdo.m_intError == 0)
            {

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["sql_command"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim().Length > 0)
                            {
                                p_oOptimizerScenarioItem.PlotTableSQLFilter = p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim();
                            }
                        }

                    }
                }
                else
                {

                    p_oOptimizerScenarioItem.PlotTableSQLFilter = "SELECT @@PlotTable@@.* FROM @@PlotTable@@ ";

                }

                p_oAdo.m_OleDbDataReader.Close();
            }



        }
        public void LoadCondFilter(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            p_oAdo.m_strSQL = "SELECT * FROM scenario_cond_filter  WHERE " +
                " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y';";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);



            if (p_oAdo.m_intError == 0)
            {

                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["sql_command"] != System.DBNull.Value)
                        {
                            if (p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim().Length > 0)
                            {
                                p_oOptimizerScenarioItem.m_oCondTableSQLFilter.SQL =
                                    p_oAdo.m_OleDbDataReader["sql_command"].ToString().Trim();
                            }
                        }

                    }
                }
                else
                {
                    p_oOptimizerScenarioItem.m_oCondTableSQLFilter.SQL =
                        "SELECT * FROM @@CondTable@@";
                }

                p_oAdo.m_OleDbDataReader.Close();
            }

            p_oAdo.m_strSQL = "SELECT * FROM scenario_cond_filter_misc WHERE " +
                             " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "';";
            p_oAdo.SqlQueryReader(p_oConn, p_oAdo.m_strSQL);



            if (p_oAdo.m_intError == 0)
            {
                //load  wind speed class definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    if (p_oAdo.m_OleDbDataReader["yard_dist"] != System.DBNull.Value)
                    {
                        if (p_oAdo.m_OleDbDataReader["yard_dist"].ToString().Trim().Length > 0)
                        {
                            p_oOptimizerScenarioItem.m_oCondTableSQLFilter.LowSlopeMaximumYardingDistanceFeet = 
                                Convert.ToString(p_oAdo.m_OleDbDataReader["yard_dist"].ToString().Trim());
                            p_oOptimizerScenarioItem.m_oCondTableSQLFilter.SteepSlopeMaximumYardingDistanceFeet=
                                Convert.ToString(p_oAdo.m_OleDbDataReader["yard_dist2"].ToString().Trim());
                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
            }
        }

        public void LoadProcessingSites(FIA_Biosum_Manager.ado_data_access p_oAdo, System.Data.OleDb.OleDbConnection p_oConn, string p_strScenarioId, FIA_Biosum_Manager.OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
			

			
			this.m_intError = 0;
            this.m_strError = "";


            int x, y;
            //
            //GET ALL THE PSITES LISTED IN THE TRAVEL TIMES TABLE
            //
            /*************************************************************************
			 **get the travel times mdb file,table, and connection strings
			 *************************************************************************/
            string strTravelTimeMDBFile = "", strTravelTimeTableName = "", strTravelTimeConn = "";
            p_oAdo.getScenarioDataSourceConnStringAndTable(ref strTravelTimeMDBFile,
                                            ref strTravelTimeTableName, ref strTravelTimeConn,
                                            "TRAVEL TIMES",
                                            p_strScenarioId,
                                            p_oConn);
            //
            //GET THE PSITE LIST
            //
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            p_oAdo.OpenConnection(strTravelTimeConn,ref oConn);
            string strList = p_oAdo.CreateCommaDelimitedList(oConn, "SELECT DISTINCT CSTR(psite_id) AS psite_id FROM " + strTravelTimeTableName, "");
            string[] strArray = frmMain.g_oUtils.ConvertListToArray(strList, ",");
            p_oAdo.CloseConnection(oConn);
            //
            //INITIALIZE PSITE COLLECTION
            //
            p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Clear();
            //
            //FIRST DELETE THE PSITE'S THAT NO LONGER EXIST
            //
            if (!String.IsNullOrEmpty(strList))
            {
                p_oAdo.m_strSQL = "DELETE FROM scenario_psites WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "' AND psite_id NOT IN (" + strList + ")";
                p_oAdo.SqlNonQuery(p_oConn, p_oAdo.m_strSQL);
                //
                //LOAD UP THE PSITES THAT EXIST
                //
                for (x = 0; x <= strArray.Length - 1; x++)
                {
                    OptimizerScenarioItem.ProcessingSiteItem oItem = new OptimizerScenarioItem.ProcessingSiteItem();
                    oItem.ProcessingSiteId = strArray[x].Trim();
                    p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Add(oItem);

                }
                /*************************************************************************
                 **get the processing sites mdb file,table, and connection strings
                 *************************************************************************/
                string strPSitesMDBFile = "", strPSitesTableName = "", strPSitesConn = "";
                p_oAdo.getScenarioDataSourceConnStringAndTable(ref strPSitesMDBFile,
                                                ref strPSitesTableName, ref strPSitesConn,
                                                "PROCESSING SITES",
                                                p_strScenarioId,
                                                p_oConn);
                //
                //GET THE PSITE LIST
                //
                oConn = new System.Data.OleDb.OleDbConnection();
                p_oAdo.OpenConnection(strPSitesConn, ref oConn);
                p_oAdo.m_strSQL = "SELECT DISTINCT p.psite_id,p.name,p.trancd,p.trancd_def,p.biocd,p.biocd_def,p.exists_yn " +
                                 "FROM " + strPSitesTableName + " p WHERE p.psite_id IN (" + strList + ")";
                p_oAdo.SqlQueryReader(oConn, p_oAdo.m_strSQL);
                if (p_oAdo.m_OleDbDataReader.HasRows == true)
                {
                    x = 0;
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        if (p_oAdo.m_OleDbDataReader["psite_id"] != System.DBNull.Value)
                        {
                            for (x = 0; x <= p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Count - 1; x++)
                            {
                                if (p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId.Trim() ==
                                    Convert.ToString(p_oAdo.m_OleDbDataReader["psite_id"]).Trim())
                                {
                                    //processing site name
                                    if (p_oAdo.m_OleDbDataReader["name"] != System.DBNull.Value)
                                    {
                                        p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteName =
                                            Convert.ToString(p_oAdo.m_OleDbDataReader["name"]).Trim();
                                    }
                                    //transportation code
                                    if (p_oAdo.m_OleDbDataReader["trancd"] != System.DBNull.Value)
                                    {
                                        p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode =
                                            Convert.ToString(p_oAdo.m_OleDbDataReader["trancd"]).Trim();
                                    }
                                    //biomass code
                                    if (p_oAdo.m_OleDbDataReader["biocd"] != System.DBNull.Value)
                                    {
                                        p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode =
                                            Convert.ToString(p_oAdo.m_OleDbDataReader["biocd"]).Trim();
                                    }
                                    //exists YN
                                    if (p_oAdo.m_OleDbDataReader["exists_yn"] != System.DBNull.Value)
                                    {
                                        p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteExistYN =
                                            Convert.ToString(p_oAdo.m_OleDbDataReader["exists_yn"]).Trim();
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
                p_oAdo.CloseConnection(oConn);
            }
            
            //
            //SET THE PREVIOUSLY SELECTED
            //
            p_oAdo.m_strSQL = "SELECT psite_id,trancd,biocd,selected_yn " +
                              "FROM scenario_psites WHERE TRIM(UCASE(scenario_id))='" + p_strScenarioId.Trim().ToUpper() + "'";
            p_oAdo.SqlQueryReader(p_oConn,p_oAdo.m_strSQL);
            if (p_oAdo.m_OleDbDataReader.HasRows == true)
            {
                x = 0;
                while (p_oAdo.m_OleDbDataReader.Read())
                {

                    if (p_oAdo.m_OleDbDataReader["psite_id"] != System.DBNull.Value)
                    {
                        for (x = 0; x <= p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Count - 1; x++)
                        {
                            if (p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId.Trim() ==
                                Convert.ToString(p_oAdo.m_OleDbDataReader["psite_id"]).Trim())
                            {
                                if (p_oAdo.m_OleDbDataReader["selected_yn"] != System.DBNull.Value)
                                {
                                    if (Convert.ToString(p_oAdo.m_OleDbDataReader["selected_yn"]).Trim() == "Y")
                                    {
                                        p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected = true;
                                    }
                                    else
                                    {
                                        p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected = false;
                                    }
                                }
                                //transportation code
                                if (p_oAdo.m_OleDbDataReader["trancd"] != System.DBNull.Value)
                                {
                                    p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode =
                                        Convert.ToString(p_oAdo.m_OleDbDataReader["trancd"]).Trim();
                                }
                                //biomass code
                                if (p_oAdo.m_OleDbDataReader["biocd"] != System.DBNull.Value)
                                {
                                    p_oOptimizerScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode =
                                        Convert.ToString(p_oAdo.m_OleDbDataReader["biocd"]).Trim();
                                }
                                break;
                            }
                        }
                    }
                }
                
            }
            p_oAdo.m_OleDbDataReader.Close();
            

			
		}
        public void LoadLandOwnerGroupFilter(ado_data_access p_oAdo, 
                                         System.Data.OleDb.OleDbConnection p_oConn,
                                         string p_strScenarioId, 
                                         OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
             p_oOptimizerScenarioItem.OwnerGroupCodeList="";
            string strSQL = "SELECT * FROM scenario_land_owner_groups WHERE " +
                " scenario_id = '" + p_strScenarioId + "';";
            p_oAdo.SqlQueryReader(p_oConn, strSQL);


            if (p_oAdo.m_intError == 0)
            {

                //load step one with wind class speed definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    if (p_oAdo.m_OleDbDataReader["owngrpcd"] != System.DBNull.Value)
                    {
                        p_oOptimizerScenarioItem.OwnerGroupCodeList =
                            p_oOptimizerScenarioItem.OwnerGroupCodeList +
                            Convert.ToString(p_oAdo.m_OleDbDataReader["owngrpcd"]).Trim() + ",";
                    }
                }
                p_oAdo.m_OleDbDataReader.Close();
                p_oAdo.m_OleDbDataReader = null;
                p_oAdo.m_OleDbCommand = null;
                if (p_oOptimizerScenarioItem.OwnerGroupCodeList.Trim().Length > 0)
                    p_oOptimizerScenarioItem.OwnerGroupCodeList =
                        p_oOptimizerScenarioItem.OwnerGroupCodeList.Substring(0, p_oOptimizerScenarioItem.OwnerGroupCodeList.Length - 1);
            }
        }
        public void LoadTransportationCosts(ado_data_access p_oAdo,
                                         System.Data.OleDb.OleDbConnection p_oConn,
                                         string p_strScenarioId,
                                         OptimizerScenarioItem p_oOptimizerScenarioItem)
        {
            p_oOptimizerScenarioItem.m_oTranCosts.RailChipTransferPerGreenTon = "";
            p_oOptimizerScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile = "";
            p_oOptimizerScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTon = "";
            p_oOptimizerScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour="";

            
            string strSQL = "SELECT * FROM scenario_costs WHERE " +
                " TRIM(UCASE(scenario_id)) = '" + p_strScenarioId.Trim().ToUpper() + "';";
            p_oAdo.SqlQueryReader(p_oConn, strSQL);

            if (p_oAdo.m_intError == 0)
            {

                //load step one with wind class speed definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    if (p_oAdo.m_OleDbDataReader["road_haul_cost_pgt_per_hour"] != System.DBNull.Value)
                    {
                        p_oOptimizerScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["road_haul_cost_pgt_per_hour"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["rail_haul_cost_pgt_per_mile"] != System.DBNull.Value)
                    {
                        p_oOptimizerScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["rail_haul_cost_pgt_per_mile"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["rail_chip_transfer_pgt"] != System.DBNull.Value)
                    {
                        p_oOptimizerScenarioItem.m_oTranCosts.RailChipTransferPerGreenTon =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["rail_chip_transfer_pgt"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["rail_merch_transfer_pgt"] != System.DBNull.Value)
                    {
                        p_oOptimizerScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTon =
                           Convert.ToString(p_oAdo.m_OleDbDataReader["rail_merch_transfer_pgt"]).Trim();
                    }

                }
                p_oAdo.m_OleDbDataReader.Close();
               
            }
        }


        public System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> LoadFvsTablesAndVariables(FIA_Biosum_Manager.ado_data_access p_oAdo)
        {
            int x, y;
            RxTools oRxTools = new RxTools();

            string strTargetMdb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb";
            
            //
            //delete the old table links first in case any are obsolete
            //
            if (System.IO.File.Exists(strTargetMdb))
            {
                oRxTools.DeleteTableLinksToFVSPrePostTables(strTargetMdb);
            }
            
            //
            //load list box with all the pre and post table columns
            //
            oRxTools.CreateTableLinksToFVSPrePostTables(strTargetMdb);
            oRxTools = null;
            p_oAdo.OpenConnection(p_oAdo.getMDBConnString(strTargetMdb, "", ""));
            System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> _dictFVSTables = 
                new System.Collections.Generic.Dictionary<string,
                System.Collections.Generic.IList<string>>();
            if (p_oAdo.m_intError == 0)
            {

                string[] strTableNamesArray = p_oAdo.getTableNames(p_oAdo.m_OleDbConnection);
                for (x = 0; x <= strTableNamesArray.Length - 1; x++)
                {
                    if (strTableNamesArray[x].ToUpper().IndexOf("PRE_", 0) == 0)
                    {
                        string strColumnNamesList = "";
                        string strDataTypesList = "";
                        p_oAdo.getFieldNamesAndDataTypes(p_oAdo.m_OleDbConnection, "SELECT * FROM " + strTableNamesArray[x], 
                            ref strColumnNamesList, ref strDataTypesList);
                        string[] strColumnNamesArray = new string[0];
                        string[] strDataTypesArray = new string[0];
                        if (!String.IsNullOrEmpty(strColumnNamesList))
                        {
                            strColumnNamesArray = strColumnNamesList.Split(",".ToCharArray());
                            strDataTypesArray = strDataTypesList.Split(",".ToCharArray());
                        }
                        System.Collections.Generic.IList<string> lstFVSFields = new System.Collections.Generic.List<string>();
                        for (y = 0; y <= strColumnNamesArray.Length - 1; y++)
                        {
                            switch (strColumnNamesArray[y].Trim().ToUpper())
                            {
                                case "BIOSUM_COND_ID":
                                    break;
                                case "RXPACKAGE":
                                    break;
                                case "RX":
                                    break;
                                case "RXCYCLE":
                                    break;
                                case "STANDID":
                                    break;
                                case "ID":
                                    break;
                                case "CASEID":
                                    break;
                                case "FVS_VARIANT":
                                    break;
                                case "YEAR":
                                    break;
                                default:
                                    // Text data types can't be have a weight applied
                                    if (!strDataTypesArray[y].Trim().ToUpper().Equals("SYSTEM.STRING"))
                                    {
                                        if (p_oAdo.ColumnExist(p_oAdo.m_OleDbConnection,
                                            "POST_" + strTableNamesArray[x].Substring(4, strTableNamesArray[x].Length - 4),
                                            strColumnNamesArray[y]))
                                        {
                                            lstFVSFields.Add(strColumnNamesArray[y]);
                                        }
                                    }
                                    break;
                            }

                        }
                        if (lstFVSFields.Count > 0)
                        {
                            string strFvsTableName = strTableNamesArray[x].Substring(4, strTableNamesArray[x].Length - 4);
                            if (!_dictFVSTables.ContainsKey(strFvsTableName))
                            {
                                _dictFVSTables.Add(strFvsTableName, lstFVSFields);
                            }
                            else
                            {
                                System.Collections.Generic.List<string> lstTemp = (System.Collections.Generic.List<string>) _dictFVSTables[strFvsTableName];
                                lstTemp.AddRange(lstFVSFields);
                            }
                        }
                    }
                }
            }
            return _dictFVSTables;
        }

        public void LoadWeightedVariables(ado_data_access p_oAdo, FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.Variable_Collection p_oWeightedVariableCollection)
        {
            p_oWeightedVariableCollection.Clear();
            p_oAdo.OpenConnection(p_oAdo.getMDBConnString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultDbFile, "", ""));
            p_oAdo.m_strSQL = "SELECT * " +
                              "FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName;
            p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);

            if (p_oAdo.m_intError == 0)
            {
                //load step one with wind class speed definitions
                while (p_oAdo.m_OleDbDataReader.Read())
                {
                    FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.VariableItem oItem = new uc_optimizer_scenario_calculated_variables.VariableItem();
                    oItem.intId = Convert.ToInt32(p_oAdo.m_OleDbDataReader["id"]);
                    oItem.strVariableName = Convert.ToString(p_oAdo.m_OleDbDataReader["variable_name"]).Trim();
                    oItem.strVariableType = Convert.ToString(p_oAdo.m_OleDbDataReader["VARIABLE_TYPE"]).Trim();
                    if (p_oAdo.m_OleDbDataReader["variable_description"] != System.DBNull.Value)
                    {
                        oItem.strVariableDescr = Convert.ToString(p_oAdo.m_OleDbDataReader["variable_description"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["BASELINE_RXPACKAGE"] != System.DBNull.Value)
                    {
                        oItem.strRxPackage = Convert.ToString(p_oAdo.m_OleDbDataReader["BASELINE_RXPACKAGE"]).Trim();
                    }
                    if (p_oAdo.m_OleDbDataReader["VARIABLE_SOURCE"] != System.DBNull.Value)
                    {
                        oItem.strVariableSource = Convert.ToString(p_oAdo.m_OleDbDataReader["VARIABLE_SOURCE"]).Trim();
                    }
                    p_oWeightedVariableCollection.Add(oItem);
                }
            }
        }

        public void loadEconomicVariableWeights(FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables.VariableItem p_oWeightedVariable)
        {
            if (p_oWeightedVariable != null)
            {
                ado_data_access oAdo = new ado_data_access();
                string strEconConn = oAdo.getMDBConnString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultDbFile, "", "");
                using (var econConn = new OleDbConnection(strEconConn))
                {
                    econConn.Open();
                    string strSql = "SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName +
                                    " WHERE calculated_variables_id = " + p_oWeightedVariable.intId +
                                    " ORDER BY rxcycle";
                    oAdo.SqlQueryReader(econConn, strSql);
                    if (oAdo.m_intError == 0)
                    {
                        System.Collections.Generic.IList<double> lstEconWeights = new System.Collections.Generic.List<double>();
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            lstEconWeights.Add(Convert.ToDouble(oAdo.m_OleDbDataReader["weight"]));
                        }
                        p_oWeightedVariable.lstWeights = lstEconWeights;
                    }
                }
            }
        }

        public string ScenarioProperties(OptimizerScenarioItem p_oScenarioItem)
        {
            int x,y;
            bool bUseRxIntensity=false;
            string strTranDef="";
            string strBioDef="";
            string strLine = "";
            FIA_Biosum_Manager.ValidateNumericValues oValidate = new ValidateNumericValues();
            strLine = "***TREATMENT OPTIMIZER SCENARIO***\r\n\r\n";
            strLine = strLine + "ID\r\n";
            strLine = strLine + "-----\r\n";
            strLine = strLine + p_oScenarioItem.ScenarioId + "\r\n\r\n";
            strLine = strLine + "Description\r\n";
            strLine = strLine + "---------------------\r\n";
            strLine = strLine + p_oScenarioItem.Description + "\r\n\r\n";
            strLine = strLine + "FVS Analysis Cycle 1\r\n";
            strLine = strLine + "-------------------------------\r\n";
           
            strLine = strLine + "//\r\n";
            strLine = strLine + "//EFFECTIVE TREATMENT EVALUATION\r\n";
            strLine = strLine + "//\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).RxCycle == "1")
                {
                    for (y = 0; y <= OptimizerScenarioItem.EffectiveVariablesItem.NUMBER_OF_VARIABLES - 1; y++)
                    {
                        strLine = strLine + "-Stand Attribute " + Convert.ToString(y + 1) + "-\r\n";
                        strLine = strLine + "Pre-Treatment Variable: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPreVarArray[y] + "\r\n";
                        strLine = strLine + "Post-Treatment Variable: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strPostVarArray[y] + "\r\n";
                        strLine = strLine + "Treatment Improvement Expression: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strBetterExpr[y] + "\r\n";
                        strLine = strLine + "Treatment Disimprovement Expression: " +
                            p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strWorseExpr[y] + "\r\n";
                        strLine = strLine + "Treatment Effective Expression: " +
                        p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strEffectiveExpr[y] + "\r\n";

                        strLine = strLine + "\r\n\r\n";

                    }
                    strLine = strLine + "Overall Effective Expression: " +
                         p_oScenarioItem.m_oEffectiveVariablesItem_Collection.Item(x).m_strOverallEffectiveExpr + "\r\n";
                }
            }
            strLine = strLine + "//\r\n";
            strLine = strLine + "//OPTIMIZATION VARIABLE\r\n";
            strLine = strLine + "//\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oOptimizationVariableItem_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).RxCycle == "1" &&
                    p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).bSelected)
                {
                    
                        
                        strLine = strLine + "Optimization Variable Name: " +
                            p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strOptimizedVariable + "\r\n";
                        strLine = strLine + "FVS Variable Name: " +
                            p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strFVSVariableName + "\r\n";
                        strLine = strLine + "Value Source: " +
                            p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strValueSource + "\r\n";
                        strLine = strLine + "Aggregate Search: ";
                        if (p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strMaxYN=="Y")
                              strLine = strLine + "Maximum\r\n";
                        else 
                              strLine = strLine + "Minimum\r\n";
                        if (p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).bUseFilter)
                        {
                            strLine = strLine + "Net Revenue Filter: ";
                            strLine = strLine + p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).strFilterOperator + " " + 
                                p_oScenarioItem.m_oOptimizationVariableItem_Collection.Item(x).dblFilterValue.ToString().Trim();
                        }
                         
                       

                        strLine = strLine + "\r\n";

                }
            }
            strLine = strLine + "//\r\n";
            strLine = strLine + "//TIE BREAKERS\r\n";
            strLine = strLine + "//\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oTieBreaker_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oTieBreaker_Collection.Item(x).RxCycle == "1" &&
                    p_oScenarioItem.m_oTieBreaker_Collection.Item(x).bSelected)
                {
                     strLine = strLine + "Method: " +
                       p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strMethod + "\r\n";

                     if (p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper() ==
                          "LAST TIE-BREAK RANK")
                     {
                         bUseRxIntensity = true;
                     }
                     else
                     {
                         strLine = strLine + "FVS Variable Name: " +
                             p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strFVSVariableName + "\r\n";
                         strLine = strLine + "Value Source: " +
                             p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strValueSource + "\r\n";
                     }
                     strLine = strLine + "Aggregate Search: ";
                     if (p_oScenarioItem.m_oTieBreaker_Collection.Item(x).strMaxYN == "Y")
                         strLine = strLine + "Maximum\r\n";
                     else
                         strLine = strLine + "Minimum\r\n";

                     
                    strLine = strLine + "\r\n\r\n";

                }
            }
            if (bUseRxIntensity)
            {
                strLine = strLine + "\r\nLast Tie-Break Rank Assignments\r\n";
                strLine = strLine + "------------------------------------\r\n";
                strLine = strLine + "--RxPackage-- --Last_TieBreak_Rank--\r\n";
                
                for (x = 0; x <= p_oScenarioItem.m_oLastTieBreakRankItem_Collection.Count - 1; x++)
                {
                    strLine = strLine + String.Format("{0,-3}{1,11}",
                              " " + p_oScenarioItem.m_oLastTieBreakRankItem_Collection.Item(x).RxPackage,
                              p_oScenarioItem.m_oLastTieBreakRankItem_Collection.Item(x).LastTieBreakRank) + "\r\n";
                }
            }
            strLine = strLine + "\r\n";
            strLine = strLine + "Processing Sites\r\n";
            strLine = strLine + "----------------\r\n";
            strLine = strLine + "--Id-- -------------Name------------- ---------------------Transportation--------------------- ---------Biomass--------\r\n";
            for (x = 0; x <= p_oScenarioItem.m_oProcessingSiteItem_Collection.Count - 1; x++)
            {
                if (p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).Selected)
                {
                    strTranDef = p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).m_strTranCdDescArray[
                        Convert.ToInt32(p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).TransportationCode) - 1, 1];
                    strBioDef = p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).m_strBioCdDescArray[
                        Convert.ToInt32(p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).BiomassCode) - 1, 1];
                    strLine = strLine + String.Format("{0,-7}{1,-31}{2,-57}{3,-53}",
                                  " " + p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteId,
                                  p_oScenarioItem.m_oProcessingSiteItem_Collection.Item(x).ProcessingSiteName,
                                  strTranDef, strBioDef) + "\r\n";
                }
            }

            strLine = strLine + "\r\n";
            strLine = strLine + "Transportation Costs\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "Road Haul Cost Per Green Ton Per Hour: ";
            oValidate.Money = true;
            oValidate.RoundDecimalLength = 2;
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RoadHaulCostPerGreenTonPerHour);
            if (oValidate.m_intError == 0)
                strLine=strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }
            strLine = strLine + "Rail Haul Cost Per Green Ton Per Mile: ";
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RailHaulCostPerGreenTonPerMile);
            if (oValidate.m_intError == 0)
                strLine = strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }
            strLine = strLine + "Transfer Merch To Rail Per Green Ton: ";
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RailMerchTransferPerGreenTon);
            if (oValidate.m_intError == 0)
                strLine = strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }
            strLine = strLine + "Transfer Chips To Rail Per Green Ton: ";
            oValidate.ValidateDecimal(p_oScenarioItem.m_oTranCosts.RailChipTransferPerGreenTon);
            if (oValidate.m_intError == 0)
                strLine = strLine + oValidate.ReturnValue + "\r\n";
            else
            {
                strLine = strLine + "NA\r\n";
            }

            strLine = strLine + "\r\n";
            strLine = strLine + "Land Ownership Groups Filter\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "--Cd-- --------Description--------\r\n";
            string[] strOwnCdArray = frmMain.g_oUtils.ConvertListToArray(p_oScenarioItem.OwnerGroupCodeList,",");
            if (strOwnCdArray != null)
            {
                for (x = 0; x <= strOwnCdArray.Length - 1; x++)
                {
                    if (strOwnCdArray[x].Trim() == "40" ||
                        strOwnCdArray[x].Trim() == "4")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "Private") + "\r\n";
                    }
                    else if (strOwnCdArray[x].Trim() == "30" ||
                        strOwnCdArray[x].Trim() == "3")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "State and Local Government") + "\r\n";
                    }
                    else if (strOwnCdArray[x].Trim() == "20" ||
                            strOwnCdArray[x].Trim() == "2")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "Federal (Non-USFS)") + "\r\n";
                    }
                    else if (strOwnCdArray[x].Trim() == "10" ||
                            strOwnCdArray[x].Trim() == "1")
                    {
                        strLine = strLine + string.Format("{0,-6}{1,-20}", "  " + strOwnCdArray[x], "U.S. Forest Service") + "\r\n";
                    }


                }
            }
            strLine = strLine + "\r\n";
            strLine = strLine + "Condition Table SQL Filter\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "SQL: " + p_oScenarioItem.m_oCondTableSQLFilter.SQL + "\r\n";
            strLine = strLine + "Maximum Low Slope Yarding Distance (Feet): " + p_oScenarioItem.m_oCondTableSQLFilter.LowSlopeMaximumYardingDistanceFeet + "\r\n"; 
            strLine = strLine + "Maximum Steep Slope Yarding Distance (Feet): " + p_oScenarioItem.m_oCondTableSQLFilter.SteepSlopeMaximumYardingDistanceFeet + "\r\n";

            strLine = strLine + "\r\n";
            strLine = strLine + "Plot Table SQL Filter\r\n";
            strLine = strLine + "----------------------------\r\n";
            strLine = strLine + "SQL: " + p_oScenarioItem.PlotTableSQLFilter;


            strLine = strLine + "\r\n\r\nEOF";

            return strLine;

        }

        public string AuditWeightedFvsVariables (string strTableName, out int intError)
        {
            intError = 0;
            string strErrorMessage = "";
            ado_data_access oAdo = new ado_data_access();
            int intFvsPreTableCount = -1;
            int intWeightedPreTableCount = -1;
            char[] preArray = "PRE_".ToCharArray();
            string strNamePart1 = strTableName.TrimStart(preArray);
            string strName = strNamePart1.Substring(0,(strNamePart1.Length - "_WEIGHTED".Length));
            string strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\fvs\\db\\PREPOST_" + strName + ".ACCDB";
            string strCalcConn = oAdo.getMDBConnString(strFvsPrePostDb, "", "");
            string strSql = "SELECT Count(*) AS N FROM (" +
                            "SELECT DISTINCT biosum_cond_id, rxpackage, fvs_variant" +
                            " from PRE_" + strName + ")";
            using (var oCalcConn = new OleDbConnection(strCalcConn))
            {
                oCalcConn.Open();
                intFvsPreTableCount = oAdo.getRecordCount(oCalcConn, strSql, "PRE_" + strName);
            }
            strFvsPrePostDb = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory +
                "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            strSql = "SELECT Count(*) AS N FROM (" +
                "SELECT DISTINCT biosum_cond_id, rxpackage, fvs_variant" +
                " from PRE_" + strTableName + ")";
            strCalcConn = oAdo.getMDBConnString(strFvsPrePostDb, "", "");
            using (var oCalcConn = new OleDbConnection(strCalcConn))
            {
                oCalcConn.Open();
                intWeightedPreTableCount = oAdo.getRecordCount(oCalcConn, strSql, strTableName);
            }

            if (intFvsPreTableCount != intWeightedPreTableCount)
            {
                intError = -1;
                strErrorMessage = "PRE_" + strName + " table has a different number of records (" + intFvsPreTableCount +
                    ") than " + strTableName + " (" + intWeightedPreTableCount + "). Weighted variables from this table" +
                    " cannot be used! \r\nConsider running the 'Recalculate All' tool from the 'Calculated Variables'" +
                    "screen to recalculate all weighted variable \r\n";
            }

            return strErrorMessage;
        }
    }

    public class GisTools
    {
        private FIA_Biosum_Manager.Datasource m_oProjectDs;
        private string m_strMasterTravelTime = Tables.TravelTime.DefaultTravelTimeTableName + "_m";
        private string m_strMasterPSite = Tables.TravelTime.DefaultProcessingSiteTableName + "_m";
        private string m_strPlotTableName = "";
        private string m_masterFolder = frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir;
        dao_data_access m_oDao;
        private string m_strTempAccdb;
        private SQLite.ADO.DataMgr _SQLite = new SQLite.ADO.DataMgr();
        public SQLite.ADO.DataMgr SQLite
        {
            get { return _SQLite; }
            set { _SQLite = value; }
        }

        public bool CheckForExistingData(string strReferenceProjectDirectory, out bool bTablesHaveData)
        {
            bool bExistingTables = false;
            bTablesHaveData = false;
            ado_data_access oAdo = new ado_data_access();

            // Load project data sources table
            m_oProjectDs = new Datasource();
            m_oProjectDs.m_strDataSourceMDBFile = strReferenceProjectDirectory + "\\db\\project.mdb";
            m_oProjectDs.m_strDataSourceTableName = "datasource";
            m_oProjectDs.m_strScenarioId = "";
            m_oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            m_oProjectDs.LoadTableRecordCount = false;
            m_oProjectDs.populate_datasource_array();

            int intTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.TravelTimes);
            string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            string strSQL = "SELECT count(*) FROM " + strTableName;

            // Check for travel times table and data
            if (strTableStatus == "F")
            {
                bExistingTables = true;
                string strTestConn = oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", "");
                using (var oTestConn = new OleDbConnection(strTestConn))
                {
                    oTestConn.Open();
                    int intRecordCount = oAdo.getRecordCount(oTestConn, strSQL, strTableName);
                    if (intRecordCount > 0)
                    {
                        bTablesHaveData = true;
                    }
                }
            }

            // If no travel times, check for psites table and data
            if (bExistingTables == false || bTablesHaveData == false)
            {
                intTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
                strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                if (strTableStatus == "F")
                {
                    bExistingTables = true;
                    string strTestConn = oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", "");
                    using (var oTestConn = new OleDbConnection(strTestConn))
                    {
                        oTestConn.Open();
                        strSQL = "SELECT count(*) FROM " + strTableName;
                        int intRecordCount = oAdo.getRecordCount(oTestConn, strSQL, strTableName);
                        if (intRecordCount > 0)
                        {
                            bTablesHaveData = true;
                        }
                    }
                }
            }

            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }
            return bExistingTables;
        }

        public bool CheckForExistingDataSqlite(string strReferenceProjectDirectory, out bool bTablesHaveData)
        {
            bool bExistingTables = false;
            bTablesHaveData = false;

            // Load project data sources table
            m_oProjectDs = new Datasource();
            m_oProjectDs.m_strDataSourceMDBFile = strReferenceProjectDirectory + "\\db\\project.mdb";
            m_oProjectDs.m_strDataSourceTableName = "datasource";
            m_oProjectDs.m_strScenarioId = "";
            m_oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            m_oProjectDs.LoadTableRecordCount = false;
            m_oProjectDs.populate_datasource_array();

            int intTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.TravelTimes);
            string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            string strSQL = $@"SELECT count(*) FROM {strTableName} LIMIT 1";

            // Check for travel times table and data
            if (strTableStatus == "F")
            {
                bExistingTables = true;
                string strTestConn = SQLite.GetConnectionString(m_oProjectDs.getFullPathAndFile(Datasource.TableTypes.TravelTimes));
                using (var oTestConn = new System.Data.SQLite.SQLiteConnection(strTestConn))
                {
                    oTestConn.Open();
                    long lngRecordCount = SQLite.getRecordCount(oTestConn, strSQL, strTableName);
                    if (lngRecordCount > 0)
                    {
                        bTablesHaveData = true;
                    }
                }
            }

            // If no travel times, check for psites table and data
            if (bExistingTables == false || bTablesHaveData == false)
            {
                intTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
                strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                if (strTableStatus == "F")
                {
                    bExistingTables = true;
                    string strTestConn = SQLite.GetConnectionString(m_oProjectDs.getFullPathAndFile(Datasource.TableTypes.ProcessingSites));
                    using (var oTestConn = new System.Data.SQLite.SQLiteConnection(strTestConn))
                    {
                        oTestConn.Open();
                        strSQL = $@"SELECT count(*) FROM {strTableName} LIMIT 1";
                        long lngRecordCount = SQLite.getRecordCount(oTestConn, strSQL, strTableName);
                        if (lngRecordCount > 0)
                        {
                            bTablesHaveData = true;
                        }
                    }
                }
            }
            return bExistingTables;
        }

        public bool BackupGisData()
        {
            string strFileSuffix = "_" + DateTime.Now.ToString("MMddyyyy");
            string strBackedUpMdb = "";
            bool bSuccess = false;
            // travel times
            int intTable = m_oProjectDs.getValidTableNameRow(Datasource.TableTypes.TravelTimes);
            string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strTableStatus == "F")
            {
                string strExtension = System.IO.Path.GetExtension(strFileName);
                string strNewFileName = System.IO.Path.GetFileNameWithoutExtension(strFileName) + strFileSuffix + strExtension;
                // Check to see if the backup database already exists; If it does, abort the process so user can delete
                if (System.IO.File.Exists(strDirectoryPath + "\\" + strNewFileName))
                {
                    MessageBox.Show("A backup database from today already exists: " + strFileName + strFileSuffix + ". Delete this database manually if you want to " +
                        "back up today's data again!!", "FIA BioSum");
                    return false;
                }
                System.IO.File.Copy(strDirectoryPath + "\\" + strFileName, strDirectoryPath + "\\" + strNewFileName);
                strBackedUpMdb = strDirectoryPath + "\\" + strFileName;
            }
 
            // processing sites
            intTable = m_oProjectDs.getValidTableNameRow(Datasource.TableTypes.ProcessingSites);
            strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            if (strTableStatus == "F")
            {
                if (strBackedUpMdb.Equals(strDirectoryPath + "\\" + strFileName))
                {
                    // Do nothing, we already made a backup for travel times
                }
                else
                {
                    string strExtension = System.IO.Path.GetExtension(strFileName);
                    string strNewFileName = System.IO.Path.GetFileNameWithoutExtension(strFileName) + strFileSuffix + strExtension;
                    // Check to see if the backup database already exists; If it does, abort the process so user can delete
                    if (System.IO.File.Exists(strDirectoryPath + "\\" + strNewFileName))
                    {
                        MessageBox.Show("A backup database from today already exists: " + strFileName + strFileSuffix + ". Delete this database manually if you want to " +
                            "back up today's data again!!", "FIA BioSum");
                        return false;
                    }
                    System.IO.File.Copy(strDirectoryPath + "\\" + strFileName, strDirectoryPath + "\\" + strNewFileName);
                }
            }

            bSuccess = true;
            return bSuccess;
        }

        public int LoadGisData()
        {
            ado_data_access oAdo = new ado_data_access();

            m_oProjectDs.populate_datasource_array();
            
            string[] arrTableTypes = { Datasource.TableTypes.TravelTimes, Datasource.TableTypes.ProcessingSites};
            string strTravelTimesTableName = "";
            string strPSitesTableName = "";
            int intRecordCount = -1;

            frmMain.g_sbpInfo.Text = "Creating travel_time and processing_site tables...Stand by";
            foreach (string strTableType in arrTableTypes)
            {
                // travel times
                int intTable = m_oProjectDs.getTableNameRow(strTableType);
                string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                string strLoadConn = oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", "");
                using (var oLoadConn = new OleDbConnection(strLoadConn))
                {
                    oLoadConn.Open();
                    if (strTableStatus == "F")
                    {
                        string strSql = "DROP TABLE " + strTableName;
                        oAdo.SqlNonQuery(oLoadConn, strSql);
                    }
                    if (oAdo.m_intError == 0)
                    {
                        if (strTableType.Equals(Datasource.TableTypes.TravelTimes))
                        {
                            frmMain.g_oTables.m_oTravelTime.CreateTravelTimeTable(oAdo, oLoadConn, strTableName);
                            strTravelTimesTableName = strTableName;
                        }
                        else
                        {
                            frmMain.g_oTables.m_oTravelTime.CreateProcessingSiteTable(oAdo, oLoadConn, strTableName);
                            strPSitesTableName = strTableName;
                        }

                    }
                }
            }

            if (oAdo.m_intError == 0)
            {
                frmMain.g_sbpInfo.Text = "Creating temp database and table links...Stand by";
                string strTempDb = CreateACCDBAndTableDataSourceLinks();
                string strLoadConn = oAdo.getMDBConnString(strTempDb, "", "");
                using (var oLoadConn = new OleDbConnection(strLoadConn))
                {
                    oLoadConn.Open();
                    // Sleep until the table link exists
                    do
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    while (!oAdo.TableExist(oLoadConn, strTravelTimesTableName));
                    
                    frmMain.g_sbpInfo.Text = "Loading travel_time table...Stand by";
                    string strSql = "INSERT INTO " + strTravelTimesTableName +
                                    " SELECT TRAVELTIME_ID, PSITE_ID, biosum_plot_id," +
                                    " COLLECTOR_ID, RAILHEAD_ID, TRAVEL_MODE, ONE_WAY_HOURS," +
                                    " PT.PLOT AS PLOT, PT.STATECD AS STATECD" +
                                    " FROM " + m_strPlotTableName + " PT" +
                                    " INNER JOIN " + m_strMasterTravelTime + " TT ON (PT.statecd = TT.STATECD) AND (PT.plot = TT.PLOT)";
                    oAdo.SqlNonQuery(oLoadConn, strSql);
                    strSql = "SELECT COUNT(*) FROM " + strTravelTimesTableName;
                    intRecordCount = oAdo.getRecordCount(oLoadConn, strSql, strTravelTimesTableName);

                    if (oAdo.m_intError == 0 && intRecordCount > 0)
                    {
                        frmMain.g_sbpInfo.Text = "Loading processing_site table...Stand by";
                        strSql = "INSERT into " + strPSitesTableName +
                                 " SELECT distinct p.psite_id, name, TRANCD, TRANCD_DEF, BIOCD," +
                                                             " BIOCD_DEF," +
                                                             " EXISTS_YN, LAT, LON, STATE, CITY, COUNTY, MILL_TYPE, STATUS" +
                                 " FROM " + m_strMasterPSite + " p" +
                                 " INNER JOIN " + strTravelTimesTableName + " tt ON p.PSITE_ID = tt.PSITE_ID" +
                                 " group by P.PSITE_ID, NAME, TRANCD, TRANCD_DEF, BIOCD, BIOCD_DEF, EXISTS_YN, LAT, LON, STATE, CITY, COUNTY, MILL_TYPE, STATUS";
                        oAdo.SqlNonQuery(oLoadConn, strSql);
                        strSql = "SELECT COUNT(*) FROM " + strPSitesTableName;
                        intRecordCount = oAdo.getRecordCount(oLoadConn, strSql, strPSitesTableName);
                    }
                }

            }

            
            

            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }

            return intRecordCount;
        }

        public int LoadSqliteGisData(bool bUpdatePlotYardingDist)
        {
            m_oProjectDs.populate_datasource_array();

            string[] arrTableTypes = { Datasource.TableTypes.TravelTimes, Datasource.TableTypes.ProcessingSites };
            string strTravelTimesTableName = "";
            string strPSitesTableName = "";
            int intRecordCount = -1;

            frmMain.g_sbpInfo.Text = "Creating travel_time and processing_site tables...Stand by";
            string strSQLiteConn = "";
            foreach (string strTableType in arrTableTypes)
            {
                // travel times
                int intTable = m_oProjectDs.getTableNameRow(strTableType);
                string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                strSQLiteConn = SQLite.GetConnectionString(strDirectoryPath + "\\" + strFileName);
                using (var oLoadConn = new System.Data.SQLite.SQLiteConnection(strSQLiteConn))
                {
                    oLoadConn.Open();
                    if (strTableStatus == "F")
                    {
                        string strSql = "DROP TABLE " + strTableName;
                        SQLite.SqlNonQuery(oLoadConn, strSql);
                    }
                    if (SQLite.m_intError == 0)
                    {
                        if (strTableType.Equals(Datasource.TableTypes.TravelTimes))
                        {
                            frmMain.g_oTables.m_oTravelTime.CreateSqliteTravelTimeTable(SQLite, oLoadConn, strTableName);
                            strTravelTimesTableName = strTableName;
                        }
                        else
                        {
                            frmMain.g_oTables.m_oTravelTime.CreateSqliteProcessingSiteTable(SQLite, oLoadConn, strTableName);
                            strPSitesTableName = strTableName;
                        }
                    }

                }
            }

            ado_data_access oAdo = new ado_data_access();
            bool bPlotRecords = false;
            if (SQLite.m_intError == 0)
            {
                frmMain.g_sbpInfo.Text = "Creating temp database and table links...Stand by";
                CreateDBAndTableDataSourceLinks();
                string strLoadConn = oAdo.getMDBConnString(m_strTempAccdb, "", "");
                string strSql = "";
                using (var oLoadConn = new OleDbConnection(strLoadConn))
                {
                    oLoadConn.Open();
                    // Sleep until the table link exists
                    do
                    {
                        System.Threading.Thread.Sleep(1000);
                    }
                    while (!oAdo.TableExist(oLoadConn, strTravelTimesTableName));

                    strSql = "INSERT INTO " + strTravelTimesTableName +
                                    " SELECT TRAVELTIME_ID, PSITE_ID, biosum_plot_id," +
                                    " COLLECTOR_ID, RAILHEAD_ID, TRAVEL_MODE, ONE_WAY_HOURS," +
                                    " PT.PLOT AS PLOT, PT.STATECD AS STATECD" +
                                    " FROM " + m_strPlotTableName + " PT" +
                                    " INNER JOIN " + m_strMasterTravelTime + " TT ON (PT.statecd = TT.STATECD) AND (PT.plot = TT.PLOT)";
                    frmMain.g_sbpInfo.Text = "Loading project travel_times table...Stand by";
                    oAdo.SqlNonQuery(oLoadConn, strSql);
                    strSql = "SELECT COUNT(*) FROM " + strTravelTimesTableName;
                    intRecordCount = oAdo.getRecordCount(oLoadConn, strSql, strTravelTimesTableName);
                    strSql = $@"SELECT COUNT(*) FROM(SELECT TOP 1 plot FROM {m_strPlotTableName})";
                    if (oAdo.getRecordCount(oLoadConn, strSql, m_strPlotTableName) > 0)
                    {
                        bPlotRecords = true;
                    } 
                }

                bool bExistsPlotGisTable = false;
                if (oAdo.m_intError == 0 && intRecordCount > 0)
                {
                    frmMain.g_sbpInfo.Text = "Loading project processing_site table...Stand by";
                    using (var oLoadConn = new System.Data.SQLite.SQLiteConnection(strSQLiteConn))
                    {
                        oLoadConn.Open();
                        strSql = $@"ATTACH '{m_masterFolder}\{Tables.TravelTime.DefaultMasterTravelTimeDbFile}' as p";
                        SQLite.SqlNonQuery(oLoadConn, strSql);
                        strSql = $@"INSERT into processing_site SELECT distinct source.psite_id, name, TRANCD, TRANCD_DEF, BIOCD, BIOCD_DEF, 
                                 EXISTS_YN, LAT, LON, STATE, CITY, COUNTY, MILL_TYPE, STATUS, '' AS NOTES 
                                 FROM p.processing_site source INNER JOIN travel_time tt ON source.PSITE_ID = tt.PSITE_ID 
                                 group by source.PSITE_ID, NAME, TRANCD, TRANCD_DEF, BIOCD, BIOCD_DEF, EXISTS_YN, LAT, LON, STATE, CITY, COUNTY, MILL_TYPE, STATUS";
                        SQLite.SqlNonQuery(oLoadConn, strSql);
                        strSql = "SELECT COUNT(*) FROM " + strPSitesTableName;
                        intRecordCount = (int) SQLite.getRecordCount(oLoadConn, strSql, strPSitesTableName);
                        if (SQLite.AttachedTableExist(oLoadConn, Tables.TravelTime.DefaultPlotGisTableName))
                        {
                            bExistsPlotGisTable = true;
                        }
                    }
                }
                if(!bExistsPlotGisTable)
                {
                    MessageBox.Show($@"The {Tables.TravelTime.DefaultPlotGisTableName} table is missing from the master travel times database. Plot yarding distances cannot be updated.", "FIA BioSum");
                }
                else if (oAdo.m_intError == 0 && bUpdatePlotYardingDist)
                {
                    //@ToDo: Code to update plot table
                    if (bExistsPlotGisTable && bPlotRecords)
                    {
                        frmMain.g_sbpInfo.Text = "Updating yarding distances on plot table...Stand by";
                        using (var oLoadConn = new OleDbConnection(strLoadConn))
                        {
                            oLoadConn.Open();
                            // Access wants a primary key so that we can update the plot table from plot_gis
                            oAdo.AddPrimaryKey(oLoadConn, Tables.TravelTime.DefaultPlotGisTableName, $@"{Tables.TravelTime.DefaultPlotGisTableName}_pk", "PLOT,STATECD");

                            // Run audit before performing the update
                            strSql = $@"INSERT INTO {Tables.TravelTime.DefaultGisPlotDistanceAuditTable} SELECT biosum_plot_id, PT.statecd, PT.plot, gis_yard_dist_ft, NEARDIST_FT
                                FROM {m_strPlotTableName} PT LEFT JOIN {Tables.TravelTime.DefaultPlotGisTableName} PG ON (PT.statecd = PG.statecd) AND (PT.plot = PG.PLOT)
                                WHERE NEARDIST_FT IS NULL";
                            oAdo.SqlNonQuery(oLoadConn, strSql);

                            strSql = $@"UPDATE {m_strPlotTableName} PT INNER JOIN {Tables.TravelTime.DefaultPlotGisTableName} PG ON 
                                PT.PLOT = PG.PLOT AND PT.STATECD = PG.STATECD SET GIS_YARD_DIST_FT = NEARDIST_FT";
                            oAdo.SqlNonQuery(oLoadConn, strSql);
                        }
                    }

                }

            }

            if (oAdo != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }



            return intRecordCount;
        }

        private string CreateACCDBAndTableDataSourceLinks()
        {
            string strTempMDB = "";
            //used to get the temporary random file name
            FIA_Biosum_Manager.utils oUtils = new FIA_Biosum_Manager.utils();
            env oEnv = new env();
            strTempMDB = oUtils.getRandomFile(oEnv.strTempDir, "accdb");

            //create a temporary mdb that will contain all 
            //the links to the scenario datasource tables
            if (m_oDao == null)
            {
                m_oDao = new dao_data_access();
            }
            m_oDao.CreateMDB(strTempMDB);

            //links to the three project tables we need
            string[] arrTableTypes = { Datasource.TableTypes.TravelTimes, Datasource.TableTypes.ProcessingSites, Datasource.TableTypes.Plot };
            foreach (string strTableType in arrTableTypes)
            {
                int intTable = m_oProjectDs.getTableNameRow(strTableType);
                string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                if (strTableStatus == "F")
                {
                    m_oDao.CreateTableLink(strTempMDB, strTableName, strDirectoryPath + "\\" + strFileName, strTableName);
                    if (strTableType.Equals(Datasource.TableTypes.Plot))
                    {
                        m_strPlotTableName = strTableName;
                    }
                }
            }
            
            // master databases            
            m_oDao.CreateTableLink(strTempMDB, m_strMasterTravelTime, m_masterFolder + "\\" + Tables.TravelTime.DefaultMasterTravelTimeAccdbFile, Tables.TravelTime.DefaultTravelTimeTableName);
            m_oDao.CreateTableLink(strTempMDB, m_strMasterPSite, m_masterFolder + "\\" + Tables.TravelTime.DefaultMasterTravelTimeAccdbFile, Tables.TravelTime.DefaultProcessingSiteTableName);

            if (m_oDao != null)
            {
                m_oDao.m_DaoWorkspace.Close();
                m_oDao = null;
            }
            return strTempMDB;
        }

        private void CreateDBAndTableDataSourceLinks()
        {
            //used to get the temporary random file name
            FIA_Biosum_Manager.utils oUtils = new FIA_Biosum_Manager.utils();
            env oEnv = new env();
            m_strTempAccdb = oUtils.getRandomFile(oEnv.strTempDir, "accdb");

            //create a temporary mdb that will contain all 
            //the links to the scenario datasource tables
            if (m_oDao == null)
            {
                m_oDao = new dao_data_access();
            }
            m_oDao.CreateMDB(m_strTempAccdb);

            // Create ODBC Data sources
            ODBCMgr odbcMgr = new ODBCMgr();
            if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.GisProjectDbDsnName))
            {
                odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.GisProjectDbDsnName);
            }
            if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.GisMasterDbDsnName))
            {
                odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.GisMasterDbDsnName);
            }
            if (odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.GisAuditDbDsnName))
            {
                odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.GisAuditDbDsnName);
            }
            int intTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.TravelTimes);
            string strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            string strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            if (strTableStatus == "F")
            {
                odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.GisProjectDbDsnName, m_oProjectDs.getFullPathAndFile(Datasource.TableTypes.TravelTimes));
            }
            m_oDao.CreateSQLiteTableLink(m_strTempAccdb, strTableName, strTableName, ODBCMgr.DSN_KEYS.GisProjectDbDsnName, m_oProjectDs.getFullPathAndFile(Datasource.TableTypes.TravelTimes));
            intTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
            strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            m_oDao.CreateSQLiteTableLink(m_strTempAccdb, strTableName, strTableName, ODBCMgr.DSN_KEYS.GisProjectDbDsnName, m_oProjectDs.getFullPathAndFile(Datasource.TableTypes.ProcessingSites));

            // Link to master travel times tables
            odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.GisMasterDbDsnName, $@"{m_masterFolder}\{Tables.TravelTime.DefaultMasterTravelTimeDbFile}");
            string strMasterConn = SQLite.GetConnectionString($@"{m_masterFolder}\{Tables.TravelTime.DefaultMasterTravelTimeDbFile}");
            using (var oMasterConn = new System.Data.SQLite.SQLiteConnection(strMasterConn))
            {
                oMasterConn.Open();
                if (SQLite.TableExist(oMasterConn, Tables.TravelTime.DefaultTravelTimeTableName))
                {
                    m_oDao.CreateSQLiteTableLink(m_strTempAccdb, Tables.TravelTime.DefaultTravelTimeTableName, m_strMasterTravelTime,
                        ODBCMgr.DSN_KEYS.GisMasterDbDsnName, $@"{m_masterFolder}\{Tables.TravelTime.DefaultMasterTravelTimeDbFile}");
                }
                if (SQLite.TableExist(oMasterConn, Tables.TravelTime.DefaultProcessingSiteTableName))
                {
                    m_oDao.CreateSQLiteTableLink(m_strTempAccdb, Tables.TravelTime.DefaultProcessingSiteTableName, m_strMasterPSite,
                        ODBCMgr.DSN_KEYS.GisMasterDbDsnName, $@"{m_masterFolder}\{Tables.TravelTime.DefaultMasterTravelTimeDbFile}");
                }
                if (SQLite.TableExist(oMasterConn, Tables.TravelTime.DefaultTravelTimeTableName))
                {
                    m_oDao.CreateSQLiteTableLink(m_strTempAccdb, Tables.TravelTime.DefaultPlotGisTableName, Tables.TravelTime.DefaultPlotGisTableName,
                        ODBCMgr.DSN_KEYS.GisMasterDbDsnName, $@"{m_masterFolder}\{Tables.TravelTime.DefaultMasterTravelTimeDbFile}");
                }
            }

            //links to the project table we need
            string[] arrTableTypes = { Datasource.TableTypes.Plot };
            foreach (string strTableType in arrTableTypes)
            {
                intTable = m_oProjectDs.getTableNameRow(strTableType);
                string strDirectoryPath = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                string strFileName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                strTableName = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                strTableStatus = m_oProjectDs.m_strDataSource[intTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                if (strTableStatus == "F")
                {
                    m_oDao.CreateTableLink(m_strTempAccdb, strTableName, strDirectoryPath + "\\" + strFileName, strTableName);
                    if (strTableType.Equals(Datasource.TableTypes.Plot))
                    {
                        m_strPlotTableName = strTableName;
                    }
                }
            }

            // Manage plot yarding distance audit table
            string strAuditDBPath = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory}\{Tables.TravelTime.DefaultGisAuditPathAndDbFile}";
            string strAuditConn = SQLite.GetConnectionString(strAuditDBPath);
            using (var oAuditConn = new System.Data.SQLite.SQLiteConnection(strAuditConn))
            {
                oAuditConn.Open();
                if (SQLite.TableExist(oAuditConn, Tables.TravelTime.DefaultGisPlotDistanceAuditTable))
                {
                    SQLite.SqlNonQuery(oAuditConn, $@"DROP TABLE {Tables.TravelTime.DefaultGisPlotDistanceAuditTable}");
                }
                frmMain.g_oTables.m_oTravelTime.CreateSqlitePlotDistanceAuditTable(SQLite, oAuditConn, Tables.TravelTime.DefaultGisPlotDistanceAuditTable);
                if (SQLite.TableExist(oAuditConn, Tables.TravelTime.DefaultGisPlotDistanceAuditTable))
                {
                    odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.GisAuditDbDsnName, strAuditDBPath);
                    m_oDao.CreateSQLiteTableLink(m_strTempAccdb, Tables.TravelTime.DefaultGisPlotDistanceAuditTable, Tables.TravelTime.DefaultGisPlotDistanceAuditTable,
                        ODBCMgr.DSN_KEYS.GisAuditDbDsnName, strAuditDBPath);
                }
            }

            if (m_oDao != null)
            {
                m_oDao.m_DaoWorkspace.Close();
                m_oDao = null;
            }
        }

        public int migrate_access_data()
        {
            // Check if Processor parameters in SQLite
            string strTest = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
            if (!System.IO.File.Exists(strTest))
            {
                MessageBox.Show("Processor parameters have not been migrated to SQLite. SQLite GIS data cannot be loaded!", "FIA Biosum");
                return -1;
            }
            // Check if Optimizer parameters in SQLite
            strTest = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}";
            if (!System.IO.File.Exists(strTest))
            {
                MessageBox.Show("Optimizer parameters have not been migrated to SQLite. SQLite GIS data cannot be loaded!", "FIA Biosum");
                return -1;
            }

            string gisPathAndDbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\" + Tables.TravelTime.DefaultTravelTimePathAndDbFile;
            if (!System.IO.File.Exists(gisPathAndDbFile))
            {
                SQLite.CreateDbFile(gisPathAndDbFile);
            }
            // Create audit db
           string strAuditDBPath = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory}\{Tables.TravelTime.DefaultGisAuditPathAndDbFile}";
            if (!System.IO.File.Exists(strAuditDBPath))
            {
                SQLite.CreateDbFile(strAuditDBPath);
            }

            // Create target tables in new database
            string strCopyConn = SQLite.GetConnectionString(gisPathAndDbFile);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strCopyConn))
            {
                conn.Open();
                frmMain.g_oTables.m_oTravelTime.CreateSqliteProcessingSiteTable(SQLite, conn, Tables.TravelTime.DefaultProcessingSiteTableName);
                frmMain.g_oTables.m_oTravelTime.CreateSqliteTravelTimeTable(SQLite, conn, Tables.TravelTime.DefaultTravelTimeTableName);
            }
            // Find path to existing tables
            m_oProjectDs = new Datasource();
            m_oProjectDs.m_strDataSourceMDBFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db\\project.mdb";
            m_oProjectDs.m_strDataSourceTableName = "datasource";
            m_oProjectDs.m_strScenarioId = "";
            m_oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            m_oProjectDs.LoadTableRecordCount = false;
            m_oProjectDs.populate_datasource_array();
            ado_data_access oAdo = new ado_data_access();

            // Travel times
            int intTravelTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.TravelTimes);
            int intPSitesTable = m_oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
            string strDirectoryPath = m_oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = m_oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.MDBFILE].Trim();
            //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
            string strTableName = m_oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strTableStatus = m_oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strCopyConn))
            {
                conn.Open();
                string[] arrUpdateTableType = new string[2];
                string[] arrUpdateTableName = new string[2];
                if (m_oProjectDs.DataSourceTableExist(intTravelTable))
                {
                    string strDbName = System.IO.Path.GetFileName(Tables.TravelTime.DefaultTravelTimePathAndDbFile);
                    string strNewDirectoryPath = System.IO.Path.GetDirectoryName(gisPathAndDbFile);
                    string strConn = oAdo.getMDBConnString(m_oProjectDs.getFullPathAndFile(Datasource.TableTypes.TravelTimes), "", "");
                    using (var pConn = new System.Data.OleDb.OleDbConnection(strConn))
                    {
                        pConn.Open();
                        oAdo.m_strSQL = $@"SELECT TRAVELTIME_ID, PSITE_ID, BIOSUM_PLOT_ID, COLLECTOR_ID,RAILHEAD_ID,
                        TRAVEL_MODE, ONE_WAY_HOURS, PLOT, STATECD FROM {strTableName}";
                        oAdo.CreateDataTable(pConn, oAdo.m_strSQL, strTableName, false);
                        using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(oAdo.m_strSQL, conn))
                        {
                            using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                            {
                                using (var transaction = conn.BeginTransaction())
                                {
                                    da.InsertCommand = cb.GetInsertCommand();
                                    int rows = da.Update(oAdo.m_DataTable);
                                    transaction.Commit();
                                }
                            }
                        }
                        m_oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.TravelTimes, strNewDirectoryPath, strDbName, strTableName);
                        arrUpdateTableType[0] = Datasource.TableTypes.TravelTimes;
                        arrUpdateTableName[0] = strTableName;
                        if (m_oProjectDs.DataSourceTableExist(intPSitesTable))
                        {
                            strTableName = m_oProjectDs.m_strDataSource[intPSitesTable, Datasource.TABLE].Trim();
                            oAdo.m_strSQL = $@"SELECT PSITE_ID,NAME,TRANCD,TRANCD_DEF,BIOCD,BIOCD_DEF,EXISTS_YN,LAT,LON,
                                            STATE,CITY,MILL_TYPE,COUNTY,STATUS,NOTES FROM {strTableName}";
                            oAdo.CreateDataTable(pConn, oAdo.m_strSQL, strTableName, false);
                            using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(oAdo.m_strSQL, conn))
                            {
                                using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                {
                                    using (var transaction = conn.BeginTransaction())
                                    {
                                        da.InsertCommand = cb.GetInsertCommand();
                                        int rows = da.Update(oAdo.m_DataTable);
                                        transaction.Commit();
                                    }
                                }
                            }
                            m_oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.ProcessingSites, strNewDirectoryPath, strDbName, strTableName);
                            arrUpdateTableType[1] = Datasource.TableTypes.ProcessingSites;
                            arrUpdateTableName[1] = strTableName;
                        }
                    }
                    strConn = SQLite.GetConnectionString($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}");
                    using (System.Data.SQLite.SQLiteConnection scenarioConn = new System.Data.SQLite.SQLiteConnection(strConn))
                    {
                        scenarioConn.Open();
                        for (int i = 0; i < arrUpdateTableType.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(arrUpdateTableType[i]))
                            {
                                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                sb.Append($@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET ");
                                sb.Append($@"PATH = '{strNewDirectoryPath}', file='{strDbName}', table_name = '{arrUpdateTableName[i]}' ");
                                sb.Append($@"WHERE TRIM(table_type) = '{arrUpdateTableType[i]}'");
                                SQLite.SqlNonQuery(scenarioConn, sb.ToString());
                            }
                        }
                    }
                    strConn = SQLite.GetConnectionString($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}");
                    using (System.Data.SQLite.SQLiteConnection scenarioConn = new System.Data.SQLite.SQLiteConnection(strConn))
                    {
                        scenarioConn.Open();
                        for (int i = 0; i < arrUpdateTableType.Length; i++)
                        {
                            if (!string.IsNullOrEmpty(arrUpdateTableType[i]))
                            {
                                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                sb.Append($@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET ");
                                sb.Append($@"PATH = '{strNewDirectoryPath}', file='{strDbName}', table_name = '{arrUpdateTableName[i]}' ");
                                sb.Append($@"WHERE TRIM(table_type) = '{arrUpdateTableType[i]}'");
                                SQLite.SqlNonQuery(scenarioConn, sb.ToString());
                            }
                        }
                    }
                }
            }
            return 0;
        }
    }
}

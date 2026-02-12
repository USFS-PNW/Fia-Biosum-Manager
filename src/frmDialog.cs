using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for frmDialog.
	/// </summary>
	public class frmDialog : System.Windows.Forms.Form
	{

		public FIA_Biosum_Manager.uc_select_list_item uc_select_list_item1;
		public FIA_Biosum_Manager.uc_project uc_project1;
		 public FIA_Biosum_Manager.uc_scenario uc_scenario1;
		public FIA_Biosum_Manager.frmDialog m_frmDialogCallingForm;
		public FIA_Biosum_Manager.frmOptimizerScenario m_frmScenarioCallingForm;
        public FIA_Biosum_Manager.frmProcessorScenario m_frmProcessorScenarioCallingForm;
		public FIA_Biosum_Manager.frmMain m_frmMain;
		public FIA_Biosum_Manager.uc_previous_expressions uc_previous_expressions1;
        public FIA_Biosum_Manager.uc_optimizer_scenario_calculated_variables uc_core_scenario_weighted_average1;
        public FIA_Biosum_Manager.uc_optimizer_load_gis_data uc_optimizer_load_gis_data1;
		public FIA_Biosum_Manager.uc_tree_volume_export uc_tree_volume_export1;
		public FIA_Biosum_Manager.txtDollarsAndCents m_txtMoney;
		public FIA_Biosum_Manager.txtNumeric m_txtNumeric;
		private System.Windows.Forms.TextBox _txtBox;
		private FIA_Biosum_Manager.txtNumeric _txtNumeric;
		private FIA_Biosum_Manager.txtDollarsAndCents _txtMoney;
		private System.Windows.Forms.ComboBox _cmbBox;
		public FIA_Biosum_Manager.uc_plot_add_edit uc_plot_add_edit1;
		public FIA_Biosum_Manager.uc_plot_input uc_plot_input1;
		public FIA_Biosum_Manager.uc_delete_conditions uc_delete_conditions;
		public FIA_Biosum_Manager.uc_delete_packages uc_delete_packages;
		public FIA_Biosum_Manager.uc_processor_scenario_tree_diam_groups_list uc_processor_scenario_tree_diam_groups_list1;
		public FIA_Biosum_Manager.uc_processor_scenario_tree_diam_groups_edit uc_processor_scenario_tree_diam_groups_edit1;
		public FIA_Biosum_Manager.uc_processor_scenario_tree_spc_groups uc_processor_scenario_tree_spc_groups1;
		public FIA_Biosum_Manager.uc_rx_list uc_rx_list1; 
		public FIA_Biosum_Manager.uc_rx_edit uc_rx_edit1;
		public FIA_Biosum_Manager.uc_rx_package_list uc_rx_package_list1;
		public FIA_Biosum_Manager.uc_fvs_input uc_fvs_input1;
        public FIA_Biosum_Manager.uc_fvs_output uc_fvs_output1;
		public FIA_Biosum_Manager.uc_processor_tree_spc uc_processor_tree_spc1;
		public FIA_Biosum_Manager.uc_processor_tree_spc_edit uc_processor_tree_spc_edit1;
        public FIA_Biosum_Manager.uc_db uc_db_1;
        //public FIA_Biosum_Manager.uc_scenario_harvest_cost_column_list uc_scenario_harvest_cost_column_list1;
        public FIA_Biosum_Manager.uc_scenario_harvest_cost_column_edit uc_scenario_harvest_cost_column_edit1;
		private bool _bDispose=false;
		private bool _bMinimizeMainForm=false;
		public FIA_Biosum_Manager.uc_filter_rows_text_datatype uc_filter_rows_text_datatype1;
		public FIA_Biosum_Manager.uc_filter_rows_numeric_datatype uc_filter_rows_numeric_datatype1;
		public FIA_Biosum_Manager.uc_gridview uc_gridview1;
        public FIA_Biosum_Manager.uc_optimizer_scenario_copy uc_scenario_optimizer_scenario_copy1;
        public FIA_Biosum_Manager.uc_processor_scenario_copy uc_processor_scenario_copy1;
        public FIA_Biosum_Manager.uc_fvs_output_prepost_seqnum uc_fvs_output_prepost_seqnum1=null;
        public Control _oParentControl = null;
        public System.Windows.Forms.FormWindowState _oLastWindowState;
		


		public string strCallingFormType;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        public frmDialog()
		{
            
			InitializeComponent();
            
            LastWindowState = this.WindowState;
            
            InitializeUserControls();
            

        }
		public frmDialog(FIA_Biosum_Manager.frmDialog p_frmDialogCallingForm)
		{
			
			strCallingFormType = "G";    //generic 
			InitializeComponent();
			this.m_frmDialogCallingForm = p_frmDialogCallingForm;
            LastWindowState = this.WindowState;

            InitializeUserControls();

			

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public frmDialog(FIA_Biosum_Manager.frmOptimizerScenario p_frmScenarioCallingForm, FIA_Biosum_Manager.frmMain p_frmMain)
		{
			strCallingFormType = "CS";   //core scenario
			InitializeComponent();
			this.m_frmScenarioCallingForm = p_frmScenarioCallingForm;
			this.m_frmMain = p_frmMain;
			InitializeUserControls();
		
		}
		public frmDialog(FIA_Biosum_Manager.frmMain p_frmMain)
		{
			strCallingFormType = "M";   //main form
			InitializeComponent();
			this.m_frmMain = p_frmMain;
			InitializeUserControls();
		}
		public frmDialog(FIA_Biosum_Manager.txtDollarsAndCents p_txtMoney)
		{
			this.m_txtMoney = p_txtMoney;
			strCallingFormType = "TM";   //main form
			InitializeComponent();
			InitializeUserControls();
		}
		public frmDialog(FIA_Biosum_Manager.txtNumeric p_txtNumeric)
		{
			this.m_txtNumeric = p_txtNumeric;
			strCallingFormType = "TN";   //main form
			InitializeComponent();
			InitializeUserControls();
		}
        public frmDialog(FIA_Biosum_Manager.frmProcessorScenario p_frmProcessorScenarioCallingForm, FIA_Biosum_Manager.frmMain p_frmMain)
        {
            strCallingFormType = "PS";   //processor scenario
            InitializeComponent();
            this.m_frmProcessorScenarioCallingForm = p_frmProcessorScenarioCallingForm;
            this.m_frmMain = p_frmMain;
            InitializeUserControls();

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
            this.SuspendLayout();
            // 
            // frmDialog
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(664, 326);
            this.Name = "frmDialog";
            this.Activated += new System.EventHandler(this.frmDialog_Activated);
            this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDialog_Closing);
            this.Resize += new System.EventHandler(this.frmDialog_Resize);
            this.ResumeLayout(false);

		}
		#endregion
		public void Initialize_User_Control(string strType)
		{
            
            if (strType.Equals("scenario"))
            {

                this.uc_scenario1.Visible = true;

                this.uc_scenario1.Width = 648;

                this.uc_scenario1.Height = 464;


                this.uc_scenario1.Top = (int)(this.ClientSize.Height * .50) - (int)(this.uc_scenario1.Height * .50);

                this.uc_scenario1.Left = (int)(this.ClientSize.Width * .50) - (int)(this.uc_scenario1.Width * .50);

                this.uc_scenario1.txtScenarioId.Visible = false;

                this.uc_scenario1.lblNewScenario.Visible = false;

                this.uc_scenario1.lblTitle.Text = "Open Scenario Or Create New Scenario";

                this.uc_project1.Visible = false;
            }
            else if (strType.Equals("PROJECT"))
            {

                this.uc_project1.lblTitle.Text = "Project Property Values";

                this.uc_project1.Dock = System.Windows.Forms.DockStyle.Fill;

                this.Text = "Project";

                this.uc_project1.Visible = true;

            }
            else
            {
            }
            
            //MessageBox.Show(this.strDialogType);
        }

		private void frmDialog_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{

            if (this.Text.Trim().ToUpper() == "PROCESSOR: TREE SPECIES GROUPS")
            {
                if (this.uc_processor_scenario_tree_spc_groups1 != null &&
                    this.uc_processor_scenario_tree_spc_groups1.btnSave.Enabled == true)
                {

                    uc_processor_scenario_tree_spc_groups1.btnClose_Click(sender, e);
                    if (this.uc_processor_scenario_tree_spc_groups1.btnSave.Enabled)
                    {
                        e.Cancel = true;
                        return;
                    }

                }

                this.Dispose();
                return;
            }
            if (uc_fvs_output_prepost_seqnum1 != null && uc_fvs_output_prepost_seqnum1.Exit==false)
            {
                uc_fvs_output_prepost_seqnum1.CloseForm();
                if (uc_fvs_output_prepost_seqnum1.Exit == false)
                {
                    e.Cancel = true;
                    return;
                }
            }
            if (this.uc_fvs_output1 != null && frmMain.g_oDelegate.CurrentThreadProcessIdle==false)
			{
				e.Cancel = true;
				return;
			}
            if (this.uc_fvs_input1 != null && frmMain.g_oDelegate.CurrentThreadProcessIdle == false)
            {
                e.Cancel = true;
                return;
            }
			if (this._bDispose == true)
			{
                if (this.uc_fvs_output1 != null) this.ParentControl.Enabled = true;
                if (this.uc_fvs_input1 != null) this.ParentControl.Enabled = true;
                if (this.ProcessorTreeSpcUserControl != null) this.ParentControl.Enabled = true;
                if (this.uc_rx_package_list1 != null) this.ParentControl.Enabled=true;
                if (this.uc_rx_list1 != null) this.ParentControl.Enabled = true;
                if (this.uc_plot_input1 != null) this.ParentControl.Enabled = true;
                if (this.uc_delete_conditions != null) this.ParentControl.Enabled = true;
                if (this.uc_delete_packages != null) this.ParentControl.Enabled = true;
				if (this.uc_optimizer_load_gis_data1 != null) this.ParentControl.Enabled = true;
				if (this.uc_tree_volume_export1 != null) this.ParentControl.Enabled = true;

				this.Dispose();
				return;
			}
			
			if (this.uc_previous_expressions1.Visible==true) this.uc_previous_expressions1.DeleteRecords();
			this.Visible=false;
			if (this.DialogResult == System.Windows.Forms.DialogResult.OK) 
			{
			}
			else 
			{
				e.Cancel = true;
			}

            
			
		}
		private void InitializeUserControls()
		{
            
            this.uc_select_list_item1 = new uc_select_list_item();
            
            this.uc_project1 = new uc_project();
            
            this.uc_previous_expressions1 = new uc_previous_expressions();
            
            this.Controls.Add(this.uc_select_list_item1);
            
            this.Controls.Add(this.uc_project1);
            
            this.Controls.Add(this.uc_previous_expressions1);
            

            this.uc_select_list_item1.Visible=false;
			this.uc_project1.Visible=false;
			this.uc_previous_expressions1.Visible=false;
			this.uc_previous_expressions1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.uc_select_list_item1.Dock = System.Windows.Forms.DockStyle.Fill;
		}

		private void frmDialog_Resize(object sender, System.EventArgs e)
		{

           

			if (this.MinimizeMainForm)
			{
				if (this.WindowState != System.Windows.Forms.FormWindowState.Minimized)
				{
					this.m_frmMain.WindowState=System.Windows.Forms.FormWindowState.Maximized;
				}
				else
				{
					this.m_frmMain.WindowState = System.Windows.Forms.FormWindowState.Minimized;
				}
				
			}
			
		
		}
        public void Initialize_FVS_Output_PREPOST_SeqNum_User_Control()
		{
			this.uc_fvs_output_prepost_seqnum1 = new uc_fvs_output_prepost_seqnum();
            this.Controls.Add(this.uc_fvs_output_prepost_seqnum1);
			uc_fvs_output_prepost_seqnum1.ReferenceDialog=this;
            this.Width = uc_fvs_output_prepost_seqnum1.Width + 10;
            this.Height = uc_fvs_output_prepost_seqnum1.Height + 40;
            uc_fvs_output_prepost_seqnum1.Dock = DockStyle.Fill;
            this.Text = "FVS: Define Table PRE/POST Sequence Numbers";
			uc_fvs_output_prepost_seqnum1.Visible=true;
		}
        //public void Initialize_Scenario_Harvest_Costs_Column_List_Control()
        //{
        //    this.uc_scenario_harvest_cost_column_list1 = new uc_scenario_harvest_cost_column_list();
        //    this.Controls.Add(this.uc_scenario_harvest_cost_column_list1);
        //    uc_scenario_harvest_cost_column_list1.ReferenceDialog=this;
        //    this.uc_scenario_harvest_cost_column_list1.Visible=true;
        //}
		public void Initialize_Scenario_Harvest_Costs_Column_Edit_Control()
		{
			this.uc_scenario_harvest_cost_column_edit1 = new uc_scenario_harvest_cost_column_edit();
			this.Controls.Add(this.uc_scenario_harvest_cost_column_edit1);
			this.uc_scenario_harvest_cost_column_edit1.Visible=true;
		}
        public void Initialize_Core_User_Variables_User_Control()
        {

            this.uc_core_scenario_weighted_average1 = new uc_optimizer_scenario_calculated_variables(this.m_frmMain);
            this.Controls.Add(this.uc_core_scenario_weighted_average1);
            this.uc_core_scenario_weighted_average1.Visible = true;

        }
        public bool Initialize_Load_Gis_Data_User_Control()
        {

            this.uc_optimizer_load_gis_data1 = new uc_optimizer_load_gis_data();
			if (this.uc_optimizer_load_gis_data1.bTerminateLoad)
			{
				return true;
			}
            else
            {
				this.Controls.Add(this.uc_optimizer_load_gis_data1);
				this.uc_optimizer_load_gis_data1.Visible = true;
				return false;
			}
        }
		public void Initialize_Tree_Volume_Export_User_Control(frmFCSTreeVolumeEdit frmTreeVolume)
		{
			this.uc_tree_volume_export1 = new uc_tree_volume_export(frmTreeVolume);
			this.Controls.Add(this.uc_tree_volume_export1);
			this.ParentControl = frmTreeVolume;
			this.uc_tree_volume_export1.Visible = true;
		}
		public void Initialize_Plot_Data_Add_Edit_User_Control()
		{

			this.uc_plot_add_edit1 = new uc_plot_add_edit();
			this.Controls.Add(this.uc_plot_add_edit1);
			this.uc_plot_add_edit1.Visible = true;
            
		}
		public void Initialize_Plot_Input_User_Control()
		{

			this.uc_plot_input1 = new uc_plot_input();
			this.Controls.Add(this.uc_plot_input1);
			this.uc_plot_input1.Visible = true;
            
		}

	    public void Initialize_Delete_Conditions_User_Control()
	    {
	        this.uc_delete_conditions = new uc_delete_conditions();
            this.Controls.Add(this.uc_delete_conditions);
	        this.uc_delete_conditions.Visible = true;
	    }

	    public void Initialize_Delete_Packages_User_Control()
	    {
            this.uc_delete_packages = new uc_delete_packages();
            this.Controls.Add(this.uc_delete_packages);
            this.uc_delete_packages.Visible = true;
	    }

		public void Initialize_Plot_Tree_Diam_User_Control()
		{

			this.uc_processor_scenario_tree_diam_groups_list1 = new uc_processor_scenario_tree_diam_groups_list();
			this.Controls.Add(this.uc_processor_scenario_tree_diam_groups_list1);
			this.uc_processor_scenario_tree_diam_groups_list1.Visible = true;
            this.uc_processor_scenario_tree_diam_groups_list1.ReferenceProcessorScenarioForm =
                this.m_frmProcessorScenarioCallingForm;
		}
		public void Initialize_Plot_Tree_Diam_Edit_User_Control()
		{

			this.uc_processor_scenario_tree_diam_groups_edit1 = new uc_processor_scenario_tree_diam_groups_edit();
			this.Controls.Add(this.uc_processor_scenario_tree_diam_groups_edit1);
			this.uc_processor_scenario_tree_diam_groups_edit1.Visible = true;
            
		}
		
		public void Initialize_Rx_User_Control()
		{

			this.uc_rx_list1 = new uc_rx_list();
			this.Controls.Add(this.uc_rx_list1);
			uc_rx_list1.ReferenceMainForm=this.m_frmMain;
			uc_rx_list1.ReferenceParentDialogForm=this;
			this.uc_rx_list1.Visible = true;
		}
		public void Initialize_Rx_Package_User_Control()
		{

			this.uc_rx_package_list1 = new uc_rx_package_list();
			this.Controls.Add(this.uc_rx_package_list1);
			this.uc_rx_package_list1.ReferenceMainForm=this.m_frmMain;
			this.uc_rx_package_list1.ReferenceParentDialogForm=this;
			this.uc_rx_package_list1.Visible = true;
		}
		public void Initialize_Filter_Rows_Text_Datatype_User_Control()
		{

			this.uc_filter_rows_text_datatype1 = new uc_filter_rows_text_datatype();
			this.Controls.Add(this.uc_filter_rows_text_datatype1);
			this.Width = this.uc_filter_rows_text_datatype1.Width + 10;
			this.Height = this.uc_filter_rows_text_datatype1.Height + 40;
			this.uc_filter_rows_text_datatype1.Visible=true;
            
		}

		public void Initialize_Filter_Rows_Numeric_Datatype_User_Control()
		{

			this.uc_filter_rows_numeric_datatype1 = new uc_filter_rows_numeric_datatype();
			this.Controls.Add(this.uc_filter_rows_numeric_datatype1);
			this.Width = this.uc_filter_rows_numeric_datatype1.Width + 10;
			this.Height = this.uc_filter_rows_numeric_datatype1.Height + 40;
			this.uc_filter_rows_numeric_datatype1.Visible=true;
            
		}
        public void Initialize_Scenario_Optimizer_Scenario_Copy(frmOptimizerScenario frmOptimizerScenario)
        {
            this.uc_scenario_optimizer_scenario_copy1 = new uc_optimizer_scenario_copy();
            this.Controls.Add(this.uc_scenario_optimizer_scenario_copy1);
            uc_scenario_optimizer_scenario_copy1.Dock = DockStyle.Fill;
            this.Width = this.uc_scenario_optimizer_scenario_copy1.Width + 10;
            this.Height = this.uc_scenario_optimizer_scenario_copy1.Height + 200;
            uc_scenario_optimizer_scenario_copy1.ReferenceDialogForm = this;
            this.uc_scenario_optimizer_scenario_copy1.ReferenceOptimizerScenarioForm =
                frmOptimizerScenario;
            uc_scenario_optimizer_scenario_copy1.Visible = true;

        }
        public void Initialize_Processor_Scenario_Copy(frmProcessorScenario frmProcessorScenario)
        {
            this.uc_processor_scenario_copy1 = new uc_processor_scenario_copy();
            this.Controls.Add(this.uc_processor_scenario_copy1);
            uc_processor_scenario_copy1.Dock = DockStyle.Fill;
            this.Width = this.uc_processor_scenario_copy1.Width + 10;
            this.Height = this.uc_processor_scenario_copy1.Height + 200;
            uc_processor_scenario_copy1.ReferenceDialogForm = this;
            this.uc_processor_scenario_copy1.ReferenceProcessorScenarioForm = frmProcessorScenario;
            uc_processor_scenario_copy1.Visible = true;

        }
        public void Initialize_Processor_Tree_Species_Groups()
        {
            this.uc_processor_scenario_tree_spc_groups1 = new uc_processor_scenario_tree_spc_groups();
            this.Controls.Add(this.uc_processor_scenario_tree_spc_groups1);
            this.uc_processor_scenario_tree_spc_groups1.Visible = true;
            this.uc_processor_scenario_tree_spc_groups1.ReferenceProcessorScenarioForm =
                this.m_frmProcessorScenarioCallingForm;

            // Set minimum size so that form can be resized larger, but not so small that controls are no longer visible
            Size minimumSize = new Size(this.uc_processor_scenario_tree_spc_groups1.Width + 20,
                                        this.uc_processor_scenario_tree_spc_groups1.Height + 40);
            this.MinimumSize = (minimumSize);
        }


		public void btnOK_Click(object sender, System.EventArgs e)
		{
		   this.DialogResult = System.Windows.Forms.DialogResult.OK;
		}
		public void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		}
		public string strTextValue
		{
			get 
			{
				string strTextValue="";
				switch (this.strCallingFormType.Trim())
				{
					case "TM":   //textbox money
						strTextValue = this.m_txtMoney.Text.Substring(1,this.m_txtMoney.Text.Trim().Length-1);
						break;
					case "TD":  //textbox decimal
						strTextValue = this.m_txtNumeric.Text ;
						break;
				}
				return strTextValue;
			}
        
		}
		public System.Windows.Forms.TextBox txtBox
		{
			set
			{
			     _txtBox = value;		
			}
			get {return _txtBox;}
		}
		public FIA_Biosum_Manager.txtNumeric txtNumeric
		{
			set
			{
				_txtNumeric = value;
			}
			get {return _txtNumeric;}
		}
		public FIA_Biosum_Manager.txtDollarsAndCents txtMoney
		{
			set
			{
				_txtMoney = value;
			}
			get {return _txtMoney;}
		}

		public System.Windows.Forms.ComboBox cmbBox
		{
			set
			{
				_cmbBox = value;		
			}
			get {return _cmbBox;}
		}
		public bool DisposeOfFormWhenClosing
		{
			set
			{
				this._bDispose = value;
			}
			get
			{
				return this._bDispose;
			}
		}
		public FIA_Biosum_Manager.uc_fvs_input FVSInputUserControl
		{
			set
			{
				this.uc_fvs_input1 = value;
			}
			get
			{
				return this.uc_fvs_input1;
			}
		}
        public FIA_Biosum_Manager.uc_fvs_output FvsOutProcessorInUserControl
		{
			set
			{
				this.uc_fvs_output1 = value;
			}
			get
			{
				return this.uc_fvs_output1;
			}
		}
		public FIA_Biosum_Manager.uc_processor_tree_spc ProcessorTreeSpcUserControl
		{
			set
			{
				this.uc_processor_tree_spc1 = value;
			}
			get
			{
				return this.uc_processor_tree_spc1;
			}
		}
		public FIA_Biosum_Manager.uc_processor_tree_spc_edit ProcessorTreeSpcEditUserControl
		{
			set
			{
				this.uc_processor_tree_spc_edit1 = value;
			}
			get
			{
				return this.uc_processor_tree_spc_edit1;
			}
		}
		
		public FIA_Biosum_Manager.uc_processor_scenario_tree_spc_groups TreeSpcGroupsUserControl
		{
			set
			{
				this.uc_processor_scenario_tree_spc_groups1 = value;
			}
			get
			{
				return this.uc_processor_scenario_tree_spc_groups1;
			}
		}
        public FIA_Biosum_Manager.uc_db DbSQLiteUserControl
        {
            set
            {
                this.uc_db_1 = value;
            }
            get
            {
                return this.uc_db_1;
            }
        }
        public bool MinimizeMainForm
		{
			set {_bMinimizeMainForm=value;}
			get {return _bMinimizeMainForm;}
		}
        public Control ParentControl
        {
            get { return _oParentControl; }
            set { _oParentControl = value; }
        }
        public System.Windows.Forms.FormWindowState LastWindowState
        {
            get { return _oLastWindowState; }
            set { _oLastWindowState = value; }
        }
        private void frmDialog_Activated(object sender, EventArgs e)
        {
            if (uc_fvs_output1 != null)
            {
                uc_fvs_output1.uc_fvs_output_Resize();

            }
            else if (uc_scenario_optimizer_scenario_copy1 != null)
            {

                uc_scenario_optimizer_scenario_copy1.panel1_Resize();
            }
        }

	}
	/// <summary>
	/// GroupBox Template that fills the form dialog box
	/// </summary>
	public class TemplateGroupBox : System.Windows.Forms.GroupBox
	{
		public TemplateGroupBox(FIA_Biosum_Manager.frmDialog p_frmDialog)
		{
			// 
			// groupBox1
			// 
			p_frmDialog.Controls.Add(this);
			this.Dock = System.Windows.Forms.DockStyle.Fill;
			this.Location = new System.Drawing.Point(0, 0);
			this.Name = "groupBox1";
			this.Size = new System.Drawing.Size(432, 368);
			this.TabIndex = 3;
			this.TabStop = false;
			this.Text="";
		}

	}
	public class TemplateTitle : System.Windows.Forms.Label
	{
		public TemplateTitle(System.Windows.Forms.GroupBox p_grpbox,int intTop,int intLeft, string strTitle)
		{
			// 
			// lblTitle
			// 
			this.Dock = System.Windows.Forms.DockStyle.Top;
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ForeColor = System.Drawing.Color.Green;
			this.Name = "lblTitle";
			int textWidth = (int)this.CreateGraphics().MeasureString(strTitle, this.Font).Width;
			int textHeight = (int)this.CreateGraphics().MeasureString(strTitle, this.Font).Height;
			this.Top = intTop;
			this.Left = intLeft;
			this.Height = textHeight;
			this.Width = textWidth;
			this.Text  = strTitle;
			p_grpbox.Controls.Add(this);
		}
	}
	public class TemplateOkCancelButtons
	{
		public System.Windows.Forms.Button btnOK;
		public System.Windows.Forms.Button btnCancel;
		public TemplateOkCancelButtons(FIA_Biosum_Manager.frmDialog p_frmDialog,
			System.Windows.Forms.GroupBox p_grpbox)

		{
			this.btnOK = new Button();
			this.btnCancel = new Button();

			// 
			// btnCancel
			// 
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(88, 48);
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(p_frmDialog.btnCancel_Click);
			// 
			// btnOK
			// 
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(88, 48);
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(p_frmDialog.btnOK_Click);

			p_grpbox.Controls.Add(this.btnOK);
			p_grpbox.Controls.Add(this.btnCancel);


		}
			                           

	}
	public class TemplateInputLabel : System.Windows.Forms.Label
	{
		public TemplateInputLabel(System.Windows.Forms.GroupBox p_grpbox,string strName,string strText)
		{
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = strName;
			this.Size = new System.Drawing.Size(184, 32);
			this.Text = strText;
			this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			p_grpbox.Controls.Add(this);
		}
	}

	public class TemplateTextBox : System.Windows.Forms.TextBox
	{
        public ValidateNumericValues m_oValidateNumericValues=null;
        private string m_strPreviousValue = "";


        public string TextBoxValue
        {
            get { return this.Text; }
            set { this.Text = value; }
        }

		public TemplateTextBox(System.Windows.Forms.GroupBox p_grpbox,string strName,string strValue)
		{
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = strName;
			this.Size = new System.Drawing.Size(184, 32);
			this.Text = strValue;
            this.m_strPreviousValue = strValue;
			p_grpbox.Controls.Add(this);
		}
        public TemplateTextBox(System.Windows.Forms.Control p_oControl, string strName, string strValue,bool p_bNumeric)
        {
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Name = strName;
            this.Size = new System.Drawing.Size(184, 32);
            this.Text = strValue;
            this.m_strPreviousValue = strValue;
            p_oControl.Controls.Add(this);
            if (p_bNumeric)
            {
                m_oValidateNumericValues = new ValidateNumericValues();
                this.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
                this.Leave += new System.EventHandler(TextBox_Leave);
            }
        }
        private void TextBox_Leave(object sender, EventArgs e)
        {
            if (m_oValidateNumericValues.Money)
            {
                this.m_oValidateNumericValues.ValidateDecimal(this.Text);
            }
            if (this.m_oValidateNumericValues.m_intError == 0)
            {
                this.Text = m_oValidateNumericValues.ReturnValue;
                this.m_strPreviousValue = this.Text;
            }
            else
            {
                this.Text = this.m_strPreviousValue;
                this.Focus();

            }

        }
        
	}

    

	public class TemplateListBox : System.Windows.Forms.ListBox
	{
		public TemplateListBox(System.Windows.Forms.GroupBox p_grpbox,string strName)
		{
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = strName;
			this.Size = new System.Drawing.Size(184, 32);
			p_grpbox.Controls.Add(this);
		}
	}


	public class TemplateComboBox : System.Windows.Forms.ComboBox
	{
		public TemplateComboBox(System.Windows.Forms.GroupBox p_grpbox,string strName,string strText)
		{
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = strName;
			this.Size = new System.Drawing.Size(48, 28);
			this.Text = strText;
			p_grpbox.Controls.Add(this);
		}
	}
	

}

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_ffe.
	/// </summary>
	public class uc_core_scenario_fvs_prepost_optimization : System.Windows.Forms.UserControl
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
		public System.Data.DataTable m_DataTable;
		public System.Data.DataRow m_DataRow;
		public int m_intError = 0;
		public string m_strError = "";
		private System.Windows.Forms.Button btnNext;
		private System.Windows.Forms.Button btnPrev;
		private string m_strCurrentIndexTypeAndClass;
		private string m_strOverallEffectiveExpression="";
		private FIA_Biosum_Manager.utils m_oUtils; 
		private bool _bTorchIndex=true;
		private bool _bCrownIndex=true;
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmCoreScenario _frmScenario=null;
		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker _uc_tiebreaker;

		
		private int m_intCurVar=-1;
		int m_intCurVariableDefinitionStepCount=1;
		string[] m_strUserNavigation=null;
		private System.Windows.Forms.Panel pnlFVSVariablesPrePostVariable;
		private System.Windows.Forms.GroupBox grpboxOptimization;
		private System.Windows.Forms.Panel pnlOptimization;
		private System.Windows.Forms.GroupBox grpboxOptimizationListValues;
		private System.Windows.Forms.GroupBox grpboxOptimizationEdit;
		private System.Windows.Forms.Button btnOptimizationEdit;
		private System.Windows.Forms.GroupBox grpboxOptimizationAudit;
		private System.Windows.Forms.Button btnOptimizationAudit;
		private System.Windows.Forms.ListView lvOptimizationListValues;
		private System.Windows.Forms.ColumnHeader lvColVariableName;
		private System.Windows.Forms.ColumnHeader lvColFVSVariableName;
		private System.Windows.Forms.ColumnHeader lvColChecked;
		private System.Windows.Forms.ColumnHeader lvColMinMax;
		private System.Windows.Forms.Label lblOptimizationVariable;
		private System.Windows.Forms.Button btnOptimiztionDone;
		private System.Windows.Forms.Button btnOptimiztionCancel;
		private System.Windows.Forms.RadioButton rdoOptimizationMaximum;
		private System.Windows.Forms.RadioButton rdoOptimizationMinimum;
		private System.Windows.Forms.TextBox txtOptimizationValue;
		private System.Windows.Forms.ComboBox cmbOptimizationOperator;

		public bool m_bSave=false;


		const int COLUMN_CHECKBOX=0;
		const int COLUMN_OPTIMIZE_VARIABLE=1;
		const int COLUMN_FVS_VARIABLE=2;
		const int COLUMN_VALUESOURCE=3;
		const int COLUMN_MAXMIN=4;
		const int COLUMN_USEFILTER=5;
		const int COLUMN_FILTER_OPERATOR=6;
		const int COLUMN_FILTER_VALUE=7;

		private System.Windows.Forms.GroupBox grpMaxMin;
		private System.Windows.Forms.GroupBox grpFilter;
		private System.Windows.Forms.CheckBox chkEnableFilter;
		private Variable_Collection m_oOldVariableCollection = new Variable_Collection();
		public Variable_Collection m_oSavVariableCollection = new Variable_Collection();
		private Variable_Collection m_oNewVariableCollection = new Variable_Collection();
		private FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables _oCurVar;
		private System.Windows.Forms.ColumnHeader lvColUseFilter;
		private System.Windows.Forms.ColumnHeader lvColFilterOperator;
		private System.Windows.Forms.ColumnHeader lvColFilterValue;
		public bool m_bFirstTime=true;
		private bool _bDisplayAuditMsg=true;
		private bool m_bIgnoreListViewItemCheck=false;
		private System.Windows.Forms.ColumnHeader lvColValueSource;
		private System.Windows.Forms.GroupBox grpboxOptimizationSettings;
		private System.Windows.Forms.GroupBox grpboxOptimizationSettingsPostPre;
		private System.Windows.Forms.ComboBox cmbOptimizationSettingsPostPreValue;
		private System.Windows.Forms.GroupBox grpboxOptimizationFVSVariable;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesOptimizationVariableValues;
		private System.Windows.Forms.Button btnFVSVariablesOptimizationVariableValues;
		private System.Windows.Forms.GroupBox grpFVSVariablesOptimizationVariableValuesSelected;
		private System.Windows.Forms.Label lblFVSVariablesOptimizationVariableValuesSelected;
		private System.Windows.Forms.Button btnOptimiztionPrev;
		private System.Windows.Forms.Button btnOptimizationFVSVariableClear;
		private System.Windows.Forms.Button btnOptimizationFVSVariableDone;
		private System.Windows.Forms.Button btnOptimizationFVSVariableCancel;
		private System.Windows.Forms.Button btnOptimizationFVSVariableNext;
		private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors=new ListViewAlternateBackgroundColors();
        private ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private Label label1;
        private GroupBox grpBoxOptimizationNetRevenue;
        private ComboBox cmbNetRevOptimzFilter;
        private TextBox txtRevenueDescr;
        private Label LblNetRevenu;
        private ListBox lstFVSFieldsList;
        private ListBox lstFVSTablesList;
        private string m_strLastValue = "";
        private GroupBox grpboxEconOptimizSettings;
        private Panel panel2;
        private GroupBox groupBox5;
        private Label label3;
        private CheckBox chkEnableEconFilter;
        private ComboBox cmbEconOptimizationOperator;
        private TextBox txtEconOptimizationValue;
        private GroupBox groupBox6;
        private RadioButton rdoEconOptimizeMinimum;
        private RadioButton rdoEconOptimizeMaximum;
        private Label lblEconomicAttribute;
        private Button BtnEconOptimizationDone;
        private Button btnEconOptimizationCancel;
        private ListBox lstEconVariables;
        private Button btnEconSelect;
        private TextBox txtEconAttribDescr;
        private Label label2;
        private TextBox txtEconRevenueDescr;
        private ComboBox cmbNetRevEconOptimzFilter;
        private TextBox txtOptVarDescr;
        private Label lblOptVarDescr;
        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> m_dictFVSTables;

		public class VariableItem
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
            public string strRevenueAttribute = "";

			public void Copy(VariableItem p_oSource,ref VariableItem p_oDest)
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
                p_oDest.strRevenueAttribute = p_oSource.strRevenueAttribute;
			}
		}

		public class Variable_Collection : System.Collections.CollectionBase
		{
			public Variable_Collection()
			{
				//
				// TODO: Add constructor logic here
				//
			}

			public void Add(FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.VariableItem m_oVariable)
			{
				// v�rify if object is not already in
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
			public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.VariableItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_optimization.VariableItem) List[Index];
			}
			public void Copy(Variable_Collection p_oSource,ref Variable_Collection p_oDest,bool p_bInitializeDest)
			{
				int x;
				if (p_bInitializeDest) p_oDest.Clear();
				for (x=0;x<=p_oSource.Count-1;x++)
				{
					VariableItem oItem = new VariableItem();
					oItem.Copy(p_oSource.Item(x),ref oItem);
					p_oDest.Add(oItem);

				}
			}


		}
	   
		
		

		

		public uc_core_scenario_fvs_prepost_optimization()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oUtils = new utils();
			this.m_strCurrentIndexTypeAndClass="";

			
			this.grpboxOptimizationSettings.Top = grpboxOptimization.Top;
			this.grpboxOptimizationSettings.Left = this.grpboxOptimization.Left;
			this.grpboxOptimizationSettings.Height = this.grpboxOptimization.Height;
			this.grpboxOptimizationSettings.Width = this.grpboxOptimization.Width;
			this.grpboxOptimizationFVSVariable.Top = grpboxOptimization.Top;
			this.grpboxOptimizationFVSVariable.Left = this.grpboxOptimization.Left;
			this.grpboxOptimizationFVSVariable.Height = this.grpboxOptimization.Height;
			this.grpboxOptimizationFVSVariable.Width = this.grpboxOptimization.Width;
            this.grpboxEconOptimizSettings.Top = grpboxOptimization.Top;
            this.grpboxEconOptimizSettings.Left = this.grpboxOptimization.Left;
            this.grpboxEconOptimizSettings.Height = this.grpboxOptimization.Height;
            this.grpboxEconOptimizSettings.Width = this.grpboxOptimization.Width;


			this.m_oLvRowColors.ReferenceListView = this.lvOptimizationListValues;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			if (frmMain.g_oGridViewFont != null) lvOptimizationListValues.Font = frmMain.g_oGridViewFont;
		    this.lvOptimizationListValues.Columns[COLUMN_CHECKBOX].Width = -2;


			this.grpboxOptimizationFVSVariable.Hide();
            this.grpboxEconOptimizSettings.Hide();


            m_oValidate.RoundDecimalLength = 0;
            m_oValidate.Money = false;
            m_oValidate.NullsAllowed = false;
            m_oValidate.TestForMaxMin = false;
            m_oValidate.MinValue = -1000;
            m_oValidate.TestForMin = true;

			

			

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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxEconOptimizSettings = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnEconSelect = new System.Windows.Forms.Button();
            this.txtEconAttribDescr = new System.Windows.Forms.TextBox();
            this.lstEconVariables = new System.Windows.Forms.ListBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtEconRevenueDescr = new System.Windows.Forms.TextBox();
            this.cmbNetRevEconOptimzFilter = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.chkEnableEconFilter = new System.Windows.Forms.CheckBox();
            this.cmbEconOptimizationOperator = new System.Windows.Forms.ComboBox();
            this.txtEconOptimizationValue = new System.Windows.Forms.TextBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.rdoEconOptimizeMinimum = new System.Windows.Forms.RadioButton();
            this.rdoEconOptimizeMaximum = new System.Windows.Forms.RadioButton();
            this.lblEconomicAttribute = new System.Windows.Forms.Label();
            this.BtnEconOptimizationDone = new System.Windows.Forms.Button();
            this.btnEconOptimizationCancel = new System.Windows.Forms.Button();
            this.grpboxOptimizationFVSVariable = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.grpboxFVSVariablesOptimizationVariableValues = new System.Windows.Forms.GroupBox();
            this.txtOptVarDescr = new System.Windows.Forms.TextBox();
            this.lblOptVarDescr = new System.Windows.Forms.Label();
            this.lstFVSFieldsList = new System.Windows.Forms.ListBox();
            this.lstFVSTablesList = new System.Windows.Forms.ListBox();
            this.btnFVSVariablesOptimizationVariableValues = new System.Windows.Forms.Button();
            this.grpFVSVariablesOptimizationVariableValuesSelected = new System.Windows.Forms.GroupBox();
            this.lblFVSVariablesOptimizationVariableValuesSelected = new System.Windows.Forms.Label();
            this.btnOptimizationFVSVariableClear = new System.Windows.Forms.Button();
            this.btnOptimizationFVSVariableDone = new System.Windows.Forms.Button();
            this.btnOptimizationFVSVariableCancel = new System.Windows.Forms.Button();
            this.btnOptimizationFVSVariableNext = new System.Windows.Forms.Button();
            this.grpboxOptimization = new System.Windows.Forms.GroupBox();
            this.pnlOptimization = new System.Windows.Forms.Panel();
            this.grpboxOptimizationListValues = new System.Windows.Forms.GroupBox();
            this.lvOptimizationListValues = new System.Windows.Forms.ListView();
            this.lvColChecked = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColVariableName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColFVSVariableName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColValueSource = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColMinMax = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColUseFilter = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColFilterOperator = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColFilterValue = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.grpboxOptimizationEdit = new System.Windows.Forms.GroupBox();
            this.btnOptimizationEdit = new System.Windows.Forms.Button();
            this.grpboxOptimizationAudit = new System.Windows.Forms.GroupBox();
            this.btnOptimizationAudit = new System.Windows.Forms.Button();
            this.grpboxOptimizationSettings = new System.Windows.Forms.GroupBox();
            this.pnlFVSVariablesPrePostVariable = new System.Windows.Forms.Panel();
            this.LblNetRevenu = new System.Windows.Forms.Label();
            this.grpBoxOptimizationNetRevenue = new System.Windows.Forms.GroupBox();
            this.txtRevenueDescr = new System.Windows.Forms.TextBox();
            this.cmbNetRevOptimzFilter = new System.Windows.Forms.ComboBox();
            this.btnOptimiztionPrev = new System.Windows.Forms.Button();
            this.grpboxOptimizationSettingsPostPre = new System.Windows.Forms.GroupBox();
            this.cmbOptimizationSettingsPostPreValue = new System.Windows.Forms.ComboBox();
            this.grpFilter = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkEnableFilter = new System.Windows.Forms.CheckBox();
            this.cmbOptimizationOperator = new System.Windows.Forms.ComboBox();
            this.txtOptimizationValue = new System.Windows.Forms.TextBox();
            this.grpMaxMin = new System.Windows.Forms.GroupBox();
            this.rdoOptimizationMinimum = new System.Windows.Forms.RadioButton();
            this.rdoOptimizationMaximum = new System.Windows.Forms.RadioButton();
            this.lblOptimizationVariable = new System.Windows.Forms.Label();
            this.btnOptimiztionDone = new System.Windows.Forms.Button();
            this.btnOptimiztionCancel = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.grpboxEconOptimizSettings.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.grpboxOptimizationFVSVariable.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpboxFVSVariablesOptimizationVariableValues.SuspendLayout();
            this.grpFVSVariablesOptimizationVariableValuesSelected.SuspendLayout();
            this.grpboxOptimization.SuspendLayout();
            this.pnlOptimization.SuspendLayout();
            this.grpboxOptimizationListValues.SuspendLayout();
            this.grpboxOptimizationEdit.SuspendLayout();
            this.grpboxOptimizationAudit.SuspendLayout();
            this.grpboxOptimizationSettings.SuspendLayout();
            this.pnlFVSVariablesPrePostVariable.SuspendLayout();
            this.grpBoxOptimizationNetRevenue.SuspendLayout();
            this.grpboxOptimizationSettingsPostPre.SuspendLayout();
            this.grpFilter.SuspendLayout();
            this.grpMaxMin.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.grpboxEconOptimizSettings);
            this.groupBox1.Controls.Add(this.grpboxOptimizationFVSVariable);
            this.groupBox1.Controls.Add(this.grpboxOptimization);
            this.groupBox1.Controls.Add(this.grpboxOptimizationSettings);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Leave += new System.EventHandler(this.groupBox1_Leave);
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // grpboxEconOptimizSettings
            // 
            this.grpboxEconOptimizSettings.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxEconOptimizSettings.Controls.Add(this.panel2);
            this.grpboxEconOptimizSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxEconOptimizSettings.ForeColor = System.Drawing.Color.Black;
            this.grpboxEconOptimizSettings.Location = new System.Drawing.Point(12, 1458);
            this.grpboxEconOptimizSettings.Name = "grpboxEconOptimizSettings";
            this.grpboxEconOptimizSettings.Size = new System.Drawing.Size(856, 448);
            this.grpboxEconOptimizSettings.TabIndex = 36;
            this.grpboxEconOptimizSettings.TabStop = false;
            this.grpboxEconOptimizSettings.Text = "Economic Attribute Optimization Settings";
            // 
            // panel2
            // 
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add(this.btnEconSelect);
            this.panel2.Controls.Add(this.txtEconAttribDescr);
            this.panel2.Controls.Add(this.lstEconVariables);
            this.panel2.Controls.Add(this.groupBox5);
            this.panel2.Controls.Add(this.groupBox6);
            this.panel2.Controls.Add(this.lblEconomicAttribute);
            this.panel2.Controls.Add(this.BtnEconOptimizationDone);
            this.panel2.Controls.Add(this.btnEconOptimizationCancel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 18);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(850, 427);
            this.panel2.TabIndex = 12;
            // 
            // btnEconSelect
            // 
            this.btnEconSelect.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEconSelect.Location = new System.Drawing.Point(486, 46);
            this.btnEconSelect.Name = "btnEconSelect";
            this.btnEconSelect.Size = new System.Drawing.Size(90, 66);
            this.btnEconSelect.TabIndex = 88;
            this.btnEconSelect.Text = "Select";
            this.btnEconSelect.Click += new System.EventHandler(this.btnEconSelect_Click);
            // 
            // txtEconAttribDescr
            // 
            this.txtEconAttribDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEconAttribDescr.Location = new System.Drawing.Point(220, 46);
            this.txtEconAttribDescr.Multiline = true;
            this.txtEconAttribDescr.Name = "txtEconAttribDescr";
            this.txtEconAttribDescr.ReadOnly = true;
            this.txtEconAttribDescr.Size = new System.Drawing.Size(242, 45);
            this.txtEconAttribDescr.TabIndex = 87;
            // 
            // lstEconVariables
            // 
            this.lstEconVariables.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstEconVariables.ItemHeight = 16;
            this.lstEconVariables.Location = new System.Drawing.Point(20, 46);
            this.lstEconVariables.Name = "lstEconVariables";
            this.lstEconVariables.Size = new System.Drawing.Size(177, 132);
            this.lstEconVariables.TabIndex = 24;
            this.lstEconVariables.SelectedIndexChanged += new System.EventHandler(this.lstEconVariables_SelectedIndexChanged);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label2);
            this.groupBox5.Controls.Add(this.txtEconRevenueDescr);
            this.groupBox5.Controls.Add(this.cmbNetRevEconOptimzFilter);
            this.groupBox5.Controls.Add(this.label3);
            this.groupBox5.Controls.Add(this.chkEnableEconFilter);
            this.groupBox5.Controls.Add(this.cmbEconOptimizationOperator);
            this.groupBox5.Controls.Add(this.txtEconOptimizationValue);
            this.groupBox5.Location = new System.Drawing.Point(20, 223);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(584, 139);
            this.groupBox5.TabIndex = 18;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Net Revenue Dollars Per Acre Filter Setting";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(43, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(158, 20);
            this.label2.TabIndex = 90;
            this.label2.Text = "Net Revenue Calculation";
            // 
            // txtEconRevenueDescr
            // 
            this.txtEconRevenueDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEconRevenueDescr.Location = new System.Drawing.Point(219, 46);
            this.txtEconRevenueDescr.Multiline = true;
            this.txtEconRevenueDescr.Name = "txtEconRevenueDescr";
            this.txtEconRevenueDescr.ReadOnly = true;
            this.txtEconRevenueDescr.Size = new System.Drawing.Size(320, 40);
            this.txtEconRevenueDescr.TabIndex = 89;
            // 
            // cmbNetRevEconOptimzFilter
            // 
            this.cmbNetRevEconOptimzFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbNetRevEconOptimzFilter.Location = new System.Drawing.Point(46, 49);
            this.cmbNetRevEconOptimzFilter.Name = "cmbNetRevEconOptimzFilter";
            this.cmbNetRevEconOptimzFilter.Size = new System.Drawing.Size(110, 24);
            this.cmbNetRevEconOptimzFilter.TabIndex = 88;
            this.cmbNetRevEconOptimzFilter.SelectedIndexChanged += new System.EventHandler(this.cmbNetRevEconOptimzFilter_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(313, 102);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(24, 25);
            this.label3.TabIndex = 18;
            this.label3.Text = "$";
            // 
            // chkEnableEconFilter
            // 
            this.chkEnableEconFilter.Location = new System.Drawing.Point(46, 98);
            this.chkEnableEconFilter.Name = "chkEnableEconFilter";
            this.chkEnableEconFilter.Size = new System.Drawing.Size(112, 32);
            this.chkEnableEconFilter.TabIndex = 17;
            this.chkEnableEconFilter.Text = "Enable Filter";
            this.chkEnableEconFilter.CheckedChanged += new System.EventHandler(this.chkEnableEconFilter_CheckedChanged);
            // 
            // cmbEconOptimizationOperator
            // 
            this.cmbEconOptimizationOperator.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbEconOptimizationOperator.Items.AddRange(new object[] {
            ">",
            "<",
            ">=",
            "<=",
            "<>"});
            this.cmbEconOptimizationOperator.Location = new System.Drawing.Point(224, 100);
            this.cmbEconOptimizationOperator.Name = "cmbEconOptimizationOperator";
            this.cmbEconOptimizationOperator.Size = new System.Drawing.Size(83, 32);
            this.cmbEconOptimizationOperator.TabIndex = 16;
            this.cmbEconOptimizationOperator.Text = ">";
            // 
            // txtEconOptimizationValue
            // 
            this.txtEconOptimizationValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEconOptimizationValue.Location = new System.Drawing.Point(339, 99);
            this.txtEconOptimizationValue.Name = "txtEconOptimizationValue";
            this.txtEconOptimizationValue.Size = new System.Drawing.Size(200, 29);
            this.txtEconOptimizationValue.TabIndex = 15;
            this.txtEconOptimizationValue.Text = "0";
            this.txtEconOptimizationValue.Leave += new System.EventHandler(this.txtEconOptimizationValue_Leave);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.rdoEconOptimizeMinimum);
            this.groupBox6.Controls.Add(this.rdoEconOptimizeMaximum);
            this.groupBox6.Location = new System.Drawing.Point(220, 133);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(351, 45);
            this.groupBox6.TabIndex = 17;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Which attribute value is best";
            // 
            // rdoEconOptimizeMinimum
            // 
            this.rdoEconOptimizeMinimum.Location = new System.Drawing.Point(197, 15);
            this.rdoEconOptimizeMinimum.Name = "rdoEconOptimizeMinimum";
            this.rdoEconOptimizeMinimum.Size = new System.Drawing.Size(122, 24);
            this.rdoEconOptimizeMinimum.TabIndex = 14;
            this.rdoEconOptimizeMinimum.Text = "Minimum Value";
            // 
            // rdoEconOptimizeMaximum
            // 
            this.rdoEconOptimizeMaximum.Checked = true;
            this.rdoEconOptimizeMaximum.Location = new System.Drawing.Point(32, 16);
            this.rdoEconOptimizeMaximum.Name = "rdoEconOptimizeMaximum";
            this.rdoEconOptimizeMaximum.Size = new System.Drawing.Size(150, 24);
            this.rdoEconOptimizeMaximum.TabIndex = 12;
            this.rdoEconOptimizeMaximum.TabStop = true;
            this.rdoEconOptimizeMaximum.Text = "Maximum Value";
            // 
            // lblEconomicAttribute
            // 
            this.lblEconomicAttribute.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblEconomicAttribute.Location = new System.Drawing.Point(7, 11);
            this.lblEconomicAttribute.Name = "lblEconomicAttribute";
            this.lblEconomicAttribute.Size = new System.Drawing.Size(472, 32);
            this.lblEconomicAttribute.TabIndex = 13;
            this.lblEconomicAttribute.Text = "Economic Attribute";
            // 
            // BtnEconOptimizationDone
            // 
            this.BtnEconOptimizationDone.Location = new System.Drawing.Point(352, 376);
            this.BtnEconOptimizationDone.Name = "BtnEconOptimizationDone";
            this.BtnEconOptimizationDone.Size = new System.Drawing.Size(88, 40);
            this.BtnEconOptimizationDone.TabIndex = 11;
            this.BtnEconOptimizationDone.Text = "Done";
            this.BtnEconOptimizationDone.Click += new System.EventHandler(this.BtnEconOptimizationDone_Click);
            // 
            // btnEconOptimizationCancel
            // 
            this.btnEconOptimizationCancel.Location = new System.Drawing.Point(440, 376);
            this.btnEconOptimizationCancel.Name = "btnEconOptimizationCancel";
            this.btnEconOptimizationCancel.Size = new System.Drawing.Size(88, 40);
            this.btnEconOptimizationCancel.TabIndex = 9;
            this.btnEconOptimizationCancel.Text = "Cancel";
            this.btnEconOptimizationCancel.Click += new System.EventHandler(this.btnEconOptimizationCancel_Click);
            // 
            // grpboxOptimizationFVSVariable
            // 
            this.grpboxOptimizationFVSVariable.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxOptimizationFVSVariable.Controls.Add(this.panel1);
            this.grpboxOptimizationFVSVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxOptimizationFVSVariable.ForeColor = System.Drawing.Color.Black;
            this.grpboxOptimizationFVSVariable.Location = new System.Drawing.Point(0, 528);
            this.grpboxOptimizationFVSVariable.Name = "grpboxOptimizationFVSVariable";
            this.grpboxOptimizationFVSVariable.Size = new System.Drawing.Size(872, 448);
            this.grpboxOptimizationFVSVariable.TabIndex = 35;
            this.grpboxOptimizationFVSVariable.TabStop = false;
            this.grpboxOptimizationFVSVariable.Text = "Stand Attribute";
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.grpboxFVSVariablesOptimizationVariableValues);
            this.panel1.Controls.Add(this.grpFVSVariablesOptimizationVariableValuesSelected);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableClear);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableDone);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableCancel);
            this.panel1.Controls.Add(this.btnOptimizationFVSVariableNext);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(866, 427);
            this.panel1.TabIndex = 12;
            // 
            // grpboxFVSVariablesOptimizationVariableValues
            // 
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.txtOptVarDescr);
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.lblOptVarDescr);
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.lstFVSFieldsList);
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.lstFVSTablesList);
            this.grpboxFVSVariablesOptimizationVariableValues.Controls.Add(this.btnFVSVariablesOptimizationVariableValues);
            this.grpboxFVSVariablesOptimizationVariableValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxFVSVariablesOptimizationVariableValues.Name = "grpboxFVSVariablesOptimizationVariableValues";
            this.grpboxFVSVariablesOptimizationVariableValues.Size = new System.Drawing.Size(816, 280);
            this.grpboxFVSVariablesOptimizationVariableValues.TabIndex = 0;
            this.grpboxFVSVariablesOptimizationVariableValues.TabStop = false;
            this.grpboxFVSVariablesOptimizationVariableValues.Text = "Stand attribute used to select the optimal silvicultural sequence";
            // 
            // txtOptVarDescr
            // 
            this.txtOptVarDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOptVarDescr.Location = new System.Drawing.Point(556, 84);
            this.txtOptVarDescr.Multiline = true;
            this.txtOptVarDescr.Name = "txtOptVarDescr";
            this.txtOptVarDescr.ReadOnly = true;
            this.txtOptVarDescr.Size = new System.Drawing.Size(258, 75);
            this.txtOptVarDescr.TabIndex = 90;
            this.txtOptVarDescr.Visible = false;
            // 
            // lblOptVarDescr
            // 
            this.lblOptVarDescr.Location = new System.Drawing.Point(477, 87);
            this.lblOptVarDescr.Name = "lblOptVarDescr";
            this.lblOptVarDescr.Size = new System.Drawing.Size(80, 24);
            this.lblOptVarDescr.TabIndex = 89;
            this.lblOptVarDescr.Text = "Description:";
            this.lblOptVarDescr.Visible = false;
            // 
            // lstFVSFieldsList
            // 
            this.lstFVSFieldsList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSFieldsList.ItemHeight = 16;
            this.lstFVSFieldsList.Location = new System.Drawing.Point(251, 21);
            this.lstFVSFieldsList.Name = "lstFVSFieldsList";
            this.lstFVSFieldsList.Size = new System.Drawing.Size(202, 180);
            this.lstFVSFieldsList.Sorted = true;
            this.lstFVSFieldsList.TabIndex = 4;
            this.lstFVSFieldsList.SelectedIndexChanged += new System.EventHandler(this.lstFVSFieldsList_SelectedIndexChanged);
            // 
            // lstFVSTablesList
            // 
            this.lstFVSTablesList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSTablesList.ItemHeight = 16;
            this.lstFVSTablesList.Location = new System.Drawing.Point(8, 21);
            this.lstFVSTablesList.Name = "lstFVSTablesList";
            this.lstFVSTablesList.Size = new System.Drawing.Size(202, 180);
            this.lstFVSTablesList.TabIndex = 3;
            this.lstFVSTablesList.SelectedIndexChanged += new System.EventHandler(this.lstFVSTablesList_SelectedIndexChanged);
            // 
            // btnFVSVariablesOptimizationVariableValues
            // 
            this.btnFVSVariablesOptimizationVariableValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesOptimizationVariableValues.Location = new System.Drawing.Point(476, 21);
            this.btnFVSVariablesOptimizationVariableValues.Name = "btnFVSVariablesOptimizationVariableValues";
            this.btnFVSVariablesOptimizationVariableValues.Size = new System.Drawing.Size(143, 57);
            this.btnFVSVariablesOptimizationVariableValues.TabIndex = 1;
            this.btnFVSVariablesOptimizationVariableValues.Text = "Select";
            this.btnFVSVariablesOptimizationVariableValues.Click += new System.EventHandler(this.btnFVSVariablesOptimizationVariableValues_Click);
            // 
            // grpFVSVariablesOptimizationVariableValuesSelected
            // 
            this.grpFVSVariablesOptimizationVariableValuesSelected.Controls.Add(this.lblFVSVariablesOptimizationVariableValuesSelected);
            this.grpFVSVariablesOptimizationVariableValuesSelected.Location = new System.Drawing.Point(16, 304);
            this.grpFVSVariablesOptimizationVariableValuesSelected.Name = "grpFVSVariablesOptimizationVariableValuesSelected";
            this.grpFVSVariablesOptimizationVariableValuesSelected.Size = new System.Drawing.Size(816, 51);
            this.grpFVSVariablesOptimizationVariableValuesSelected.TabIndex = 4;
            this.grpFVSVariablesOptimizationVariableValuesSelected.TabStop = false;
            this.grpFVSVariablesOptimizationVariableValuesSelected.Text = "Currently Active Optimization Stand Attribute";
            // 
            // lblFVSVariablesOptimizationVariableValuesSelected
            // 
            this.lblFVSVariablesOptimizationVariableValuesSelected.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFVSVariablesOptimizationVariableValuesSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFVSVariablesOptimizationVariableValuesSelected.Location = new System.Drawing.Point(3, 18);
            this.lblFVSVariablesOptimizationVariableValuesSelected.Name = "lblFVSVariablesOptimizationVariableValuesSelected";
            this.lblFVSVariablesOptimizationVariableValuesSelected.Size = new System.Drawing.Size(810, 30);
            this.lblFVSVariablesOptimizationVariableValuesSelected.TabIndex = 2;
            this.lblFVSVariablesOptimizationVariableValuesSelected.Text = "Not Defined";
            // 
            // btnOptimizationFVSVariableClear
            // 
            this.btnOptimizationFVSVariableClear.Location = new System.Drawing.Point(24, 376);
            this.btnOptimizationFVSVariableClear.Name = "btnOptimizationFVSVariableClear";
            this.btnOptimizationFVSVariableClear.Size = new System.Drawing.Size(72, 40);
            this.btnOptimizationFVSVariableClear.TabIndex = 5;
            this.btnOptimizationFVSVariableClear.Text = "Clear";
            // 
            // btnOptimizationFVSVariableDone
            // 
            this.btnOptimizationFVSVariableDone.Location = new System.Drawing.Point(352, 376);
            this.btnOptimizationFVSVariableDone.Name = "btnOptimizationFVSVariableDone";
            this.btnOptimizationFVSVariableDone.Size = new System.Drawing.Size(88, 40);
            this.btnOptimizationFVSVariableDone.TabIndex = 11;
            this.btnOptimizationFVSVariableDone.Text = "Done";
            this.btnOptimizationFVSVariableDone.Click += new System.EventHandler(this.btnOptimizationFVSVariableDone_Click);
            // 
            // btnOptimizationFVSVariableCancel
            // 
            this.btnOptimizationFVSVariableCancel.Location = new System.Drawing.Point(440, 376);
            this.btnOptimizationFVSVariableCancel.Name = "btnOptimizationFVSVariableCancel";
            this.btnOptimizationFVSVariableCancel.Size = new System.Drawing.Size(88, 40);
            this.btnOptimizationFVSVariableCancel.TabIndex = 9;
            this.btnOptimizationFVSVariableCancel.Text = "Cancel";
            this.btnOptimizationFVSVariableCancel.Click += new System.EventHandler(this.btnOptimizationFVSVariableCancel_Click);
            // 
            // btnOptimizationFVSVariableNext
            // 
            this.btnOptimizationFVSVariableNext.Location = new System.Drawing.Point(616, 376);
            this.btnOptimizationFVSVariableNext.Name = "btnOptimizationFVSVariableNext";
            this.btnOptimizationFVSVariableNext.Size = new System.Drawing.Size(88, 40);
            this.btnOptimizationFVSVariableNext.TabIndex = 8;
            this.btnOptimizationFVSVariableNext.Text = "Next-->";
            this.btnOptimizationFVSVariableNext.Click += new System.EventHandler(this.btnOptimizationFVSVariableNext_Click);
            // 
            // grpboxOptimization
            // 
            this.grpboxOptimization.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxOptimization.Controls.Add(this.pnlOptimization);
            this.grpboxOptimization.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxOptimization.ForeColor = System.Drawing.Color.Black;
            this.grpboxOptimization.Location = new System.Drawing.Point(8, 48);
            this.grpboxOptimization.Name = "grpboxOptimization";
            this.grpboxOptimization.Size = new System.Drawing.Size(856, 472);
            this.grpboxOptimization.TabIndex = 32;
            this.grpboxOptimization.TabStop = false;
            this.grpboxOptimization.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePost_Resize);
            // 
            // pnlOptimization
            // 
            this.pnlOptimization.AutoScroll = true;
            this.pnlOptimization.Controls.Add(this.grpboxOptimizationListValues);
            this.pnlOptimization.Controls.Add(this.grpboxOptimizationEdit);
            this.pnlOptimization.Controls.Add(this.grpboxOptimizationAudit);
            this.pnlOptimization.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlOptimization.Location = new System.Drawing.Point(3, 18);
            this.pnlOptimization.Name = "pnlOptimization";
            this.pnlOptimization.Size = new System.Drawing.Size(850, 451);
            this.pnlOptimization.TabIndex = 70;
            this.pnlOptimization.Resize += new System.EventHandler(this.pnlOptimization_Resize);
            // 
            // grpboxOptimizationListValues
            // 
            this.grpboxOptimizationListValues.Controls.Add(this.lvOptimizationListValues);
            this.grpboxOptimizationListValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxOptimizationListValues.Name = "grpboxOptimizationListValues";
            this.grpboxOptimizationListValues.Size = new System.Drawing.Size(840, 256);
            this.grpboxOptimizationListValues.TabIndex = 67;
            this.grpboxOptimizationListValues.TabStop = false;
            this.grpboxOptimizationListValues.Text = "Step 1: Select 1 Optimization Variable";
            // 
            // lvOptimizationListValues
            // 
            this.lvOptimizationListValues.CheckBoxes = true;
            this.lvOptimizationListValues.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lvColChecked,
            this.lvColVariableName,
            this.lvColFVSVariableName,
            this.lvColValueSource,
            this.lvColMinMax,
            this.lvColUseFilter,
            this.lvColFilterOperator,
            this.lvColFilterValue});
            this.lvOptimizationListValues.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvOptimizationListValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lvOptimizationListValues.GridLines = true;
            this.lvOptimizationListValues.HideSelection = false;
            this.lvOptimizationListValues.LabelEdit = true;
            this.lvOptimizationListValues.Location = new System.Drawing.Point(3, 18);
            this.lvOptimizationListValues.MultiSelect = false;
            this.lvOptimizationListValues.Name = "lvOptimizationListValues";
            this.lvOptimizationListValues.Size = new System.Drawing.Size(834, 235);
            this.lvOptimizationListValues.TabIndex = 67;
            this.lvOptimizationListValues.UseCompatibleStateImageBehavior = false;
            this.lvOptimizationListValues.View = System.Windows.Forms.View.Details;
            this.lvOptimizationListValues.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lvOptimizationListValues_ItemCheck);
            this.lvOptimizationListValues.SelectedIndexChanged += new System.EventHandler(this.lvOptimizationListValues_SelectedIndexChanged);
            this.lvOptimizationListValues.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvOptimizationListValues_MouseUp);
            // 
            // lvColChecked
            // 
            this.lvColChecked.Text = "";
            // 
            // lvColVariableName
            // 
            this.lvColVariableName.Text = "Optimization";
            this.lvColVariableName.Width = 174;
            // 
            // lvColFVSVariableName
            // 
            this.lvColFVSVariableName.Text = "FVS Variable";
            this.lvColFVSVariableName.Width = 191;
            // 
            // lvColValueSource
            // 
            this.lvColValueSource.Text = "Value Source";
            this.lvColValueSource.Width = 134;
            // 
            // lvColMinMax
            // 
            this.lvColMinMax.Text = "Max/Min";
            this.lvColMinMax.Width = 100;
            // 
            // lvColUseFilter
            // 
            this.lvColUseFilter.Text = "Enable Filter";
            this.lvColUseFilter.Width = 133;
            // 
            // lvColFilterOperator
            // 
            this.lvColFilterOperator.Text = "Filter Operator";
            this.lvColFilterOperator.Width = 149;
            // 
            // lvColFilterValue
            // 
            this.lvColFilterValue.Text = "Filter Value";
            this.lvColFilterValue.Width = 150;
            // 
            // grpboxOptimizationEdit
            // 
            this.grpboxOptimizationEdit.Controls.Add(this.btnOptimizationEdit);
            this.grpboxOptimizationEdit.Location = new System.Drawing.Point(8, 288);
            this.grpboxOptimizationEdit.Name = "grpboxOptimizationEdit";
            this.grpboxOptimizationEdit.Size = new System.Drawing.Size(832, 72);
            this.grpboxOptimizationEdit.TabIndex = 68;
            this.grpboxOptimizationEdit.TabStop = false;
            this.grpboxOptimizationEdit.Text = "Step 2: Edit";
            // 
            // btnOptimizationEdit
            // 
            this.btnOptimizationEdit.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOptimizationEdit.Location = new System.Drawing.Point(14, 24);
            this.btnOptimizationEdit.Name = "btnOptimizationEdit";
            this.btnOptimizationEdit.Size = new System.Drawing.Size(810, 32);
            this.btnOptimizationEdit.TabIndex = 0;
            this.btnOptimizationEdit.Text = "Edit";
            this.btnOptimizationEdit.Click += new System.EventHandler(this.btnOptimizationEdit_Click);
            // 
            // grpboxOptimizationAudit
            // 
            this.grpboxOptimizationAudit.Controls.Add(this.btnOptimizationAudit);
            this.grpboxOptimizationAudit.Location = new System.Drawing.Point(8, 368);
            this.grpboxOptimizationAudit.Name = "grpboxOptimizationAudit";
            this.grpboxOptimizationAudit.Size = new System.Drawing.Size(832, 72);
            this.grpboxOptimizationAudit.TabIndex = 69;
            this.grpboxOptimizationAudit.TabStop = false;
            this.grpboxOptimizationAudit.Text = "Step 3: Audit";
            // 
            // btnOptimizationAudit
            // 
            this.btnOptimizationAudit.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOptimizationAudit.Location = new System.Drawing.Point(16, 24);
            this.btnOptimizationAudit.Name = "btnOptimizationAudit";
            this.btnOptimizationAudit.Size = new System.Drawing.Size(800, 32);
            this.btnOptimizationAudit.TabIndex = 0;
            this.btnOptimizationAudit.Text = "Audit";
            this.btnOptimizationAudit.Click += new System.EventHandler(this.btnOptimizationAudit_Click);
            // 
            // grpboxOptimizationSettings
            // 
            this.grpboxOptimizationSettings.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxOptimizationSettings.Controls.Add(this.pnlFVSVariablesPrePostVariable);
            this.grpboxOptimizationSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxOptimizationSettings.ForeColor = System.Drawing.Color.Black;
            this.grpboxOptimizationSettings.Location = new System.Drawing.Point(16, 992);
            this.grpboxOptimizationSettings.Name = "grpboxOptimizationSettings";
            this.grpboxOptimizationSettings.Size = new System.Drawing.Size(856, 448);
            this.grpboxOptimizationSettings.TabIndex = 30;
            this.grpboxOptimizationSettings.TabStop = false;
            this.grpboxOptimizationSettings.Text = "Stand Attribute Settings";
            this.grpboxOptimizationSettings.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePostVariable_Resize);
            // 
            // pnlFVSVariablesPrePostVariable
            // 
            this.pnlFVSVariablesPrePostVariable.AutoScroll = true;
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.LblNetRevenu);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpBoxOptimizationNetRevenue);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnOptimiztionPrev);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpboxOptimizationSettingsPostPre);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpFilter);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpMaxMin);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.lblOptimizationVariable);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnOptimiztionDone);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnOptimiztionCancel);
            this.pnlFVSVariablesPrePostVariable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFVSVariablesPrePostVariable.Location = new System.Drawing.Point(3, 18);
            this.pnlFVSVariablesPrePostVariable.Name = "pnlFVSVariablesPrePostVariable";
            this.pnlFVSVariablesPrePostVariable.Size = new System.Drawing.Size(850, 427);
            this.pnlFVSVariablesPrePostVariable.TabIndex = 12;
            this.pnlFVSVariablesPrePostVariable.Resize += new System.EventHandler(this.pnlFVSVariablesPrePostVariable_Resize);
            // 
            // LblNetRevenu
            // 
            this.LblNetRevenu.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblNetRevenu.Location = new System.Drawing.Point(24, 176);
            this.LblNetRevenu.Name = "LblNetRevenu";
            this.LblNetRevenu.Size = new System.Drawing.Size(182, 32);
            this.LblNetRevenu.TabIndex = 23;
            this.LblNetRevenu.Text = "Net Revenue";
            // 
            // grpBoxOptimizationNetRevenue
            // 
            this.grpBoxOptimizationNetRevenue.Controls.Add(this.txtRevenueDescr);
            this.grpBoxOptimizationNetRevenue.Controls.Add(this.cmbNetRevOptimzFilter);
            this.grpBoxOptimizationNetRevenue.Location = new System.Drawing.Point(80, 215);
            this.grpBoxOptimizationNetRevenue.Name = "grpBoxOptimizationNetRevenue";
            this.grpBoxOptimizationNetRevenue.Size = new System.Drawing.Size(464, 69);
            this.grpBoxOptimizationNetRevenue.TabIndex = 22;
            this.grpBoxOptimizationNetRevenue.TabStop = false;
            this.grpBoxOptimizationNetRevenue.Text = "Net Revenue Calculation";
            this.grpBoxOptimizationNetRevenue.Visible = false;
            // 
            // txtRevenueDescr
            // 
            this.txtRevenueDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRevenueDescr.Location = new System.Drawing.Point(189, 21);
            this.txtRevenueDescr.Multiline = true;
            this.txtRevenueDescr.Name = "txtRevenueDescr";
            this.txtRevenueDescr.ReadOnly = true;
            this.txtRevenueDescr.Size = new System.Drawing.Size(259, 40);
            this.txtRevenueDescr.TabIndex = 87;
            // 
            // cmbNetRevOptimzFilter
            // 
            this.cmbNetRevOptimzFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbNetRevOptimzFilter.Location = new System.Drawing.Point(16, 24);
            this.cmbNetRevOptimzFilter.Name = "cmbNetRevOptimzFilter";
            this.cmbNetRevOptimzFilter.Size = new System.Drawing.Size(110, 24);
            this.cmbNetRevOptimzFilter.TabIndex = 20;
            this.cmbNetRevOptimzFilter.SelectedIndexChanged += new System.EventHandler(this.cmbNetRevOptimzFilter_SelectedIndexChanged);
            // 
            // btnOptimiztionPrev
            // 
            this.btnOptimiztionPrev.Location = new System.Drawing.Point(528, 376);
            this.btnOptimiztionPrev.Name = "btnOptimiztionPrev";
            this.btnOptimiztionPrev.Size = new System.Drawing.Size(88, 40);
            this.btnOptimiztionPrev.TabIndex = 21;
            this.btnOptimiztionPrev.Text = "<--Previous";
            this.btnOptimiztionPrev.Click += new System.EventHandler(this.btnOptimiztionPrev_Click);
            // 
            // grpboxOptimizationSettingsPostPre
            // 
            this.grpboxOptimizationSettingsPostPre.Controls.Add(this.cmbOptimizationSettingsPostPreValue);
            this.grpboxOptimizationSettingsPostPre.Location = new System.Drawing.Point(80, 55);
            this.grpboxOptimizationSettingsPostPre.Name = "grpboxOptimizationSettingsPostPre";
            this.grpboxOptimizationSettingsPostPre.Size = new System.Drawing.Size(344, 72);
            this.grpboxOptimizationSettingsPostPre.TabIndex = 20;
            this.grpboxOptimizationSettingsPostPre.TabStop = false;
            this.grpboxOptimizationSettingsPostPre.Text = "Post Treatment Stand Attribute Or Pre/Post Treatment Change";
            // 
            // cmbOptimizationSettingsPostPreValue
            // 
            this.cmbOptimizationSettingsPostPreValue.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbOptimizationSettingsPostPreValue.Items.AddRange(new object[] {
            "Post Value",
            "Post - Pre  Change Value"});
            this.cmbOptimizationSettingsPostPreValue.Location = new System.Drawing.Point(16, 40);
            this.cmbOptimizationSettingsPostPreValue.Name = "cmbOptimizationSettingsPostPreValue";
            this.cmbOptimizationSettingsPostPreValue.Size = new System.Drawing.Size(320, 24);
            this.cmbOptimizationSettingsPostPreValue.TabIndex = 0;
            // 
            // grpFilter
            // 
            this.grpFilter.Controls.Add(this.label1);
            this.grpFilter.Controls.Add(this.chkEnableFilter);
            this.grpFilter.Controls.Add(this.cmbOptimizationOperator);
            this.grpFilter.Controls.Add(this.txtOptimizationValue);
            this.grpFilter.Location = new System.Drawing.Point(80, 290);
            this.grpFilter.Name = "grpFilter";
            this.grpFilter.Size = new System.Drawing.Size(584, 64);
            this.grpFilter.TabIndex = 18;
            this.grpFilter.TabStop = false;
            this.grpFilter.Text = "Net Revenue Dollars Per Acre Filter Setting";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(326, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(24, 25);
            this.label1.TabIndex = 18;
            this.label1.Text = "$";
            // 
            // chkEnableFilter
            // 
            this.chkEnableFilter.Location = new System.Drawing.Point(46, 22);
            this.chkEnableFilter.Name = "chkEnableFilter";
            this.chkEnableFilter.Size = new System.Drawing.Size(112, 32);
            this.chkEnableFilter.TabIndex = 17;
            this.chkEnableFilter.Text = "Enable Filter";
            this.chkEnableFilter.CheckedChanged += new System.EventHandler(this.chkEnableFilter_CheckedChanged);
            // 
            // cmbOptimizationOperator
            // 
            this.cmbOptimizationOperator.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOptimizationOperator.Items.AddRange(new object[] {
            ">",
            "<",
            ">=",
            "<=",
            "<>"});
            this.cmbOptimizationOperator.Location = new System.Drawing.Point(237, 26);
            this.cmbOptimizationOperator.Name = "cmbOptimizationOperator";
            this.cmbOptimizationOperator.Size = new System.Drawing.Size(83, 32);
            this.cmbOptimizationOperator.TabIndex = 16;
            this.cmbOptimizationOperator.Text = ">";
            // 
            // txtOptimizationValue
            // 
            this.txtOptimizationValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOptimizationValue.Location = new System.Drawing.Point(352, 25);
            this.txtOptimizationValue.Name = "txtOptimizationValue";
            this.txtOptimizationValue.Size = new System.Drawing.Size(200, 29);
            this.txtOptimizationValue.TabIndex = 15;
            this.txtOptimizationValue.Text = "0";
            this.txtOptimizationValue.Leave += new System.EventHandler(this.txtOptimizationValue_Leave);
            // 
            // grpMaxMin
            // 
            this.grpMaxMin.Controls.Add(this.rdoOptimizationMinimum);
            this.grpMaxMin.Controls.Add(this.rdoOptimizationMaximum);
            this.grpMaxMin.Location = new System.Drawing.Point(457, 55);
            this.grpMaxMin.Name = "grpMaxMin";
            this.grpMaxMin.Size = new System.Drawing.Size(388, 45);
            this.grpMaxMin.TabIndex = 17;
            this.grpMaxMin.TabStop = false;
            this.grpMaxMin.Text = "Which attribute value is best";
            // 
            // rdoOptimizationMinimum
            // 
            this.rdoOptimizationMinimum.Location = new System.Drawing.Point(210, 16);
            this.rdoOptimizationMinimum.Name = "rdoOptimizationMinimum";
            this.rdoOptimizationMinimum.Size = new System.Drawing.Size(150, 24);
            this.rdoOptimizationMinimum.TabIndex = 14;
            this.rdoOptimizationMinimum.Text = "Minimum Value";
            // 
            // rdoOptimizationMaximum
            // 
            this.rdoOptimizationMaximum.Checked = true;
            this.rdoOptimizationMaximum.Location = new System.Drawing.Point(32, 16);
            this.rdoOptimizationMaximum.Name = "rdoOptimizationMaximum";
            this.rdoOptimizationMaximum.Size = new System.Drawing.Size(150, 24);
            this.rdoOptimizationMaximum.TabIndex = 12;
            this.rdoOptimizationMaximum.TabStop = true;
            this.rdoOptimizationMaximum.Text = "Maximum Value";
            // 
            // lblOptimizationVariable
            // 
            this.lblOptimizationVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOptimizationVariable.Location = new System.Drawing.Point(24, 24);
            this.lblOptimizationVariable.Name = "lblOptimizationVariable";
            this.lblOptimizationVariable.Size = new System.Drawing.Size(472, 32);
            this.lblOptimizationVariable.TabIndex = 13;
            this.lblOptimizationVariable.Text = "Optimization Variable";
            // 
            // btnOptimiztionDone
            // 
            this.btnOptimiztionDone.Location = new System.Drawing.Point(352, 376);
            this.btnOptimiztionDone.Name = "btnOptimiztionDone";
            this.btnOptimiztionDone.Size = new System.Drawing.Size(88, 40);
            this.btnOptimiztionDone.TabIndex = 11;
            this.btnOptimiztionDone.Text = "Done";
            this.btnOptimiztionDone.Click += new System.EventHandler(this.btnOptimiztionDone_Click);
            // 
            // btnOptimiztionCancel
            // 
            this.btnOptimiztionCancel.Location = new System.Drawing.Point(440, 376);
            this.btnOptimiztionCancel.Name = "btnOptimiztionCancel";
            this.btnOptimiztionCancel.Size = new System.Drawing.Size(88, 40);
            this.btnOptimiztionCancel.TabIndex = 9;
            this.btnOptimiztionCancel.Text = "Cancel";
            this.btnOptimiztionCancel.Click += new System.EventHandler(this.btnOptimiztionCancel_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(866, 32);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Optimization Settings";
            // 
            // uc_core_scenario_fvs_prepost_optimization
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_core_scenario_fvs_prepost_optimization";
            this.Size = new System.Drawing.Size(872, 2000);
            this.groupBox1.ResumeLayout(false);
            this.grpboxEconOptimizSettings.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.grpboxOptimizationFVSVariable.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.grpboxFVSVariablesOptimizationVariableValues.ResumeLayout(false);
            this.grpboxFVSVariablesOptimizationVariableValues.PerformLayout();
            this.grpFVSVariablesOptimizationVariableValuesSelected.ResumeLayout(false);
            this.grpboxOptimization.ResumeLayout(false);
            this.pnlOptimization.ResumeLayout(false);
            this.grpboxOptimizationListValues.ResumeLayout(false);
            this.grpboxOptimizationEdit.ResumeLayout(false);
            this.grpboxOptimizationAudit.ResumeLayout(false);
            this.grpboxOptimizationSettings.ResumeLayout(false);
            this.pnlFVSVariablesPrePostVariable.ResumeLayout(false);
            this.grpBoxOptimizationNetRevenue.ResumeLayout(false);
            this.grpBoxOptimizationNetRevenue.PerformLayout();
            this.grpboxOptimizationSettingsPostPre.ResumeLayout(false);
            this.grpFilter.ResumeLayout(false);
            this.grpFilter.PerformLayout();
            this.grpMaxMin.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

        public void loadvalues_FromProperties()
        {
            int x;
            this.m_oOldVariableCollection.Clear();
            this.m_oSavVariableCollection.Clear();
            if (ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0) != null)
            {
                for (x = 0; x <= ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Count - 1; x++)
                {
                    VariableItem oItem = new VariableItem();
                    oItem.strOptimizedVariable =
                       ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strOptimizedVariable;
					//optimization variable
                    oItem.strFVSVariableName =
                        ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strFVSVariableName;
        			//value source (POST or POST-PRE)
                    oItem.strValueSource =
                        ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strValueSource;								
                    //max value
                    oItem.strMaxYN =
                        ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strMaxYN;
					//min value
                    oItem.strMinYN =
                        ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strMinYN;
                    //enable filter
                    oItem.bUseFilter = 
                        ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).bUseFilter;
                    //filter operator
                    oItem.strFilterOperator =
                         ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strFilterOperator;
                    //filter value
                    oItem.dblFilterValue =
                         ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).dblFilterValue;
					//item checked
                    oItem.bSelected =
                        ReferenceCoreScenarioForm.m_oCoreAnalysisScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).bSelected;

                    if (oItem.bSelected)
                    {
                        if (oItem.strOptimizedVariable.Trim().ToUpper() == "STAND ATTRIBUTE")
                        {
                            ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText(oItem.strFVSVariableName);
                        }
                        else
                        {
                            ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText(oItem.strOptimizedVariable);
                        }

                    }
                    this.m_oOldVariableCollection.Add(oItem);
                
						
                }
                this.m_oOldVariableCollection.Copy(this.m_oOldVariableCollection, ref this.m_oSavVariableCollection, true);
                this.UpdateValues();
            }

        }
        public void loadvalues(System.Collections.Generic.IDictionary<string, System.Collections.Generic.IList<String>> p_dictFVSTables)
		{
			
			this.m_intError=0;
			this.m_strError="";

			int y;
            lstFVSTablesList.Items.Clear();
            m_dictFVSTables = new System.Collections.Generic.Dictionary<string,
                System.Collections.Generic.IList<string>>(p_dictFVSTables);
            foreach (string strKey in m_dictFVSTables.Keys)
            {
                lstFVSTablesList.Items.Add(strKey);
            }

            //Load economic variables list
            lstEconVariables.Items.Clear();
            foreach (uc_core_scenario_weighted_average.VariableItem oItem in this.ReferenceCoreScenarioForm.m_oWeightedVariableCollection)
            {
                if (oItem.strVariableType.Equals("ECON"))
                {
                    lstEconVariables.Items.Add(oItem.strVariableName);
                    if (oItem.strVariableName.IndexOf("revenue") > -1 || oItem.strVariableName.IndexOf("onsite_treatment") > -1)
                    {
                        cmbNetRevOptimzFilter.Items.Add(oItem.strVariableName);
                        cmbNetRevEconOptimzFilter.Items.Add(oItem.strVariableName);
                    }
                }
            }

			this.lvOptimizationListValues.Items.Clear();
			this.m_oLvRowColors.InitializeRowCollection();
			this.m_oLvRowColors.m_intSelectedRow=-1;

			CreateListViewOptimizationRow();
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(lvOptimizationListValues.Items.Count-1,lvOptimizationListValues.Columns.Count);
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text = "Revenue";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FVS_VARIABLE].Text = "NA";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_VALUESOURCE].Text = "NA";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_USEFILTER].Text = "No";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_OPERATOR].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_VALUE].Text = "Not Defined";

			CreateListViewOptimizationRow();
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(lvOptimizationListValues.Items.Count-1,lvOptimizationListValues.Columns.Count);
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text = "Merchantable Volume";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FVS_VARIABLE].Text = "NA";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_VALUESOURCE].Text = "NA";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_USEFILTER].Text = "No";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_OPERATOR].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_VALUE].Text = "Not Defined";

			CreateListViewOptimizationRow();
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(lvOptimizationListValues.Items.Count-1,lvOptimizationListValues.Columns.Count);
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text = "Stand Attribute";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FVS_VARIABLE].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_VALUESOURCE].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_USEFILTER].Text = "No";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_OPERATOR].Text = "Not Defined";
			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_VALUE].Text = "Not Defined";

            // New row for economic weighted variables
            CreateListViewOptimizationRow();
            this.m_oLvRowColors.AddRow();
            this.m_oLvRowColors.AddColumns(lvOptimizationListValues.Items.Count - 1, lvOptimizationListValues.Columns.Count);
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text = "Economic Attribute";
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FVS_VARIABLE].Text = "NA";
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_VALUESOURCE].Text = "NA";
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "Not Defined";
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_USEFILTER].Text = "No";
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_OPERATOR].Text = "Not Defined";
            this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count - 1].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_VALUE].Text = "Not Defined";


			//
			//load previous scenario values
			//
			if (this.m_bFirstTime)
			{
				ado_data_access oAdo = new ado_data_access();
				string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
				string strScenarioMDB = 
					frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
					"\\core\\db\\scenario_core_rule_definitions.mdb";
				oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
				if (oAdo.m_intError==0)
				{
					int intVarNum=0;

					oAdo.m_strSQL = "SELECT * " + 
						"FROM scenario_fvs_variables_optimization " + 
						"WHERE TRIM(scenario_id)='" + strScenarioId.Trim() + "' AND " + 
						"current_yn='Y'";

					oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);
					if (oAdo.m_OleDbDataReader.HasRows)
					{
						while (oAdo.m_OleDbDataReader.Read())
						{
						    
							if (oAdo.m_OleDbDataReader["optimization_variable"] != System.DBNull.Value)
							{
								VariableItem oItem = new VariableItem();
								oItem.strOptimizedVariable = Convert.ToString(oAdo.m_OleDbDataReader["optimization_variable"]);
								//optimization variable
								if (oAdo.m_OleDbDataReader["fvs_variable_name"] != System.DBNull.Value)
								{
									oItem.strFVSVariableName = Convert.ToString(oAdo.m_OleDbDataReader["fvs_variable_name"]).Trim();
								}
								else
								{
									oItem.strFVSVariableName = "NA";
								}
								//value source (POST or POST-PRE)
								if (oAdo.m_OleDbDataReader["value_source"] != System.DBNull.Value)
								{
									oItem.strValueSource = Convert.ToString(oAdo.m_OleDbDataReader["value_source"]).Trim();
								}
								else
								{
									oItem.strValueSource = "Not Defined";
								}

								//max value
								if (oAdo.m_OleDbDataReader["max_yn"] != System.DBNull.Value)
								{
									oItem.strMaxYN = Convert.ToString(oAdo.m_OleDbDataReader["max_yn"]).Trim();
								}
								else
								{
									oItem.strMaxYN = "N";
								}
								//min value
								if (oAdo.m_OleDbDataReader["min_yn"] != System.DBNull.Value)
								{
									oItem.strMinYN = Convert.ToString(oAdo.m_OleDbDataReader["min_yn"]).Trim();
								}
								else
								{
									oItem.strMinYN = "N";
								}
								//enable filter
								if (oAdo.m_OleDbDataReader["filter_enabled_yn"] != System.DBNull.Value)
								{
									if (Convert.ToString(oAdo.m_OleDbDataReader["filter_enabled_yn"]).Trim()=="Y")
										oItem.bUseFilter=true;
									else
										oItem.bUseFilter=false;
										 
								}
								else
								{
									oItem.bUseFilter=false;
								}
								//filter operator
								if (oAdo.m_OleDbDataReader["filter_operator"] != System.DBNull.Value)
								{
									
									oItem.strFilterOperator = Convert.ToString(oAdo.m_OleDbDataReader["filter_operator"]).Trim();	 
								}
								else
								{
									oItem.strFilterOperator="";
								}
								//filter value
								if (oAdo.m_OleDbDataReader["filter_value"] != System.DBNull.Value)
								{
									oItem.dblFilterValue = Convert.ToDouble(oAdo.m_OleDbDataReader["filter_value"]);	 
								}
								//filter operator
								if (oAdo.m_OleDbDataReader["checked_yn"] != System.DBNull.Value)
								{
                                    if (Convert.ToString(oAdo.m_OleDbDataReader["checked_yn"]).Trim() == "Y")
                                    {
                                        oItem.bSelected = true;
                                        if (oItem.strOptimizedVariable.Trim().ToUpper() == "STAND ATTRIBUTE")
                                        {
                                            ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText(oItem.strFVSVariableName);
                                        }
                                        else
                                        {
                                            ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText(oItem.strOptimizedVariable);
                                        }
                                    }
                                    else
                                        oItem.bSelected = false;
								}
								else
								{
									oItem.bSelected=false;
								}
                                //revenue attribute
                                if (oAdo.m_OleDbDataReader["revenue_attribute"] != System.DBNull.Value)
                                {
                                    oItem.strRevenueAttribute = Convert.ToString(oAdo.m_OleDbDataReader["revenue_attribute"]).Trim();
                                }
								this.m_oOldVariableCollection.Add(oItem);
							}
						}
						this.m_oOldVariableCollection.Copy(this.m_oOldVariableCollection,ref this.m_oSavVariableCollection,true);
					}
					oAdo.m_OleDbDataReader.Close();
					oAdo.CloseConnection(oAdo.m_OleDbConnection);
					oAdo.m_OleDbConnection.Dispose();
				}



			
			
				this.m_intError=oAdo.m_intError;
				this.m_strError=oAdo.m_strError;
				oAdo=null;
				this.m_bFirstTime=false;
			}
			this.UpdateValues();

			
		}

		
		public void UpdateValues()
		{
			int x,y;
			//
			//update list view items
			//
				bool bFound;
			    for (x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
					this.m_oSavVariableCollection.Item(x).intListViewIndex=-1;
				for (y=0;y<=this.lvOptimizationListValues.Items.Count-1;y++)
				{
					bFound=false;
                    string strOptimizeVariable = lvOptimizationListValues.Items[y].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text.Trim().ToUpper();
                    // It is the Revenue or Merchantable Volume line
                    if (strOptimizeVariable.IndexOf("ATTRIBUTE") < 0)
					{
						for (x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
						{
							if (lvOptimizationListValues.Items[y].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text.Trim().ToUpper() == 
								m_oSavVariableCollection.Item(x).strOptimizedVariable.Trim().ToUpper())
							{
								bFound=true;
								break;
							}
						}
					}
					// It is the Stand or Economic Attribute line
                    else
					{
						for (x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
						{
                            string strTest = m_oSavVariableCollection.Item(x).strOptimizedVariable.Trim().ToUpper();
                            if ((strOptimizeVariable == "STAND ATTRIBUTE" &&
                                m_oSavVariableCollection.Item(x).strOptimizedVariable.Trim().ToUpper() == "STAND ATTRIBUTE") ||
                                (strOptimizeVariable == "ECONOMIC ATTRIBUTE" &&
                                m_oSavVariableCollection.Item(x).strOptimizedVariable.Trim().ToUpper() == "ECONOMIC ATTRIBUTE"))
							{
								bFound=true;
								break;
							}
						}
					}
					if (bFound)
					{
						UpdateListViewItem(lvOptimizationListValues.Items[y],m_oSavVariableCollection.Item(x));
					}
					else
					{
						VariableItem oItem = new VariableItem();
						oItem.intListViewIndex = y;
						oItem.strOptimizedVariable = this.lvOptimizationListValues.Items[y].SubItems[COLUMN_OPTIMIZE_VARIABLE].Text.Trim();
						oItem.strFVSVariableName = this.lvOptimizationListValues.Items[y].SubItems[COLUMN_FVS_VARIABLE].Text.Trim();
						oItem.strValueSource = lvOptimizationListValues.Items[y].SubItems[COLUMN_VALUESOURCE].Text.Trim();
						oItem.bSelected = this.lvOptimizationListValues.Items[y].Checked;
						if (this.lvOptimizationListValues.Items[y].SubItems[COLUMN_USEFILTER].Text.Trim() == "Yes")
							oItem.bUseFilter=true;
						else
							oItem.bUseFilter=false;
						
						if (this.lvOptimizationListValues.Items[y].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text.Trim() == "MAX")
						{
							oItem.strMaxYN="Y";
						}
						else
							oItem.strMaxYN="N";

						if (this.lvOptimizationListValues.Items[y].SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text.Trim() == "MIN")
						{
							oItem.strMinYN="Y";
						}
						else
							oItem.strMinYN="N";

						this.m_oSavVariableCollection.Add(oItem);

					}
				}
			
				for (x=m_oSavVariableCollection.Count-1;x>=0;x--)
				{
					if (this.m_oSavVariableCollection.Item(x).intListViewIndex==-1)
						this.m_oSavVariableCollection.Remove(x);
				}
			
			
		}
		private void UpdateListViewItem(System.Windows.Forms.ListViewItem p_lvItem,VariableItem p_oVariableItem)
		{
			p_oVariableItem.intListViewIndex=p_lvItem.Index;
			p_lvItem.Checked=p_oVariableItem.bSelected;

            if (p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text.Trim().ToUpper() == "STAND ATTRIBUTE")
            {
                p_lvItem.SubItems[COLUMN_FVS_VARIABLE].Text = p_oVariableItem.strFVSVariableName;
                p_oVariableItem.strOptimizedVariable = p_lvItem.SubItems[COLUMN_OPTIMIZE_VARIABLE].Text.Trim();
                p_lvItem.SubItems[COLUMN_VALUESOURCE].Text = p_oVariableItem.strValueSource;
            }
            else if (p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_OPTIMIZE_VARIABLE].Text.Trim().ToUpper() == "ECONOMIC ATTRIBUTE")
            {
                p_lvItem.SubItems[COLUMN_FVS_VARIABLE].Text = p_oVariableItem.strFVSVariableName;
                p_oVariableItem.strOptimizedVariable = p_lvItem.SubItems[COLUMN_OPTIMIZE_VARIABLE].Text.Trim();
                // No current need to set Value Source (POST, POST-PRE) for economic attributes
            }


			if (p_oVariableItem.bUseFilter)
				p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_USEFILTER].Text = "Yes";
			else
				p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_USEFILTER].Text = "No";

			if (p_oVariableItem.strMaxYN=="Y")
				p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "MAX";
			else if (p_oVariableItem.strMinYN=="Y")
				p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "MIN";
			else
				p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_MAXMIN].Text = "Not Defined";

			if (p_oVariableItem.strFilterOperator.Trim().Length > 0)
			{
				p_lvItem.SubItems[uc_core_scenario_fvs_prepost_optimization.COLUMN_FILTER_OPERATOR].Text = 
					p_oVariableItem.strFilterOperator;
			}
            m_oValidate.ValidateDecimal(Convert.ToString(p_oVariableItem.dblFilterValue).Trim());
            p_lvItem.SubItems[COLUMN_FILTER_VALUE].Text = m_oValidate.ReturnValue;
		}
		public int savevalues()
		{
			int x;
			string strValues="";
			string strColumns="";
			string strWhere="";
			ado_data_access oAdo = new ado_data_access();
			string strScenarioId = this.ReferenceCoreScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
			if (oAdo.m_intError==0)
			{
				oAdo.m_strSQL = "SELECT COUNT(*) FROM scenario_fvs_variables_optimization WHERE " + 
					" scenario_id = '" + strScenarioId + "' AND current_yn = 'Y';";
				if ((int)oAdo.getRecordCount(oAdo.m_OleDbConnection,oAdo.m_strSQL,"scenario_fvs_variables_optimization")> 0)
				{
					oAdo.m_strSQL = "UPDATE scenario_fvs_variables_optimization SET current_yn = 'N'" + 
						" WHERE scenario_id = '" + strScenarioId + "' AND current_yn = 'Y';";
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				}

				if (oAdo.m_intError < 0)
				{
					oAdo.m_OleDbConnection.Close();
					x=oAdo.m_intError;
					oAdo = null;
					return x;
				}
				strColumns = "scenario_id,rxcycle,optimization_variable,fvs_variable_name,value_source,max_yn,min_yn,filter_enabled_yn,filter_operator,filter_value,checked_yn,current_yn,revenue_attribute";
				
				
				for (x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
				{
					strValues = "'" + strScenarioId.Trim() + "','1'";
					strWhere = "TRIM(scenario_id)='" + strScenarioId.Trim() + "' AND rxcycle='1' ";
					strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strOptimizedVariable + "'";
					strWhere = strWhere + " AND UCASE(TRIM(optimization_variable))='" + m_oSavVariableCollection.Item(x).strOptimizedVariable.Trim().ToUpper() + "'";
					strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strFVSVariableName + "'";
					strWhere = strWhere + " AND UCASE(TRIM(fvs_variable_name))='" + this.m_oSavVariableCollection.Item(x).strFVSVariableName.Trim().ToUpper() + "'";	

					strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strValueSource + "'";
					strWhere = strWhere + " AND UCASE(TRIM(value_source))='" + this.m_oSavVariableCollection.Item(x).strValueSource.Trim().ToUpper() + "'";	
				
					strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strMaxYN + "'";
					strWhere = strWhere + " AND UCASE(TRIM(max_yn))='" + this.m_oSavVariableCollection.Item(x).strMaxYN.Trim().ToUpper() + "'";					
					strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strMinYN + "'";
					strWhere = strWhere + " AND UCASE(TRIM(min_yn))='" + this.m_oSavVariableCollection.Item(x).strMinYN.Trim().ToUpper() + "'";					
					if (this.m_oSavVariableCollection.Item(x).bUseFilter)
					{
						strValues = strValues + ",'Y'";
						strWhere = strWhere + " AND UCASE(TRIM(filter_enabled_yn))='Y'";					
					}
					else
					{
						strValues = strValues + ",'N'";
						strWhere = strWhere + " AND UCASE(TRIM(filter_enabled_yn))='N'";					
					}

					
					strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strFilterOperator + "'";
					strWhere = strWhere + " AND UCASE(TRIM(filter_operator))='" + this.m_oSavVariableCollection.Item(x).strFilterOperator.Trim().ToUpper() + "'";					
					strValues = strValues + "," + this.m_oSavVariableCollection.Item(x).dblFilterValue.ToString().Trim();
                    strWhere = strWhere + " AND filter_value=" + this.m_oSavVariableCollection.Item(x).dblFilterValue.ToString().Trim();					
					if (this.m_oSavVariableCollection.Item(x).bSelected)
						strValues = strValues + ",'Y'";
					else
						strValues = strValues + ",'N'";

					strValues = strValues + ",'Y'";
                    if (!String.IsNullOrEmpty(this.m_oSavVariableCollection.Item(x).strRevenueAttribute.Trim()))
                    {
                        strValues = strValues + ",'" + this.m_oSavVariableCollection.Item(x).strRevenueAttribute.Trim() + "'";
                    }
                    else
                    {
                        strValues = strValues + ",''";
                    }

					//delete duplicates
					oAdo.m_strSQL = "DELETE FROM scenario_fvs_variables_optimization WHERE " + strWhere;
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);




					oAdo.m_strSQL = "INSERT INTO scenario_fvs_variables_optimization " + 
							"(" + strColumns + ") VALUES " + 
							"(" + strValues + ")";
					oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				}
				

			}
			return 1;
			
		}
		private void CreateListViewOptimizationRow()
		{
			this.lvOptimizationListValues.Items.Add(" ");
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(lvOptimizationListValues.Items.Count-1,lvOptimizationListValues.Columns.Count);

			this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].UseItemStyleForSubItems=false;
			this.m_oLvRowColors.ListViewSubItem(lvOptimizationListValues.Items.Count-1,0,
				                                lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems.Count-1],false);
			for (int x=1;x<=lvOptimizationListValues.Columns.Count-1;x++)
			{
				lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems.Add(" ");
				this.m_oLvRowColors.ListViewSubItem(lvOptimizationListValues.Items.Count-1,x,
					lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems[this.lvOptimizationListValues.Items[lvOptimizationListValues.Items.Count-1].SubItems.Count-1],false);
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
			catch  (Exception caught)
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
		
		

		
		public void main_resize()
		{
			
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{

			
			this.grpboxOptimization.Height = this.ClientSize.Height - this.grpboxOptimization.Top - 5;
			grpboxOptimization.Width = this.ClientSize.Width - (grpboxOptimization.Left * 2) ;
		    grpboxOptimizationSettings.Height = grpboxOptimization.Height;
			grpboxOptimizationSettings.Width =  grpboxOptimization.Width;
			this.grpboxOptimizationFVSVariable.Height = grpboxOptimization.Height;
			this.grpboxOptimizationFVSVariable.Width = grpboxOptimization.Width;
            this.grpboxEconOptimizSettings.Height = grpboxOptimization.Height;
            this.grpboxEconOptimizSettings.Width = grpboxOptimization.Width;

		
		}

	
		

		private void btnFVSVariablesPrePostVariableValue_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnFVSVariablesPrePostVariableClearAll_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostVariableNext_Click(object sender, System.EventArgs e)
		{
			
		}
		private void RollBack()
		{
			RollBack_variable1();
			RollBack_variable2();
			RollBack_variable3();
			RollBack_SqlBetter();
			RollBack_SqlWorse();
			RollBack_Overall();
		}
		private void RollBack_variable1()
		{
		}
		private void RollBack_variable2()
		{
		}
		private void RollBack_variable3()
		{
		}
		private void RollBack_SqlBetter()
		{
		}
		private void RollBack_SqlWorse()
		{
		}
		private void RollBack_Overall()
		{
		}

		private void grpboxFVSVariablesPrePostVariable_Resize(object sender, System.EventArgs e)
		{
			
		}

		private void grpboxFVSVariablesPrePostExpression_Resize(object sender, System.EventArgs e)
		{
		}

		private void grpboxFVSVariablesPrePostVariableValues_Resize(object sender, System.EventArgs e)
		{

		}

		private void btnFVSVariablesPrePostExpressionPrevious_Click(object sender, System.EventArgs e)
		{
			
			
			
		}

	


		private void Go()
		{

		}
		
		
		private void ShowGroupBox(string p_strName)
		{
			int x;
			//System.Windows.Forms.Control oControl;
			for (x=0;x<=groupBox1.Controls.Count-1;x++)
			{
				if (groupBox1.Controls[x].Name.Substring(0,3)=="grp")
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

		private void btnFVSVariablesPrePostValuesButtonsEdit_Click(object sender, System.EventArgs e)
		{

		}


		private void btnFVSVariablesPrePost2Overall_Click(object sender, System.EventArgs e)
		{
		}

		private void btnFVSVariablesPrePostValuesButtonsClear_Click(object sender, System.EventArgs e)
		{
		}
		

		

		private void grpboxFVSVariablesPrePost_Resize(object sender, System.EventArgs e)
		{
			   
		}

		

		private void lvFVSVariablesPrePostValues_DoubleClick(object sender, System.EventArgs e)
		{
		
		}

		private void pnlFVSVariablesPrePostExpression_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private void pnlFVSVariablesPrePostVariable_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private void btnOptimiztionCancel_Click(object sender, System.EventArgs e)
		{
			this.EnableTabs(true);
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Hide();
            this.grpboxEconOptimizSettings.Hide();
			this.grpboxOptimization.Show();
			
		}

		private void btnOptimiztionDone_Click(object sender, System.EventArgs e)
		{
			WizardSave();
		}
		private void WizardSave()
		{
			this.ReferenceCoreScenarioForm.m_bSave=true;
			this.EnableTabs(true);
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Hide();
            this.grpboxEconOptimizSettings.Hide();
			this.grpboxOptimization.Show();
			for (int x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
			{
				// We only update the selected listview item
                if (this.m_oSavVariableCollection.Item(x).intListViewIndex==this.lvOptimizationListValues.SelectedItems[0].Index)
				{
					// ECONOMIC ATTRIBUTE
                    if (this.m_oSavVariableCollection.Item(x).strOptimizedVariable.Trim().ToUpper().Equals("ECONOMIC ATTRIBUTE"))
                    {
                        if (this.rdoEconOptimizeMaximum.Checked)
                        {
                            this.m_oSavVariableCollection.Item(x).strMaxYN = "Y";
                            this.m_oSavVariableCollection.Item(x).strMinYN = "N";
                        }
                        else
                        {
                            this.m_oSavVariableCollection.Item(x).strMaxYN = "N";
                            this.m_oSavVariableCollection.Item(x).strMinYN = "Y";
                        }
                        this.m_oSavVariableCollection.Item(x).strFVSVariableName =
                                this.lblEconomicAttribute.Text.Trim();
                        this.m_oSavVariableCollection.Item(x).strValueSource = "NA";
                        this.m_oSavVariableCollection.Item(x).bUseFilter = this.chkEnableEconFilter.Checked;
                        this.m_oSavVariableCollection.Item(x).strFilterOperator = this.cmbEconOptimizationOperator.Text.Trim();
                        this.m_oSavVariableCollection.Item(x).dblFilterValue = Convert.ToDouble(ValidateNumeric(this.txtEconOptimizationValue.Text.Trim()));
                        this.m_oSavVariableCollection.Item(x).strRevenueAttribute = this.cmbNetRevEconOptimzFilter.Text.Trim();
                    }
                    else
                    {					
                        if (this.rdoOptimizationMaximum.Checked)
					    {
						    this.m_oSavVariableCollection.Item(x).strMaxYN="Y";
						    this.m_oSavVariableCollection.Item(x).strMinYN="N";
					    }
					    else
					    {
						    this.m_oSavVariableCollection.Item(x).strMaxYN="N";
						    this.m_oSavVariableCollection.Item(x).strMinYN="Y";

					    }
					    if (m_oSavVariableCollection.Item(x).intListViewIndex > 1)
					    {
						    this.m_oSavVariableCollection.Item(x).strFVSVariableName=
							    this.lblFVSVariablesOptimizationVariableValuesSelected.Text.Trim();
						    if (this.cmbOptimizationSettingsPostPreValue.Text.Trim().ToUpper()==
							    this.cmbOptimizationSettingsPostPreValue.Items[0].ToString().Trim().ToUpper())
							    this.m_oSavVariableCollection.Item(x).strValueSource="POST";
						    else if (this.cmbOptimizationSettingsPostPreValue.Text.Trim().ToUpper()==
							    this.cmbOptimizationSettingsPostPreValue.Items[1].ToString().Trim().ToUpper())
							    this.m_oSavVariableCollection.Item(x).strValueSource="POST-PRE";
						    else
							    this.m_oSavVariableCollection.Item(x).strValueSource="Not Defined";
					    }
					    else
						    m_oSavVariableCollection.Item(x).strValueSource="NA";
					    this.m_oSavVariableCollection.Item(x).bUseFilter=this.chkEnableFilter.Checked;
					    this.m_oSavVariableCollection.Item(x).strFilterOperator=this.cmbOptimizationOperator.Text.Trim();
                        this.m_oSavVariableCollection.Item(x).dblFilterValue = Convert.ToDouble(ValidateNumeric(this.txtOptimizationValue.Text.Trim()));
                        this.m_oSavVariableCollection.Item(x).strRevenueAttribute = this.cmbNetRevOptimzFilter.Text.Trim();
                    }
                    this.UpdateListViewItem(lvOptimizationListValues.SelectedItems[0], m_oSavVariableCollection.Item(x));
                    break;
				}
			}

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
		private void lvOptimizationListValues_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			bool bValid=true;
			int x,y;
            int intSelectedIndex = -1;
			
			this.ReferenceCoreScenarioForm.m_bSave=true;
            if (!m_bIgnoreListViewItemCheck) this.lvOptimizationListValues.Items[e.Index].Selected = true;
            
                
            
			if (e.CurrentValue==CheckState.Checked) 
			{
				for (x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
				{
					if (e.Index == this.m_oSavVariableCollection.Item(x).intListViewIndex)
					{
						this.m_oSavVariableCollection.Item(x).bSelected=false;
						break;
					}
				}
				//find the current index
				
			}
			else 
			{
				
					for (x=0;x<=this.lvOptimizationListValues.Items.Count-1;x++)
					{
						
						if (x!=e.Index)
						{
							m_bIgnoreListViewItemCheck=true;
							this.lvOptimizationListValues.Items[x].Checked=false;
							m_bIgnoreListViewItemCheck=false;
							for (y=0;y<=this.m_oSavVariableCollection.Count-1;y++)
							{
								if (this.m_oSavVariableCollection.Item(y).intListViewIndex==x)
								{
									this.m_oSavVariableCollection.Item(y).bSelected=false;
								}
							}
						}
						else
						{
							for (y=0;y<=this.m_oSavVariableCollection.Count-1;y++)
							{
								if (this.m_oSavVariableCollection.Item(y).intListViewIndex==x)
								{
									this.m_oSavVariableCollection.Item(y).bSelected=true;
                                    intSelectedIndex = y;
								}
							}
						}
						
					}
				
			}
            if (!m_bIgnoreListViewItemCheck)
            {
                if (intSelectedIndex == -1)
                {
                    this.ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText("None Selected");
                }
                else
                {
                    if (m_oSavVariableCollection.Item(intSelectedIndex).strOptimizedVariable.Trim().ToUpper() == "STAND ATTRIBUTE")
                        this.ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText(m_oSavVariableCollection.Item(intSelectedIndex).strFVSVariableName);
                    else
                        this.ReferenceCoreScenarioForm.uc_scenario_run1.UpdateOptimizationVariableGroupboxText(m_oSavVariableCollection.Item(intSelectedIndex).strOptimizedVariable);
                }
            }
		}

        private void pnlOptimization_Resize(object sender, System.EventArgs e)
		{
			grpboxOptimizationListValues.Width = this.pnlOptimization.ClientSize.Width - (grpboxOptimizationListValues.Left * 2);
		}

		private void btnOptimizationEdit_Click(object sender, System.EventArgs e)
		{
			if (this.lvOptimizationListValues.SelectedItems.Count==0) return;

            // Get item out of memory so we can access the revenue attribute
            VariableItem oSelectedItem = null;
            foreach (VariableItem oItem in m_oSavVariableCollection)
            {
                if (oItem.strOptimizedVariable.Equals(
                    this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_OPTIMIZE_VARIABLE].Text.Trim()))
                {
                    oSelectedItem = oItem;
                    break;
                }
            }

			//Either Revenue or Merch Volume has been selected
            if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FVS_VARIABLE].Text.Trim()=="NA")
			{
				btnOptimiztionPrev.Hide();
				this.lblOptimizationVariable.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_OPTIMIZE_VARIABLE].Text.Trim();
                if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_OPERATOR].Text.Trim().ToUpper() != "NOT DEFINED")
                {
                    this.cmbOptimizationOperator.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_OPERATOR].Text.Trim();
                }
                this.txtOptimizationValue.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_VALUE].Text.Trim();
                this.m_strLastValue = this.txtOptimizationValue.Text;

                if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_USEFILTER].Text.Trim() == "Yes")
                {
                    this.chkEnableFilter.Checked = true;
                }
                else
                {
                    this.chkEnableFilter.Checked = false;
                }
                this.grpMaxMin.Location = grpboxOptimizationSettingsPostPre.Location;
                // Hide the second (FVS) layer so the third is visible
                this.grpboxOptimizationSettingsPostPre.Hide();
                // Hide the fourth (Economic) layer
                this.grpboxEconOptimizSettings.Hide();
                // Show the third layer
                this.grpboxOptimizationSettings.Show();
			}
            //Economic attribute was selected
            else if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_OPTIMIZE_VARIABLE].Text.Trim() == "Economic Attribute")
            {
                if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FVS_VARIABLE].Text.Trim() != "Not Defined")
                {
                    for (int index = 0; index < lstEconVariables.Items.Count + 1; index++)
                    {
                        string item = lstEconVariables.Items[index].ToString();
                        if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FVS_VARIABLE].Text.Trim() == item)
                        {
                            lstEconVariables.SelectedIndex = index;
                            break;
                        }
                    }
                }

                if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_MAXMIN].Text.Trim() == "MAX")
                    this.rdoEconOptimizeMaximum.Checked = true;
                else if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_MAXMIN].Text.Trim() == "MIN")
                    this.rdoEconOptimizeMinimum.Checked = true;
                else
                    this.rdoEconOptimizeMaximum.Checked = true;

                // Revenue filter settings
                this.cmbNetRevEconOptimzFilter.SelectedIndex = -1;
                this.cmbEconOptimizationOperator.Text = ">";
                if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_USEFILTER].Text.Trim() == "Yes")
                {
                    this.chkEnableEconFilter.Checked = true;
                    if (!String.IsNullOrEmpty(oSelectedItem.strRevenueAttribute))
                    {
                        cmbNetRevEconOptimzFilter.SelectedItem = oSelectedItem.strRevenueAttribute;
                        if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_OPERATOR].Text.Trim().ToUpper() != "NOT DEFINED")
                        {
                            this.cmbEconOptimizationOperator.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_OPERATOR].Text.Trim();
                        }
                    }
                    this.txtEconOptimizationValue.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_VALUE].Text.Trim();
                    this.m_strLastValue = this.txtEconOptimizationValue.Text;
                }
                else
                {
                    this.chkEnableEconFilter.Checked = false;
                }
                
                // Hide the second (FVS) layer so the fourth is visible
                this.grpboxOptimizationSettingsPostPre.Hide();
                // Hide the third layer; This layer replaces it
                this.grpboxOptimizationSettings.Hide();
                // Show the fourth layer
                this.grpboxEconOptimizSettings.Show();
            }
			else
			{
				// Populate the fields on the second (FVS) layer
                btnOptimiztionPrev.Show();
                this.loadFVSTableAndField();
                this.lblFVSVariablesOptimizationVariableValuesSelected.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FVS_VARIABLE].Text.Trim();
                this.lblOptimizationVariable.Text = this.lblFVSVariablesOptimizationVariableValuesSelected.Text;
                this.grpMaxMin.Location = new System.Drawing.Point(457, 55);
                // Show the second layer on top
                this.grpboxOptimizationFVSVariable.Show();
				if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_VALUESOURCE].Text.Trim()=="POST")
					this.cmbOptimizationSettingsPostPreValue.Text = 
						this.cmbOptimizationSettingsPostPreValue.Items[0].ToString().Trim();
				else if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_VALUESOURCE].Text.Trim()=="POST-PRE")
					this.cmbOptimizationSettingsPostPreValue.Text =
						this.cmbOptimizationSettingsPostPreValue.Items[1].ToString().Trim();
				else
					this.cmbOptimizationSettingsPostPreValue.Text = 
						this.cmbOptimizationSettingsPostPreValue.Items[0].ToString().Trim();

                // Revenue filter settings
                this.cmbNetRevOptimzFilter.SelectedIndex = -1;
                this.cmbOptimizationOperator.Text = ">";
                if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_USEFILTER].Text.Trim() == "Yes")
                {
                    this.chkEnableFilter.Checked = true;
                    if (!String.IsNullOrEmpty(oSelectedItem.strRevenueAttribute))
                    {
                        cmbNetRevOptimzFilter.SelectedItem = oSelectedItem.strRevenueAttribute;
                        if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_OPERATOR].Text.Trim().ToUpper() != "NOT DEFINED")
                        {
                            this.cmbOptimizationOperator.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_OPERATOR].Text.Trim();
                        }
                    }
                    this.txtOptimizationValue.Text = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FILTER_VALUE].Text.Trim();
                    this.m_strLastValue = this.txtOptimizationValue.Text;
                }
                else
                {
                    this.chkEnableFilter.Checked = false;
                }
			}

			// These fields are common to FVS variable, Merchantable Volume, and Revenue options. Not breaking these out for now.
            // The interface only saves the values for the selected item, so setting these incidentally shouldn't be a problem
            if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_MAXMIN].Text.Trim()=="MAX")
				this.rdoOptimizationMaximum.Checked=true;
			else if (this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_MAXMIN].Text.Trim()=="MIN")
				this.rdoOptimizationMinimum.Checked=true;
			else
				this.rdoOptimizationMaximum.Checked=true;

			// Hide the top Layer
            this.grpboxOptimization.Hide();
            // Fourth layer should be hidden because this is third
			this.EnableTabs(false);
			
			
		}

		private void groupBox1_Leave(object sender, System.EventArgs e)
		{
			
		}
		private void EnableTabs(bool p_bEnable)
		{
            int x;
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlScenario,"tbdesc,tbnotes,tbdatasources",p_bEnable);
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlRules,"tbpsites,tbowners,tbcost,tbtreatmentintensity,tbfilterplots,tbrun",p_bEnable);
			ReferenceCoreScenarioForm.EnableTabPage(ReferenceCoreScenarioForm.tabControlFVSVariables,"tbeffective,tbtiebreaker",p_bEnable);
            for (x = 0; x <= ReferenceCoreScenarioForm.tlbScenario.Buttons.Count - 1; x++)
            {
                ReferenceCoreScenarioForm.tlbScenario.Buttons[x].Enabled = p_bEnable;
            }
            frmMain.g_oFrmMain.grpboxLeft.Enabled = p_bEnable;
            frmMain.g_oFrmMain.tlbMain.Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[0].Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[1].Enabled = p_bEnable;
            frmMain.g_oFrmMain.mnuMain.MenuItems[2].Enabled = p_bEnable;

		}

		private void btnOptimizationAudit_Click(object sender, System.EventArgs e)
		{
			this.DisplayAuditMessage=true;
			Audit();
		}
		public void Audit()
		{
			
			
			int x;
			this.m_intError=0;
			m_strError="";
			if (DisplayAuditMessage)
			{
				this.m_strError="Audit Results \r\n";
				this.m_strError=m_strError + "-------------\r\n\r\n";
			}

			

			for (x=0;x<=this.m_oSavVariableCollection.Count-1;x++)
			{
				if (this.m_oSavVariableCollection.Item(x).bSelected) break;
			}
			if (x > this.m_oSavVariableCollection.Count-1)
			{
				m_intError=-1;
				m_strError = m_strError + "No optimization variable was checked. \r\n";
			}
			else
			{

				if (this.m_oSavVariableCollection.Item(x).strFVSVariableName.Trim().ToUpper() == "NOT DEFINED" &&
					this.m_oSavVariableCollection.Item(x).bSelected)
				{
					m_intError=-1;
                    m_strError = m_strError + "Stand Attribute is selected but is not defined\r\n";
				}
				else
				{
					if (this.m_oSavVariableCollection.Item(x).strFVSVariableName.Trim().ToUpper() != "NA" && 
						this.m_oSavVariableCollection.Item(x).strFVSVariableName.Trim().ToUpper() != "NOT DEFINED" &&
                        this.m_oSavVariableCollection.Item(x).strFVSVariableName.IndexOf(".") > -1)
					{
						if (System.IO.File.Exists(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb"))
						{
							ado_data_access oAdo = new ado_data_access();
							oAdo.OpenConnection(oAdo.getMDBConnString(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\fvs\\db\\biosum_fvsout_prepost_rx.mdb","",""));
							if (oAdo.m_intError==0)
							{
								string[] strTableColumnArray = frmMain.g_oUtils.ConvertListToArray(this.m_oSavVariableCollection.Item(x).strFVSVariableName,".");
								string strTable = strTableColumnArray[0].Trim();
								string strColumn = strTableColumnArray[1].Trim();
								if (oAdo.TableExist(oAdo.m_OleDbConnection,"post_" + strTable))
								{
									if (oAdo.ColumnExist(oAdo.m_OleDbConnection,"post_" + strTable,strColumn)==false)
									{
										m_intError=-1;
										m_strError=m_strError + "Table column post_" + strTable + "." + strColumn + " does not exist in Db file biosum_fvsout_prepost_rx.mdb\r\n";
									}
								}
								else
								{
									m_intError=-1;
									m_strError = m_strError + "Table post_" + strTable + " does not exist in Db file biosum_fvsout_prepost_rx.mdb\r\n";
								}
							
								if (this.m_oSavVariableCollection.Item(x).strValueSource.Trim().ToUpper()=="POST-PRE")
								{
									if (oAdo.TableExist(oAdo.m_OleDbConnection,"pre_" + strTable))
									{
										if (oAdo.ColumnExist(oAdo.m_OleDbConnection,"pre_" + strTable,strColumn)==false)
										{
											m_intError=-1;
											m_strError=m_strError + "Table column pre_" + strTable + "." + strColumn + " does not exist in Db file biosum_fvsout_prepost_rx.mdb\r\n";
										}
									}
									else
									{
										m_intError=-1;
										m_strError = m_strError + "Table pre_" + strTable + " does not exist in Db file biosum_fvsout_prepost_rx.mdb\r\n";
									}
								}
							
								oAdo.CloseConnection(oAdo.m_OleDbConnection);

							}
							else
							{
								m_intError=-1;
								m_strError= m_strError + "Error making a db connection to biosum_fvsout_prepost_rx.mdb\r\n";
							}
						}
						else
						{
							m_intError=-1;
							m_strError = m_strError + "Db file biosum_fvsout_prepost_rx.mdb does not exist\r\n";
						}

					}

					if (this.m_oSavVariableCollection.Item(x).strMaxYN=="N" &&
						this.m_oSavVariableCollection.Item(x).strMinYN=="N")
					{
						m_intError=-1;
						m_strError = m_strError + "The optimization variable requires a MAXIMUM or MINIMUM definition. \r\n";
					}

					if (this.m_oSavVariableCollection.Item(x).bUseFilter)
					{
						switch (this.m_oSavVariableCollection.Item(x).strFilterOperator.Trim())
						{ 
							case ">": break;
							case "<": break;
							case "<>": break;
							case ">=": break;
							case "<=": break;
							default:
								m_intError=-1;
								m_strError = m_strError + "Invalid filter operator for the optimization variable. \r\n";
								break;
						}
                        if (String.IsNullOrEmpty(this.m_oSavVariableCollection.Item(x).strRevenueAttribute))
                        {
                            m_intError = -1;
                            m_strError = m_strError + "Net Revenue calculation is required when using Net Revenue filter \r\n";
                        }

					}
				}
				this.ReferenceTieBreaker.AuditCheckOptimiztionAndTieBreakerVariable(ref m_intError,ref m_strError);
			}
			
			if (DisplayAuditMessage)
			{
				if (m_intError==0) this.m_strError=m_strError + "Passed Audit";
				else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";
				MessageBox.Show(m_strError,"FIA Biosum");
			}

		}

		private void lvOptimizationListValues_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			int x;
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lvOptimizationListValues.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lvOptimizationListValues.Items[lvOptimizationListValues.TopItem.Index + (int)dblRow-1].Selected=true;
				}
			}
			catch 
			{
			}
		}

		private void lvOptimizationListValues_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lvOptimizationListValues.SelectedItems.Count > 0)
				this.m_oLvRowColors.DelegateListViewItem(this.lvOptimizationListValues.SelectedItems[0]);
		}

		private void btnFVSVariablesOptimizationVariableValues_Click(object sender, System.EventArgs e)
		{
            if (this.lstFVSTablesList.SelectedItems.Count == 0 || this.lstFVSFieldsList.SelectedItems.Count == 0) return;
            this.lblFVSVariablesOptimizationVariableValuesSelected.Text =
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();
		}

		private void btnOptimizationFVSVariableDone_Click(object sender, System.EventArgs e)
		{
			WizardSave();
		}

		private void btnOptimizationFVSVariableCancel_Click(object sender, System.EventArgs e)
		{
			this.EnableTabs(true);
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Hide();
			this.grpboxOptimization.Show();

		}

		private void btnOptimizationFVSVariableNext_Click(object sender, System.EventArgs e)
		{
            if (this.lblFVSVariablesOptimizationVariableValuesSelected.Text == "Not Defined")
            {
                System.Windows.Forms.MessageBox.Show("Use the Select button to choose a stand attribute before proceeding!", "FIA Biosum");
                return;
            }
            this.lblOptimizationVariable.Text = this.lblFVSVariablesOptimizationVariableValuesSelected.Text;
            this.grpboxOptimizationSettingsPostPre.Show();
            this.grpboxOptimizationFVSVariable.Hide();
			this.grpboxOptimizationSettings.Show();
            this.grpBoxOptimizationNetRevenue.Show();

		}

		private void btnOptimiztionPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxOptimizationSettings.Hide();
			this.grpboxOptimizationFVSVariable.Show();			
			
		}


        private void loadFVSTableAndField()
        {
            string[] strPieces = this.lvOptimizationListValues.SelectedItems[0].SubItems[COLUMN_FVS_VARIABLE].Text.Trim().Split('.');
            //set FVS table and variable
            if (strPieces.Length == 2)
            {
                for (int index = 0; index < lstFVSTablesList.Items.Count + 1; index++)
                {
                    string item = lstFVSTablesList.Items[index].ToString();
                    if (strPieces[0] == item)
                    {
                        lstFVSTablesList.SelectedIndex = index;
                        break;
                    }
                }
                for (int index = 0; index < lstFVSFieldsList.Items.Count + 1; index++)
                {
                    string item = lstFVSFieldsList.Items[index].ToString();
                    if (strPieces[1] == item)
                    {
                        lstFVSFieldsList.SelectedIndex = index;
                        break;
                    }
                }
            }
        }
		
		public bool DisplayAuditMessage
		{
			get {return _bDisplayAuditMsg;}
			set {_bDisplayAuditMsg=value;}
		}
		public FIA_Biosum_Manager.frmCoreScenario ReferenceCoreScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_effective.Variables ReferenceFVSVariables
		{
			get {return this._oCurVar;}
			set {_oCurVar=value;}
		}
		public FIA_Biosum_Manager.uc_core_scenario_fvs_prepost_variables_tiebreaker ReferenceTieBreaker
		{
			get {return _uc_tiebreaker;}
			set {_uc_tiebreaker=value;}
		}

        private void txtOptimizationValue_Leave(object sender, EventArgs e)
        {
            this.m_oValidate.ValidateDecimal(txtOptimizationValue.Text);
            if (m_oValidate.m_intError == 0)
            {
                this.txtOptimizationValue.Text = m_oValidate.ReturnValue;
                this.m_strLastValue = m_oValidate.ReturnValue;
            }
            else
            {
                this.txtOptimizationValue.Text = this.m_strLastValue;
            }

        }

        private void txtEconOptimizationValue_Leave(object sender, EventArgs e)
        {
            this.m_oValidate.ValidateDecimal(txtEconOptimizationValue.Text);
            if (m_oValidate.m_intError == 0)
            {
                this.txtEconOptimizationValue.Text = m_oValidate.ReturnValue;
                this.m_strLastValue = m_oValidate.ReturnValue;
            }
            else
            {
                this.txtEconOptimizationValue.Text = this.m_strLastValue;
            }

        }

        private void lstFVSTablesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstFVSFieldsList.Items.Clear();
            txtOptVarDescr.Text = "";
            this.lblFVSVariablesOptimizationVariableValuesSelected.Text = "Not Defined";
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
                // Control visibility of the weighted variable description fields
                bool bDisplayWeightedFields = this.lstFVSTablesList.SelectedItems[0].ToString().ToUpper().Contains("_WEIGHTED");
                lblOptVarDescr.Visible = bDisplayWeightedFields;
                txtOptVarDescr.Visible = bDisplayWeightedFields;
            }
            else
            {
                lblOptVarDescr.Visible = false;
                txtOptVarDescr.Visible = false;
            }
        }

        private void lstFVSFieldsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.lblFVSVariablesOptimizationVariableValuesSelected.Text = "Not Defined";
            this.txtOptVarDescr.Text = "";
            if (this.lstFVSFieldsList.SelectedIndex > -1)
            {
                this.btnFVSVariablesOptimizationVariableValues.Enabled = true;
                if (this.lstFVSTablesList.SelectedItems[0].ToString().ToUpper().Contains("_WEIGHTED") == true)
                {
                    foreach (uc_core_scenario_weighted_average.VariableItem oItem in this.ReferenceCoreScenarioForm.m_oWeightedVariableCollection)
                    {
                        if (oItem.strVariableName.Equals(Convert.ToString(this.lstFVSFieldsList.SelectedItem)))
                        {
                            if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                                txtOptVarDescr.Text = oItem.strVariableDescr;
                            break;
                        }
                    }
                }
            }
            else
            {
                this.btnFVSVariablesOptimizationVariableValues.Enabled = false;
            }
        }

        private void btnEconSelect_Click(object sender, EventArgs e)
        {
            if (this.lstEconVariables.SelectedItems.Count == 0 || this.lstEconVariables.SelectedItems.Count == 0) return;
            this.lblEconomicAttribute.Text =
                this.lstEconVariables.SelectedItems[0].ToString();
        }

        private void btnEconOptimizationCancel_Click(object sender, EventArgs e)
        {
            this.EnableTabs(true);
            this.grpboxOptimizationSettings.Hide();
            this.grpboxOptimizationFVSVariable.Hide();
            this.grpboxEconOptimizSettings.Hide();
            this.grpboxOptimization.Show();
        }

        private void BtnEconOptimizationDone_Click(object sender, EventArgs e)
        {
            WizardSave();
        }

        private void lstEconVariables_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.lblEconomicAttribute.Text = "Not Defined";
            this.txtEconAttribDescr.Text = "";
            if (this.lstEconVariables.SelectedIndex > -1)
            {
                this.btnEconSelect.Enabled = true;
                foreach (uc_core_scenario_weighted_average.VariableItem oItem in this.ReferenceCoreScenarioForm.m_oWeightedVariableCollection)
                {
                    if (oItem.strVariableName.Equals(Convert.ToString(this.lstEconVariables.SelectedItem)))
                    {
                        if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                        {
                            this.txtEconAttribDescr.Text = oItem.strVariableDescr;
                            this.lblEconomicAttribute.Text = oItem.strVariableName;
                            break;
                        }
                    }
                }
            }
            else
            {
                this.btnEconSelect.Enabled = false;
            }
        }

        private void chkEnableEconFilter_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEnableEconFilter.Checked == false)
            {
                this.cmbNetRevEconOptimzFilter.SelectedIndex = -1;
                this.cmbEconOptimizationOperator.Text = ">";
                this.txtEconOptimizationValue.Text = "0";
            }
        }

        private void chkEnableFilter_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEnableFilter.Checked == false)
            {
                this.cmbNetRevOptimzFilter.SelectedIndex = -1;
                this.cmbOptimizationOperator.Text = ">";
                this.txtOptimizationValue.Text = "0";
            }
        }

        private void cmbNetRevEconOptimzFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.txtEconRevenueDescr.Text = "";
            if (this.cmbNetRevEconOptimzFilter.SelectedIndex > -1)
            {
                foreach (uc_core_scenario_weighted_average.VariableItem oItem in this.ReferenceCoreScenarioForm.m_oWeightedVariableCollection)
                {
                    if (oItem.strVariableName.Equals(Convert.ToString(this.cmbNetRevEconOptimzFilter.SelectedItem)))
                    {
                        if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                        {
                            this.txtEconRevenueDescr.Text = oItem.strVariableDescr;
                            break;
                        }
                    }
                }
            }
        }

        private void cmbNetRevOptimzFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.txtRevenueDescr.Text = "";
            if (this.cmbNetRevOptimzFilter.SelectedIndex > -1)
            {
                foreach (uc_core_scenario_weighted_average.VariableItem oItem in this.ReferenceCoreScenarioForm.m_oWeightedVariableCollection)
                {
                    if (oItem.strVariableName.Equals(Convert.ToString(this.cmbNetRevOptimzFilter.SelectedItem)))
                    {
                        if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                        {
                            this.txtRevenueDescr.Text = oItem.strVariableDescr;
                            break;
                        }
                    }
                }
            }
        }
	
	}
}

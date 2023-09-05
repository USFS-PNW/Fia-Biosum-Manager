using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using SQLite.ADO;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_ffe.
	/// </summary>
	public class uc_optimizer_scenario_fvs_prepost_variables_tiebreaker: System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
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
		public System.Windows.Forms.Label lblTitle;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_optimization _uc_optimization;


				
		private int m_intCurVar=-1;
		int m_intCurVariableDefinitionStepCount=1;
		string[] m_strUserNavigation=null;
		

		public const byte NUMBER_OF_VARIABLES=4;
		const byte VARIABLE_DEFINITION_STEPS=4;


		const int WIZARD_STEP_VARIABLES_DEFINED=0;
		const int WIZARD_STEP_VARIABLE_SELECT=1;
		const int WIZARD_STEP_VARIABLE_BETTER=2;
		const int WIZARD_STEP_VARIABLE_WORSE=3;
		const int WIZARD_STEP_VARIABLE_EFFECTIVE=4;
		const int WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE=5;

		
		//const int COLUMN_OPTIMIZE=0;
		const int COLUMN_CHECKBOX=0;
		const int COLUMN_METHOD=1;
		const int COLUMN_FVSVARIABLE=2;
		const int COLUMN_VALUESOURCE=3;
		const int COLUMN_MAXMIN=4;

        private string m_strHelpChapter = "TIEBREAKER_SETTINGS";

		public bool m_bSave=false;
		private bool _bDisplayAuditMsg=true;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreaker;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.ColumnHeader lvColChecked;
		private System.Windows.Forms.ColumnHeader lvColMethod;
		private System.Windows.Forms.ColumnHeader lvColStandAttribute;
		private System.Windows.Forms.ColumnHeader lvColMinMax;
		private System.Windows.Forms.GroupBox grpboxStandAttributeTieBreakerVariable;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerVariableValues;
        private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableValues;
		private System.Windows.Forms.GroupBox grpFVSVariablesTieBreakerVariableValuesSelected;
		private System.Windows.Forms.Label lblFVSVariablesTieBreakerVariableValuesSelected;
		private System.Windows.Forms.GroupBox grpMaxMin;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerLastTieBreakRank;
		private System.Windows.Forms.Panel panel2;
		public FIA_Biosum_Manager.uc_optimizer_scenario_last_tiebreak_rank uc_scenario_last_tiebreak_rank1;
		private System.Windows.Forms.Panel pnlTieBreaker;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerValues;
		private System.Windows.Forms.ListView lvFVSVariablesTieBreakerValues;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerEdit;
        private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors= new ListViewAlternateBackgroundColors();

		
		private System.Windows.Forms.RadioButton rdoFVSVariablesTieBreakerVariableValuesSelectedMin;
		private System.Windows.Forms.RadioButton rdoFVSVariablesTieBreakerVariableValuesSelectedMax;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityPrev;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityClear;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityDone;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerTreatmentIntensityCancel;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableClear;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableDone;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableCancel;
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerVariableNext;
		private TieBreaker_Collection m_oNewTieBreakerCollection = new TieBreaker_Collection();
		private TieBreaker_Collection m_oOldTieBreakerCollection = new TieBreaker_Collection();
		private System.Windows.Forms.Button btnFVSVariablesTieBreakerAudit;
		private System.Windows.Forms.ColumnHeader lvColFieldSource;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesTieBreakerVariableValueSource;
		private System.Windows.Forms.ComboBox cmbFVSVariablesTieBreakerVariableValueSource;
        private Label label1;
        private Label label2;
        private ListBox lstFVSFieldsList;
        private ListBox lstFVSTablesList;
		public TieBreaker_Collection m_oSavTieBreakerCollection = new TieBreaker_Collection();
        private TextBox txtTieBreakVarDescr;
        private Label lblTieBreakVarDescr;
        private ListBox lstEconVariablesList;
        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> _dictFVSTables;
        private Point _objGrpMaxMinLocation;
        private Point _objLblTieBreakVarDescrLocation;
        private Button BtnTiebreakerClear;
        private Point _objtxtTieBreakVarDescrLocation;


		public class TieBreakerItem
		{
			public bool bSelected=true;
			public string strMethod="";
			public string strFVSVariableName="";
			public string strMaxYN="N";
			public string strMinYN="N";
			public string strValueSource="";
			public int intListViewIndex=-1;
			
			

			public void Copy(TieBreakerItem p_oSource,ref TieBreakerItem p_oDest)
			{
				p_oDest.bSelected = p_oSource.bSelected;
				p_oDest.strMethod=p_oSource.strMethod;
				p_oDest.strFVSVariableName = p_oSource.strFVSVariableName;
				p_oDest.strMaxYN = p_oSource.strMaxYN;
				p_oDest.strMinYN = p_oSource.strMinYN;
				p_oDest.strValueSource=p_oSource.strValueSource;
				p_oDest.intListViewIndex = p_oSource.intListViewIndex;
				

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

			public void Add(FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem m_oTieBreaker)
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
			public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem Item(int Index)
			{
				// The appropriate item is retrieved from the List object and
				// explicitly cast to the Widget type, then returned to the 
				// caller.
				return (FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem) List[Index];
			}
			public void Copy(TieBreaker_Collection p_oSource,ref TieBreaker_Collection p_oDest,bool p_bInitializeDest)
			{
				int x;
				if (p_bInitializeDest) p_oDest.Clear();
				for (x=0;x<=p_oSource.Count-1;x++)
				{
					TieBreakerItem oItem = new TieBreakerItem();
					oItem.Copy(p_oSource.Item(x),ref oItem);
					p_oDest.Add(oItem);

				}
			}


		}

		public uc_optimizer_scenario_fvs_prepost_variables_tiebreaker()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			if (frmMain.g_oGridViewFont != null) this.lvFVSVariablesTieBreakerValues.Font = frmMain.g_oGridViewFont;
			
			for (int x=0;x<=this.lvFVSVariablesTieBreakerValues.Items.Count-1;x++)
				this.lvFVSVariablesTieBreakerValues.Items[x].UseItemStyleForSubItems=false;
			this.m_oLvRowColors.InitializeRowCollection();
			this.m_oLvRowColors.ReferenceAlternateBackgroundColor=frmMain.g_oGridViewAlternateRowBackgroundColor;
            this.m_oLvRowColors.ReferenceSelectedRowBackgroundColor = frmMain.g_oGridViewSelectedRowBackgroundColor;
			this.m_oLvRowColors.ReferenceBackgroundColor = frmMain.g_oGridViewRowBackgroundColor;
			this.m_oLvRowColors.ReferenceListView = this.lvFVSVariablesTieBreakerValues;
			this.m_oLvRowColors.CustomFullRowSelect=true;
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(0,this.lvFVSVariablesTieBreakerValues.Columns.Count);
			this.m_oLvRowColors.AddRow();
			this.m_oLvRowColors.AddColumns(1,this.lvFVSVariablesTieBreakerValues.Columns.Count);
            this.m_oLvRowColors.AddRow();
            this.m_oLvRowColors.AddColumns(2, this.lvFVSVariablesTieBreakerValues.Columns.Count);
			this.m_oLvRowColors.ListView();

			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Hide();
			this.grpboxStandAttributeTieBreakerVariable.Hide();
            this.lstEconVariablesList.Location = this.lstFVSTablesList.Location;
			this.grpboxFVSVariablesTieBreaker.Show();
        }

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(uc_optimizer_scenario_fvs_prepost_variables_tiebreaker));
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "Stand Attribute",
            "Not Defined",
            "Not Defined",
            "Not Defined"}, -1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "Economic Attribute",
            "Not Defined",
            "NA",
            "Not Defined"}, -1);
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "Last Tie-Break Rank",
            "NA",
            "NA",
            "MIN"}, -1);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank = new System.Windows.Forms.GroupBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel = new System.Windows.Forms.Button();
            this.grpboxStandAttributeTieBreakerVariable = new System.Windows.Forms.GroupBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.grpboxFVSVariablesTieBreakerVariableValueSource = new System.Windows.Forms.GroupBox();
            this.cmbFVSVariablesTieBreakerVariableValueSource = new System.Windows.Forms.ComboBox();
            this.grpMaxMin = new System.Windows.Forms.GroupBox();
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin = new System.Windows.Forms.RadioButton();
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax = new System.Windows.Forms.RadioButton();
            this.grpboxFVSVariablesTieBreakerVariableValues = new System.Windows.Forms.GroupBox();
            this.lstEconVariablesList = new System.Windows.Forms.ListBox();
            this.txtTieBreakVarDescr = new System.Windows.Forms.TextBox();
            this.lblTieBreakVarDescr = new System.Windows.Forms.Label();
            this.lstFVSFieldsList = new System.Windows.Forms.ListBox();
            this.btnFVSVariablesTieBreakerVariableValues = new System.Windows.Forms.Button();
            this.lstFVSTablesList = new System.Windows.Forms.ListBox();
            this.grpFVSVariablesTieBreakerVariableValuesSelected = new System.Windows.Forms.GroupBox();
            this.lblFVSVariablesTieBreakerVariableValuesSelected = new System.Windows.Forms.Label();
            this.btnFVSVariablesTieBreakerVariableClear = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerVariableDone = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerVariableCancel = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerVariableNext = new System.Windows.Forms.Button();
            this.grpboxFVSVariablesTieBreaker = new System.Windows.Forms.GroupBox();
            this.pnlTieBreaker = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.grpboxFVSVariablesTieBreakerValues = new System.Windows.Forms.GroupBox();
            this.lvFVSVariablesTieBreakerValues = new System.Windows.Forms.ListView();
            this.lvColChecked = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColMethod = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColStandAttribute = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColFieldSource = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColMinMax = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnFVSVariablesTieBreakerAudit = new System.Windows.Forms.Button();
            this.btnFVSVariablesTieBreakerEdit = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.BtnTiebreakerClear = new System.Windows.Forms.Button();
            this.uc_scenario_last_tiebreak_rank1 = new FIA_Biosum_Manager.uc_optimizer_scenario_last_tiebreak_rank();
            this.groupBox1.SuspendLayout();
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.SuspendLayout();
            this.panel2.SuspendLayout();
            this.grpboxStandAttributeTieBreakerVariable.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpboxFVSVariablesTieBreakerVariableValueSource.SuspendLayout();
            this.grpMaxMin.SuspendLayout();
            this.grpboxFVSVariablesTieBreakerVariableValues.SuspendLayout();
            this.grpFVSVariablesTieBreakerVariableValuesSelected.SuspendLayout();
            this.grpboxFVSVariablesTieBreaker.SuspendLayout();
            this.pnlTieBreaker.SuspendLayout();
            this.grpboxFVSVariablesTieBreakerValues.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.grpboxFVSVariablesTieBreakerLastTieBreakRank);
            this.groupBox1.Controls.Add(this.grpboxStandAttributeTieBreakerVariable);
            this.groupBox1.Controls.Add(this.grpboxFVSVariablesTieBreaker);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(900, 2000);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // grpboxFVSVariablesTieBreakerLastTieBreakRank
            // 
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Controls.Add(this.panel2);
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.ForeColor = System.Drawing.Color.Black;
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Location = new System.Drawing.Point(16, 1016);
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Name = "grpboxFVSVariablesTieBreakerLastTieBreakRank";
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Size = new System.Drawing.Size(872, 448);
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.TabIndex = 35;
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.TabStop = false;
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Text = "Last Tie-Break Rank";
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Visible = false;
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.VisibleChanged += new System.EventHandler(this.grpboxFVSVariablesTieBreakerLastTieBreakRank_VisibleChanged);
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Resize += new System.EventHandler(this.grpboxFVSVariablesTieBreakerTreatmentIntensity_Resize);
            // 
            // panel2
            // 
            this.panel2.AutoScroll = true;
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.uc_scenario_last_tiebreak_rank1);
            this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityPrev);
            this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityClear);
            this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityDone);
            this.panel2.Controls.Add(this.btnFVSVariablesTieBreakerTreatmentIntensityCancel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(3, 18);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(866, 427);
            this.panel2.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(9, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(837, 17);
            this.label1.TabIndex = 14;
            this.label1.Text = "Assign integer ranks for silvicultural sequences; lowest numbered sequence will b" +
    "e best when there is more than one optimal sequence";
            // 
            // btnFVSVariablesTieBreakerTreatmentIntensityPrev
            // 
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Location = new System.Drawing.Point(528, 376);
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Name = "btnFVSVariablesTieBreakerTreatmentIntensityPrev";
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.TabIndex = 12;
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Text = "<--Previous";
            this.btnFVSVariablesTieBreakerTreatmentIntensityPrev.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityPrev_Click);
            // 
            // btnFVSVariablesTieBreakerTreatmentIntensityClear
            // 
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Location = new System.Drawing.Point(100, 376);
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Name = "btnFVSVariablesTieBreakerTreatmentIntensityClear";
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Size = new System.Drawing.Size(72, 40);
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear.TabIndex = 5;
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Text = "Clear";
            this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityClear_Click);
            // 
            // btnFVSVariablesTieBreakerTreatmentIntensityDone
            // 
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Location = new System.Drawing.Point(352, 376);
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Name = "btnFVSVariablesTieBreakerTreatmentIntensityDone";
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone.TabIndex = 11;
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Text = "Done";
            this.btnFVSVariablesTieBreakerTreatmentIntensityDone.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityDone_Click);
            // 
            // btnFVSVariablesTieBreakerTreatmentIntensityCancel
            // 
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Location = new System.Drawing.Point(440, 376);
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Name = "btnFVSVariablesTieBreakerTreatmentIntensityCancel";
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.TabIndex = 9;
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Text = "Cancel";
            this.btnFVSVariablesTieBreakerTreatmentIntensityCancel.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerTreatmentIntensityCancel_Click);
            // 
            // grpboxStandAttributeTieBreakerVariable
            // 
            this.grpboxStandAttributeTieBreakerVariable.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxStandAttributeTieBreakerVariable.Controls.Add(this.panel1);
            this.grpboxStandAttributeTieBreakerVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxStandAttributeTieBreakerVariable.ForeColor = System.Drawing.Color.Black;
            this.grpboxStandAttributeTieBreakerVariable.Location = new System.Drawing.Point(16, 544);
            this.grpboxStandAttributeTieBreakerVariable.Name = "grpboxStandAttributeTieBreakerVariable";
            this.grpboxStandAttributeTieBreakerVariable.Size = new System.Drawing.Size(872, 448);
            this.grpboxStandAttributeTieBreakerVariable.TabIndex = 34;
            this.grpboxStandAttributeTieBreakerVariable.TabStop = false;
            this.grpboxStandAttributeTieBreakerVariable.Text = "Stand Attribute";
            this.grpboxStandAttributeTieBreakerVariable.Visible = false;
            this.grpboxStandAttributeTieBreakerVariable.VisibleChanged += new System.EventHandler(this.grpboxStandAttributeTieBreakerVariable_VisibleChanged);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.grpboxFVSVariablesTieBreakerVariableValueSource);
            this.panel1.Controls.Add(this.grpMaxMin);
            this.panel1.Controls.Add(this.grpboxFVSVariablesTieBreakerVariableValues);
            this.panel1.Controls.Add(this.grpFVSVariablesTieBreakerVariableValuesSelected);
            this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableClear);
            this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableDone);
            this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableCancel);
            this.panel1.Controls.Add(this.btnFVSVariablesTieBreakerVariableNext);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 18);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(866, 427);
            this.panel1.TabIndex = 12;
            // 
            // grpboxFVSVariablesTieBreakerVariableValueSource
            // 
            this.grpboxFVSVariablesTieBreakerVariableValueSource.Controls.Add(this.cmbFVSVariablesTieBreakerVariableValueSource);
            this.grpboxFVSVariablesTieBreakerVariableValueSource.Location = new System.Drawing.Point(8, 296);
            this.grpboxFVSVariablesTieBreakerVariableValueSource.Name = "grpboxFVSVariablesTieBreakerVariableValueSource";
            this.grpboxFVSVariablesTieBreakerVariableValueSource.Size = new System.Drawing.Size(344, 72);
            this.grpboxFVSVariablesTieBreakerVariableValueSource.TabIndex = 19;
            this.grpboxFVSVariablesTieBreakerVariableValueSource.TabStop = false;
            this.grpboxFVSVariablesTieBreakerVariableValueSource.Text = "Post Treatment Variable Or Pre/Post Treatment Change";
            // 
            // cmbFVSVariablesTieBreakerVariableValueSource
            // 
            this.cmbFVSVariablesTieBreakerVariableValueSource.Items.AddRange(new object[] {
            "Post Value",
            "Post - Pre  Change Value"});
            this.cmbFVSVariablesTieBreakerVariableValueSource.Location = new System.Drawing.Point(16, 40);
            this.cmbFVSVariablesTieBreakerVariableValueSource.Name = "cmbFVSVariablesTieBreakerVariableValueSource";
            this.cmbFVSVariablesTieBreakerVariableValueSource.Size = new System.Drawing.Size(320, 24);
            this.cmbFVSVariablesTieBreakerVariableValueSource.TabIndex = 0;
            this.cmbFVSVariablesTieBreakerVariableValueSource.Text = "Post Value";
            // 
            // grpMaxMin
            // 
            this.grpMaxMin.Controls.Add(this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin);
            this.grpMaxMin.Controls.Add(this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax);
            this.grpMaxMin.Location = new System.Drawing.Point(360, 296);
            this.grpMaxMin.Name = "grpMaxMin";
            this.grpMaxMin.Size = new System.Drawing.Size(464, 64);
            this.grpMaxMin.TabIndex = 18;
            this.grpMaxMin.TabStop = false;
            this.grpMaxMin.Text = "Which attribute value is best";
            // 
            // rdoFVSVariablesTieBreakerVariableValuesSelectedMin
            // 
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Location = new System.Drawing.Point(256, 16);
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Name = "rdoFVSVariablesTieBreakerVariableValuesSelectedMin";
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Size = new System.Drawing.Size(176, 40);
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.TabIndex = 14;
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Text = "Minimum value";
            // 
            // rdoFVSVariablesTieBreakerVariableValuesSelectedMax
            // 
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked = true;
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Location = new System.Drawing.Point(32, 16);
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Name = "rdoFVSVariablesTieBreakerVariableValuesSelectedMax";
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Size = new System.Drawing.Size(176, 40);
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.TabIndex = 12;
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.TabStop = true;
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Text = "Maximum value";
            // 
            // grpboxFVSVariablesTieBreakerVariableValues
            // 
            this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.lstEconVariablesList);
            this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.txtTieBreakVarDescr);
            this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.lblTieBreakVarDescr);
            this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.lstFVSFieldsList);
            this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.btnFVSVariablesTieBreakerVariableValues);
            this.grpboxFVSVariablesTieBreakerVariableValues.Controls.Add(this.lstFVSTablesList);
            this.grpboxFVSVariablesTieBreakerVariableValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxFVSVariablesTieBreakerVariableValues.Name = "grpboxFVSVariablesTieBreakerVariableValues";
            this.grpboxFVSVariablesTieBreakerVariableValues.Size = new System.Drawing.Size(816, 216);
            this.grpboxFVSVariablesTieBreakerVariableValues.TabIndex = 0;
            this.grpboxFVSVariablesTieBreakerVariableValues.TabStop = false;
            this.grpboxFVSVariablesTieBreakerVariableValues.Text = "Attribute for resolving ties when there is more than one optimal silvicultural se" +
    "quence";
            // 
            // lstEconVariablesList
            // 
            this.lstEconVariablesList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstEconVariablesList.ItemHeight = 16;
            this.lstEconVariablesList.Location = new System.Drawing.Point(134, 38);
            this.lstEconVariablesList.Name = "lstEconVariablesList";
            this.lstEconVariablesList.Size = new System.Drawing.Size(202, 180);
            this.lstEconVariablesList.Sorted = true;
            this.lstEconVariablesList.TabIndex = 91;
            this.lstEconVariablesList.Visible = false;
            this.lstEconVariablesList.SelectedIndexChanged += new System.EventHandler(this.lstEconVariablesList_SelectedIndexChanged);
            // 
            // txtTieBreakVarDescr
            // 
            this.txtTieBreakVarDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTieBreakVarDescr.Location = new System.Drawing.Point(553, 85);
            this.txtTieBreakVarDescr.Multiline = true;
            this.txtTieBreakVarDescr.Name = "txtTieBreakVarDescr";
            this.txtTieBreakVarDescr.ReadOnly = true;
            this.txtTieBreakVarDescr.Size = new System.Drawing.Size(258, 75);
            this.txtTieBreakVarDescr.TabIndex = 90;
            this.txtTieBreakVarDescr.Visible = false;
            // 
            // lblTieBreakVarDescr
            // 
            this.lblTieBreakVarDescr.Location = new System.Drawing.Point(472, 88);
            this.lblTieBreakVarDescr.Name = "lblTieBreakVarDescr";
            this.lblTieBreakVarDescr.Size = new System.Drawing.Size(80, 24);
            this.lblTieBreakVarDescr.TabIndex = 89;
            this.lblTieBreakVarDescr.Text = "Description:";
            this.lblTieBreakVarDescr.Visible = false;
            // 
            // lstFVSFieldsList
            // 
            this.lstFVSFieldsList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSFieldsList.ItemHeight = 16;
            this.lstFVSFieldsList.Location = new System.Drawing.Point(251, 21);
            this.lstFVSFieldsList.Name = "lstFVSFieldsList";
            this.lstFVSFieldsList.Size = new System.Drawing.Size(202, 180);
            this.lstFVSFieldsList.Sorted = true;
            this.lstFVSFieldsList.TabIndex = 2;
            this.lstFVSFieldsList.SelectedIndexChanged += new System.EventHandler(this.lstFVSFieldsList_SelectedIndexChanged);
            // 
            // btnFVSVariablesTieBreakerVariableValues
            // 
            this.btnFVSVariablesTieBreakerVariableValues.Enabled = false;
            this.btnFVSVariablesTieBreakerVariableValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesTieBreakerVariableValues.Location = new System.Drawing.Point(476, 21);
            this.btnFVSVariablesTieBreakerVariableValues.Name = "btnFVSVariablesTieBreakerVariableValues";
            this.btnFVSVariablesTieBreakerVariableValues.Size = new System.Drawing.Size(143, 57);
            this.btnFVSVariablesTieBreakerVariableValues.TabIndex = 1;
            this.btnFVSVariablesTieBreakerVariableValues.Text = "Select";
            this.btnFVSVariablesTieBreakerVariableValues.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableValues_Click);
            // 
            // lstFVSTablesList
            // 
            this.lstFVSTablesList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSTablesList.ItemHeight = 16;
            this.lstFVSTablesList.Location = new System.Drawing.Point(8, 21);
            this.lstFVSTablesList.Name = "lstFVSTablesList";
            this.lstFVSTablesList.Size = new System.Drawing.Size(202, 180);
            this.lstFVSTablesList.TabIndex = 0;
            this.lstFVSTablesList.SelectedIndexChanged += new System.EventHandler(this.lstFVSTablesList_SelectedIndexChanged);
            // 
            // grpFVSVariablesTieBreakerVariableValuesSelected
            // 
            this.grpFVSVariablesTieBreakerVariableValuesSelected.Controls.Add(this.lblFVSVariablesTieBreakerVariableValuesSelected);
            this.grpFVSVariablesTieBreakerVariableValuesSelected.Location = new System.Drawing.Point(8, 240);
            this.grpFVSVariablesTieBreakerVariableValuesSelected.Name = "grpFVSVariablesTieBreakerVariableValuesSelected";
            this.grpFVSVariablesTieBreakerVariableValuesSelected.Size = new System.Drawing.Size(816, 51);
            this.grpFVSVariablesTieBreakerVariableValuesSelected.TabIndex = 4;
            this.grpFVSVariablesTieBreakerVariableValuesSelected.TabStop = false;
            this.grpFVSVariablesTieBreakerVariableValuesSelected.Text = "Currently Active Tie Breaker";
            // 
            // lblFVSVariablesTieBreakerVariableValuesSelected
            // 
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Location = new System.Drawing.Point(3, 18);
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Name = "lblFVSVariablesTieBreakerVariableValuesSelected";
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Size = new System.Drawing.Size(810, 30);
            this.lblFVSVariablesTieBreakerVariableValuesSelected.TabIndex = 2;
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = "Not Defined";
            // 
            // btnFVSVariablesTieBreakerVariableClear
            // 
            this.btnFVSVariablesTieBreakerVariableClear.Location = new System.Drawing.Point(100, 376);
            this.btnFVSVariablesTieBreakerVariableClear.Name = "btnFVSVariablesTieBreakerVariableClear";
            this.btnFVSVariablesTieBreakerVariableClear.Size = new System.Drawing.Size(72, 40);
            this.btnFVSVariablesTieBreakerVariableClear.TabIndex = 5;
            this.btnFVSVariablesTieBreakerVariableClear.Text = "Clear";
            this.btnFVSVariablesTieBreakerVariableClear.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableClear_Click);
            // 
            // btnFVSVariablesTieBreakerVariableDone
            // 
            this.btnFVSVariablesTieBreakerVariableDone.Location = new System.Drawing.Point(352, 376);
            this.btnFVSVariablesTieBreakerVariableDone.Name = "btnFVSVariablesTieBreakerVariableDone";
            this.btnFVSVariablesTieBreakerVariableDone.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesTieBreakerVariableDone.TabIndex = 11;
            this.btnFVSVariablesTieBreakerVariableDone.Text = "Done";
            this.btnFVSVariablesTieBreakerVariableDone.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableDone_Click);
            // 
            // btnFVSVariablesTieBreakerVariableCancel
            // 
            this.btnFVSVariablesTieBreakerVariableCancel.Location = new System.Drawing.Point(440, 376);
            this.btnFVSVariablesTieBreakerVariableCancel.Name = "btnFVSVariablesTieBreakerVariableCancel";
            this.btnFVSVariablesTieBreakerVariableCancel.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesTieBreakerVariableCancel.TabIndex = 9;
            this.btnFVSVariablesTieBreakerVariableCancel.Text = "Cancel";
            this.btnFVSVariablesTieBreakerVariableCancel.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableCancel_Click);
            // 
            // btnFVSVariablesTieBreakerVariableNext
            // 
            this.btnFVSVariablesTieBreakerVariableNext.Location = new System.Drawing.Point(616, 376);
            this.btnFVSVariablesTieBreakerVariableNext.Name = "btnFVSVariablesTieBreakerVariableNext";
            this.btnFVSVariablesTieBreakerVariableNext.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesTieBreakerVariableNext.TabIndex = 8;
            this.btnFVSVariablesTieBreakerVariableNext.Text = "Next-->";
            this.btnFVSVariablesTieBreakerVariableNext.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerVariableNext_Click);
            // 
            // grpboxFVSVariablesTieBreaker
            // 
            this.grpboxFVSVariablesTieBreaker.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxFVSVariablesTieBreaker.Controls.Add(this.pnlTieBreaker);
            this.grpboxFVSVariablesTieBreaker.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxFVSVariablesTieBreaker.ForeColor = System.Drawing.Color.Black;
            this.grpboxFVSVariablesTieBreaker.Location = new System.Drawing.Point(16, 56);
            this.grpboxFVSVariablesTieBreaker.Name = "grpboxFVSVariablesTieBreaker";
            this.grpboxFVSVariablesTieBreaker.Size = new System.Drawing.Size(872, 448);
            this.grpboxFVSVariablesTieBreaker.TabIndex = 33;
            this.grpboxFVSVariablesTieBreaker.TabStop = false;
            this.grpboxFVSVariablesTieBreaker.Resize += new System.EventHandler(this.grpboxFVSVariablesTieBreaker_Resize);
            // 
            // pnlTieBreaker
            // 
            this.pnlTieBreaker.AutoScroll = true;
            this.pnlTieBreaker.Controls.Add(this.BtnTiebreakerClear);
            this.pnlTieBreaker.Controls.Add(this.label2);
            this.pnlTieBreaker.Controls.Add(this.grpboxFVSVariablesTieBreakerValues);
            this.pnlTieBreaker.Controls.Add(this.groupBox3);
            this.pnlTieBreaker.Controls.Add(this.btnFVSVariablesTieBreakerEdit);
            this.pnlTieBreaker.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlTieBreaker.Location = new System.Drawing.Point(3, 18);
            this.pnlTieBreaker.Name = "pnlTieBreaker";
            this.pnlTieBreaker.Size = new System.Drawing.Size(866, 427);
            this.pnlTieBreaker.TabIndex = 70;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 154);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(825, 35);
            this.label2.TabIndex = 70;
            this.label2.Text = resources.GetString("label2.Text");
            // 
            // grpboxFVSVariablesTieBreakerValues
            // 
            this.grpboxFVSVariablesTieBreakerValues.Controls.Add(this.lvFVSVariablesTieBreakerValues);
            this.grpboxFVSVariablesTieBreakerValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxFVSVariablesTieBreakerValues.Name = "grpboxFVSVariablesTieBreakerValues";
            this.grpboxFVSVariablesTieBreakerValues.Size = new System.Drawing.Size(840, 135);
            this.grpboxFVSVariablesTieBreakerValues.TabIndex = 67;
            this.grpboxFVSVariablesTieBreakerValues.TabStop = false;
            this.grpboxFVSVariablesTieBreakerValues.Text = "Step 1: Rules for choosing best silvicultural sequence when more than one are opt" +
    "imal according to the optimization rule";
            // 
            // lvFVSVariablesTieBreakerValues
            // 
            this.lvFVSVariablesTieBreakerValues.CheckBoxes = true;
            this.lvFVSVariablesTieBreakerValues.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lvColChecked,
            this.lvColMethod,
            this.lvColStandAttribute,
            this.lvColFieldSource,
            this.lvColMinMax});
            this.lvFVSVariablesTieBreakerValues.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvFVSVariablesTieBreakerValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lvFVSVariablesTieBreakerValues.GridLines = true;
            this.lvFVSVariablesTieBreakerValues.HideSelection = false;
            listViewItem1.StateImageIndex = 0;
            listViewItem2.StateImageIndex = 0;
            listViewItem3.Checked = true;
            listViewItem3.StateImageIndex = 1;
            this.lvFVSVariablesTieBreakerValues.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3});
            this.lvFVSVariablesTieBreakerValues.Location = new System.Drawing.Point(3, 18);
            this.lvFVSVariablesTieBreakerValues.MultiSelect = false;
            this.lvFVSVariablesTieBreakerValues.Name = "lvFVSVariablesTieBreakerValues";
            this.lvFVSVariablesTieBreakerValues.Size = new System.Drawing.Size(834, 114);
            this.lvFVSVariablesTieBreakerValues.TabIndex = 67;
            this.lvFVSVariablesTieBreakerValues.UseCompatibleStateImageBehavior = false;
            this.lvFVSVariablesTieBreakerValues.View = System.Windows.Forms.View.Details;
            this.lvFVSVariablesTieBreakerValues.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lvFVSVariablesTieBreakerValues_ItemCheck);
            this.lvFVSVariablesTieBreakerValues.SelectedIndexChanged += new System.EventHandler(this.lvFVSVariablesTieBreakerValues_SelectedIndexChanged);
            this.lvFVSVariablesTieBreakerValues.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvFVSVariablesTieBreakerValues_MouseUp);
            // 
            // lvColChecked
            // 
            this.lvColChecked.Text = "";
            this.lvColChecked.Width = 21;
            // 
            // lvColMethod
            // 
            this.lvColMethod.Text = "Method";
            this.lvColMethod.Width = 176;
            // 
            // lvColStandAttribute
            // 
            this.lvColStandAttribute.Text = "Stand Attribute";
            this.lvColStandAttribute.Width = 249;
            // 
            // lvColFieldSource
            // 
            this.lvColFieldSource.Text = "Value Source";
            this.lvColFieldSource.Width = 169;
            // 
            // lvColMinMax
            // 
            this.lvColMinMax.Text = "Max/Min";
            this.lvColMinMax.Width = 166;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnFVSVariablesTieBreakerAudit);
            this.groupBox3.Location = new System.Drawing.Point(24, 312);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(832, 72);
            this.groupBox3.TabIndex = 69;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Step 2: Audit";
            // 
            // btnFVSVariablesTieBreakerAudit
            // 
            this.btnFVSVariablesTieBreakerAudit.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesTieBreakerAudit.Location = new System.Drawing.Point(16, 24);
            this.btnFVSVariablesTieBreakerAudit.Name = "btnFVSVariablesTieBreakerAudit";
            this.btnFVSVariablesTieBreakerAudit.Size = new System.Drawing.Size(800, 32);
            this.btnFVSVariablesTieBreakerAudit.TabIndex = 0;
            this.btnFVSVariablesTieBreakerAudit.Text = "Audit Tie Breaker Specifications";
            this.btnFVSVariablesTieBreakerAudit.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerAudit_Click);
            // 
            // btnFVSVariablesTieBreakerEdit
            // 
            this.btnFVSVariablesTieBreakerEdit.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesTieBreakerEdit.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesTieBreakerEdit.ForeColor = System.Drawing.Color.Black;
            this.btnFVSVariablesTieBreakerEdit.Location = new System.Drawing.Point(299, 201);
            this.btnFVSVariablesTieBreakerEdit.Name = "btnFVSVariablesTieBreakerEdit";
            this.btnFVSVariablesTieBreakerEdit.Size = new System.Drawing.Size(128, 40);
            this.btnFVSVariablesTieBreakerEdit.TabIndex = 36;
            this.btnFVSVariablesTieBreakerEdit.Text = "Edit";
            this.btnFVSVariablesTieBreakerEdit.UseVisualStyleBackColor = false;
            this.btnFVSVariablesTieBreakerEdit.Click += new System.EventHandler(this.btnFVSVariablesTieBreakerEdit_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(894, 32);
            this.lblTitle.TabIndex = 27;
            this.lblTitle.Text = "Tie Breaker Settings";
            // 
            // BtnTiebreakerClear
            // 
            this.BtnTiebreakerClear.BackColor = System.Drawing.SystemColors.Control;
            this.BtnTiebreakerClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnTiebreakerClear.ForeColor = System.Drawing.Color.Black;
            this.BtnTiebreakerClear.Location = new System.Drawing.Point(451, 201);
            this.BtnTiebreakerClear.Name = "BtnTiebreakerClear";
            this.BtnTiebreakerClear.Size = new System.Drawing.Size(128, 40);
            this.BtnTiebreakerClear.TabIndex = 71;
            this.BtnTiebreakerClear.Text = "Clear";
            this.BtnTiebreakerClear.UseVisualStyleBackColor = false;
            this.BtnTiebreakerClear.Click += new System.EventHandler(this.BtnTiebreakerClear_Click);
            // 
            // uc_scenario_last_tiebreak_rank1
            // 
            this.uc_scenario_last_tiebreak_rank1.BackColor = System.Drawing.SystemColors.Control;
            this.uc_scenario_last_tiebreak_rank1.Location = new System.Drawing.Point(8, 25);
            this.uc_scenario_last_tiebreak_rank1.Name = "uc_scenario_treatment_intensity1";
            this.uc_scenario_last_tiebreak_rank1.ReferenceOptimizerScenarioForm = null;
            this.uc_scenario_last_tiebreak_rank1.Size = new System.Drawing.Size(840, 320);
            this.uc_scenario_last_tiebreak_rank1.TabIndex = 13;
            this.uc_scenario_last_tiebreak_rank1.Load += new System.EventHandler(this.uc_scenario_treatment_intensity1_Load);
            // 
            // uc_optimizer_scenario_fvs_prepost_variables_tiebreaker
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_scenario_fvs_prepost_variables_tiebreaker";
            this.Size = new System.Drawing.Size(900, 2000);
            this.groupBox1.ResumeLayout(false);
            this.grpboxFVSVariablesTieBreakerLastTieBreakRank.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.grpboxStandAttributeTieBreakerVariable.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.grpboxFVSVariablesTieBreakerVariableValueSource.ResumeLayout(false);
            this.grpMaxMin.ResumeLayout(false);
            this.grpboxFVSVariablesTieBreakerVariableValues.ResumeLayout(false);
            this.grpboxFVSVariablesTieBreakerVariableValues.PerformLayout();
            this.grpFVSVariablesTieBreakerVariableValuesSelected.ResumeLayout(false);
            this.grpboxFVSVariablesTieBreaker.ResumeLayout(false);
            this.pnlTieBreaker.ResumeLayout(false);
            this.grpboxFVSVariablesTieBreakerValues.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void grpboxFFEIndices_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxFFE_Resize(object sender, System.EventArgs e)
		{
			try
			{
				
			}
			catch
			{
			}
		}

		

		private void cmbFFE_TI1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbFFE_TI2_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkFFE_TI1_Click(object sender, System.EventArgs e)
		{
            
		}
		

		private void label6_Click(object sender, System.EventArgs e)
		{
		
		}

		

		private void txtFFE_TI1_Leave(object sender, System.EventArgs e)
		{
		   
		}

		private void txtFFE_TI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);

		}
		protected void AllowNumericOnly(System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				// Digits are OK
				//if (((frmScenario)this.ParentForm).btnSave.Enabled == false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmOptimizerScenario)this.ParentForm).m_bSave=true;
			}
			else if (e.KeyChar == '\b') 
			{
				//back space is okay
				//if (((frmScenario)this.ParentForm).btnSave.Enabled == false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmOptimizerScenario)this.ParentForm).m_bSave=true;
			}
			else
			{
				e.Handled = true;
			}
		}

		private void txtFFE_TI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_TI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_TI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_TI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}
        public void loadvalues_FromProperties()
        {
            this.m_intError = 0;
            this.m_strError = "";

            int x;




            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0) != null)
            {
                for (x = 0; x <= ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Count - 1; x++)
                {

                    if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper().IndexOf("ATTRIBUTE") > -1)
                    {
                        //idxAttribute = 0 for FVS Variables and 1 for Economic Variables; Most fields are similar
                        int idxAttribute = 0;
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper().Equals("ECONOMIC ATTRIBUTE"))
                        {
                            idxAttribute = 1;
                        }
                        
                        //attribute name
                        string strAttrib = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strFVSVariableName.Trim();
                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName =
                            ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strFVSVariableName.Trim();
                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_FVSVARIABLE].Text =
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName;

                        //fvs value source (POST or POST/PRE change)
                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource =
                            ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strValueSource.Trim();
                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_VALUESOURCE].Text =
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource;

                        //MAX or MIN	
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMaxYN == "Y")
                        {
                            lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MAX";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "Y";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "N";
                            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked = true;
                        }
                        else
                        {
                            lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MIN";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "Y";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "N";
                            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked = true;
                        }
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).bSelected)
                        {
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = true;

                        }
                        else
                        {
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = false;
                        }
                        this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].Checked = 
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                        // Select stand attribute on the listView by default if is enabled
                        if (this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected == true)
                        {
                            lvFVSVariablesTieBreakerValues.Items[idxAttribute].Selected = this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                            lvFVSVariablesTieBreakerValues.Select();
                        }
                    }
                    else if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper() == "LAST TIE-BREAK RANK")
                    {
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).bSelected)
                        {
                            this.m_oOldTieBreakerCollection.Item(2).bSelected = true;

                        }
                        else
                        {
                            this.m_oOldTieBreakerCollection.Item(2).bSelected = false;
                        }
                        this.lvFVSVariablesTieBreakerValues.Items[2].Checked = this.m_oOldTieBreakerCollection.Item(2).bSelected;
                    }

                }
                this.m_oSavTieBreakerCollection.Copy(this.m_oOldTieBreakerCollection, ref this.m_oSavTieBreakerCollection, true);
            }
            this.uc_scenario_last_tiebreak_rank1.loadgrid(true);

        }
        public void loadvalues_FromPropertiesSqlite()
        {
            this.m_intError = 0;
            this.m_strError = "";

            int x;

            if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0) != null)
            {
                for (x = 0; x <= ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Count - 1; x++)
                {

                    if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper().IndexOf("ATTRIBUTE") > -1)
                    {
                        //idxAttribute = 0 for FVS Variables and 1 for Economic Variables; Most fields are similar
                        int idxAttribute = 0;
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper().Equals("ECONOMIC ATTRIBUTE"))
                        {
                            idxAttribute = 1;
                        }

                        //attribute name
                        string strAttrib = ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strFVSVariableName.Trim();
                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName =
                            ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strFVSVariableName.Trim();
                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_FVSVARIABLE].Text =
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName;

                        //fvs value source (POST or POST/PRE change)
                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource =
                            ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strValueSource.Trim();
                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_VALUESOURCE].Text =
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource;

                        //MAX or MIN	
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMaxYN == "Y")
                        {
                            lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MAX";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "Y";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "N";
                            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked = true;
                        }
                        else
                        {
                            lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MIN";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "Y";
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "N";
                            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked = true;
                        }
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).bSelected)
                        {
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = true;

                        }
                        else
                        {
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = false;
                        }
                        this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].Checked =
                            this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                        // Select stand attribute on the listView by default if is enabled
                        if (this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected == true)
                        {
                            lvFVSVariablesTieBreakerValues.Items[idxAttribute].Selected = this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                            lvFVSVariablesTieBreakerValues.Select();
                        }
                    }
                    else if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).strMethod.Trim().ToUpper() == "LAST TIE-BREAK RANK")
                    {
                        if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oTieBreaker_Collection.Item(x).bSelected)
                        {
                            this.m_oOldTieBreakerCollection.Item(2).bSelected = true;

                        }
                        else
                        {
                            this.m_oOldTieBreakerCollection.Item(2).bSelected = false;
                        }
                        this.lvFVSVariablesTieBreakerValues.Items[2].Checked = this.m_oOldTieBreakerCollection.Item(2).bSelected;
                    }

                }
                this.m_oSavTieBreakerCollection.Copy(this.m_oOldTieBreakerCollection, ref this.m_oSavTieBreakerCollection, true);
            }
            this.uc_scenario_last_tiebreak_rank1.loadgridsqlite(true);

        }
        public void loadvalues(System.Collections.Generic.IDictionary<string, System.Collections.Generic.IList<String>> p_dictFVSTables)
		{

			
			this.m_intError=0;
			this.m_strError="";

            lstFVSTablesList.Items.Clear();
            _dictFVSTables = new System.Collections.Generic.Dictionary<string,
                System.Collections.Generic.IList<string>>(p_dictFVSTables);           
            foreach (string strKey in _dictFVSTables.Keys)
            {
                lstFVSTablesList.Items.Add(strKey);
            }

            lstEconVariablesList.Items.Clear();
            foreach (uc_optimizer_scenario_calculated_variables.VariableItem vItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
            {
                if (vItem.strVariableType.Equals("ECON"))
                {
                    lstEconVariablesList.Items.Add(vItem.strVariableName);
                }
            }

			//Stash original location of some controls so we can move them later
            _objGrpMaxMinLocation = this.grpMaxMin.Location;
            _objLblTieBreakVarDescrLocation = this.lblTieBreakVarDescr.Location;
            _objtxtTieBreakVarDescrLocation = this.txtTieBreakVarDescr.Location;
            
            this.m_oOldTieBreakerCollection.Clear();

			// Stand Attribute
            uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem = new TieBreakerItem();
			oItem.intListViewIndex=0;
			oItem.strFVSVariableName=this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
			oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim();
			oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim();   //POST or PRE - POST
			if (lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper()=="MAX")
			{
				oItem.strMaxYN="Y"; oItem.strMinYN="N";
			}
			else if (lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper()=="MIN")
			{
				oItem.strMaxYN="N"; oItem.strMinYN="Y";
			}
		    oItem.bSelected=this.lvFVSVariablesTieBreakerValues.Items[0].Checked;
			this.m_oOldTieBreakerCollection.Add(oItem);

            // Economic Attribute
            oItem = new TieBreakerItem();
            oItem.intListViewIndex = 1;
            oItem.strFVSVariableName = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
            oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim();
            oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim();
            if (lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MAX")
            {
                oItem.strMaxYN = "Y"; oItem.strMinYN = "N";
            }
            else if (lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MIN")
            {
                oItem.strMaxYN = "N"; oItem.strMinYN = "Y";
            }
            oItem.bSelected = this.lvFVSVariablesTieBreakerValues.Items[1].Checked;
            this.m_oOldTieBreakerCollection.Add(oItem);

			// Last Tie-Break Rank
            oItem = new TieBreakerItem();
			oItem.intListViewIndex=2;
			oItem.strFVSVariableName=this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
			oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_METHOD].Text.Trim();
			oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_VALUESOURCE].Text.Trim();
			oItem.strMaxYN="N"; oItem.strMinYN="N";
			oItem.bSelected=this.lvFVSVariablesTieBreakerValues.Items[2].Checked;
			this.m_oOldTieBreakerCollection.Add(oItem);
            // Make row appear read-only
            // @ToDo: Save off in case we want to reformat
            //for (int i = COLUMN_CHECKBOX; i <= COLUMN_MAXMIN; i++)
            //{
            //    this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[i].ForeColor = SystemColors.GrayText;
            //}
			
			ado_data_access oAdo = new ado_data_access();

			string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
			if (oAdo.m_intError==0)
			{

				if (!oAdo.TableExist(oAdo.m_OleDbConnection,Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName))
					 frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateScenarioFVSVariablesTieBreakerTable(oAdo,oAdo.m_OleDbConnection,
						 Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);

				oAdo.m_strSQL = "SELECT * FROM " + 
					            Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
					            "WHERE TRIM(scenario_id)='" + strScenarioId.Trim() + "'";


				oAdo.SqlQueryReader(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				if (oAdo.m_OleDbDataReader.HasRows)
				{
					while (oAdo.m_OleDbDataReader.Read())
					{
 

						if (oAdo.m_OleDbDataReader["tiebreaker_method"] != System.DBNull.Value)
						{
                            if (oAdo.m_OleDbDataReader["tiebreaker_method"].ToString().Trim().ToUpper().IndexOf("ATTRIBUTE") > -1)
							{
								//idxAttribute = 0 for FVS Variables and 1 for Economic Variables; Most fields are similar
                                int idxAttribute = 0;
                                if (oAdo.m_OleDbDataReader["tiebreaker_method"].ToString().Trim().ToUpper().Equals("ECONOMIC ATTRIBUTE"))
                                {
                                    idxAttribute = 1;
                                }
                                
                                //fvs variable name
								this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName=oAdo.m_OleDbDataReader["fvs_variable_name"].ToString().Trim();
                                lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_FVSVARIABLE].Text =
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName;

								//fvs value source (POST or POST/PRE change)
                                this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource = oAdo.m_OleDbDataReader["value_source"].ToString().Trim();
                                lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_VALUESOURCE].Text =
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource;

								//MAX or MIN	
								if (oAdo.m_OleDbDataReader["max_yn"].ToString().Trim().ToUpper()=="Y")
								{
                                    lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MAX";
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "Y";
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "N";
									this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked=true;
								}
                                else if (oAdo.m_OleDbDataReader["min_yn"].ToString().Trim().ToUpper() == "Y")
								{
                                    lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MIN";
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "Y";
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "N";
									this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked=true;
								}
								if (oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper()=="Y")
								{
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = true;
									
								}
								else
								{
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = false;
								}
                                this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].Checked = this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                                // Select stand attribute on the listView by default if is enabled
                                if (this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected == true)
                                {
                                    lvFVSVariablesTieBreakerValues.Items[idxAttribute].Selected = this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                                    lvFVSVariablesTieBreakerValues.Select();
                                }
							}
                            else if (oAdo.m_OleDbDataReader["tiebreaker_method"].ToString().Trim().ToUpper()=="LAST TIE-BREAK RANK")
							{
								if (oAdo.m_OleDbDataReader["checked_yn"].ToString().Trim().ToUpper()=="Y")
								{
									this.m_oOldTieBreakerCollection.Item(2).bSelected=true;
									
								}
								else
								{
									this.m_oOldTieBreakerCollection.Item(2).bSelected=false;
								}
								this.lvFVSVariablesTieBreakerValues.Items[2].Checked = this.m_oOldTieBreakerCollection.Item(2).bSelected;
							}
						}

					}
				}
				
				oAdo.m_OleDbDataReader.Close();
				oAdo.CloseConnection(oAdo.m_OleDbConnection);
				oAdo.m_OleDbConnection.Dispose();
				this.m_oSavTieBreakerCollection.Copy(this.m_oOldTieBreakerCollection,ref this.m_oSavTieBreakerCollection,true);
					            
			}
			this.m_intError=oAdo.m_intError;
			this.m_strError=oAdo.m_strError;
			oAdo=null;

            this.uc_scenario_last_tiebreak_rank1.loadgrid(false);
			
		}
        public void loadvaluessqlite(System.Collections.Generic.IDictionary<string, System.Collections.Generic.IList<String>> p_dictFVSTables)
        {
            this.m_intError = 0;
            this.m_strError = "";

            lstFVSTablesList.Items.Clear();
            _dictFVSTables = new System.Collections.Generic.Dictionary<string,
                System.Collections.Generic.IList<string>>(p_dictFVSTables);
            foreach (string strKey in _dictFVSTables.Keys)
            {
                lstFVSTablesList.Items.Add(strKey);
            }

            lstEconVariablesList.Items.Clear();
            foreach (uc_optimizer_scenario_calculated_variables.VariableItem vItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
            {
                if (vItem.strVariableType.Equals("ECON"))
                {
                    lstEconVariablesList.Items.Add(vItem.strVariableName);
                }
            }

            //Stash original location of some controls so we can move them later
            _objGrpMaxMinLocation = this.grpMaxMin.Location;
            _objLblTieBreakVarDescrLocation = this.lblTieBreakVarDescr.Location;
            _objtxtTieBreakVarDescrLocation = this.txtTieBreakVarDescr.Location;

            this.m_oOldTieBreakerCollection.Clear();

            // Stand Attribute
            uc_optimizer_scenario_fvs_prepost_variables_tiebreaker.TieBreakerItem oItem = new TieBreakerItem();
            oItem.intListViewIndex = 0;
            oItem.strFVSVariableName = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
            oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim();
            oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim();   //POST or PRE - POST
            if (lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MAX")
            {
                oItem.strMaxYN = "Y"; oItem.strMinYN = "N";
            }
            else if (lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MIN")
            {
                oItem.strMaxYN = "N"; oItem.strMinYN = "Y";
            }
            oItem.bSelected = this.lvFVSVariablesTieBreakerValues.Items[0].Checked;
            this.m_oOldTieBreakerCollection.Add(oItem);

            // Economic Attribute
            oItem = new TieBreakerItem();
            oItem.intListViewIndex = 1;
            oItem.strFVSVariableName = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
            oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim();
            oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim();
            if (lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MAX")
            {
                oItem.strMaxYN = "Y"; oItem.strMinYN = "N";
            }
            else if (lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MIN")
            {
                oItem.strMaxYN = "N"; oItem.strMinYN = "Y";
            }
            oItem.bSelected = this.lvFVSVariablesTieBreakerValues.Items[1].Checked;
            this.m_oOldTieBreakerCollection.Add(oItem);

            // Last Tie-Break Rank
            oItem = new TieBreakerItem();
            oItem.intListViewIndex = 2;
            oItem.strFVSVariableName = this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
            oItem.strMethod = this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_METHOD].Text.Trim();
            oItem.strValueSource = this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_VALUESOURCE].Text.Trim();
            oItem.strMaxYN = "N"; oItem.strMinYN = "N";
            oItem.bSelected = this.lvFVSVariablesTieBreakerValues.Items[2].Checked;
            this.m_oOldTieBreakerCollection.Add(oItem);

            DataMgr oDataMgr = new DataMgr();

            string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;

            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
            {
                conn.Open();
                if (oDataMgr.m_intError == 0)
                {
                    if (!oDataMgr.TableExist(conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName))
                    {
                        frmMain.g_oTables.m_oOptimizerScenarioRuleDef.CreateSqliteScenarioFVSVariablesTieBreakerTable(oDataMgr, conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName);
                    }
                    oDataMgr.m_strSQL = "SELECT * FROM " +
                                    Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                    "WHERE TRIM(UPPER(scenario_id))='" + strScenarioId.Trim().ToUpper() + "'";

                    oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_DataReader.HasRows)
                    {
                        while (oDataMgr.m_DataReader.Read())
                        {
                            if (oDataMgr.m_DataReader["tiebreaker_method"] != System.DBNull.Value)
                            {
                                if (oDataMgr.m_DataReader["tiebreaker_method"].ToString().Trim().ToUpper().IndexOf("ATTRIBUTE") > -1)
                                {
                                    //idxAttribute = 0 for FVS Variables and 1 for Economic Variables; Most fields are similar
                                    int idxAttribute = 0;
                                    if (oDataMgr.m_DataReader["tiebreaker_method"].ToString().Trim().ToUpper().Equals("ECONOMIC ATTRIBUTE"))
                                    {
                                        idxAttribute = 1;
                                    }

                                    //fvs variable name
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName = oDataMgr.m_DataReader["fvs_variable_name"].ToString().Trim();
                                    lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_FVSVARIABLE].Text =
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strFVSVariableName;

                                    //fvs value source (POST or POST/PRE change)
                                    this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource = oDataMgr.m_DataReader["value_source"].ToString().Trim();
                                    lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_VALUESOURCE].Text =
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strValueSource;

                                    //MAX or MIN	
                                    if (oDataMgr.m_DataReader["max_yn"].ToString().Trim().ToUpper() == "Y")
                                    {
                                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MAX";
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "Y";
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "N";
                                        this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked = true;
                                    }
                                    else if (oDataMgr.m_DataReader["min_yn"].ToString().Trim().ToUpper() == "Y")
                                    {
                                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text = "MIN";
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strMinYN = "Y";
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).strMaxYN = "N";
                                        this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked = true;
                                    }
                                    if (oDataMgr.m_DataReader["checked_yn"].ToString().Trim().ToUpper() == "Y")
                                    {
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = true;

                                    }
                                    else
                                    {
                                        this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected = false;
                                    }
                                    this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].Checked = this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                                    // Select stand attribute on the listView by default if is enabled
                                    if (this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected == true)
                                    {
                                        lvFVSVariablesTieBreakerValues.Items[idxAttribute].Selected = this.m_oOldTieBreakerCollection.Item(idxAttribute).bSelected;
                                        lvFVSVariablesTieBreakerValues.Select();
                                    }
                                }
                                else if (oDataMgr.m_DataReader["tiebreaker_method"].ToString().Trim().ToUpper() == "LAST TIE-BREAK RANK")
                                {
                                    if (oDataMgr.m_DataReader["checked_yn"].ToString().Trim().ToUpper() == "Y")
                                    {
                                        this.m_oOldTieBreakerCollection.Item(2).bSelected = true;

                                    }
                                    else
                                    {
                                        this.m_oOldTieBreakerCollection.Item(2).bSelected = false;
                                    }
                                    this.lvFVSVariablesTieBreakerValues.Items[2].Checked = this.m_oOldTieBreakerCollection.Item(2).bSelected;
                                }
                            }
                        }
                    }
                    oDataMgr.m_DataReader.Close();
                    this.m_oSavTieBreakerCollection.Copy(this.m_oOldTieBreakerCollection, ref this.m_oSavTieBreakerCollection, true);
                }
                conn.Close();
            }
            this.m_intError = oDataMgr.m_intError;
            this.m_strError = oDataMgr.m_strError;
            this.uc_scenario_last_tiebreak_rank1.loadgridsqlite(false);
        }

        public int savevalues()
		{
			int x=0;
			string strColumns="";
			string strValues="";
			ado_data_access oAdo = new ado_data_access();
			string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = 
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
			oAdo.OpenConnection(oAdo.getMDBConnString(strScenarioMDB,"",""));
			if (oAdo.m_intError==0)
			{
				//delete all records from the scenario fvs variables table
				oAdo.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " + 
					            "WHERE LCASE(TRIM(scenario_id)) = '" + strScenarioId + "';";

				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);
				if (oAdo.m_intError < 0)
				{
					oAdo.m_OleDbConnection.Close();
					x=oAdo.m_intError;
					oAdo = null;
					return x;
				}
				strColumns="scenario_id,rxcycle,tiebreaker_method,fvs_variable_name,value_source,max_yn,min_yn,checked_yn";
				//
				//FVS VARIABLE
				//
				//scenario id
				strValues = "'" + strScenarioId + "','1',";
				//method
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim()  + "',";
				this.m_oSavTieBreakerCollection.Item(0).strMethod = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim();
				//fvs variable name
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
				this.m_oSavTieBreakerCollection.Item(0).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
				//value source
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
				m_oSavTieBreakerCollection.Item(0).strValueSource = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim();
				//aggregate MAX or MIN
				if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper()=="MAX")
				{
					strValues = strValues + "'Y','N',";
					this.m_oSavTieBreakerCollection.Item(0).strMaxYN="Y";
					this.m_oSavTieBreakerCollection.Item(0).strMinYN="N";
				}
				else
				{
					strValues = strValues + "'N','Y',";
					this.m_oSavTieBreakerCollection.Item(0).strMaxYN="N";
					this.m_oSavTieBreakerCollection.Item(0).strMinYN="Y";
				}
				//checked
				if (this.lvFVSVariablesTieBreakerValues.Items[0].Checked)
				{
					strValues = strValues + "'Y'";
					this.m_oSavTieBreakerCollection.Item(0).bSelected=true;
				}
				else
				{
					strValues = strValues + "'N'";
					this.m_oSavTieBreakerCollection.Item(0).bSelected=false;
				}

				oAdo.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " + 
					            "(" + strColumns + ") VALUES " + 
					            "(" + strValues + ")";

				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

                //
                //ECONOMIC VARIABLE
                //
                //scenario id
                strValues = "'" + strScenarioId + "','1',";
                //method
                strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim() + "',";
                this.m_oSavTieBreakerCollection.Item(1).strMethod = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim();
                //fvs variable name
                strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
                this.m_oSavTieBreakerCollection.Item(1).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
                //value source
                strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
                m_oSavTieBreakerCollection.Item(1).strValueSource = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim();
                //aggregate MAX or MIN
                if (this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MAX")
                {
                    strValues = strValues + "'Y','N',";
                    this.m_oSavTieBreakerCollection.Item(1).strMaxYN = "Y";
                    this.m_oSavTieBreakerCollection.Item(1).strMinYN = "N";
                }
                else
                {
                    strValues = strValues + "'N','Y',";
                    this.m_oSavTieBreakerCollection.Item(1).strMaxYN = "N";
                    this.m_oSavTieBreakerCollection.Item(1).strMinYN = "Y";
                }
                //checked
                if (this.lvFVSVariablesTieBreakerValues.Items[1].Checked)
                {
                    strValues = strValues + "'Y'";
                    this.m_oSavTieBreakerCollection.Item(1).bSelected = true;
                }
                else
                {
                    strValues = strValues + "'N'";
                    this.m_oSavTieBreakerCollection.Item(1).bSelected = false;
                }

                oAdo.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                "(" + strColumns + ") VALUES " +
                                "(" + strValues + ")";

                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);

				//
                //LAST TIE-BREAK RANK
				//
				//scenario id
				strValues = "'" + strScenarioId + "','1',";
				//method
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_METHOD].Text.Trim()  + "',";
				this.m_oSavTieBreakerCollection.Item(2).strMethod = lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_METHOD].Text.Trim();
				//fvs variable name
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
				this.m_oSavTieBreakerCollection.Item(2).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
				//value source
				strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
				m_oSavTieBreakerCollection.Item(2).strValueSource = lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_VALUESOURCE].Text.Trim();
				//aggregate MAX or MIN
				strValues = strValues + "'N','Y',";
                this.m_oSavTieBreakerCollection.Item(1).strMaxYN = "N";
				this.m_oSavTieBreakerCollection.Item(1).strMinYN = "Y";
				//checked
				if (this.lvFVSVariablesTieBreakerValues.Items[2].Checked)
				{
					strValues = strValues + "'Y'";
					this.m_oSavTieBreakerCollection.Item(2).bSelected=true;
				}
				else
				{
					strValues = strValues + "'N'";
					this.m_oSavTieBreakerCollection.Item(2).bSelected=false;
				}

				oAdo.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " + 
					"(" + strColumns + ") VALUES " + 
					"(" + strValues + ")";

				oAdo.SqlNonQuery(oAdo.m_OleDbConnection,oAdo.m_strSQL);

				oAdo.CloseConnection(oAdo.m_OleDbConnection);
			}
            this.uc_scenario_last_tiebreak_rank1.savevalues();
			return 0;
		}


        public int savevaluessqlite()
        {
            int x = 0;
            string strColumns = "";
            string strValues = "";
            DataMgr oDataMgr = new DataMgr();
            string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            string strScenarioDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
            {
                conn.Open();
                if (oDataMgr.m_intError == 0)
                {
                    //delete all records from the scenario fvs variables table
                    oDataMgr.m_strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                    "WHERE LOWER(TRIM(scenario_id)) = '" + strScenarioId + "';";

                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    if (oDataMgr.m_intError < 0)
                    {
                        conn.Close();
                        x = oDataMgr.m_intError;
                        oDataMgr = null;
                        return x;
                    }
                    strColumns = "scenario_id,rxcycle,tiebreaker_method,fvs_variable_name,value_source,max_yn,min_yn,checked_yn";
                    //
                    //FVS VARIABLE
                    //
                    //scenario id
                    strValues = "'" + strScenarioId + "','1',";
                    //method
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim() + "',";
                    this.m_oSavTieBreakerCollection.Item(0).strMethod = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_METHOD].Text.Trim();
                    //fvs variable name
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
                    this.m_oSavTieBreakerCollection.Item(0).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
                    //value source
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
                    m_oSavTieBreakerCollection.Item(0).strValueSource = lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text.Trim();
                    //aggregate MAX or MIN
                    if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MAX")
                    {
                        strValues = strValues + "'Y','N',";
                        this.m_oSavTieBreakerCollection.Item(0).strMaxYN = "Y";
                        this.m_oSavTieBreakerCollection.Item(0).strMinYN = "N";
                    }
                    else
                    {
                        strValues = strValues + "'N','Y',";
                        this.m_oSavTieBreakerCollection.Item(0).strMaxYN = "N";
                        this.m_oSavTieBreakerCollection.Item(0).strMinYN = "Y";
                    }
                    //checked
                    if (this.lvFVSVariablesTieBreakerValues.Items[0].Checked)
                    {
                        strValues = strValues + "'Y'";
                        this.m_oSavTieBreakerCollection.Item(0).bSelected = true;
                    }
                    else
                    {
                        strValues = strValues + "'N'";
                        this.m_oSavTieBreakerCollection.Item(0).bSelected = false;
                    }

                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                    "(" + strColumns + ") VALUES " +
                                    "(" + strValues + ")";

                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                    //
                    //ECONOMIC VARIABLE
                    //
                    //scenario id
                    strValues = "'" + strScenarioId + "','1',";
                    //method
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim() + "',";
                    this.m_oSavTieBreakerCollection.Item(1).strMethod = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_METHOD].Text.Trim();
                    //fvs variable name
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
                    this.m_oSavTieBreakerCollection.Item(1).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
                    //value source
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
                    m_oSavTieBreakerCollection.Item(1).strValueSource = lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_VALUESOURCE].Text.Trim();
                    //aggregate MAX or MIN
                    if (this.lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_MAXMIN].Text.Trim().ToUpper() == "MAX")
                    {
                        strValues = strValues + "'Y','N',";
                        this.m_oSavTieBreakerCollection.Item(1).strMaxYN = "Y";
                        this.m_oSavTieBreakerCollection.Item(1).strMinYN = "N";
                    }
                    else
                    {
                        strValues = strValues + "'N','Y',";
                        this.m_oSavTieBreakerCollection.Item(1).strMaxYN = "N";
                        this.m_oSavTieBreakerCollection.Item(1).strMinYN = "Y";
                    }
                    //checked
                    if (this.lvFVSVariablesTieBreakerValues.Items[1].Checked)
                    {
                        strValues = strValues + "'Y'";
                        this.m_oSavTieBreakerCollection.Item(1).bSelected = true;
                    }
                    else
                    {
                        strValues = strValues + "'N'";
                        this.m_oSavTieBreakerCollection.Item(1).bSelected = false;
                    }

                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                                    "(" + strColumns + ") VALUES " +
                                    "(" + strValues + ")";

                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                    //
                    //LAST TIE-BREAK RANK
                    //
                    //scenario id
                    strValues = "'" + strScenarioId + "','1',";
                    //method
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_METHOD].Text.Trim() + "',";
                    this.m_oSavTieBreakerCollection.Item(2).strMethod = lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_METHOD].Text.Trim();
                    //fvs variable name
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_FVSVARIABLE].Text.Trim() + "',";
                    this.m_oSavTieBreakerCollection.Item(2).strFVSVariableName = lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
                    //value source
                    strValues = strValues + "'" + this.lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_VALUESOURCE].Text.Trim() + "',";
                    m_oSavTieBreakerCollection.Item(2).strValueSource = lvFVSVariablesTieBreakerValues.Items[2].SubItems[COLUMN_VALUESOURCE].Text.Trim();
                    //aggregate MAX or MIN
                    strValues = strValues + "'N','Y',";
                    this.m_oSavTieBreakerCollection.Item(1).strMaxYN = "N";
                    this.m_oSavTieBreakerCollection.Item(1).strMinYN = "Y";
                    //checked
                    if (this.lvFVSVariablesTieBreakerValues.Items[2].Checked)
                    {
                        strValues = strValues + "'Y'";
                        this.m_oSavTieBreakerCollection.Item(2).bSelected = true;
                    }
                    else
                    {
                        strValues = strValues + "'N'";
                        this.m_oSavTieBreakerCollection.Item(2).bSelected = false;
                    }

                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " " +
                        "(" + strColumns + ") VALUES " +
                        "(" + strValues + ")";

                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                conn.Close();
            }
            this.uc_scenario_last_tiebreak_rank1.savevaluessqlite();
            return 0;
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
			//p_oButton.Left = this.btnNext.Left - p_oButton.Width;  //p_oGb.Width - this.btnNext.Width - 5 - p_oButton.Width;
			p_oButton.Top = this.btnNext.Top;
			p_oButton.Height = this.btnNext.Height;
			p_oButton.Width = this.btnNext.Width;
			p_oButton.Left = this.btnNext.Left - p_oButton.Width;
			p_oButton.Name = strButtonName;	
		}

		

		

		
		
		private void cmbFFECIBackslide2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIHazardOperator_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIHazardWindSpeedClass_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIHazardOperator_SelectedValueChanged(object sender, System.EventArgs e)
		{
			//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
			//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreaker.Height = this.ClientSize.Height - this.grpboxFVSVariablesTieBreaker.Top - 5;
			this.grpboxFVSVariablesTieBreaker.Width = this.ClientSize.Width - (this.grpboxFVSVariablesTieBreaker.Left * 2) ;
			this.grpboxStandAttributeTieBreakerVariable.Height = this.grpboxFVSVariablesTieBreaker.Height;
			this.grpboxStandAttributeTieBreakerVariable.Width =  this.grpboxStandAttributeTieBreakerVariable.Width;
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Height = this.grpboxFVSVariablesTieBreaker.Height;
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Width = this.grpboxFVSVariablesTieBreaker.Width;
			
			//grpboxFVSVariablesPrePost.Height = this.ClientSize.Height - grpboxFVSVariablesPrePost.Top - 5;
			//grpboxFVSVariablesPrePost.Width = this.ClientSize.Width - (grpboxFVSVariablesPrePost.Left * 2) ;
		    //grpboxFVSVariablesPrePostVariable.Height = grpboxFVSVariablesPrePost.Height;
			//grpboxFVSVariablesPrePostVariable.Width =  grpboxFVSVariablesPrePost.Width;
			//grpboxFVSVariablesPrePostExpression.Height = grpboxFVSVariablesPrePost.Height;
			//grpboxFVSVariablesPrePostExpression.Width = grpboxFVSVariablesPrePost.Width;
			
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

		
		private void InitializeUserNavigation()
		{
			for (int x=0;x<=this.m_strUserNavigation.Length - 1;x++)
			{
				this.m_strUserNavigation[x]="";
			}

		}
		private void AddUserNavigation(string p_strMenuItem)
		{

		}

		

		
		private void SaveToCurrentVariables()
		{
			
			
		}
		private void UpdateListViewVariableItem(int p_intListViewItem,int p_intVarArrayItem,uc_optimizer_scenario_fvs_prepost_variables_effective.Variables p_oVar)
		{
			
		}
		
			

		private void grpboxFVSVariablesPrePostOverallExp_Resize(object sender, System.EventArgs e)
		{
			
		}

		
		private void EnableAllMacroButtons(bool p_bEnable)
		{
			

		}

		private void btnFVSVariablesPrePostExpressionMacroVarPre1_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPre2_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPre3_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPost1_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPost2_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVarPost3_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar1Better_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar2Better_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar3Better_Click(object sender, System.EventArgs e)
		{
		  	
		}

		private void btnFVSVariablesPrePostExpressionMacroVar1Worse_Click(object sender, System.EventArgs e)
		{
  		  
		}

		private void btnFVSVariablesPrePostExpressionMacroVar2Worse_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionMacroVar3Worse_Click(object sender, System.EventArgs e)
		{
			
		}

		

		
		private void btnFVSVariablesPrePostExpressionMacroVar1Change_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionMacroVar2Change_Click(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionMacroVar3Change_Click(object sender, System.EventArgs e)
		{
			
		}

		
	


		private void ShowGroupBox(string p_strName)
		{
			int x;
			
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
		private bool Modified()
		{
			
			return true;
            

		}

		private void btnFVSVariablesPrePostValuesButtonsEdit_Click(object sender, System.EventArgs e)
		{
			EditVariable();

		}
		private void EditVariable()
		{

		}
		

		
		private void btnFVSVariablesPrePostValuesButtonsClear_Click(object sender, System.EventArgs e)
		{
			
		}

		private void RemoveVariable(int p_intIndex)
		{


		}

		private void btnFVSVariablesPrePostValuesButtonsClearAll_Click(object sender, System.EventArgs e)
		{

			DialogResult result = MessageBox.Show("Are you sure you wish to delete all the variables? (YN)","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
			if (result == System.Windows.Forms.DialogResult.Yes) 
			{
				for (int x=0;x<=NUMBER_OF_VARIABLES-1;x++)
				{
					RemoveVariable(x);
				}
			}
		}

		
		private void grpboxFVSVariablesPrePost_Resize(object sender, System.EventArgs e)
		{
		}

		private void btnFVSVariablesPrePostAudit_Click(object sender, System.EventArgs e)
		{
			DisplayAuditMessage=true;
			Audit();

			
		}
		public int Audit(bool bDisplayAuditMsg)
		{
			this.DisplayAuditMessage=bDisplayAuditMsg;
			Audit();
			return m_intError;
		}
		public void Audit()
		{
			int x = 0;
			this.m_intError=0;
			this.m_strError="Audit Results \r\n";
			this.m_strError=m_strError + "-------------\r\n\r\n";

            // Validate STAND ATTRIBUTE if it is checked
            if (this.lvFVSVariablesTieBreakerValues.Items[0].Checked)
            {
                string[] strPieces = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim().ToUpper().Split('.');
                if (strPieces.Length == 2)
                {
                    if (strPieces[0].ToUpper().Contains("_WEIGHTED"))
                    {
                        FIA_Biosum_Manager.OptimizerScenarioTools oOptimizerScenarioTools = new OptimizerScenarioTools();
                        string strWeightedError = oOptimizerScenarioTools.AuditWeightedFvsVariables(strPieces[0], out m_intError);
                        if (m_intError != 0)
                        {
                            m_strError = m_strError + strWeightedError;
                        }
                    }
                }
            }

            // Validate LAST TIE-BREAK RANK if it is checked
            if (this.lvFVSVariablesTieBreakerValues.Items[2].Checked)
            {
                x = this.uc_scenario_last_tiebreak_rank1.Val_Last_Tiebreak_Rank(false);
            }

            if (x<0)
			{
                if (x == -1) m_strError = m_strError + "No treatments defined\r\n";
                else if (x == -2) m_strError = m_strError + "Last Tie-Break Rankings must be unique\r\n";
                else if (x == -3) m_strError = m_strError + "Last Tie-Break Rankings cannot be null in value\r\n";
				m_intError=x;
			}
			if (this.lvFVSVariablesTieBreakerValues.CheckedItems.Count==0)
			{
				m_intError=-1;
				m_strError=m_strError + "No tie breaker methods are checked\r\n";
			}
			else
			{
				this.AuditCheckOptimiztionAndTieBreakerVariable(ref m_intError,ref m_strError);
				
			}
			if (m_intError==0) this.m_strError=m_strError + "Passed Audit";
			else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";

			if (this.DisplayAuditMessage)
				MessageBox.Show(m_strError,"FIA Biosum");
		}
		public void AuditCheckOptimiztionAndTieBreakerVariable(ref int p_intError,ref string p_strError)
		{
			int x;
			if (this.lvFVSVariablesTieBreakerValues.Items[0].Checked)
			{
				if (this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim().ToUpper()=="NOT DEFINED")
				{
					p_intError=-1;
					p_strError=p_strError + "Tie Breaker Stand Attribute not selected\r\n";
				}
				else
				{
					//cannot be equal to the stand attribute optimization variable
                    for (x = 0; x <= this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Count - 1; x++)
					{
                        if (this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).bSelected)
						{
                            if (this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(x).strFVSVariableName.Trim().ToUpper() == 
								lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim().ToUpper())
							{
								p_intError=-1;
								p_strError=p_strError + lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_FVSVARIABLE].Text + " cannot be both optimization variable and tie breaker variable\r\n";
								break;
							}
						}
					}

				}
			}
            //cannot be equal to the economic optimization variable
            if (this.lvFVSVariablesTieBreakerValues.Items[1].Checked)
            {
                //cannot be equal to the economic optimization variable
                if (this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(3).bSelected == true)
                {
                    if (this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection.Item(3).strFVSVariableName.Trim().ToUpper() ==
                        lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text.Trim().ToUpper())
                    {
                        p_intError = -1;
                        p_strError = p_strError + lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text + " cannot be both optimization variable and tie breaker variable\r\n";
                    }
                }
            }
		}
		
		private void lvFVSVariablesPrePostValues_DoubleClick(object sender, System.EventArgs e)
		{
			EditVariable();
		}

		private void pnlFVSVariablesPrePostExpression_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		private void pnlFVSVariablesPrePostVariable_Resize(object sender, System.EventArgs e)
		{
			
			
		}

		

		private void grpboxFVSVariablesPrePost_VisibleChanged(object sender, System.EventArgs e)
		{

		}


		private void EnableTabs(bool p_bEnable)
		{
            int x;
			ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlScenario,"tbdesc,tbnotes,tbdatasources",p_bEnable);
			ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlRules,"tbpsites,tbowners,tbcost,tbtreatmentintensity,tbfilterplots,tbrun",p_bEnable);
			ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlFVSVariables,"tboptimization,tbeffective",p_bEnable);
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

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnFVSVariablesTieBreakerEdit_Click(object sender, System.EventArgs e)
		{
			if (this.lvFVSVariablesTieBreakerValues.SelectedItems.Count==0) return;

            //idxAttribute = 0 for FVS Variables and 1 for Economic Variables; They share grpboxStandAttributeTieBreakerVariable
            int idxAttribute = this.lvFVSVariablesTieBreakerValues.SelectedItems[0].Index;
            grpboxStandAttributeTieBreakerVariable.Text = lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_METHOD].Text;            
            if (idxAttribute == 0)      //FVS ATTRIBUTE
            {
                this.lstFVSTablesList.Visible = true;
                this.lstFVSFieldsList.Visible = true;
                this.grpboxFVSVariablesTieBreakerVariableValueSource.Visible = true;
                this.lstEconVariablesList.Visible = false;
                this.lblTieBreakVarDescr.Visible = false;
                this.txtTieBreakVarDescr.Visible = false;
                this.grpMaxMin.Location = _objGrpMaxMinLocation;
                this.lblTieBreakVarDescr.Location = _objLblTieBreakVarDescrLocation;
                this.txtTieBreakVarDescr.Location = _objtxtTieBreakVarDescrLocation;
                this.loadFVSTableAndField(lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_FVSVARIABLE].Text);
                if (lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_VALUESOURCE].Text.Trim().ToUpper() == "POST")
                    this.cmbFVSVariablesTieBreakerVariableValueSource.Text =
                        this.cmbFVSVariablesTieBreakerVariableValueSource.Items[0].ToString();
                else if (lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_VALUESOURCE].Text.Trim().ToUpper() == "POST-PRE")
                    this.cmbFVSVariablesTieBreakerVariableValueSource.Text =
                        this.cmbFVSVariablesTieBreakerVariableValueSource.Items[1].ToString();
            }
            else if (idxAttribute == 1)     //ECONOMIC ATTRIBUTE
            {
                this.lstFVSTablesList.Visible = false;
                this.lstFVSFieldsList.Visible = false;
                this.grpboxFVSVariablesTieBreakerVariableValueSource.Visible = false;
                this.lstEconVariablesList.Visible = true;
                this.lblTieBreakVarDescr.Visible = true;
                this.txtTieBreakVarDescr.Visible = true;
                this.grpMaxMin.Location = new Point(this.grpboxFVSVariablesTieBreakerVariableValueSource.Location.X, 
                    this.grpMaxMin.Location.Y);
                this.lblTieBreakVarDescr.Location = new Point(this.lstFVSFieldsList.Location.X,
                    this.lblTieBreakVarDescr.Location.Y);
                this.txtTieBreakVarDescr.Location = new Point(this.lblTieBreakVarDescr.Location.X + 80,
                    this.txtTieBreakVarDescr.Location.Y);
                lstEconVariablesList.SelectedIndex = -1;
                for (int index = 0; index < lstEconVariablesList.Items.Count; index++)
                {
                    string item = lstEconVariablesList.Items[index].ToString();
                    if (lvFVSVariablesTieBreakerValues.Items[1].SubItems[COLUMN_FVSVARIABLE].Text == item)
                    {
                        lstEconVariablesList.SelectedIndex = index;
                        break;
                    }
                }
            }
            // Do this after the lists are set so it doesn't get reset
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_FVSVARIABLE].Text.Trim();
            
            if (this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text.Trim() == "MAX")
				this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked=true;
			else if (this.lvFVSVariablesTieBreakerValues.Items[idxAttribute].SubItems[COLUMN_MAXMIN].Text.Trim()=="MIN")
				this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked=true;
			else
				this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked=true;


			this.grpboxFVSVariablesTieBreaker.Hide();
			this.EnableTabs(false);
			if (idxAttribute == 0 || idxAttribute == 1)
			{
				this.grpboxStandAttributeTieBreakerVariable.Show();
			}
			else
			{
				this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Show();
			}
		
		}

		private void grpboxFVSVariablesTieBreaker_Resize(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Top = this.grpboxFVSVariablesTieBreaker.Top;
			this.grpboxStandAttributeTieBreakerVariable.Top = this.grpboxFVSVariablesTieBreaker.Top;
		}

		private void uc_scenario_treatment_intensity1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void grpboxFVSVariablesTieBreakerTreatmentIntensity_Resize(object sender, System.EventArgs e)
		{
            this.uc_scenario_last_tiebreak_rank1.Width = grpboxFVSVariablesTieBreakerLastTieBreakRank.ClientSize.Width - this.uc_scenario_last_tiebreak_rank1.Left * 2;
            this.uc_scenario_last_tiebreak_rank1.Height = this.btnFVSVariablesTieBreakerTreatmentIntensityClear.Top - this.uc_scenario_last_tiebreak_rank1.Top;
		}

		private void btnFVSVariablesTieBreakerVariableClear_Click(object sender, System.EventArgs e)
		{
            this.lstFVSTablesList.ClearSelected();
            this.cmbFVSVariablesTieBreakerVariableValueSource.Text = "Post Value";
            this.rdoFVSVariablesTieBreakerVariableValuesSelectedMin.Checked = true;
		}

		private void btnFVSVariablesTieBreakerVariableValues_Click(object sender, System.EventArgs e)
		{
            if (this.lvFVSVariablesTieBreakerValues.SelectedItems[0].Index == 0)
            {
                if (this.lstFVSTablesList.SelectedItems.Count == 0 || this.lstFVSFieldsList.SelectedItems.Count == 0) return;
                this.lblFVSVariablesTieBreakerVariableValuesSelected.Text =
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();
            }
            else if (this.lvFVSVariablesTieBreakerValues.SelectedItems[0].Index == 1)
            {
                if (this.lstEconVariablesList.SelectedItems.Count == 0) return;
                this.lblFVSVariablesTieBreakerVariableValuesSelected.Text =
                this.lstEconVariablesList.SelectedItems[0].ToString();
            }
		}

		private void btnFVSVariablesTieBreakerVariableNext_Click(object sender, System.EventArgs e)
		{
			this.grpboxStandAttributeTieBreakerVariable.Hide();
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Show();
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityCancel_Click(object sender, System.EventArgs e)
		{
            this.uc_scenario_last_tiebreak_rank1.m_DataSet.RejectChanges();
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Hide();
			this.EnableTabs(true);
			grpboxFVSVariablesTieBreaker.Show();
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityClear_Click(object sender, System.EventArgs e)
		{
            for (int x = 0; x <= uc_scenario_last_tiebreak_rank1.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows.Count - 1; x++)
                this.uc_scenario_last_tiebreak_rank1.m_DataSet.Tables["scenario_last_tiebreak_rank"].Rows[x]["last_tiebreak_rank"] = System.DBNull.Value;
		}

		private void btnFVSVariablesTieBreakerVariableCancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxStandAttributeTieBreakerVariable.Hide();
			this.EnableTabs(true);
			grpboxFVSVariablesTieBreaker.Show();
		}

		private void btnFVSVariablesTieBreakerVariableDone_Click(object sender, System.EventArgs e)
		{
			this.ReferenceOptimizerScenarioForm.m_bSave=true;
			UpdateFVSVariableListView();

			this.grpboxStandAttributeTieBreakerVariable.Hide();
			this.EnableTabs(true);
			this.grpboxFVSVariablesTieBreaker.Show();
		}
		
		private void UpdateFVSVariableListView()
		{
            int idxListView = 0;
            if (grpboxStandAttributeTieBreakerVariable.Text.ToUpper().Equals("STAND ATTRIBUTE"))
            {
                if (this.cmbFVSVariablesTieBreakerVariableValueSource.Text.Trim().ToUpper() == "POST VALUE")
                {
                    this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text = "POST";
                }
                else
                {
                    this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text = "POST-PRE";
                }
                this.m_oSavTieBreakerCollection.Item(0).strValueSource = this.lvFVSVariablesTieBreakerValues.Items[0].SubItems[COLUMN_VALUESOURCE].Text;
            }
            else if (grpboxStandAttributeTieBreakerVariable.Text.ToUpper().Equals("ECONOMIC ATTRIBUTE"))
            {
                idxListView = 1;
            }
            if (!this.lblFVSVariablesTieBreakerVariableValuesSelected.Text.ToUpper().Equals("NA"))
            {
                this.lvFVSVariablesTieBreakerValues.Items[idxListView].SubItems[COLUMN_FVSVARIABLE].Text = this.lblFVSVariablesTieBreakerVariableValuesSelected.Text;
            }
            this.m_oSavTieBreakerCollection.Item(idxListView).strFVSVariableName = this.lblFVSVariablesTieBreakerVariableValuesSelected.Text;
			if (this.rdoFVSVariablesTieBreakerVariableValuesSelectedMax.Checked)
			{
                this.lvFVSVariablesTieBreakerValues.Items[idxListView].SubItems[COLUMN_MAXMIN].Text = 
					"MAX";
                this.m_oSavTieBreakerCollection.Item(idxListView).strMaxYN = "Y";
                this.m_oSavTieBreakerCollection.Item(idxListView).strMinYN = "N";
			}
			else
			{
                this.lvFVSVariablesTieBreakerValues.Items[idxListView].SubItems[COLUMN_MAXMIN].Text = 
					"MIN";
                this.m_oSavTieBreakerCollection.Item(idxListView).strMaxYN = "N";
                this.m_oSavTieBreakerCollection.Item(idxListView).strMinYN = "Y";

			}
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityDone_Click(object sender, System.EventArgs e)
		{
			this.ReferenceOptimizerScenarioForm.m_bSave=true;
            this.uc_scenario_last_tiebreak_rank1.m_DataSet.AcceptChanges();
			this.UpdateFVSVariableListView();

			this.EnableTabs(true);
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Hide();

			this.grpboxFVSVariablesTieBreaker.Show();
		}

		private void lvFVSVariablesTieBreakerValues_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
            string strMethod = this.lvFVSVariablesTieBreakerValues.Items[e.Index].SubItems[COLUMN_METHOD].Text.Trim();
            //Tie-Break Rank is read-only and cannot be unchecked
            if (strMethod.Equals("Last Tie-Break Rank") && e.NewValue != CheckState.Checked)
            {
                e.NewValue = e.CurrentValue;
                MessageBox.Show("Last Tie-Break Rank is a required Tie Breaker. It cannot be unselected.", 
                    "FIA Biosum");
            }
            else
            {
                bool bStandChecked = false;
                bool bEconomicChecked = false;
                if (this.lvFVSVariablesTieBreakerValues.Items[0] != null)
                {
                    bStandChecked = this.lvFVSVariablesTieBreakerValues.Items[0].Checked;
                }
                if (this.lvFVSVariablesTieBreakerValues.Items[1] != null)
                {
                    bEconomicChecked = this.lvFVSVariablesTieBreakerValues.Items[1].Checked;
                }
                if (strMethod.Equals("Stand Attribute") && e.NewValue == CheckState.Checked)
                {
                    if (bEconomicChecked == true)
                    {
                        e.NewValue = e.CurrentValue;
                        MessageBox.Show("!! Stand Attribute and Economic Attribute cannot both be selected as tiebreakers. Uncheck " +
                        "Economic Attribute before checking Stand Attribute! ", "FIA Biosum");
                    }
                }
                else if (strMethod.Equals("Economic Attribute") && e.NewValue == CheckState.Checked)
                {
                    if (bStandChecked == true)
                    {
                        e.NewValue = e.CurrentValue;
                        MessageBox.Show("!! Stand Attribute and Economic Attribute cannot both be selected as tiebreakers. Uncheck " +
                            "Stand Attribute before checking Economic Attribute! ", "FIA Biosum");
                    }
                }
                this.ReferenceOptimizerScenarioForm.m_bSave = true;
            }
		}

		private void btnFVSVariablesTieBreakerAudit_Click(object sender, System.EventArgs e)
		{
			Audit(true);
		}

		private void lvFVSVariablesTieBreakerValues_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = this.lvFVSVariablesTieBreakerValues.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lvFVSVariablesTieBreakerValues.Items[this.lvFVSVariablesTieBreakerValues.TopItem.Index + (int)dblRow-1].Selected=true;
					
				}
			}
			catch 
			{
			}
		}

        private void loadFVSTableAndField(string strFvsVariableName)
        {
            string[] strPieces = strFvsVariableName.Split('.');
            //set FVS table and variable
            if (strPieces.Length == 2)
            {
                for (int index = 0; index < lstFVSTablesList.Items.Count; index++)
                {
                    string item = lstFVSTablesList.Items[index].ToString();
                    if (strPieces[0] == item)
                    {
                        lstFVSTablesList.SelectedIndex = index;
                        break;
                    }
                }
                for (int index = 0; index < lstFVSFieldsList.Items.Count; index++)
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

		private void lvFVSVariablesTieBreakerValues_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lvFVSVariablesTieBreakerValues.SelectedItems.Count > 0)
				this.m_oLvRowColors.DelegateListViewItem(this.lvFVSVariablesTieBreakerValues.SelectedItems[0]);
		}

		private void btnFVSVariablesTieBreakerTreatmentIntensityPrev_Click(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesTieBreakerLastTieBreakRank.Hide();
			this.grpboxStandAttributeTieBreakerVariable.Show();
			
		}
	
		public bool DisplayAuditMessage
		{
			get {return _bDisplayAuditMsg;}
			set {_bDisplayAuditMsg=value;}
		}
		
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmScenario;}
			set {_frmScenario=value;}
		}
		public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_optimization ReferenceOptimizationUserControl
		{
			get {return _uc_optimization;}
			set {_uc_optimization=value;}
		}

        private void lstFVSTablesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstFVSFieldsList.Items.Clear();
            txtTieBreakVarDescr.Text = "";
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = "Not Defined";
            if (this.lstFVSTablesList.SelectedIndex > -1)
            {
                System.Collections.Generic.IList<string> lstFields = 
                    _dictFVSTables[Convert.ToString(this.lstFVSTablesList.SelectedItem)];
                if (lstFields != null)
                {
                    foreach (string strField in lstFields)
                    {
                        lstFVSFieldsList.Items.Add(strField);
                    }
                }
                // Control visibility of the weighted variable description fields
                bool bDisplayWeightedFields = this.lstFVSTablesList.SelectedItems[0].ToString().ToUpper().Contains("_WEIGHTED");
                lblTieBreakVarDescr.Visible = bDisplayWeightedFields;
                txtTieBreakVarDescr.Visible = bDisplayWeightedFields;
            }
            else
            {
                lblTieBreakVarDescr.Visible = false;
                txtTieBreakVarDescr.Visible = false;
            }
        }

        private void lstFVSFieldsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = "Not Defined";
            this.txtTieBreakVarDescr.Text = "";
            if (this.lstFVSFieldsList.SelectedIndex > -1)
            {
                this.btnFVSVariablesTieBreakerVariableValues.Enabled = true;
                if (this.lstFVSTablesList.SelectedItems[0].ToString().ToUpper().Contains("_WEIGHTED") == true)
                {
                    foreach (uc_optimizer_scenario_calculated_variables.VariableItem oItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
                    {
                        if (oItem.strVariableName.Equals(Convert.ToString(this.lstFVSFieldsList.SelectedItem)))
                        {
                            if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                                txtTieBreakVarDescr.Text = oItem.strVariableDescr;
                            break;
                        }
                    }
                }
            }
            else
            {
                this.btnFVSVariablesTieBreakerVariableValues.Enabled = false;
            }
        }

        private void lstEconVariablesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.lblFVSVariablesTieBreakerVariableValuesSelected.Text = "Not Defined";
            this.txtTieBreakVarDescr.Text = "";
            if (this.lstEconVariablesList.SelectedIndex > -1)
            {
                this.btnFVSVariablesTieBreakerVariableValues.Enabled = true;
                foreach (uc_optimizer_scenario_calculated_variables.VariableItem oItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
                {
                    if (oItem.strVariableName.Equals(Convert.ToString(this.lstEconVariablesList.SelectedItem)))
                    {
                        if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                            txtTieBreakVarDescr.Text = oItem.strVariableDescr;
                        break;
                    }
                }
            }
        }

        public string HelpChapter
        {
            get
            {
                return m_strHelpChapter;
            }
        }

        private void grpboxStandAttributeTieBreakerVariable_VisibleChanged(object sender, EventArgs e)
        {
            if (grpboxStandAttributeTieBreakerVariable.Visible == false)
            {
                m_strHelpChapter = "TIEBREAKER_SETTINGS";
            }
            else
            {
                m_strHelpChapter = "EDIT_TIEBREAKER_ATTRIBUTE";
            }
        }

        private void grpboxFVSVariablesTieBreakerLastTieBreakRank_VisibleChanged(object sender, EventArgs e)
        {
            if (grpboxFVSVariablesTieBreakerLastTieBreakRank.Visible == false)
            {
                m_strHelpChapter = "TIEBREAKER_SETTINGS";
            }
            else
            {
                m_strHelpChapter = "EDIT_LAST_TIEBREAK_RANK";
            }
        }

        private void ClearVariable(int p_intIndex)
        {
            for (int x = 0; x <= this.m_oSavTieBreakerCollection.Count - 1; x++)
            {
                // We only update the selected listview item
                if (this.m_oSavTieBreakerCollection.Item(x).intListViewIndex == this.lvFVSVariablesTieBreakerValues.SelectedItems[0].Index)
                {
                    TieBreakerItem oClearItem = this.m_oSavTieBreakerCollection.Item(x);
                    // FVS ATTRIBUTE
                    if (oClearItem.strMethod.Trim().ToUpper().Equals("STAND ATTRIBUTE"))
                    {
                        oClearItem.strValueSource = "Not Defined";
                    }
                    // ECONOMIC ATTRIBUTE
                    else
                    {
                        oClearItem.strValueSource = "NA";
                    }
                    oClearItem.strFVSVariableName = "Not Defined";
                    oClearItem.strMaxYN = "N";
                    oClearItem.strMinYN = "N";
                    oClearItem.bSelected = false;
                }
            }

            if (this.lvFVSVariablesTieBreakerValues.Items[p_intIndex].SubItems[COLUMN_METHOD].Text.Trim().ToUpper() == "STAND ATTRIBUTE")
            {
                this.lvFVSVariablesTieBreakerValues.Items[p_intIndex].SubItems[COLUMN_VALUESOURCE].Text = "Not Defined";
            }
            else
            {
                this.lvFVSVariablesTieBreakerValues.Items[p_intIndex].SubItems[COLUMN_VALUESOURCE].Text = "NA";
            }

            this.lvFVSVariablesTieBreakerValues.Items[p_intIndex].SubItems[COLUMN_FVSVARIABLE].Text = "Not Defined";
            this.lvFVSVariablesTieBreakerValues.Items[p_intIndex].SubItems[COLUMN_MAXMIN].Text = "Not Defined";
            this.lvFVSVariablesTieBreakerValues.Items[p_intIndex].Checked = false;
        }

        private void BtnTiebreakerClear_Click(object sender, EventArgs e)
        {
            if (this.lvFVSVariablesTieBreakerValues.SelectedItems.Count == 0)
            {
                MessageBox.Show("!! No variable selected to clear!", "FIA Biosum", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            if (this.lvFVSVariablesTieBreakerValues.SelectedItems[0].SubItems[COLUMN_METHOD].Text.Trim().ToUpper() == "LAST TIE-BREAK RANK")
            {
                MessageBox.Show("!! Last Tie-Break Rank is required and cannot be cleared!", "FIA Biosum", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (this.lvFVSVariablesTieBreakerValues.SelectedItems[0].SubItems[COLUMN_FVSVARIABLE].Text.Trim() == "Not Defined") return;

            DialogResult result = MessageBox.Show("Are you sure you wish to delete this Tie-Breaker variable ? (YN)", "FIA Biosum",
                System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (result == System.Windows.Forms.DialogResult.Yes)
                ClearVariable(this.lvFVSVariablesTieBreakerValues.SelectedItems[0].Index);
        }
	}
}

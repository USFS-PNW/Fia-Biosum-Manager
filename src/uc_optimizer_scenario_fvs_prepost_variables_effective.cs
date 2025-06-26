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
	public class uc_optimizer_scenario_fvs_prepost_variables_effective : System.Windows.Forms.UserControl
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.ComponentModel.IContainer components;
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
		public System.Windows.Forms.Label lblTitle;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostVariable;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostVariablePostSelected;
		private System.Windows.Forms.Label lblFVSVariablesPrePostVariablePostSelected;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostVariablePreSelected;
		private System.Windows.Forms.Label lblFVSVariablesPrePostVariablePreSelected;
		private System.Windows.Forms.Button btnFVSVariablesPrePostVariableClearAll;
		private System.Windows.Forms.Button btnFVSVariablesPrePostVariableNext;
		private System.Windows.Forms.Button btnFVSVariablesPrePostVariableCancel;
        private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostVariableValues;
		private System.Windows.Forms.Button btnN;
		private System.Windows.Forms.Button btnY;
		private System.Windows.Forms.Button btnFFENote;
		private System.Windows.Forms.Button btnFFEOr;
		private System.Windows.Forms.Button btnFFEAnd;
		private System.Windows.Forms.Button btnRightBracket;
		private System.Windows.Forms.Button btnLeftBracket;
		private System.Windows.Forms.Button btnLessThanOrEqualTo;
		private System.Windows.Forms.Button btnFFEMoreThanOrEqualTo;
		private System.Windows.Forms.Button btnFFELessThan;
		private System.Windows.Forms.Button btnFFEMoreThan;
		private System.Windows.Forms.Button btnFFENotEqual;
		private System.Windows.Forms.Button btnFFEEqual;
		private System.Windows.Forms.Button btnFVSVariablesPrePostVariableValue;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostExpression;
		private FIA_Biosum_Manager.frmOptimizerScenario _frmScenario=null;
		private FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_optimization _uc_optimization;
		private FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker _uc_tiebreaker;
		
		private System.Windows.Forms.Label lblFVSVariablesPrePostExpression;
		private System.Windows.Forms.TextBox txtExpression;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionPrev;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionClear;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionTest;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionPrevious;
		private System.Windows.Forms.GroupBox grpboxExpressionsBtns;
		private System.Windows.Forms.ComboBox cmbFVSVariablesPrePost;
		private System.Windows.Forms.Button btnFVSVariablesPrePostGo;

		
		private int m_intCurVar=-1;
		public  FIA_Biosum_Manager.SQLMacroSubstitutionVariable_Collection m_oSQLMacroSubstitutionVariable_Collection= new SQLMacroSubstitutionVariable_Collection();
		private FIA_Biosum_Manager.macrosubst m_oMacroVar = new macrosubst();
		int m_intCurVariableDefinitionStepCount=1;
		string[] m_strUserNavigation=null;
		
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionCancel;
		private System.Windows.Forms.Button btnFVSVariablesPrePostVariableDone;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionDone;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionNext;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostExpressionSelectedVariables;
		private System.Windows.Forms.ListBox lstFVSVariablesPrePostExpressionSelectedVariables;
		private Variables m_oOldVar=null;
		private Variables m_oCurVar=null;
		public Variables m_oSavVar=null;

		public const byte NUMBER_OF_VARIABLES=4;
		const byte VARIABLE_DEFINITION_STEPS=4;


		const int WIZARD_STEP_VARIABLES_DEFINED=0;
		const int WIZARD_STEP_VARIABLE_SELECT=1;
		const int WIZARD_STEP_VARIABLE_BETTER=2;
		const int WIZARD_STEP_VARIABLE_WORSE=3;
		const int WIZARD_STEP_VARIABLE_EFFECTIVE=4;
		const int WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE=5;

		
		//const int COLUMN_OPTIMIZE=0;
		const int COLUMN_NULL=0;
		const int COLUMN_VARID=1;
		const int COLUMN_PREVAR=2;
		const int COLUMN_POSTVAR=3;
		const int COLUMN_BETTER=4;
		const int COLUMN_WORSE=5;
		const int COLUMN_EFFECTIVE=6;


		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblFVSVariablesPrePostExpressionVariable1;
		private System.Windows.Forms.Label lblFVSVariablesPrePostExpressionVariable2;
		private System.Windows.Forms.Button btnFVSVariablesPrePostExpressionDefault;
		private System.Windows.Forms.Label lblFVSVariablesPrePostExpressionVariable3;
		private System.Windows.Forms.Label lblFVSVariablesPrePostExpressionVariable4;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostValuesButtons;
		private System.Windows.Forms.Button btnFVSVariablesPrePostValuesButtonsClearAll;
		private System.Windows.Forms.Button btnFVSVariablesPrePostValuesButtonsClear;
		private System.Windows.Forms.Button btnFVSVariablesPrePostValuesButtonsEdit;
		private System.Windows.Forms.ListView lvFVSVariablesPrePostValues;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ColumnHeader lvColVariableId;
		private System.Windows.Forms.ColumnHeader lvColPreVariableName;
		private System.Windows.Forms.ColumnHeader lvColPostVariableName;
		private System.Windows.Forms.ColumnHeader lvColVariableImprovement;
		private System.Windows.Forms.ColumnHeader lvColVariableDisimprovement;
		private System.Windows.Forms.ColumnHeader lvColVariableEffective;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostAudit;
		private System.Windows.Forms.Button btnFVSVariablesPrePostAudit;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePost;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostOverall;
		private System.Windows.Forms.Button btnFVSVariablesPrePostOverall;
		private System.Windows.Forms.GroupBox grpboxFVSVariablesPrePostValues;
		private System.Windows.Forms.Panel pnlFVSVariablesPrePost;
		private System.Windows.Forms.Panel pnlFVSVariablesPrePostVariable;
		private System.Windows.Forms.Panel pnlFVSVariablesPrePostExpression;

		public bool m_bSave=false;
		private bool _bDisplayAuditMsg=true;
        private System.Windows.Forms.ColumnHeader lvColVariableNull;
        private FIA_Biosum_Manager.ListViewAlternateBackgroundColors m_oLvRowColors = new ListViewAlternateBackgroundColors();
        private TextBox txtEffVarDescr;
        private Label lblEffVarDescr;
        private FIA_Biosum_Manager.ValidateNumericValues m_oValidate = new ValidateNumericValues();
        private ListBox lstFVSTablesList;
        private ListBox lstFVSFieldsList;
        private FIA_Biosum_Manager.OptimizerScenarioTools m_oOptimizerScenarioTools = new OptimizerScenarioTools();
        private System.Collections.Generic.Dictionary<string, System.Collections.Generic.IList<String>> m_dictFVSTables;
        private string m_strHelpChapter = "EFFECTIVE_SETTINGS";
        
        public class Variables
		{
			public string[] m_strPreVarArray = new string[NUMBER_OF_VARIABLES];
			public string[] m_strPostVarArray = new string[NUMBER_OF_VARIABLES];
			public string[] m_strBetterExpr = new string[NUMBER_OF_VARIABLES];
			public string[] m_strWorseExpr = new string[NUMBER_OF_VARIABLES];
			public string[] m_strEffectiveExpr = new string[NUMBER_OF_VARIABLES];
			public string m_strOverallEffectiveExpr="";

			public Variables()
			{
				for (int x=0;x<=NUMBER_OF_VARIABLES-1;x++)
				{
					m_strPreVarArray[x]="Not Defined";
					m_strPostVarArray[x] = "Not Defined";
					m_strBetterExpr[x]="";
					m_strWorseExpr[x]="";
					m_strEffectiveExpr[x]="";
				}
			}
			public void Copy(Variables p_oSource, ref Variables p_oDest)
			{
				for (int x=0;x<=NUMBER_OF_VARIABLES-1;x++)
				{
				    p_oDest.m_strPreVarArray[x] = p_oSource.m_strPreVarArray[x];
					p_oDest.m_strPostVarArray[x] = p_oSource.m_strPostVarArray[x];
					p_oDest.m_strBetterExpr[x] = p_oSource.m_strBetterExpr[x];
					p_oDest.m_strWorseExpr[x] = p_oSource.m_strWorseExpr[x];
					p_oDest.m_strEffectiveExpr[x] = p_oSource.m_strEffectiveExpr[x];

				}
				p_oDest.m_strOverallEffectiveExpr = p_oSource.m_strOverallEffectiveExpr;
			}
			/// <summary>
			/// return the table names found in the either m_strPreVarArray or m_strPostVarArray variables
			/// </summary>
			/// <param name="p_oVarArray">m_strPreVarArray or m_strPostVarArray values</param>
			/// <returns></returns>
			public string[] TableNames(string[] p_oVarArray)
			{
				string strTable="";
				string strTableList=",";
				string[] strTableArray=null;
				for (int x=0;x<=p_oVarArray.Length-1;x++)
				{
					strTable=this.TableName(p_oVarArray[x]);
					if (strTable.Trim().Length > 0)
					{
						if (strTableList.Trim().ToUpper().IndexOf("," + strTable.ToUpper().Trim().ToUpper() + ",",0) != 0)
						{
							strTableList = strTableList + strTable.Trim() + ",";
						}

					}
				}
			
				if (strTableList.Trim().Length > 1)
				{
					strTableList = strTableList.Substring(1,strTableList.Length - 2);
					strTableArray = frmMain.g_oUtils.ConvertListToArray(strTableList,",");
					
				}
				else 
				{
					strTableList="";
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
				string strTableName="";
				if (p_strValue.Trim().ToUpper() != "NOT DEFINED")
				{
					if (p_strValue.IndexOf(".",0) > 0)
					{
						string[] strArray = frmMain.g_oUtils.ConvertListToArray(p_strValue,".");
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
				string strColumnName="";
				if (p_strValue.Trim().ToUpper() != "NOT DEFINED")
				{
					if (p_strValue.IndexOf(".",0) > 0)
					{
						string[] strArray = frmMain.g_oUtils.ConvertListToArray(p_strValue,".");
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


			public bool Modified(Variables p_oCompare)
			{
				if (m_strOverallEffectiveExpr.Trim().ToUpper() != p_oCompare.m_strOverallEffectiveExpr.Trim().ToUpper()) return true;

				for (int x=0;x<=NUMBER_OF_VARIABLES-1;x++)
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
	
		

		public uc_optimizer_scenario_fvs_prepost_variables_effective()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.m_oUtils = new utils();
			this.m_strCurrentIndexTypeAndClass="";

			
			

			this.grpboxFVSVariablesPrePostVariable.Top = this.grpboxFVSVariablesPrePost.Top;
			this.grpboxFVSVariablesPrePostVariable.Left = this.grpboxFVSVariablesPrePost.Left;
			this.grpboxFVSVariablesPrePostVariable.Height = this.grpboxFVSVariablesPrePost.Height;
			this.grpboxFVSVariablesPrePostVariable.Width = this.grpboxFVSVariablesPrePost.Width;
			this.grpboxFVSVariablesPrePostVariable.Hide();

			this.grpboxFVSVariablesPrePostExpression.Top = this.grpboxFVSVariablesPrePost.Top;
			this.grpboxFVSVariablesPrePostExpression.Left = this.grpboxFVSVariablesPrePost.Left;
			this.grpboxFVSVariablesPrePostExpression.Height = this.grpboxFVSVariablesPrePost.Height;
			this.grpboxFVSVariablesPrePostExpression.Width = this.grpboxFVSVariablesPrePost.Width;
			this.grpboxFVSVariablesPrePostExpression.Hide();


			FIA_Biosum_Manager.SQLMacroSubstitutionVariableItem oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "pre_v1";
			oItem.Index = 0;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "pre_v2";
			oItem.Index = 1;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "pre_v3";
			oItem.Index = 2;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "post_v1";
			oItem.Index = 3;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);


			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "post_v2";
			oItem.Index = 4;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "post_v3";
			oItem.Index = 5;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v1_better";
			
			oItem.SQLVariableSubstitutionString = "effective.variable1_better_yn";
			oItem.Index = 6;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v1_worse";
			oItem.SQLVariableSubstitutionString = "effective.variable1_worse_yn";
			oItem.Index = 7;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v2_better";
			oItem.SQLVariableSubstitutionString = "effective.variable2_better_yn";
			oItem.Index = 8;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v2_worse";
			oItem.SQLVariableSubstitutionString = "effective.variable2_worse_yn";
			oItem.Index = 9;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v3_better";
			oItem.SQLVariableSubstitutionString = "effective.variable3_better_yn";
			oItem.Index = 10;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v3_worse";
			oItem.SQLVariableSubstitutionString = "effective.variable3_worse_yn";
			oItem.Index = 11;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v1_change";
			oItem.SQLVariableSubstitutionString = "effective.variable1_change";
			oItem.Index = 12;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v2_change";
			oItem.SQLVariableSubstitutionString = "effective.variable2_change";
			oItem.Index = 13;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v3_change";
			oItem.SQLVariableSubstitutionString = "effective.variable3_change";
			oItem.Index = 14;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v1_effective";
			oItem.SQLVariableSubstitutionString = "effective.variable1_better_yn='Y' AND effective.variable1_worse_yn='N'";
			oItem.Index = 15;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v2_effective";
			oItem.SQLVariableSubstitutionString = "effective.variable2_better_yn='Y' AND effective.variable2_worse_yn='N'";
			oItem.Index = 16;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			oItem = new SQLMacroSubstitutionVariableItem();
			oItem.VariableName = "v3_effective";
			oItem.SQLVariableSubstitutionString = "effective.variable3_better_yn='Y' AND effective.variable3_worse_yn='N'";
			oItem.Index = 17;
			this.m_oSQLMacroSubstitutionVariable_Collection.Add(oItem);

			EnableAllMacroButtons(false);


			oItem=null;

			
			this.m_strOverallEffectiveExpression="";

			this.m_oMacroVar.ReferenceSQLMacroSubstitutionVariableCollection = this.m_oSQLMacroSubstitutionVariable_Collection;

			this.m_strUserNavigation = new string[this.cmbFVSVariablesPrePost.Items.Count];

			InitializeUserNavigation();
			
			


			if (frmMain.g_oGridViewFont != null) lvFVSVariablesPrePostValues.Font = frmMain.g_oGridViewFont;



			for (int x=COLUMN_PREVAR;x<=this.lvFVSVariablesPrePostValues.Columns.Count-1;x++)
			{
				int textWidth = (int)(this.CreateGraphics().MeasureString(lvFVSVariablesPrePostValues.Columns[x].Text, this.lvFVSVariablesPrePostValues.Font).Width * 1.5);
				lvFVSVariablesPrePostValues.Columns[x].Width = textWidth;
			}

			this.m_oLvRowColors.InitializeRowCollection();
			this.m_oLvRowColors.ReferenceListView = this.lvFVSVariablesPrePostValues;
			this.m_oLvRowColors.CustomFullRowSelect=true;

			for (int x=0;x<=this.lvFVSVariablesPrePostValues.Items.Count-1;x++)
			{
				this.lvFVSVariablesPrePostValues.Items[x].UseItemStyleForSubItems=false;

				this.m_oLvRowColors.AddRow();
				this.m_oLvRowColors.AddColumns(x,this.lvFVSVariablesPrePostValues.Columns.Count);

			    //lvFVSVariablesPrePostValues.Items[x].SubItems[COLUMN_BETTER].ForeColor = Color.Gray;
				lvFVSVariablesPrePostValues.Items[x].SubItems[COLUMN_BETTER].Text = "No";

				//lvFVSVariablesPrePostValues.Items[x].SubItems[COLUMN_WORSE].ForeColor = Color.Gray;
				lvFVSVariablesPrePostValues.Items[x].SubItems[COLUMN_WORSE].Text = "No";

				//lvFVSVariablesPrePostValues.Items[x].SubItems[COLUMN_EFFECTIVE].ForeColor = Color.Gray;
				lvFVSVariablesPrePostValues.Items[x].SubItems[COLUMN_EFFECTIVE].Text = "No";


			}
			this.m_oLvRowColors.ListView();

            lblFVSVariablesPrePostExpressionVariable2.Text = "";


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
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "1",
            "Not Defined",
            "Not Defined",
            "No",
            "No",
            "No"}, -1);
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "2",
            "Not Defined",
            "Not Defined",
            "No",
            "No",
            "No"}, -1);
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "3",
            "Not Defined",
            "Not Defined",
            "No",
            "No",
            "No"}, -1);
            System.Windows.Forms.ListViewItem listViewItem4 = new System.Windows.Forms.ListViewItem(new string[] {
            "",
            "4",
            "Not Defined",
            "Not Defined",
            "No",
            "No",
            "No"}, -1);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.grpboxFVSVariablesPrePost = new System.Windows.Forms.GroupBox();
            this.pnlFVSVariablesPrePost = new System.Windows.Forms.Panel();
            this.grpboxFVSVariablesPrePostValues = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lvFVSVariablesPrePostValues = new System.Windows.Forms.ListView();
            this.lvColVariableNull = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColVariableId = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColPreVariableName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColPostVariableName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColVariableImprovement = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColVariableDisimprovement = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvColVariableEffective = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.grpboxFVSVariablesPrePostValuesButtons = new System.Windows.Forms.GroupBox();
            this.btnFVSVariablesPrePostValuesButtonsClearAll = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostValuesButtonsClear = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostValuesButtonsEdit = new System.Windows.Forms.Button();
            this.grpboxFVSVariablesPrePostOverall = new System.Windows.Forms.GroupBox();
            this.btnFVSVariablesPrePostOverall = new System.Windows.Forms.Button();
            this.grpboxFVSVariablesPrePostAudit = new System.Windows.Forms.GroupBox();
            this.btnFVSVariablesPrePostAudit = new System.Windows.Forms.Button();
            this.grpboxFVSVariablesPrePostExpression = new System.Windows.Forms.GroupBox();
            this.pnlFVSVariablesPrePostExpression = new System.Windows.Forms.Panel();
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables = new System.Windows.Forms.GroupBox();
            this.lstFVSVariablesPrePostExpressionSelectedVariables = new System.Windows.Forms.ListBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnFFEEqual = new System.Windows.Forms.Button();
            this.btnFFENote = new System.Windows.Forms.Button();
            this.btnY = new System.Windows.Forms.Button();
            this.btnN = new System.Windows.Forms.Button();
            this.btnFFENotEqual = new System.Windows.Forms.Button();
            this.btnLeftBracket = new System.Windows.Forms.Button();
            this.btnRightBracket = new System.Windows.Forms.Button();
            this.btnFFEMoreThan = new System.Windows.Forms.Button();
            this.btnFFEMoreThanOrEqualTo = new System.Windows.Forms.Button();
            this.btnFFELessThan = new System.Windows.Forms.Button();
            this.btnLessThanOrEqualTo = new System.Windows.Forms.Button();
            this.btnFFEOr = new System.Windows.Forms.Button();
            this.btnFFEAnd = new System.Windows.Forms.Button();
            this.lblFVSVariablesPrePostExpressionVariable1 = new System.Windows.Forms.Label();
            this.lblFVSVariablesPrePostExpressionVariable2 = new System.Windows.Forms.Label();
            this.lblFVSVariablesPrePostExpressionVariable3 = new System.Windows.Forms.Label();
            this.lblFVSVariablesPrePostExpressionVariable4 = new System.Windows.Forms.Label();
            this.lblFVSVariablesPrePostExpression = new System.Windows.Forms.Label();
            this.txtExpression = new System.Windows.Forms.TextBox();
            this.grpboxExpressionsBtns = new System.Windows.Forms.GroupBox();
            this.btnFVSVariablesPrePostExpressionDefault = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionPrev = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionClear = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionTest = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionDone = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionCancel = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionPrevious = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostExpressionNext = new System.Windows.Forms.Button();
            this.grpboxFVSVariablesPrePostVariable = new System.Windows.Forms.GroupBox();
            this.pnlFVSVariablesPrePostVariable = new System.Windows.Forms.Panel();
            this.grpboxFVSVariablesPrePostVariableValues = new System.Windows.Forms.GroupBox();
            this.lstFVSFieldsList = new System.Windows.Forms.ListBox();
            this.lstFVSTablesList = new System.Windows.Forms.ListBox();
            this.txtEffVarDescr = new System.Windows.Forms.TextBox();
            this.lblEffVarDescr = new System.Windows.Forms.Label();
            this.btnFVSVariablesPrePostVariableValue = new System.Windows.Forms.Button();
            this.grpboxFVSVariablesPrePostVariablePreSelected = new System.Windows.Forms.GroupBox();
            this.lblFVSVariablesPrePostVariablePreSelected = new System.Windows.Forms.Label();
            this.grpboxFVSVariablesPrePostVariablePostSelected = new System.Windows.Forms.GroupBox();
            this.lblFVSVariablesPrePostVariablePostSelected = new System.Windows.Forms.Label();
            this.btnFVSVariablesPrePostVariableClearAll = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostVariableDone = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostVariableCancel = new System.Windows.Forms.Button();
            this.btnFVSVariablesPrePostVariableNext = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.cmbFVSVariablesPrePost = new System.Windows.Forms.ComboBox();
            this.btnFVSVariablesPrePostGo = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.grpboxFVSVariablesPrePost.SuspendLayout();
            this.pnlFVSVariablesPrePost.SuspendLayout();
            this.grpboxFVSVariablesPrePostValues.SuspendLayout();
            this.grpboxFVSVariablesPrePostValuesButtons.SuspendLayout();
            this.grpboxFVSVariablesPrePostOverall.SuspendLayout();
            this.grpboxFVSVariablesPrePostAudit.SuspendLayout();
            this.grpboxFVSVariablesPrePostExpression.SuspendLayout();
            this.pnlFVSVariablesPrePostExpression.SuspendLayout();
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpboxExpressionsBtns.SuspendLayout();
            this.grpboxFVSVariablesPrePostVariable.SuspendLayout();
            this.pnlFVSVariablesPrePostVariable.SuspendLayout();
            this.grpboxFVSVariablesPrePostVariableValues.SuspendLayout();
            this.grpboxFVSVariablesPrePostVariablePreSelected.SuspendLayout();
            this.grpboxFVSVariablesPrePostVariablePostSelected.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.grpboxFVSVariablesPrePost);
            this.groupBox1.Controls.Add(this.grpboxFVSVariablesPrePostExpression);
            this.groupBox1.Controls.Add(this.grpboxFVSVariablesPrePostVariable);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.cmbFVSVariablesPrePost);
            this.groupBox1.Controls.Add(this.btnFVSVariablesPrePostGo);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(872, 2500);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Resize += new System.EventHandler(this.groupBox1_Resize);
            // 
            // grpboxFVSVariablesPrePost
            // 
            this.grpboxFVSVariablesPrePost.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxFVSVariablesPrePost.Controls.Add(this.pnlFVSVariablesPrePost);
            this.grpboxFVSVariablesPrePost.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxFVSVariablesPrePost.ForeColor = System.Drawing.Color.Black;
            this.grpboxFVSVariablesPrePost.Location = new System.Drawing.Point(8, 88);
            this.grpboxFVSVariablesPrePost.Name = "grpboxFVSVariablesPrePost";
            this.grpboxFVSVariablesPrePost.Size = new System.Drawing.Size(856, 472);
            this.grpboxFVSVariablesPrePost.TabIndex = 32;
            this.grpboxFVSVariablesPrePost.TabStop = false;
            this.grpboxFVSVariablesPrePost.VisibleChanged += new System.EventHandler(this.grpboxFVSVariablesPrePost_VisibleChanged);
            this.grpboxFVSVariablesPrePost.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePost_Resize);
            // 
            // pnlFVSVariablesPrePost
            // 
            this.pnlFVSVariablesPrePost.AutoScroll = true;
            this.pnlFVSVariablesPrePost.Controls.Add(this.grpboxFVSVariablesPrePostValues);
            this.pnlFVSVariablesPrePost.Controls.Add(this.grpboxFVSVariablesPrePostOverall);
            this.pnlFVSVariablesPrePost.Controls.Add(this.grpboxFVSVariablesPrePostAudit);
            this.pnlFVSVariablesPrePost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFVSVariablesPrePost.Location = new System.Drawing.Point(3, 18);
            this.pnlFVSVariablesPrePost.Name = "pnlFVSVariablesPrePost";
            this.pnlFVSVariablesPrePost.Size = new System.Drawing.Size(850, 451);
            this.pnlFVSVariablesPrePost.TabIndex = 70;
            // 
            // grpboxFVSVariablesPrePostValues
            // 
            this.grpboxFVSVariablesPrePostValues.Controls.Add(this.label1);
            this.grpboxFVSVariablesPrePostValues.Controls.Add(this.lvFVSVariablesPrePostValues);
            this.grpboxFVSVariablesPrePostValues.Controls.Add(this.grpboxFVSVariablesPrePostValuesButtons);
            this.grpboxFVSVariablesPrePostValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxFVSVariablesPrePostValues.Name = "grpboxFVSVariablesPrePostValues";
            this.grpboxFVSVariablesPrePostValues.Size = new System.Drawing.Size(840, 256);
            this.grpboxFVSVariablesPrePostValues.TabIndex = 67;
            this.grpboxFVSVariablesPrePostValues.TabStop = false;
            this.grpboxFVSVariablesPrePostValues.Text = "Step 1: Define Variables";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(232, 24);
            this.label1.TabIndex = 66;
            this.label1.Text = "Maximum Of 4 Variables";
            // 
            // lvFVSVariablesPrePostValues
            // 
            this.lvFVSVariablesPrePostValues.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lvColVariableNull,
            this.lvColVariableId,
            this.lvColPreVariableName,
            this.lvColPostVariableName,
            this.lvColVariableImprovement,
            this.lvColVariableDisimprovement,
            this.lvColVariableEffective});
            this.lvFVSVariablesPrePostValues.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lvFVSVariablesPrePostValues.GridLines = true;
            this.lvFVSVariablesPrePostValues.HideSelection = false;
            listViewItem1.StateImageIndex = 0;
            listViewItem2.StateImageIndex = 0;
            listViewItem3.StateImageIndex = 0;
            listViewItem4.StateImageIndex = 0;
            this.lvFVSVariablesPrePostValues.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3,
            listViewItem4});
            this.lvFVSVariablesPrePostValues.Location = new System.Drawing.Point(8, 48);
            this.lvFVSVariablesPrePostValues.MultiSelect = false;
            this.lvFVSVariablesPrePostValues.Name = "lvFVSVariablesPrePostValues";
            this.lvFVSVariablesPrePostValues.Size = new System.Drawing.Size(824, 144);
            this.lvFVSVariablesPrePostValues.TabIndex = 65;
            this.lvFVSVariablesPrePostValues.UseCompatibleStateImageBehavior = false;
            this.lvFVSVariablesPrePostValues.View = System.Windows.Forms.View.Details;
            this.lvFVSVariablesPrePostValues.SelectedIndexChanged += new System.EventHandler(this.lvFVSVariablesPrePostValues_SelectedIndexChanged);
            this.lvFVSVariablesPrePostValues.DoubleClick += new System.EventHandler(this.lvFVSVariablesPrePostValues_DoubleClick);
            this.lvFVSVariablesPrePostValues.MouseUp += new System.Windows.Forms.MouseEventHandler(this.lvFVSVariablesPrePostValues_MouseUp);
            // 
            // lvColVariableNull
            // 
            this.lvColVariableNull.Text = "";
            this.lvColVariableNull.Width = 2;
            // 
            // lvColVariableId
            // 
            this.lvColVariableId.Text = "Variable";
            this.lvColVariableId.Width = 153;
            // 
            // lvColPreVariableName
            // 
            this.lvColPreVariableName.Text = "Pre-Treatment Variable Name";
            this.lvColPreVariableName.Width = 254;
            // 
            // lvColPostVariableName
            // 
            this.lvColPostVariableName.Text = "Post-Treatment Variable Name";
            this.lvColPostVariableName.Width = 272;
            // 
            // lvColVariableImprovement
            // 
            this.lvColVariableImprovement.Text = "Improvement Expression Defined";
            this.lvColVariableImprovement.Width = 217;
            // 
            // lvColVariableDisimprovement
            // 
            this.lvColVariableDisimprovement.Text = "Disimprovement Expression Defined";
            this.lvColVariableDisimprovement.Width = 112;
            // 
            // lvColVariableEffective
            // 
            this.lvColVariableEffective.Text = "Effective Expression Defined";
            // 
            // grpboxFVSVariablesPrePostValuesButtons
            // 
            this.grpboxFVSVariablesPrePostValuesButtons.Controls.Add(this.btnFVSVariablesPrePostValuesButtonsClearAll);
            this.grpboxFVSVariablesPrePostValuesButtons.Controls.Add(this.btnFVSVariablesPrePostValuesButtonsClear);
            this.grpboxFVSVariablesPrePostValuesButtons.Controls.Add(this.btnFVSVariablesPrePostValuesButtonsEdit);
            this.grpboxFVSVariablesPrePostValuesButtons.Location = new System.Drawing.Point(168, 200);
            this.grpboxFVSVariablesPrePostValuesButtons.Name = "grpboxFVSVariablesPrePostValuesButtons";
            this.grpboxFVSVariablesPrePostValuesButtons.Size = new System.Drawing.Size(512, 48);
            this.grpboxFVSVariablesPrePostValuesButtons.TabIndex = 61;
            this.grpboxFVSVariablesPrePostValuesButtons.TabStop = false;
            // 
            // btnFVSVariablesPrePostValuesButtonsClearAll
            // 
            this.btnFVSVariablesPrePostValuesButtonsClearAll.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostValuesButtonsClearAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesPrePostValuesButtonsClearAll.Location = new System.Drawing.Point(16, 16);
            this.btnFVSVariablesPrePostValuesButtonsClearAll.Name = "btnFVSVariablesPrePostValuesButtonsClearAll";
            this.btnFVSVariablesPrePostValuesButtonsClearAll.Size = new System.Drawing.Size(160, 24);
            this.btnFVSVariablesPrePostValuesButtonsClearAll.TabIndex = 52;
            this.btnFVSVariablesPrePostValuesButtonsClearAll.Text = "Clear All Variables";
            this.btnFVSVariablesPrePostValuesButtonsClearAll.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostValuesButtonsClearAll.Click += new System.EventHandler(this.btnFVSVariablesPrePostValuesButtonsClearAll_Click);
            // 
            // btnFVSVariablesPrePostValuesButtonsClear
            // 
            this.btnFVSVariablesPrePostValuesButtonsClear.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostValuesButtonsClear.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesPrePostValuesButtonsClear.Location = new System.Drawing.Point(192, 16);
            this.btnFVSVariablesPrePostValuesButtonsClear.Name = "btnFVSVariablesPrePostValuesButtonsClear";
            this.btnFVSVariablesPrePostValuesButtonsClear.Size = new System.Drawing.Size(152, 24);
            this.btnFVSVariablesPrePostValuesButtonsClear.TabIndex = 51;
            this.btnFVSVariablesPrePostValuesButtonsClear.Text = "Clear Variable";
            this.btnFVSVariablesPrePostValuesButtonsClear.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostValuesButtonsClear.Click += new System.EventHandler(this.btnFVSVariablesPrePostValuesButtonsClear_Click);
            // 
            // btnFVSVariablesPrePostValuesButtonsEdit
            // 
            this.btnFVSVariablesPrePostValuesButtonsEdit.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostValuesButtonsEdit.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesPrePostValuesButtonsEdit.Location = new System.Drawing.Point(368, 16);
            this.btnFVSVariablesPrePostValuesButtonsEdit.Name = "btnFVSVariablesPrePostValuesButtonsEdit";
            this.btnFVSVariablesPrePostValuesButtonsEdit.Size = new System.Drawing.Size(128, 24);
            this.btnFVSVariablesPrePostValuesButtonsEdit.TabIndex = 36;
            this.btnFVSVariablesPrePostValuesButtonsEdit.Text = "Edit Variable";
            this.btnFVSVariablesPrePostValuesButtonsEdit.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostValuesButtonsEdit.Click += new System.EventHandler(this.btnFVSVariablesPrePostValuesButtonsEdit_Click);
            // 
            // grpboxFVSVariablesPrePostOverall
            // 
            this.grpboxFVSVariablesPrePostOverall.Controls.Add(this.btnFVSVariablesPrePostOverall);
            this.grpboxFVSVariablesPrePostOverall.Location = new System.Drawing.Point(8, 288);
            this.grpboxFVSVariablesPrePostOverall.Name = "grpboxFVSVariablesPrePostOverall";
            this.grpboxFVSVariablesPrePostOverall.Size = new System.Drawing.Size(832, 72);
            this.grpboxFVSVariablesPrePostOverall.TabIndex = 68;
            this.grpboxFVSVariablesPrePostOverall.TabStop = false;
            this.grpboxFVSVariablesPrePostOverall.Text = "Step 2: Define Overall Effectiveness";
            // 
            // btnFVSVariablesPrePostOverall
            // 
            this.btnFVSVariablesPrePostOverall.Enabled = false;
            this.btnFVSVariablesPrePostOverall.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesPrePostOverall.Location = new System.Drawing.Point(14, 24);
            this.btnFVSVariablesPrePostOverall.Name = "btnFVSVariablesPrePostOverall";
            this.btnFVSVariablesPrePostOverall.Size = new System.Drawing.Size(810, 32);
            this.btnFVSVariablesPrePostOverall.TabIndex = 0;
            this.btnFVSVariablesPrePostOverall.Text = "Overall Effective Expression";
            this.btnFVSVariablesPrePostOverall.Click += new System.EventHandler(this.btnFVSVariablesPrePostOverall_Click);
            // 
            // grpboxFVSVariablesPrePostAudit
            // 
            this.grpboxFVSVariablesPrePostAudit.Controls.Add(this.btnFVSVariablesPrePostAudit);
            this.grpboxFVSVariablesPrePostAudit.Location = new System.Drawing.Point(8, 368);
            this.grpboxFVSVariablesPrePostAudit.Name = "grpboxFVSVariablesPrePostAudit";
            this.grpboxFVSVariablesPrePostAudit.Size = new System.Drawing.Size(832, 72);
            this.grpboxFVSVariablesPrePostAudit.TabIndex = 69;
            this.grpboxFVSVariablesPrePostAudit.TabStop = false;
            this.grpboxFVSVariablesPrePostAudit.Text = "Step 3: Audit";
            // 
            // btnFVSVariablesPrePostAudit
            // 
            this.btnFVSVariablesPrePostAudit.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesPrePostAudit.Location = new System.Drawing.Point(16, 24);
            this.btnFVSVariablesPrePostAudit.Name = "btnFVSVariablesPrePostAudit";
            this.btnFVSVariablesPrePostAudit.Size = new System.Drawing.Size(800, 32);
            this.btnFVSVariablesPrePostAudit.TabIndex = 0;
            this.btnFVSVariablesPrePostAudit.Text = "Audit";
            this.btnFVSVariablesPrePostAudit.Click += new System.EventHandler(this.btnFVSVariablesPrePostAudit_Click);
            // 
            // grpboxFVSVariablesPrePostExpression
            // 
            this.grpboxFVSVariablesPrePostExpression.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxFVSVariablesPrePostExpression.Controls.Add(this.pnlFVSVariablesPrePostExpression);
            this.grpboxFVSVariablesPrePostExpression.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxFVSVariablesPrePostExpression.ForeColor = System.Drawing.Color.Black;
            this.grpboxFVSVariablesPrePostExpression.Location = new System.Drawing.Point(8, 1040);
            this.grpboxFVSVariablesPrePostExpression.Name = "grpboxFVSVariablesPrePostExpression";
            this.grpboxFVSVariablesPrePostExpression.Size = new System.Drawing.Size(856, 448);
            this.grpboxFVSVariablesPrePostExpression.TabIndex = 31;
            this.grpboxFVSVariablesPrePostExpression.TabStop = false;
            this.grpboxFVSVariablesPrePostExpression.Text = "Expression Builder";
            this.grpboxFVSVariablesPrePostExpression.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePostExpression_Resize);
            this.grpboxFVSVariablesPrePostExpression.VisibleChanged += new System.EventHandler(this.grpboxFVSVariablesPrePostExpression_VisibleChanged);
            // 
            // pnlFVSVariablesPrePostExpression
            // 
            this.pnlFVSVariablesPrePostExpression.AutoScroll = true;
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.grpboxFVSVariablesPrePostExpressionSelectedVariables);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.groupBox2);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.lblFVSVariablesPrePostExpressionVariable1);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.lblFVSVariablesPrePostExpressionVariable2);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.lblFVSVariablesPrePostExpressionVariable3);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.lblFVSVariablesPrePostExpressionVariable4);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.lblFVSVariablesPrePostExpression);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.txtExpression);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.grpboxExpressionsBtns);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.btnFVSVariablesPrePostExpressionDone);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.btnFVSVariablesPrePostExpressionCancel);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.btnFVSVariablesPrePostExpressionPrevious);
            this.pnlFVSVariablesPrePostExpression.Controls.Add(this.btnFVSVariablesPrePostExpressionNext);
            this.pnlFVSVariablesPrePostExpression.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFVSVariablesPrePostExpression.Location = new System.Drawing.Point(3, 18);
            this.pnlFVSVariablesPrePostExpression.Name = "pnlFVSVariablesPrePostExpression";
            this.pnlFVSVariablesPrePostExpression.Size = new System.Drawing.Size(850, 427);
            this.pnlFVSVariablesPrePostExpression.TabIndex = 71;
            this.pnlFVSVariablesPrePostExpression.Resize += new System.EventHandler(this.pnlFVSVariablesPrePostExpression_Resize);
            // 
            // grpboxFVSVariablesPrePostExpressionSelectedVariables
            // 
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.Controls.Add(this.lstFVSVariablesPrePostExpressionSelectedVariables);
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.Location = new System.Drawing.Point(8, 8);
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.Name = "grpboxFVSVariablesPrePostExpressionSelectedVariables";
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.Size = new System.Drawing.Size(216, 152);
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.TabIndex = 65;
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.TabStop = false;
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.Text = "Available Variables(s)";
            // 
            // lstFVSVariablesPrePostExpressionSelectedVariables
            // 
            this.lstFVSVariablesPrePostExpressionSelectedVariables.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstFVSVariablesPrePostExpressionSelectedVariables.HorizontalScrollbar = true;
            this.lstFVSVariablesPrePostExpressionSelectedVariables.ItemHeight = 16;
            this.lstFVSVariablesPrePostExpressionSelectedVariables.Location = new System.Drawing.Point(3, 18);
            this.lstFVSVariablesPrePostExpressionSelectedVariables.Name = "lstFVSVariablesPrePostExpressionSelectedVariables";
            this.lstFVSVariablesPrePostExpressionSelectedVariables.Size = new System.Drawing.Size(210, 131);
            this.lstFVSVariablesPrePostExpressionSelectedVariables.TabIndex = 66;
            this.lstFVSVariablesPrePostExpressionSelectedVariables.DoubleClick += new System.EventHandler(this.lstFVSVariablesPrePostExpressionSelectedVariables_DoubleClick);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnFFEEqual);
            this.groupBox2.Controls.Add(this.btnFFENote);
            this.groupBox2.Controls.Add(this.btnY);
            this.groupBox2.Controls.Add(this.btnN);
            this.groupBox2.Controls.Add(this.btnFFENotEqual);
            this.groupBox2.Controls.Add(this.btnLeftBracket);
            this.groupBox2.Controls.Add(this.btnRightBracket);
            this.groupBox2.Controls.Add(this.btnFFEMoreThan);
            this.groupBox2.Controls.Add(this.btnFFEMoreThanOrEqualTo);
            this.groupBox2.Controls.Add(this.btnFFELessThan);
            this.groupBox2.Controls.Add(this.btnLessThanOrEqualTo);
            this.groupBox2.Controls.Add(this.btnFFEOr);
            this.groupBox2.Controls.Add(this.btnFFEAnd);
            this.groupBox2.Location = new System.Drawing.Point(232, 8);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(216, 152);
            this.groupBox2.TabIndex = 66;
            this.groupBox2.TabStop = false;
            // 
            // btnFFEEqual
            // 
            this.btnFFEEqual.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFEEqual.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFEEqual.Location = new System.Drawing.Point(16, 24);
            this.btnFFEEqual.Name = "btnFFEEqual";
            this.btnFFEEqual.Size = new System.Drawing.Size(56, 24);
            this.btnFFEEqual.TabIndex = 39;
            this.btnFFEEqual.Text = "Equal";
            this.btnFFEEqual.UseVisualStyleBackColor = false;
            this.btnFFEEqual.Click += new System.EventHandler(this.btnFFEEqual_Click);
            // 
            // btnFFENote
            // 
            this.btnFFENote.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFENote.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFENote.Location = new System.Drawing.Point(72, 24);
            this.btnFFENote.Name = "btnFFENote";
            this.btnFFENote.Size = new System.Drawing.Size(40, 24);
            this.btnFFENote.TabIndex = 50;
            this.btnFFENote.Text = "NOT";
            this.btnFFENote.UseVisualStyleBackColor = false;
            this.btnFFENote.Click += new System.EventHandler(this.btnFFENote_Click);
            // 
            // btnY
            // 
            this.btnY.BackColor = System.Drawing.SystemColors.Control;
            this.btnY.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnY.Location = new System.Drawing.Point(112, 24);
            this.btnY.Name = "btnY";
            this.btnY.Size = new System.Drawing.Size(32, 24);
            this.btnY.TabIndex = 52;
            this.btnY.Text = "\'Y\'";
            this.btnY.UseVisualStyleBackColor = false;
            this.btnY.Click += new System.EventHandler(this.btnY_Click);
            // 
            // btnN
            // 
            this.btnN.BackColor = System.Drawing.SystemColors.Control;
            this.btnN.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnN.Location = new System.Drawing.Point(144, 24);
            this.btnN.Name = "btnN";
            this.btnN.Size = new System.Drawing.Size(32, 24);
            this.btnN.TabIndex = 53;
            this.btnN.Text = "\'N\'";
            this.btnN.UseVisualStyleBackColor = false;
            this.btnN.Click += new System.EventHandler(this.btnN_Click);
            // 
            // btnFFENotEqual
            // 
            this.btnFFENotEqual.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFENotEqual.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFENotEqual.Location = new System.Drawing.Point(16, 48);
            this.btnFFENotEqual.Name = "btnFFENotEqual";
            this.btnFFENotEqual.Size = new System.Drawing.Size(56, 24);
            this.btnFFENotEqual.TabIndex = 40;
            this.btnFFENotEqual.Text = "Not Equal";
            this.btnFFENotEqual.UseVisualStyleBackColor = false;
            this.btnFFENotEqual.Click += new System.EventHandler(this.btnFFENotEqual_Click);
            // 
            // btnLeftBracket
            // 
            this.btnLeftBracket.BackColor = System.Drawing.SystemColors.Control;
            this.btnLeftBracket.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLeftBracket.Location = new System.Drawing.Point(72, 48);
            this.btnLeftBracket.Name = "btnLeftBracket";
            this.btnLeftBracket.Size = new System.Drawing.Size(40, 24);
            this.btnLeftBracket.TabIndex = 45;
            this.btnLeftBracket.Text = "(";
            this.btnLeftBracket.UseVisualStyleBackColor = false;
            this.btnLeftBracket.Click += new System.EventHandler(this.btnLeftBracket_Click);
            // 
            // btnRightBracket
            // 
            this.btnRightBracket.BackColor = System.Drawing.SystemColors.Control;
            this.btnRightBracket.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRightBracket.Location = new System.Drawing.Point(112, 48);
            this.btnRightBracket.Name = "btnRightBracket";
            this.btnRightBracket.Size = new System.Drawing.Size(32, 24);
            this.btnRightBracket.TabIndex = 46;
            this.btnRightBracket.Text = ")";
            this.btnRightBracket.UseVisualStyleBackColor = false;
            this.btnRightBracket.Click += new System.EventHandler(this.btnRightBracket_Click);
            // 
            // btnFFEMoreThan
            // 
            this.btnFFEMoreThan.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFEMoreThan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFEMoreThan.Location = new System.Drawing.Point(16, 72);
            this.btnFFEMoreThan.Name = "btnFFEMoreThan";
            this.btnFFEMoreThan.Size = new System.Drawing.Size(64, 24);
            this.btnFFEMoreThan.TabIndex = 41;
            this.btnFFEMoreThan.Text = "More Than";
            this.btnFFEMoreThan.UseVisualStyleBackColor = false;
            this.btnFFEMoreThan.Click += new System.EventHandler(this.btnFFEMoreThan_Click);
            // 
            // btnFFEMoreThanOrEqualTo
            // 
            this.btnFFEMoreThanOrEqualTo.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFEMoreThanOrEqualTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFEMoreThanOrEqualTo.Location = new System.Drawing.Point(80, 72);
            this.btnFFEMoreThanOrEqualTo.Name = "btnFFEMoreThanOrEqualTo";
            this.btnFFEMoreThanOrEqualTo.Size = new System.Drawing.Size(128, 24);
            this.btnFFEMoreThanOrEqualTo.TabIndex = 43;
            this.btnFFEMoreThanOrEqualTo.Text = "More Than Or Equal To";
            this.btnFFEMoreThanOrEqualTo.UseVisualStyleBackColor = false;
            this.btnFFEMoreThanOrEqualTo.Click += new System.EventHandler(this.btnFFEMoreThanOrEqualTo_Click);
            // 
            // btnFFELessThan
            // 
            this.btnFFELessThan.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFELessThan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFELessThan.Location = new System.Drawing.Point(16, 96);
            this.btnFFELessThan.Name = "btnFFELessThan";
            this.btnFFELessThan.Size = new System.Drawing.Size(64, 24);
            this.btnFFELessThan.TabIndex = 42;
            this.btnFFELessThan.Text = "Less Than";
            this.btnFFELessThan.UseVisualStyleBackColor = false;
            this.btnFFELessThan.Click += new System.EventHandler(this.btnFFELessThan_Click);
            // 
            // btnLessThanOrEqualTo
            // 
            this.btnLessThanOrEqualTo.BackColor = System.Drawing.SystemColors.Control;
            this.btnLessThanOrEqualTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLessThanOrEqualTo.Location = new System.Drawing.Point(80, 96);
            this.btnLessThanOrEqualTo.Name = "btnLessThanOrEqualTo";
            this.btnLessThanOrEqualTo.Size = new System.Drawing.Size(128, 24);
            this.btnLessThanOrEqualTo.TabIndex = 44;
            this.btnLessThanOrEqualTo.Text = "Less Than Or Equal To";
            this.btnLessThanOrEqualTo.UseVisualStyleBackColor = false;
            this.btnLessThanOrEqualTo.Click += new System.EventHandler(this.btnLessThanOrEqualTo_Click);
            // 
            // btnFFEOr
            // 
            this.btnFFEOr.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFEOr.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFEOr.Location = new System.Drawing.Point(16, 120);
            this.btnFFEOr.Name = "btnFFEOr";
            this.btnFFEOr.Size = new System.Drawing.Size(40, 24);
            this.btnFFEOr.TabIndex = 48;
            this.btnFFEOr.Text = "OR";
            this.btnFFEOr.UseVisualStyleBackColor = false;
            this.btnFFEOr.Click += new System.EventHandler(this.btnFFEOr_Click);
            // 
            // btnFFEAnd
            // 
            this.btnFFEAnd.BackColor = System.Drawing.SystemColors.Control;
            this.btnFFEAnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFFEAnd.Location = new System.Drawing.Point(56, 120);
            this.btnFFEAnd.Name = "btnFFEAnd";
            this.btnFFEAnd.Size = new System.Drawing.Size(40, 24);
            this.btnFFEAnd.TabIndex = 47;
            this.btnFFEAnd.Text = "AND";
            this.btnFFEAnd.UseVisualStyleBackColor = false;
            this.btnFFEAnd.Click += new System.EventHandler(this.btnFFEAnd_Click);
            // 
            // lblFVSVariablesPrePostExpressionVariable1
            // 
            this.lblFVSVariablesPrePostExpressionVariable1.Location = new System.Drawing.Point(464, 16);
            this.lblFVSVariablesPrePostExpressionVariable1.Name = "lblFVSVariablesPrePostExpressionVariable1";
            this.lblFVSVariablesPrePostExpressionVariable1.Size = new System.Drawing.Size(368, 24);
            this.lblFVSVariablesPrePostExpressionVariable1.TabIndex = 67;
            this.lblFVSVariablesPrePostExpressionVariable1.Text = "Variable 1:";
            // 
            // lblFVSVariablesPrePostExpressionVariable2
            // 
            this.lblFVSVariablesPrePostExpressionVariable2.Location = new System.Drawing.Point(464, 40);
            this.lblFVSVariablesPrePostExpressionVariable2.Name = "lblFVSVariablesPrePostExpressionVariable2";
            this.lblFVSVariablesPrePostExpressionVariable2.Size = new System.Drawing.Size(368, 24);
            this.lblFVSVariablesPrePostExpressionVariable2.TabIndex = 68;
            this.lblFVSVariablesPrePostExpressionVariable2.Text = "Variable 2:";
            // 
            // lblFVSVariablesPrePostExpressionVariable3
            // 
            this.lblFVSVariablesPrePostExpressionVariable3.Location = new System.Drawing.Point(464, 64);
            this.lblFVSVariablesPrePostExpressionVariable3.Name = "lblFVSVariablesPrePostExpressionVariable3";
            this.lblFVSVariablesPrePostExpressionVariable3.Size = new System.Drawing.Size(368, 24);
            this.lblFVSVariablesPrePostExpressionVariable3.TabIndex = 69;
            this.lblFVSVariablesPrePostExpressionVariable3.Text = "Variable 3:";
            // 
            // lblFVSVariablesPrePostExpressionVariable4
            // 
            this.lblFVSVariablesPrePostExpressionVariable4.Location = new System.Drawing.Point(464, 88);
            this.lblFVSVariablesPrePostExpressionVariable4.Name = "lblFVSVariablesPrePostExpressionVariable4";
            this.lblFVSVariablesPrePostExpressionVariable4.Size = new System.Drawing.Size(368, 24);
            this.lblFVSVariablesPrePostExpressionVariable4.TabIndex = 70;
            this.lblFVSVariablesPrePostExpressionVariable4.Text = "Variable 4:";
            // 
            // lblFVSVariablesPrePostExpression
            // 
            this.lblFVSVariablesPrePostExpression.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFVSVariablesPrePostExpression.Location = new System.Drawing.Point(8, 183);
            this.lblFVSVariablesPrePostExpression.Name = "lblFVSVariablesPrePostExpression";
            this.lblFVSVariablesPrePostExpression.Size = new System.Drawing.Size(816, 16);
            this.lblFVSVariablesPrePostExpression.TabIndex = 59;
            this.lblFVSVariablesPrePostExpression.Text = "Expression To Test If Treatment Resulted In An Improvement";
            // 
            // txtExpression
            // 
            this.txtExpression.Location = new System.Drawing.Point(8, 202);
            this.txtExpression.Multiline = true;
            this.txtExpression.Name = "txtExpression";
            this.txtExpression.Size = new System.Drawing.Size(816, 120);
            this.txtExpression.TabIndex = 32;
            // 
            // grpboxExpressionsBtns
            // 
            this.grpboxExpressionsBtns.Controls.Add(this.btnFVSVariablesPrePostExpressionDefault);
            this.grpboxExpressionsBtns.Controls.Add(this.btnFVSVariablesPrePostExpressionPrev);
            this.grpboxExpressionsBtns.Controls.Add(this.btnFVSVariablesPrePostExpressionClear);
            this.grpboxExpressionsBtns.Controls.Add(this.btnFVSVariablesPrePostExpressionTest);
            this.grpboxExpressionsBtns.Location = new System.Drawing.Point(120, 328);
            this.grpboxExpressionsBtns.Name = "grpboxExpressionsBtns";
            this.grpboxExpressionsBtns.Size = new System.Drawing.Size(632, 48);
            this.grpboxExpressionsBtns.TabIndex = 61;
            this.grpboxExpressionsBtns.TabStop = false;
            // 
            // btnFVSVariablesPrePostExpressionDefault
            // 
            this.btnFVSVariablesPrePostExpressionDefault.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostExpressionDefault.Location = new System.Drawing.Point(16, 16);
            this.btnFVSVariablesPrePostExpressionDefault.Name = "btnFVSVariablesPrePostExpressionDefault";
            this.btnFVSVariablesPrePostExpressionDefault.Size = new System.Drawing.Size(128, 24);
            this.btnFVSVariablesPrePostExpressionDefault.TabIndex = 52;
            this.btnFVSVariablesPrePostExpressionDefault.Text = "Default Expression";
            this.btnFVSVariablesPrePostExpressionDefault.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostExpressionDefault.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionDefault_Click);
            // 
            // btnFVSVariablesPrePostExpressionPrev
            // 
            this.btnFVSVariablesPrePostExpressionPrev.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostExpressionPrev.Location = new System.Drawing.Point(152, 16);
            this.btnFVSVariablesPrePostExpressionPrev.Name = "btnFVSVariablesPrePostExpressionPrev";
            this.btnFVSVariablesPrePostExpressionPrev.Size = new System.Drawing.Size(152, 24);
            this.btnFVSVariablesPrePostExpressionPrev.TabIndex = 51;
            this.btnFVSVariablesPrePostExpressionPrev.Text = "Previous Expressions";
            this.btnFVSVariablesPrePostExpressionPrev.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostExpressionPrev.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionPrev_Click);
            // 
            // btnFVSVariablesPrePostExpressionClear
            // 
            this.btnFVSVariablesPrePostExpressionClear.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostExpressionClear.Location = new System.Drawing.Point(312, 15);
            this.btnFVSVariablesPrePostExpressionClear.Name = "btnFVSVariablesPrePostExpressionClear";
            this.btnFVSVariablesPrePostExpressionClear.Size = new System.Drawing.Size(128, 24);
            this.btnFVSVariablesPrePostExpressionClear.TabIndex = 36;
            this.btnFVSVariablesPrePostExpressionClear.Text = "Clear Expression";
            this.btnFVSVariablesPrePostExpressionClear.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostExpressionClear.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionClear_Click);
            // 
            // btnFVSVariablesPrePostExpressionTest
            // 
            this.btnFVSVariablesPrePostExpressionTest.BackColor = System.Drawing.SystemColors.Control;
            this.btnFVSVariablesPrePostExpressionTest.Location = new System.Drawing.Point(544, 16);
            this.btnFVSVariablesPrePostExpressionTest.Name = "btnFVSVariablesPrePostExpressionTest";
            this.btnFVSVariablesPrePostExpressionTest.Size = new System.Drawing.Size(72, 24);
            this.btnFVSVariablesPrePostExpressionTest.TabIndex = 33;
            this.btnFVSVariablesPrePostExpressionTest.Text = "Test";
            this.btnFVSVariablesPrePostExpressionTest.UseVisualStyleBackColor = false;
            this.btnFVSVariablesPrePostExpressionTest.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionTest_Click);
            // 
            // btnFVSVariablesPrePostExpressionDone
            // 
            this.btnFVSVariablesPrePostExpressionDone.Location = new System.Drawing.Point(352, 384);
            this.btnFVSVariablesPrePostExpressionDone.Name = "btnFVSVariablesPrePostExpressionDone";
            this.btnFVSVariablesPrePostExpressionDone.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostExpressionDone.TabIndex = 63;
            this.btnFVSVariablesPrePostExpressionDone.Text = "Done";
            this.btnFVSVariablesPrePostExpressionDone.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionDone_Click);
            // 
            // btnFVSVariablesPrePostExpressionCancel
            // 
            this.btnFVSVariablesPrePostExpressionCancel.Location = new System.Drawing.Point(440, 384);
            this.btnFVSVariablesPrePostExpressionCancel.Name = "btnFVSVariablesPrePostExpressionCancel";
            this.btnFVSVariablesPrePostExpressionCancel.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostExpressionCancel.TabIndex = 62;
            this.btnFVSVariablesPrePostExpressionCancel.Text = "Cancel";
            this.btnFVSVariablesPrePostExpressionCancel.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionCancel_Click);
            // 
            // btnFVSVariablesPrePostExpressionPrevious
            // 
            this.btnFVSVariablesPrePostExpressionPrevious.ForeColor = System.Drawing.Color.Black;
            this.btnFVSVariablesPrePostExpressionPrevious.Location = new System.Drawing.Point(528, 384);
            this.btnFVSVariablesPrePostExpressionPrevious.Name = "btnFVSVariablesPrePostExpressionPrevious";
            this.btnFVSVariablesPrePostExpressionPrevious.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostExpressionPrevious.TabIndex = 8;
            this.btnFVSVariablesPrePostExpressionPrevious.Text = "<--Previous";
            this.btnFVSVariablesPrePostExpressionPrevious.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionPrevious_Click);
            // 
            // btnFVSVariablesPrePostExpressionNext
            // 
            this.btnFVSVariablesPrePostExpressionNext.Location = new System.Drawing.Point(616, 384);
            this.btnFVSVariablesPrePostExpressionNext.Name = "btnFVSVariablesPrePostExpressionNext";
            this.btnFVSVariablesPrePostExpressionNext.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostExpressionNext.TabIndex = 64;
            this.btnFVSVariablesPrePostExpressionNext.Text = "Next-->";
            this.btnFVSVariablesPrePostExpressionNext.Click += new System.EventHandler(this.btnFVSVariablesPrePostExpressionNext_Click);
            // 
            // grpboxFVSVariablesPrePostVariable
            // 
            this.grpboxFVSVariablesPrePostVariable.BackColor = System.Drawing.SystemColors.Control;
            this.grpboxFVSVariablesPrePostVariable.Controls.Add(this.pnlFVSVariablesPrePostVariable);
            this.grpboxFVSVariablesPrePostVariable.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxFVSVariablesPrePostVariable.ForeColor = System.Drawing.Color.Black;
            this.grpboxFVSVariablesPrePostVariable.Location = new System.Drawing.Point(8, 576);
            this.grpboxFVSVariablesPrePostVariable.Name = "grpboxFVSVariablesPrePostVariable";
            this.grpboxFVSVariablesPrePostVariable.Size = new System.Drawing.Size(856, 448);
            this.grpboxFVSVariablesPrePostVariable.TabIndex = 30;
            this.grpboxFVSVariablesPrePostVariable.TabStop = false;
            this.grpboxFVSVariablesPrePostVariable.Text = "Attribute";
            this.grpboxFVSVariablesPrePostVariable.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePostVariable_Resize);
            // 
            // pnlFVSVariablesPrePostVariable
            // 
            this.pnlFVSVariablesPrePostVariable.AutoScroll = true;
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpboxFVSVariablesPrePostVariableValues);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpboxFVSVariablesPrePostVariablePreSelected);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.grpboxFVSVariablesPrePostVariablePostSelected);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnFVSVariablesPrePostVariableClearAll);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnFVSVariablesPrePostVariableDone);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnFVSVariablesPrePostVariableCancel);
            this.pnlFVSVariablesPrePostVariable.Controls.Add(this.btnFVSVariablesPrePostVariableNext);
            this.pnlFVSVariablesPrePostVariable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlFVSVariablesPrePostVariable.Location = new System.Drawing.Point(3, 18);
            this.pnlFVSVariablesPrePostVariable.Name = "pnlFVSVariablesPrePostVariable";
            this.pnlFVSVariablesPrePostVariable.Size = new System.Drawing.Size(850, 427);
            this.pnlFVSVariablesPrePostVariable.TabIndex = 12;
            this.pnlFVSVariablesPrePostVariable.Resize += new System.EventHandler(this.pnlFVSVariablesPrePostVariable_Resize);
            // 
            // grpboxFVSVariablesPrePostVariableValues
            // 
            this.grpboxFVSVariablesPrePostVariableValues.Controls.Add(this.lstFVSFieldsList);
            this.grpboxFVSVariablesPrePostVariableValues.Controls.Add(this.lstFVSTablesList);
            this.grpboxFVSVariablesPrePostVariableValues.Controls.Add(this.txtEffVarDescr);
            this.grpboxFVSVariablesPrePostVariableValues.Controls.Add(this.lblEffVarDescr);
            this.grpboxFVSVariablesPrePostVariableValues.Controls.Add(this.btnFVSVariablesPrePostVariableValue);
            this.grpboxFVSVariablesPrePostVariableValues.Location = new System.Drawing.Point(8, 16);
            this.grpboxFVSVariablesPrePostVariableValues.Name = "grpboxFVSVariablesPrePostVariableValues";
            this.grpboxFVSVariablesPrePostVariableValues.Size = new System.Drawing.Size(816, 216);
            this.grpboxFVSVariablesPrePostVariableValues.TabIndex = 0;
            this.grpboxFVSVariablesPrePostVariableValues.TabStop = false;
            this.grpboxFVSVariablesPrePostVariableValues.Text = "Stand Attribute";
            this.grpboxFVSVariablesPrePostVariableValues.Resize += new System.EventHandler(this.grpboxFVSVariablesPrePostVariableValues_Resize);
            // 
            // lstFVSFieldsList
            // 
            this.lstFVSFieldsList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSFieldsList.ItemHeight = 16;
            this.lstFVSFieldsList.Location = new System.Drawing.Point(252, 21);
            this.lstFVSFieldsList.Name = "lstFVSFieldsList";
            this.lstFVSFieldsList.Size = new System.Drawing.Size(202, 180);
            this.lstFVSFieldsList.Sorted = true;
            this.lstFVSFieldsList.TabIndex = 90;
            this.lstFVSFieldsList.SelectedIndexChanged += new System.EventHandler(this.lstFVSFieldsList_SelectedIndexChanged);
            // 
            // lstFVSTablesList
            // 
            this.lstFVSTablesList.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFVSTablesList.ItemHeight = 16;
            this.lstFVSTablesList.Location = new System.Drawing.Point(11, 21);
            this.lstFVSTablesList.Name = "lstFVSTablesList";
            this.lstFVSTablesList.Size = new System.Drawing.Size(202, 180);
            this.lstFVSTablesList.TabIndex = 89;
            this.lstFVSTablesList.SelectedIndexChanged += new System.EventHandler(this.lstFVSTablesList_SelectedIndexChanged);
            // 
            // txtEffVarDescr
            // 
            this.txtEffVarDescr.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEffVarDescr.Location = new System.Drawing.Point(552, 89);
            this.txtEffVarDescr.Multiline = true;
            this.txtEffVarDescr.Name = "txtEffVarDescr";
            this.txtEffVarDescr.ReadOnly = true;
            this.txtEffVarDescr.Size = new System.Drawing.Size(258, 75);
            this.txtEffVarDescr.TabIndex = 88;
            this.txtEffVarDescr.Visible = false;
            // 
            // lblEffVarDescr
            // 
            this.lblEffVarDescr.Location = new System.Drawing.Point(471, 92);
            this.lblEffVarDescr.Name = "lblEffVarDescr";
            this.lblEffVarDescr.Size = new System.Drawing.Size(80, 24);
            this.lblEffVarDescr.TabIndex = 87;
            this.lblEffVarDescr.Text = "Description:";
            this.lblEffVarDescr.Visible = false;
            // 
            // btnFVSVariablesPrePostVariableValue
            // 
            this.btnFVSVariablesPrePostVariableValue.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFVSVariablesPrePostVariableValue.Location = new System.Drawing.Point(476, 21);
            this.btnFVSVariablesPrePostVariableValue.Name = "btnFVSVariablesPrePostVariableValue";
            this.btnFVSVariablesPrePostVariableValue.Size = new System.Drawing.Size(143, 57);
            this.btnFVSVariablesPrePostVariableValue.TabIndex = 1;
            this.btnFVSVariablesPrePostVariableValue.Text = "Select";
            this.btnFVSVariablesPrePostVariableValue.Click += new System.EventHandler(this.btnFVSVariablesPrePostVariableValue_Click);
            // 
            // grpboxFVSVariablesPrePostVariablePreSelected
            // 
            this.grpboxFVSVariablesPrePostVariablePreSelected.Controls.Add(this.lblFVSVariablesPrePostVariablePreSelected);
            this.grpboxFVSVariablesPrePostVariablePreSelected.Location = new System.Drawing.Point(24, 248);
            this.grpboxFVSVariablesPrePostVariablePreSelected.Name = "grpboxFVSVariablesPrePostVariablePreSelected";
            this.grpboxFVSVariablesPrePostVariablePreSelected.Size = new System.Drawing.Size(664, 51);
            this.grpboxFVSVariablesPrePostVariablePreSelected.TabIndex = 3;
            this.grpboxFVSVariablesPrePostVariablePreSelected.TabStop = false;
            this.grpboxFVSVariablesPrePostVariablePreSelected.Text = "Pre-Treatment Stand Attribute";
            // 
            // lblFVSVariablesPrePostVariablePreSelected
            // 
            this.lblFVSVariablesPrePostVariablePreSelected.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFVSVariablesPrePostVariablePreSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFVSVariablesPrePostVariablePreSelected.Location = new System.Drawing.Point(3, 18);
            this.lblFVSVariablesPrePostVariablePreSelected.Name = "lblFVSVariablesPrePostVariablePreSelected";
            this.lblFVSVariablesPrePostVariablePreSelected.Size = new System.Drawing.Size(658, 30);
            this.lblFVSVariablesPrePostVariablePreSelected.TabIndex = 1;
            this.lblFVSVariablesPrePostVariablePreSelected.Text = "Not Defined";
            // 
            // grpboxFVSVariablesPrePostVariablePostSelected
            // 
            this.grpboxFVSVariablesPrePostVariablePostSelected.Controls.Add(this.lblFVSVariablesPrePostVariablePostSelected);
            this.grpboxFVSVariablesPrePostVariablePostSelected.Location = new System.Drawing.Point(24, 312);
            this.grpboxFVSVariablesPrePostVariablePostSelected.Name = "grpboxFVSVariablesPrePostVariablePostSelected";
            this.grpboxFVSVariablesPrePostVariablePostSelected.Size = new System.Drawing.Size(664, 51);
            this.grpboxFVSVariablesPrePostVariablePostSelected.TabIndex = 4;
            this.grpboxFVSVariablesPrePostVariablePostSelected.TabStop = false;
            this.grpboxFVSVariablesPrePostVariablePostSelected.Text = "Post-Treatment Stand Attribute";
            // 
            // lblFVSVariablesPrePostVariablePostSelected
            // 
            this.lblFVSVariablesPrePostVariablePostSelected.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFVSVariablesPrePostVariablePostSelected.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFVSVariablesPrePostVariablePostSelected.Location = new System.Drawing.Point(3, 18);
            this.lblFVSVariablesPrePostVariablePostSelected.Name = "lblFVSVariablesPrePostVariablePostSelected";
            this.lblFVSVariablesPrePostVariablePostSelected.Size = new System.Drawing.Size(658, 30);
            this.lblFVSVariablesPrePostVariablePostSelected.TabIndex = 2;
            this.lblFVSVariablesPrePostVariablePostSelected.Text = "Not Defined";
            // 
            // btnFVSVariablesPrePostVariableClearAll
            // 
            this.btnFVSVariablesPrePostVariableClearAll.Location = new System.Drawing.Point(99, 376);
            this.btnFVSVariablesPrePostVariableClearAll.Name = "btnFVSVariablesPrePostVariableClearAll";
            this.btnFVSVariablesPrePostVariableClearAll.Size = new System.Drawing.Size(72, 40);
            this.btnFVSVariablesPrePostVariableClearAll.TabIndex = 5;
            this.btnFVSVariablesPrePostVariableClearAll.Text = "Clear";
            this.btnFVSVariablesPrePostVariableClearAll.Click += new System.EventHandler(this.btnFVSVariablesPrePostVariableClearAll_Click);
            // 
            // btnFVSVariablesPrePostVariableDone
            // 
            this.btnFVSVariablesPrePostVariableDone.Location = new System.Drawing.Point(352, 376);
            this.btnFVSVariablesPrePostVariableDone.Name = "btnFVSVariablesPrePostVariableDone";
            this.btnFVSVariablesPrePostVariableDone.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostVariableDone.TabIndex = 11;
            this.btnFVSVariablesPrePostVariableDone.Text = "Done";
            this.btnFVSVariablesPrePostVariableDone.Click += new System.EventHandler(this.btnFVSVariablesPrePostVariableDone_Click);
            // 
            // btnFVSVariablesPrePostVariableCancel
            // 
            this.btnFVSVariablesPrePostVariableCancel.Location = new System.Drawing.Point(440, 376);
            this.btnFVSVariablesPrePostVariableCancel.Name = "btnFVSVariablesPrePostVariableCancel";
            this.btnFVSVariablesPrePostVariableCancel.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostVariableCancel.TabIndex = 9;
            this.btnFVSVariablesPrePostVariableCancel.Text = "Cancel";
            this.btnFVSVariablesPrePostVariableCancel.Click += new System.EventHandler(this.btnFVSVariablesPrePostVariableCancel_Click);
            // 
            // btnFVSVariablesPrePostVariableNext
            // 
            this.btnFVSVariablesPrePostVariableNext.Location = new System.Drawing.Point(616, 376);
            this.btnFVSVariablesPrePostVariableNext.Name = "btnFVSVariablesPrePostVariableNext";
            this.btnFVSVariablesPrePostVariableNext.Size = new System.Drawing.Size(88, 40);
            this.btnFVSVariablesPrePostVariableNext.TabIndex = 8;
            this.btnFVSVariablesPrePostVariableNext.Text = "Next-->";
            this.btnFVSVariablesPrePostVariableNext.Click += new System.EventHandler(this.btnFVSVariablesPrePostVariableNext_Click);
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
            this.lblTitle.Text = "Effective Settings";
            // 
            // cmbFVSVariablesPrePost
            // 
            this.cmbFVSVariablesPrePost.Items.AddRange(new object[] {
            "Select Variable 1",
            "Define Expression For What Constitutes Variable 1 Post-Treatment Improvement (Bet" +
                "ter)",
            "Define Expression For What Constitutes Variable 1 Post-Treatment Disimprovement (" +
                "Worse)",
            "Define Expression For What Constitutes Variable 1 Effective Treatment",
            "Select Variable 2",
            "Define Expression For What Constitutes Variable 2 Post-Treatment Improvement (Bet" +
                "ter)",
            "Define Expression For What Constitutes Variable 2 Post-Treatment Disimprovement (" +
                "Worse)",
            "Define Expression For What Constitutes Variable 2 Effective Treatment",
            "Select Variable 3",
            "Define Expression For What Constitutes Variable 3 Post-Treatment Improvement (Bet" +
                "ter)",
            "Define Expression For What Constitutes Variable 3 Post-Treatment Disimprovement (" +
                "Worse)",
            "Define Expression For What Constitutes Variable 3 Effective Treatment",
            "Select Variable 4",
            "Define Expression For What Constitutes Variable 4 Post-Treatment Improvement (Bet" +
                "ter)",
            "Define Expression For What Constitutes Variable 4 Post-Treatment Disimprovement (" +
                "Worse)",
            "Define Expression For What Constitutes Variable 4 Effective Treatment",
            "Define Expression For What Constitutes Overall Effectiveness"});
            this.cmbFVSVariablesPrePost.Location = new System.Drawing.Point(8, 50);
            this.cmbFVSVariablesPrePost.Name = "cmbFVSVariablesPrePost";
            this.cmbFVSVariablesPrePost.Size = new System.Drawing.Size(544, 21);
            this.cmbFVSVariablesPrePost.TabIndex = 5;
            // 
            // btnFVSVariablesPrePostGo
            // 
            this.btnFVSVariablesPrePostGo.ForeColor = System.Drawing.Color.Black;
            this.btnFVSVariablesPrePostGo.Location = new System.Drawing.Point(560, 50);
            this.btnFVSVariablesPrePostGo.Name = "btnFVSVariablesPrePostGo";
            this.btnFVSVariablesPrePostGo.Size = new System.Drawing.Size(32, 24);
            this.btnFVSVariablesPrePostGo.TabIndex = 6;
            this.btnFVSVariablesPrePostGo.Text = "Go";
            this.btnFVSVariablesPrePostGo.Click += new System.EventHandler(this.btnFVSVariablesPrePostGo_Click);
            // 
            // uc_core_scenario_fvs_prepost_variables_effective
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_core_scenario_fvs_prepost_variables_effective";
            this.Size = new System.Drawing.Size(872, 2500);
            this.Resize += new System.EventHandler(this.uc_scenario_fvs_prepost_variables_Resize);
            this.groupBox1.ResumeLayout(false);
            this.grpboxFVSVariablesPrePost.ResumeLayout(false);
            this.pnlFVSVariablesPrePost.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostValues.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostValuesButtons.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostOverall.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostAudit.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostExpression.ResumeLayout(false);
            this.pnlFVSVariablesPrePostExpression.ResumeLayout(false);
            this.pnlFVSVariablesPrePostExpression.PerformLayout();
            this.grpboxFVSVariablesPrePostExpressionSelectedVariables.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.grpboxExpressionsBtns.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostVariable.ResumeLayout(false);
            this.pnlFVSVariablesPrePostVariable.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostVariableValues.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostVariableValues.PerformLayout();
            this.grpboxFVSVariablesPrePostVariablePreSelected.ResumeLayout(false);
            this.grpboxFVSVariablesPrePostVariablePostSelected.ResumeLayout(false);
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

			int x;

			for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
				this.RemoveVariable(x);

			m_oOldVar = new Variables();
			m_oSavVar = new Variables();
			//
			//effective variables
			//
			if (ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Count > 0 &&
				ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0) != null)
			{
				for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
				{
					m_oOldVar.m_strPreVarArray[x] =
						ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0).m_strPreVarArray[x].Trim();
					m_oOldVar.m_strPostVarArray[x] =
						ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0).m_strPostVarArray[x].Trim();
					m_oOldVar.m_strBetterExpr[x] =
						ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0).m_strBetterExpr[x].Trim();
					m_oOldVar.m_strWorseExpr[x] =
						ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0).m_strWorseExpr[x].Trim();
					m_oOldVar.m_strEffectiveExpr[x] =
						ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0).m_strEffectiveExpr[x].Trim();
					this.UpdateListViewVariableItem(x, x + 1, m_oOldVar);

				}
				//
				//overall effective
				//
				m_oOldVar.m_strOverallEffectiveExpr =
						  ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem.m_oEffectiveVariablesItem_Collection.Item(0).m_strOverallEffectiveExpr;

			}

			for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
			{
				if (m_oOldVar.m_strPreVarArray[x].Trim().Length > 0 &&
					m_oOldVar.m_strPreVarArray[x].Trim().ToUpper() != "NOT DEFINED")
				{
					this.btnFVSVariablesPrePostOverall.Enabled = true;
					break;
				}
			}
			if (x > NUMBER_OF_VARIABLES - 1) btnFVSVariablesPrePostOverall.Enabled = false;

			m_oOldVar.Copy(m_oOldVar, ref m_oSavVar);
			if (m_oCurVar == null) m_oCurVar = new Variables();
			m_oOldVar.Copy(m_oOldVar, ref m_oCurVar);

			this.ReferenceOptimizationUserControl.loadvalues_FromProperties();
			this.ReferenceTieBreakerUserControl.loadvalues_FromProperties();



		}
		
		public void loadvalues()
        {
			this.m_intError = 0;
			this.m_strError = "";

			DataMgr oDataMgr = new DataMgr();

			this.lstFVSTablesList.Items.Clear();
			this.lstFVSTablesList.Sorted = true;
			m_dictFVSTables = m_oOptimizerScenarioTools.LoadFvsTablesAndVariables();
			foreach (string strKey in m_dictFVSTables.Keys)
			{
				lstFVSTablesList.Items.Add(strKey);
			}
			m_oOldVar = new Variables();
			m_oSavVar = new Variables();

			int x = 0;
			for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
				this.RemoveVariable(x);

			string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioDB =
				frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" +
				Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDB)))
			{
				conn.Open();

				int intVarNum = 0;

				//fvs variables
				oDataMgr.m_strSQL = "SELECT * " +
								"FROM scenario_fvs_variables " +
								"WHERE TRIM(UPPER(scenario_id))='" + strScenarioId.Trim().ToUpper() + "' AND " +
								"current_yn='Y' " +
								"ORDER BY variable_number";
				oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
				if (oDataMgr.m_DataReader.HasRows)
				{
					while (oDataMgr.m_DataReader.Read())
					{
						intVarNum = Convert.ToInt32(oDataMgr.m_DataReader["variable_number"]) - 1;

						m_oOldVar.m_strPreVarArray[intVarNum] =
							Convert.ToString(oDataMgr.m_DataReader["pre_fvs_variable"]).Trim();
						m_oOldVar.m_strPostVarArray[intVarNum] =
							Convert.ToString(oDataMgr.m_DataReader["post_fvs_variable"]).Trim();

						if (oDataMgr.m_DataReader["better_expression"] != System.DBNull.Value)
							m_oOldVar.m_strBetterExpr[intVarNum] =
								Convert.ToString(oDataMgr.m_DataReader["better_expression"]).Trim();

						if (oDataMgr.m_DataReader["worse_expression"] != System.DBNull.Value)
							m_oOldVar.m_strWorseExpr[intVarNum] =
								Convert.ToString(oDataMgr.m_DataReader["worse_expression"]).Trim();

						if (oDataMgr.m_DataReader["effective_expression"] != System.DBNull.Value)
							m_oOldVar.m_strEffectiveExpr[intVarNum] =
								Convert.ToString(oDataMgr.m_DataReader["effective_expression"]).Trim();


						this.UpdateListViewVariableItem(intVarNum, intVarNum + 1, m_oOldVar);
					}
				}
				oDataMgr.m_DataReader.Close();

				//overall expression
				oDataMgr.m_strSQL = "SELECT overall_effective_expression, current_yn " +
								"FROM scenario_fvs_variables_overall_effective " +
								"WHERE TRIM(UPPER(scenario_id))='" + strScenarioId.Trim().ToUpper() + "' AND " +
								"current_yn='Y'";

				oDataMgr.SqlQueryReader(conn, oDataMgr.m_strSQL);
				if (oDataMgr.m_DataReader.HasRows)
				{
					while (oDataMgr.m_DataReader.Read())
					{
						if (oDataMgr.m_DataReader["overall_effective_expression"] != System.DBNull.Value &&
							m_oOldVar.m_strOverallEffectiveExpr.Trim().Length == 0)
                        {
							m_oOldVar.m_strOverallEffectiveExpr = Convert.ToString(oDataMgr.m_DataReader["overall_effective_expression"]).Trim();
                        }
					}
				}
				oDataMgr.m_DataReader.Close();
			}

			for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
			{
				if (m_oOldVar.m_strPreVarArray[x].Trim().Length > 0 &&
					m_oOldVar.m_strPreVarArray[x].Trim().ToUpper() != "NOT DEFINED")
					this.btnFVSVariablesPrePostOverall.Enabled = true;
			}
			m_oOldVar.Copy(m_oOldVar, ref m_oSavVar);

			this.ReferenceOptimizationUserControl.loadvalues(m_dictFVSTables);
			this.ReferenceTieBreakerUserControl.loadvalues(m_dictFVSTables);
			this.m_intError = oDataMgr.m_intError;
			this.m_strError = oDataMgr.m_strError;
		}

        private void loadFVSTableAndField()
        {
            lstFVSTablesList.ClearSelected();
            string[] strPieces = this.lvFVSVariablesPrePostValues.SelectedItems[0].SubItems[COLUMN_PREVAR].Text.Trim().Split('.');
            //set FVS table and variable
            if (strPieces.Length == 2 && !String.IsNullOrEmpty(strPieces[0]))
            {
                string strTable = strPieces[0].TrimStart("PRE_".ToCharArray());
                for (int index = 0; index < lstFVSTablesList.Items.Count; index++)
                {
                    string item = lstFVSTablesList.Items[index].ToString();
                    if (strTable == item)
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

		public int savevalues()
		{
			int x;
			string strColumns = "";
			string strValues = "";
			string strWhere = "";


			string strColumn1 = "";
			string strColumn2 = "";
			string strColumn3 = "";
			string strColumn4 = "";
			string strFVSVariableList = "";

			string[] strArray = null;


			this.m_intCurVariableDefinitionStepCount = WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE;

			strArray = this.m_oUtils.ConvertListToArray(m_oSavVar.m_strPreVarArray[0], ".");

			if (strArray.Length > 0) strColumn1 = strArray[strArray.Length - 1];

			strArray = this.m_oUtils.ConvertListToArray(m_oSavVar.m_strPreVarArray[1], ".");

			if (strArray.Length > 0) strColumn2 = strArray[strArray.Length - 1];

			strArray = this.m_oUtils.ConvertListToArray(m_oSavVar.m_strPreVarArray[2], ".");

			if (strArray.Length > 0) strColumn3 = strArray[strArray.Length - 1];

			strArray = this.m_oUtils.ConvertListToArray(m_oSavVar.m_strPreVarArray[3], ".");

			if (strArray.Length > 0) strColumn4 = strArray[strArray.Length - 1];

			if (strColumn1.Trim().Length > 0 && strColumn1.Trim().ToUpper() != "NOT DEFINED") strFVSVariableList = "1-" + strColumn1 + ",";
			if (strColumn2.Trim().Length > 0 && strColumn2.Trim().ToUpper() != "NOT DEFINED") strFVSVariableList = strFVSVariableList + "2-" + strColumn2 + ",";
			if (strColumn3.Trim().Length > 0 && strColumn3.Trim().ToUpper() != "NOT DEFINED") strFVSVariableList = strFVSVariableList + "3-" + strColumn3 + ",";
			if (strColumn4.Trim().Length > 0 && strColumn4.Trim().ToUpper() != "NOT DEFINED") strFVSVariableList = strFVSVariableList + "4-" + strColumn4 + ",";
			if (strFVSVariableList.Trim().Length > 0)
				strFVSVariableList = strFVSVariableList.Substring(0, strFVSVariableList.Length - 1);

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
					oDataMgr.m_strSQL = "SELECT COUNT(*) FROM scenario_fvs_variables WHERE " +
						" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y' AND rxcycle='1';";
					if ((int)oDataMgr.getRecordCount(conn, oDataMgr.m_strSQL, "scenario_fvs_variables") > 0)
					{
						oDataMgr.m_strSQL = "UPDATE scenario_fvs_variables SET current_yn = 'N'" +
							" WHERE TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y' AND rxcycle='1';";
						oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
					}

					if (oDataMgr.m_intError < 0)
					{
						conn.Close();
						x = oDataMgr.m_intError;
						oDataMgr = null;
						return x;
					}

					oDataMgr.m_strSQL = "SELECT COUNT(*) FROM scenario_fvs_variables_overall_effective WHERE " +
						" TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y' AND rxcycle='1';";
					if ((int)oDataMgr.getRecordCount(conn, oDataMgr.m_strSQL, "scenario_fvs_variables_overall_effective") > 0)
					{
						oDataMgr.m_strSQL = "UPDATE scenario_fvs_variables_overall_effective SET current_yn = 'N'" +
							" WHERE TRIM(UPPER(scenario_id)) = '" + strScenarioId.Trim().ToUpper() + "' AND current_yn = 'Y' AND rxcycle='1';";
						oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
					}
					if (oDataMgr.m_intError < 0)
					{
						conn.Close();
						x = oDataMgr.m_intError;
						oDataMgr = null;
						return x;
					}

					strColumns = "scenario_id,rxcycle,variable_number,fvs_variables_list,pre_fvs_variable,post_fvs_variable,better_expression,worse_expression,effective_expression,current_yn";
					for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
					{
						strValues = "'" + strScenarioId.Trim() + "','1'";

						strWhere = "TRIM(scenario_id)='" + strScenarioId.Trim() + "' AND rxcycle='1' ";
						if (m_oSavVar.m_strPreVarArray[x].Trim().Length > 0 &&
							m_oSavVar.m_strPreVarArray[x].Trim().ToUpper() != "NOT DEFINED")
						{
							strValues = strValues + "," + Convert.ToString(x + 1).Trim();
							strWhere = strWhere + " AND variable_number=" + Convert.ToString(x + 1).Trim();
							strValues = strValues + ",'" + strFVSVariableList + "'";
							strWhere = strWhere + " AND UPPER(TRIM(fvs_variables_list))='" + strFVSVariableList.Trim().ToUpper() + "'";
							strValues = strValues + ",'" + m_oSavVar.m_strPreVarArray[x].Trim() + "'";
							strWhere = strWhere + " AND UPPER(TRIM(pre_fvs_variable))='" + m_oSavVar.m_strPreVarArray[x].Trim().ToUpper() + "'";
							strValues = strValues + ",'" + m_oSavVar.m_strPostVarArray[x].Trim() + "'";
							strWhere = strWhere + " AND UPPER(TRIM(post_fvs_variable))='" + m_oSavVar.m_strPostVarArray[x].Trim().ToUpper() + "'";
							strValues = strValues + ",'" + oDataMgr.FixString(m_oSavVar.m_strBetterExpr[x].Trim(), "'", "''") + "'";
							strWhere = strWhere + " AND UPPER(TRIM(better_expression))='" + oDataMgr.FixString(m_oSavVar.m_strBetterExpr[x].Trim(), "'", "''").ToUpper() + "'";
							strValues = strValues + ",'" + oDataMgr.FixString(m_oSavVar.m_strWorseExpr[x].Trim(), "'", "''") + "'";
							strWhere = strWhere + " AND UPPER(TRIM(worse_expression))='" + oDataMgr.FixString(m_oSavVar.m_strWorseExpr[x].Trim(), "'", "''").ToUpper() + "'";
							strValues = strValues + ",'" + oDataMgr.FixString(m_oSavVar.m_strEffectiveExpr[x].Trim(), "'", "''") + "'";
							strWhere = strWhere + " AND UPPER(TRIM(effective_expression))='" + oDataMgr.FixString(m_oSavVar.m_strEffectiveExpr[x].Trim(), "'", "''").ToUpper() + "'";
							strValues = strValues + ",'Y'";

							//delete duplicates
							oDataMgr.m_strSQL = "DELETE FROM scenario_fvs_variables WHERE " + strWhere;
							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

							//insert the current item
							oDataMgr.m_strSQL = "INSERT INTO scenario_fvs_variables " +
											"(" + strColumns + ") VALUES " +
											"(" + strValues + ")";
							oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
						}
					}

					if (m_oSavVar.m_strOverallEffectiveExpr.Trim().Length > 0)
					{
						strColumns = "scenario_id,rxcycle,fvs_variables_list,overall_effective_expression,current_yn";
						strValues = "'" + strScenarioId.Trim() + "','1'";
						strValues = strValues + ",'" + strFVSVariableList + "'";
						strValues = strValues + ",'" + oDataMgr.FixString(this.m_oSavVar.m_strOverallEffectiveExpr.Trim(), "'", "''") + "'";
						strValues = strValues + ",'Y'";
						oDataMgr.m_strSQL = "INSERT INTO scenario_fvs_variables_overall_effective " +
							"(" + strColumns + ") VALUES " +
							"(" + strValues + ")";
						oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
					}
				}
			}
			return 1;
		}


		private void txtFFE_CI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		   this.AllowNumericOnly(e);
		}

		private void txtFFE_CI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_CI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_CI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		private void txtFFE_CI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.AllowNumericOnly(e);
		}

		


		

		private void btnFFEExpressionBuilderDefault_Click(object sender, System.EventArgs e)
		{
			this.txtExpression.Text = "";
			if (this.btnPrev.Name == "overalleffective_prev")
			{
				this.txtExpression.AppendText(this.getDefaultExpression("overall"));
				//if (((frmScenario)this.ParentForm).btnSave.Enabled==false) 
				//	((frmScenario)this.ParentForm).btnSave.Enabled=true;
				((frmOptimizerScenario)this.ParentForm).m_bSave=true;
				this.m_strOverallEffectiveExpression=this.txtExpression.Text.Trim();
			}
			else 
			{
				this.txtExpression.AppendText(this.getDefaultExpression(this.m_strCurrentIndexTypeAndClass));
			}
		}
		protected string getDefaultExpression(string str)
		{
			switch (str)
			{
				case "T1":
					return "(PRE_TI_CL = 1 AND ((TI_CHANGE > 10 AND POST_TI_CL >= 2) OR (TI_CHANGE >= 20)))";
				case "T2":
					return "(PRE_TI_CL = 2 AND TI_CHANGE >= 15)";
				case "T3":
					return "(PRE_TI_CL = 3 AND TI_CHANGE >= 20)";
				case "T4":
					return "(PRE_TI_CL = 4 AND TI_CHANGE >= 25)";
				case "T5":
					return "(PRE_TI_CL = 5 AND TI_CHANGE >= 30)";
				case "C1":
					return " (PRE_CI_CL = 1 AND ((TI_CHANGE > 10 AND POST_CI_CL >= 2) OR (CI_CHANGE > 20)))";
				case "C2":
					return "(PRE_CI_CL = 2 AND CI_CHANGE >= 15)";
				case "C3":
					return "(PRE_CI_CL = 3 AND CI_CHANGE >= 20)";
				case "C4":
					return "(PRE_CI_CL = 4 AND CI_CHANGE >= 25)";
				case "C5":
					return  "(PRE_CI_CL = 5 AND CI_CHANGE >= 30)";
                case "overall":
					return "((TI_EFFECTIVE_YN = 'Y' AND (CI_BACKSLIDE_YN = 'N') AND (POST_CI_CL > 1)) " + 
						                       " OR " + 
					        "(CI_EFFECTIVE_YN = 'Y' AND (TI_BACKSLIDE_YN = 'N') AND (POST_TI_CL > 1)))";
                default:
					return "";
			}

		}

		//private void lstFFEAvailableFields_DoubleClick(object sender, System.EventArgs e)
		//{
		//	this.SendKeyStrokes(this.txtExpression," " + this.lstFFEAvailableFields.SelectedItem.ToString().Trim() + " ");
		//}
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

		private void btnFFEEqual_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression," = ");
		}

		private void btnFFENotEqual_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression," <> ");
		}

		private void btnLeftBracket_Click(object sender, System.EventArgs e)
		{
			//string str= "{(}";
			this.SendKeyStrokes(this.txtExpression, " {(}");
			
		}

		private void btnRightBracket_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, "{)} ");
		}

		private void btnFFEMoreThan_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " > ");
		}

		private void btnFFEMoreThanOrEqualTo_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " >= ");
		}

		private void btnFFEAnd_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, " and ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, " AND ");
			}
		}

		private void btnFFELessThan_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " < ");
		}

		private void btnLessThanOrEqualTo_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression, " <= ");
		}

		private void btnFFEOr_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, " or ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, " OR ");
			}
		}

		private void btnFFENote_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, " not ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, " NOT ");
			}
		}

		private void btnFFEExpressionBuilderClear_Click(object sender, System.EventArgs e)
		{
			string strKeyStrokes="";
			this.txtExpression.Text="";
			if (this.btnPrev.Name != "overalleffective_prev")
			{
				switch (this.m_strCurrentIndexTypeAndClass)
				{
					case "T1":
						strKeyStrokes= "(PRE_TI_CL = 1 AND ";
						break;
					case "T2":
						strKeyStrokes = "(PRE_TI_CL = 2 AND ";
						break;
					case "T3":
						strKeyStrokes = "(PRE_TI_CL = 3 AND ";
						break;
					case "T4":
						strKeyStrokes = "(PRE_TI_CL = 4 AND ";
						break;
					case "T5":
						strKeyStrokes = "(PRE_TI_CL = 5 AND ";
						break;
					case "C1":
						strKeyStrokes = "(PRE_CI_CL = 1 AND ";
						break;
					case "C2":
						strKeyStrokes = "(PRE_CI_CL = 2 AND ";
						break;
					case "C3":
						strKeyStrokes = "(PRE_CI_CL = 3 AND ";
						break;
					case "C4":
						strKeyStrokes = "(PRE_CI_CL = 4 AND ";
						break;
					case "C5":
						strKeyStrokes = "(PRE_CI_CL = 5 AND ";
						break;
				}
				this.txtExpression.Focus();
				this.txtExpression.AppendText(strKeyStrokes);
			}


		}

		private void btnFFEExpressionBuilderTest_Click(object sender, System.EventArgs e)
		{
			string strSQL="";
			//int intArrayCount;
			//int x=0;
			string strConn="";
			//string strCommand="";
			//string str="";

			ado_data_access p_ado = new ado_data_access();

			string strScenarioId = this.ReferenceOptimizerScenarioForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
            
			//scenario mdb connection
			string strOptimizerResultsAccdb = 
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory +
                "\\optimizer\\" + strScenarioId + "\\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile;

			this.m_OleDbConnectionScenario = new System.Data.OleDb.OleDbConnection();
			strConn=p_ado.getMDBConnString(strOptimizerResultsAccdb,"admin","");
			p_ado.OpenConnection(strConn, ref this.m_OleDbConnectionScenario);	
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return;
			}
			strSQL = "SELECT * FROM effective WHERE " +
				this.txtExpression.Text + ";";
			p_ado.SqlQueryReader(this.m_OleDbConnectionScenario, strSQL);

				  
			if (p_ado.m_intError == 0)
			{
				

				MessageBox.Show("Valid Syntax");
				
				p_ado.m_OleDbDataReader.Close();
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				
			}
			p_ado = null;
			this.m_OleDbConnectionScenario.Close();

		}

		

		

		


		public int val_overall_effective_expression()
		{
			//int x=0;
			this.m_intError=0;
			this.m_strError = "Run Scenario Failed: ";

			if (this.m_strOverallEffectiveExpression.Trim().Length==0)
			{
				this.m_intError=-1;
				MessageBox.Show(this.m_strError + "Need user defined expression for overall effective treatment in <Fuel and Fire Effects>");
				return this.m_intError;
			}   
			return this.m_intError;
			
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

		

		

		
		
		private void cmbFFECIBackslide2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide3_SelectedValueChanged(object sender, System.EventArgs e)
		{
		
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide3_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide2_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIBackslide_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIBackslide_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIHazardOperator_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFETIHazardWindSpeedClass_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIHazardOperator_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void cmbFFECIHazardWindSpeedClass_SelectedValueChanged(object sender, System.EventArgs e)
		{
			
			((frmOptimizerScenario)this.ParentForm).m_bSave=true;
		}

		private void txtExpression_Leave(object sender, System.EventArgs e)
		{
			if (this.btnPrev.Name == "overalleffective_prev")
			{
				this.m_strOverallEffectiveExpression=this.txtExpression.Text;
			}
		}

		

		private void cmbFFE_TI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_TI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI3_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbFFE_CI3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI4_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFE_CI5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIBackslide_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIBackslide_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIBackslide2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIBackslide2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIBackslide3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIBackslide3_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIHazardOperator_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIHazardOperator_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFETIHazardWindSpeedClass_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		private void cmbFFECIHazardWindSpeedClass_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			e.Handled=true;
		}

		

		
		private void btnY_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, "'y' ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, "'Y' ");
			}
		}

		private void btnN_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtExpression, "'n' ");
			}
			else
			{
				this.SendKeyStrokes(this.txtExpression, "'N' ");
			}
		}

		private void cmdFFE_Click(object sender, System.EventArgs e)
		{

		}

		private void listBox1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void uc_scenario_fvs_prepost_variables_Resize(object sender, System.EventArgs e)
		{
			main_resize();
		}
		public void main_resize()
		{
			
		}

		private void groupBox1_Resize(object sender, System.EventArgs e)
		{

			
			grpboxFVSVariablesPrePost.Height = this.ClientSize.Height - grpboxFVSVariablesPrePost.Top - 5;
			grpboxFVSVariablesPrePost.Width = this.ClientSize.Width - (grpboxFVSVariablesPrePost.Left * 2) ;
		    grpboxFVSVariablesPrePostVariable.Height = grpboxFVSVariablesPrePost.Height;
			grpboxFVSVariablesPrePostVariable.Width =  grpboxFVSVariablesPrePost.Width;
			grpboxFVSVariablesPrePostExpression.Height = grpboxFVSVariablesPrePost.Height;
			grpboxFVSVariablesPrePostExpression.Width = grpboxFVSVariablesPrePost.Width;
			grpboxFVSVariablesPrePostValues.Width = this.pnlFVSVariablesPrePost.Width - (this.grpboxFVSVariablesPrePostValues.Left * 2)-20;
			this.lvFVSVariablesPrePostValues.Width = grpboxFVSVariablesPrePostValues.Width - (this.lvFVSVariablesPrePostValues.Left * 2); 
			
		}

	
		

		private void btnFVSVariablesPrePostVariableValue_Click(object sender, System.EventArgs e)
		{
            if (this.lstFVSTablesList.SelectedItems.Count == 0 || this.lstFVSFieldsList.SelectedItems.Count == 0) return;
            this.lblFVSVariablesPrePostVariablePreSelected.Text = "PRE_" +
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();

            this.lblFVSVariablesPrePostVariablePostSelected.Text = "POST_" +
                this.lstFVSTablesList.SelectedItems[0].ToString() + "." + this.lstFVSFieldsList.SelectedItems[0].ToString();
		}

		private void btnFVSVariablesPrePostVariableClearAll_Click(object sender, System.EventArgs e)
		{
			this.lblFVSVariablesPrePostVariablePreSelected.Text = "Not Defined";
			this.lblFVSVariablesPrePostVariablePostSelected.Text = "Not Defined";
            this.lstFVSTablesList.ClearSelected();
		}

		private void btnFVSVariablesPrePostVariableNext_Click(object sender, System.EventArgs e)
		{
			if (this.lblFVSVariablesPrePostVariablePreSelected.Text.Trim().ToUpper() != "NOT DEFINED")
			{
				this.m_oCurVar.m_strPreVarArray[m_intCurVar-1] = this.lblFVSVariablesPrePostVariablePreSelected.Text;
				this.m_oCurVar.m_strPostVarArray[m_intCurVar-1] = this.lblFVSVariablesPrePostVariablePostSelected.Text;
				this.loadvalues_variable((byte)(m_intCurVar-1),2);
				this.grpboxFVSVariablesPrePostVariable.Hide();
				this.grpboxFVSVariablesPrePostExpression.Show();
				this.txtExpression.Focus();
				this.txtExpression.Select(this.txtExpression.Text.Length,0);
			}
		}

		
		/// <summary>
		/// Load the Pre-Post variable values
		/// Each pre-post variable set has these variables
		///1. Pre variable name
		///2. Post variable name
		///3. Better SQL 
		///4. Worse  SQL
		///5. Effective SQL
		/// </summary>
		/// <param name="p_intVariable">Current Pre-Post Variable Set</param>
		/// <param name="p_intStep">The variable item within the set to load</param>
		private void loadvalues_variable(int p_intVariable,int p_intStep)
		{
			int intCurrentSet; //=p_intVariable+1;
			int intItemsWithinASet=5;		
			int intCurrentItem=p_intStep;
			string strColumn1="";
			string strColumn2="";
			string strColumn3="";
			string strColumn4="";

			string[] strArray=null; 


			
			intCurrentSet = p_intVariable;


			this.m_intCurVar=p_intVariable+1;

			m_intCurVariableDefinitionStepCount=p_intStep;

			this.cmbFVSVariablesPrePost.Text = 
				this.cmbFVSVariablesPrePost.Items[
				    (intCurrentSet * intItemsWithinASet) - (intCurrentSet - intCurrentItem)-1].ToString();

			grpboxFVSVariablesPrePostVariable.Text = "Variable " + m_intCurVar.ToString();
			this.lblFVSVariablesPrePostVariablePreSelected.Text = this.m_oCurVar.m_strPreVarArray[p_intVariable];
			this.lblFVSVariablesPrePostVariablePostSelected.Text = this.m_oCurVar.m_strPostVarArray[p_intVariable];
			this.grpboxFVSVariablesPrePostExpression.Text = "Expression Builder: Variable " + m_intCurVar.ToString();

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[0],".");

			if (strArray.Length > 0) strColumn1 = strArray[strArray.Length-1];

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[1],".");

			if (strArray.Length > 0) strColumn2 = strArray[strArray.Length-1];

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[2],".");

			if (strArray.Length > 0) strColumn3 = strArray[strArray.Length-1];

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[3],".");

			if (strArray.Length > 0) strColumn4 = strArray[strArray.Length-1];
			
			this.lblFVSVariablesPrePostExpression.Text = cmbFVSVariablesPrePost.Text;


			this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Clear();
			this.btnFVSVariablesPrePostExpressionNext.Show();
            this.btnFVSVariablesPrePostExpressionNext.Top = this.btnFVSVariablesPrePostExpressionCancel.Top;
			if (p_intStep == WIZARD_STEP_VARIABLE_BETTER ||
				p_intStep == WIZARD_STEP_VARIABLE_WORSE)
			{
				switch (p_intVariable)
				{
					case 0:
						this.lblFVSVariablesPrePostExpressionVariable1.Text = "Variable 1 = " + strColumn1;
                        this.lblFVSVariablesPrePostExpressionVariable2.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable3.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable4.Text = "";
						break;
					case 1:
						this.lblFVSVariablesPrePostExpressionVariable2.Text = "Variable 2 = " + strColumn2;
						this.lblFVSVariablesPrePostExpressionVariable1.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable3.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable4.Text = "";
						break;
					case 2:
						this.lblFVSVariablesPrePostExpressionVariable3.Text = "Variable 3 = " + strColumn3;
						this.lblFVSVariablesPrePostExpressionVariable1.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable2.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable4.Text = "";
						break;
					case 3:
						this.lblFVSVariablesPrePostExpressionVariable4.Text = "Variable 4 = " + strColumn4;
						this.lblFVSVariablesPrePostExpressionVariable1.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable2.Text = "";
                        this.lblFVSVariablesPrePostExpressionVariable3.Text = "";
						break;


				}
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable" + m_intCurVar.ToString().Trim() + "_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable" + m_intCurVar.ToString().Trim() + "_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable" + m_intCurVar.ToString().Trim() + "_change");

				this.btnFVSVariablesPrePostExpressionNext.Show();
                this.btnFVSVariablesPrePostExpressionNext.Top = this.btnFVSVariablesPrePostExpressionCancel.Top;
				if (p_intStep == WIZARD_STEP_VARIABLE_BETTER)
				{
					this.txtExpression.Text = this.m_oCurVar.m_strBetterExpr[m_intCurVar-1];
				}
				else
				{
					this.txtExpression.Text = this.m_oCurVar.m_strWorseExpr[m_intCurVar-1];
				}
			}
            else if (p_intStep == WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE || 
				     p_intStep == WIZARD_STEP_VARIABLE_EFFECTIVE)
			{
                this.lblFVSVariablesPrePostExpressionVariable1.Text = "";
                this.lblFVSVariablesPrePostExpressionVariable2.Text = "";
                this.lblFVSVariablesPrePostExpressionVariable3.Text = "";
                this.lblFVSVariablesPrePostExpressionVariable4.Text = "";
				if (p_intStep == WIZARD_STEP_VARIABLE_EFFECTIVE)
				{
					this.txtExpression.Text = this.m_oCurVar.m_strEffectiveExpr[m_intCurVar-1];
					this.btnFVSVariablesPrePostExpressionNext.Hide();
				}

				if (m_oCurVar.m_strPreVarArray[0].Trim().ToUpper() != "NOT DEFINED")
				{
					this.lblFVSVariablesPrePostExpressionVariable1.Text = "Variable 1 = " + strColumn1;
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable1_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable1_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_change");
					if (m_oCurVar.m_strBetterExpr[0].Trim().Length > 0)
					   this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_better_yn");
					if (this.m_oCurVar.m_strWorseExpr[0].Trim().Length > 0)
					   this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_worse_yn");
					if (p_intStep == WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
					{
						if (this.m_oCurVar.m_strEffectiveExpr[0].Trim().Length > 0)
							this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_effective_yn");
					}

				}
				

				if (m_oCurVar.m_strPreVarArray[1].Trim().ToUpper() != "NOT DEFINED")
				{
					this.lblFVSVariablesPrePostExpressionVariable2.Text = "Variable 2 = " + strColumn2;
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable2_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable2_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_change");
					if (this.m_oCurVar.m_strBetterExpr[1].Trim().Length > 0)
						this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_better_yn");
					if (this.m_oCurVar.m_strWorseExpr[1].Trim().Length > 0)
						this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_worse_yn");
					if (p_intStep == WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
					{
						if (this.m_oCurVar.m_strEffectiveExpr[1].Trim().Length > 0)
							this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_effective_yn");
					}

				}
				if (m_oCurVar.m_strPreVarArray[2].Trim().ToUpper() != "NOT DEFINED")
				{
					this.lblFVSVariablesPrePostExpressionVariable3.Text = "Variable 3 = " + strColumn3;
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable3_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable3_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_change");
					if (this.m_oCurVar.m_strBetterExpr[2].Trim().Length > 0)
					    this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_better_yn");
					if (this.m_oCurVar.m_strWorseExpr[2].Trim().Length > 0)
						this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_worse_yn");
					if (p_intStep == WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
					{
						if (this.m_oCurVar.m_strEffectiveExpr[2].Trim().Length > 0)
							this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_effective_yn");
					}
				}
				if (m_oCurVar.m_strPreVarArray[3].Trim().ToUpper() != "NOT DEFINED")
				{
					this.lblFVSVariablesPrePostExpressionVariable4.Text = "Variable 4 = " + strColumn4;
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable4_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable4_value");
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_change");
					if (this.m_oCurVar.m_strBetterExpr[3].Trim().Length > 0)
						this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_better_yn");
					if (this.m_oCurVar.m_strWorseExpr[3].Trim().Length > 0)
						this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_worse_yn");
					if (p_intStep == WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
					{
						if (this.m_oCurVar.m_strEffectiveExpr[3].Trim().Length > 0)
							this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_effective_yn");
					}
				}
			}
		}
		
		private void loadvalues_overall()
		{
			string strColumn1="";
			string strColumn2="";
			string strColumn3="";
			string strColumn4="";

			string[] strArray=null;
            this.lblFVSVariablesPrePostExpressionVariable1.Text = "";
            this.lblFVSVariablesPrePostExpressionVariable2.Text = "";
            this.lblFVSVariablesPrePostExpressionVariable3.Text = "";
            this.lblFVSVariablesPrePostExpressionVariable4.Text = "";
			
			if (m_oCurVar==null)
			{
				this.m_oCurVar = new Variables();
				this.m_oCurVar.Copy(this.m_oSavVar,ref m_oCurVar);
			}

			this.m_intCurVariableDefinitionStepCount = WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE;

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[0],".");

			if (strArray.Length > 0) strColumn1 = strArray[strArray.Length-1];

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[1],".");

			if (strArray.Length > 0) strColumn2 = strArray[strArray.Length-1];

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[2],".");

			if (strArray.Length > 0) strColumn3 = strArray[strArray.Length-1];

			strArray = this.m_oUtils.ConvertListToArray(m_oCurVar.m_strPreVarArray[3],".");

			if (strArray.Length > 0) strColumn4 = strArray[strArray.Length-1];

			
			this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Clear();

			if (m_oCurVar.m_strPreVarArray[0].Trim().ToUpper() != "NOT DEFINED")
			{
				this.lblFVSVariablesPrePostExpressionVariable1.Text = "Variable 1 = " + strColumn1;
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable1_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable1_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_change");
				if (this.m_oCurVar.m_strBetterExpr[0].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_better_yn");
				if (this.m_oCurVar.m_strWorseExpr[0].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_worse_yn");
				if (this.m_oCurVar.m_strEffectiveExpr[0].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable1_effective_yn");
			}
			if (m_oCurVar.m_strPreVarArray[1].Trim().ToUpper() != "NOT DEFINED")
			{
				this.lblFVSVariablesPrePostExpressionVariable2.Text = "Variable 2 = " + strColumn2;
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable2_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable2_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_change");
				if (this.m_oCurVar.m_strBetterExpr[1].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_better_yn");
				if (this.m_oCurVar.m_strWorseExpr[1].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_worse_yn");
				if (this.m_oCurVar.m_strEffectiveExpr[1].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable2_effective_yn");
			}
			if (m_oCurVar.m_strPreVarArray[2].Trim().ToUpper() != "NOT DEFINED")
			{
				this.lblFVSVariablesPrePostExpressionVariable3.Text = "Variable 3 = " + strColumn3;
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable3_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable3_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_change");
				if (this.m_oCurVar.m_strBetterExpr[2].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_better_yn");
				if (this.m_oCurVar.m_strWorseExpr[2].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_worse_yn");
				if (this.m_oCurVar.m_strEffectiveExpr[2].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable3_effective_yn");

			}
			if (m_oCurVar.m_strPreVarArray[3].Trim().ToUpper() != "NOT DEFINED")
			{
				this.lblFVSVariablesPrePostExpressionVariable4.Text = "Variable 4 = " + strColumn4;
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("pre_variable4_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("post_variable4_value");
				this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_change");
				if (this.m_oCurVar.m_strBetterExpr[3].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_better_yn");
				if (this.m_oCurVar.m_strWorseExpr[3].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_worse_yn");
				if (this.m_oCurVar.m_strEffectiveExpr[3].Trim().Length > 0)
					this.lstFVSVariablesPrePostExpressionSelectedVariables.Items.Add("variable4_effective_yn");

			}
			this.cmbFVSVariablesPrePost.Text = cmbFVSVariablesPrePost.Items[cmbFVSVariablesPrePost.Items.Count-1].ToString();

			this.txtExpression.Text = this.m_oCurVar.m_strOverallEffectiveExpr;

			this.lblFVSVariablesPrePostExpression.Text = "Define Expression for Overall Effectiveness";
			this.grpboxFVSVariablesPrePostExpression.Text = "Overall Effectiveness";
			this.btnFVSVariablesPrePostExpressionPrevious.Hide();
			this.btnFVSVariablesPrePostExpressionNext.Hide();
		
			this.grpboxFVSVariablesPrePost.Hide();
			this.grpboxFVSVariablesPrePostExpression.Show();
		}
		
		private void loadmacro(string p_strVariable,string p_strValue)
		{
			for (int x=0;x<=this.m_oSQLMacroSubstitutionVariable_Collection.Count-1;x++)
			{
				if (this.m_oSQLMacroSubstitutionVariable_Collection.Item(x).VariableName.Trim().ToUpper()==
					p_strVariable.Trim().ToUpper())
				{
					this.m_oSQLMacroSubstitutionVariable_Collection.Item(x).SQLVariableSubstitutionString=p_strValue;
					break;
				}
			}
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
			switch (this.m_intCurVariableDefinitionStepCount)
			{
				case WIZARD_STEP_VARIABLE_BETTER:
					this.m_oCurVar.m_strBetterExpr[this.m_intCurVar-1]=this.txtExpression.Text;
					this.loadvalues_variable(m_intCurVar-1,m_intCurVariableDefinitionStepCount-1);
					this.grpboxFVSVariablesPrePostVariable.Show();
					this.grpboxFVSVariablesPrePostExpression.Hide();
					break;
                case WIZARD_STEP_VARIABLE_WORSE:
					this.m_oCurVar.m_strWorseExpr[this.m_intCurVar-1] = this.txtExpression.Text;
					this.loadvalues_variable(m_intCurVar-1,m_intCurVariableDefinitionStepCount-1);
					break;
				case WIZARD_STEP_VARIABLE_EFFECTIVE:
					this.m_oCurVar.m_strEffectiveExpr[this.m_intCurVar-1] = this.txtExpression.Text;
					this.loadvalues_variable(m_intCurVar-1,m_intCurVariableDefinitionStepCount-1);
					break;
			}
		}


		private void btnFVSVariablesPrePostExpressionTest_Click(object sender, System.EventArgs e)
        {
			string strSQL = this.txtExpression.Text;
			DataMgr oDataMgr = new DataMgr();
			string strTempFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
			oDataMgr.CreateDbFile(strTempFile);

			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTempFile)))
            {
				conn.Open();

				// Query the optimization uc for the selected revenue attribute so we can pass it to the table
				OptimizerScenarioItem.OptimizationVariableItem_Collection oOptimizationVariableCollection =
					this.ReferenceOptimizerScenarioForm.m_oOptimizerScenarioItem_Collection.Item(0).m_oOptimizationVariableItem_Collection;
				string strColumnFilterName = "";
				for (int x = 0; x <= oOptimizationVariableCollection.Count - 1; x++)
				{
					if (oOptimizationVariableCollection.Item(x).bSelected)
					{
						if (oOptimizationVariableCollection.Item(x).bUseFilter == true)
							strColumnFilterName = oOptimizationVariableCollection.Item(x).strRevenueAttribute;
						break;
					}
				}

				string strPrefix = "cycle1";
				for (int i = 0; i < lvFVSVariablesPrePostValues.Items.Count; i++)
				{
					string strPreVariable = Convert.ToString(lvFVSVariablesPrePostValues.Items[i].SubItems[COLUMN_PREVAR]).Trim();
					string[] strPieces = strPreVariable.Split('.');
					if (strPieces.Length == 2 && !String.IsNullOrEmpty(strPieces[0]))
					{
						if (strPieces[0].ToUpper().Contains("_WEIGHTED"))
						{
							strPrefix = "all_cycles";
							break;
						}
					}

				}

				string strEffectiveTableName = strPrefix + "_effective";
				frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEffectiveTable(oDataMgr, conn,
					strPrefix, strColumnFilterName);

				strSQL = "SELECT * FROM " + strEffectiveTableName + " WHERE " +
					strSQL + ";";
				oDataMgr.SqlQueryReader(conn, strSQL);

				if (oDataMgr.m_intError == 0)
				{
					MessageBox.Show("Valid Syntax");

					oDataMgr.m_DataReader.Close();
				}
			}
		}
		private void Go()
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

		private void btnFVSVariablesPrePostVariableCancel_Click(object sender, System.EventArgs e)
		{

			this.grpboxFVSVariablesPrePostVariable.Hide();
			
			ReferenceOptimizerScenarioForm.m_bEnableSelectedTabPage=true;
		    ReferenceOptimizerScenarioForm.m_strCurrentEditTabControlName="";
			this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabPageText="";
			this.EnableTabs(true);
		    this.m_intCurVar=-1;
			this.grpboxFVSVariablesPrePost.Show();
		}

		private void btnFVSVariablesPrePostVariableDone_Click(object sender, System.EventArgs e)
		{
			this.ReferenceOptimizerScenarioForm.m_bSave=true;
			SaveToCurrentVariables();
			
			this.grpboxFVSVariablesPrePostVariable.Hide();
			
			ReferenceOptimizerScenarioForm.m_bEnableSelectedTabPage=true;
			ReferenceOptimizerScenarioForm.m_strCurrentEditTabControlName="";
			this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabPageText="";
			m_intCurVar=-1;
			EnableTabs(true);
			this.grpboxFVSVariablesPrePost.Show();
		}
		private void SaveToCurrentVariables()
		{
			if (this.m_intCurVar != -1)
			{
				this.m_oCurVar.m_strPostVarArray[this.m_intCurVar-1] = this.lblFVSVariablesPrePostVariablePostSelected.Text;
				this.m_oCurVar.m_strPreVarArray[this.m_intCurVar-1] = this.lblFVSVariablesPrePostVariablePreSelected.Text;
			}
			this.m_oCurVar.Copy(m_oCurVar,ref m_oSavVar);
			if (this.m_intCurVar != -1)
			{
                if (lvFVSVariablesPrePostValues.SelectedItems.Count == 0 ||
                    lvFVSVariablesPrePostValues.SelectedItems[0].Index != m_intCurVar-1)
                {
                    lvFVSVariablesPrePostValues.Items[m_intCurVar - 1].Selected = true;
                }
				UpdateListViewVariableItem(this.lvFVSVariablesPrePostValues.SelectedItems[0].Index,m_intCurVar,m_oCurVar);
				
			}
    		EnableOverallExpr();
		}
		private void UpdateListViewVariableItem(int p_intListViewItem,int p_intVarArrayItem,uc_optimizer_scenario_fvs_prepost_variables_effective.Variables p_oVar)
		{
			this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_PREVAR].Text = p_oVar.m_strPreVarArray[p_intVarArrayItem-1];
			this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_POSTVAR].Text = p_oVar.m_strPostVarArray[p_intVarArrayItem-1];
			if (p_oVar.m_strBetterExpr[p_intVarArrayItem-1].Trim().Length > 0)
			{

				
				this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_BETTER].Text = "Yes";
			}
			else
			{
				
				this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_BETTER].Text = "No";
			}
			if (p_oVar.m_strWorseExpr[p_intVarArrayItem-1].Trim().Length > 0)
			{
				
				this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_WORSE].Text = "Yes";
			}
			else
			{
				
				this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_WORSE].Text = "No";
			}
			if (p_oVar.m_strEffectiveExpr[p_intVarArrayItem-1].Trim().Length > 0)
			{
				
				this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_EFFECTIVE].Text = "Yes";
			}
			else
			{
				
				this.lvFVSVariablesPrePostValues.Items[p_intListViewItem].SubItems[COLUMN_EFFECTIVE].Text = "No";
			}
		    

		}
		private void EnableOverallExpr()
		{

			if ((this.lvFVSVariablesPrePostValues.Items[0].SubItems[COLUMN_POSTVAR].Text.Trim().ToUpper() != 
				"NOT DEFINED" && 
				this.lvFVSVariablesPrePostValues.Items[0].SubItems[COLUMN_PREVAR].Text.Trim().ToUpper() !=
				"NOT DEFINED") ||
				(this.lvFVSVariablesPrePostValues.Items[1].SubItems[COLUMN_POSTVAR].Text.Trim().ToUpper() != 
				"NOT DEFINED" && 
				this.lvFVSVariablesPrePostValues.Items[1].SubItems[COLUMN_PREVAR].Text.Trim().ToUpper() !=
				"NOT DEFINED") ||
				(this.lvFVSVariablesPrePostValues.Items[2].SubItems[COLUMN_POSTVAR].Text.Trim().ToUpper() != 
				"NOT DEFINED" && 
				this.lvFVSVariablesPrePostValues.Items[2].SubItems[COLUMN_PREVAR].Text.Trim().ToUpper() !=
				"NOT DEFINED") ||
				(this.lvFVSVariablesPrePostValues.Items[3].SubItems[COLUMN_POSTVAR].Text.Trim().ToUpper() != 
				"NOT DEFINED" && 
				this.lvFVSVariablesPrePostValues.Items[3].SubItems[COLUMN_PREVAR].Text.Trim().ToUpper() !=
				"NOT DEFINED"))
				   
				   this.btnFVSVariablesPrePostOverall.Enabled=true;
				   
			else
				this.btnFVSVariablesPrePostOverall.Enabled=false;
				   

		}
			

		private void grpboxFVSVariablesPrePostOverallExp_Resize(object sender, System.EventArgs e)
		{
			
		}

		private void btnFVSVariablesPrePostExpressionNext_Click(object sender, System.EventArgs e)
		{
			
			if (this.lblFVSVariablesPrePostVariablePreSelected.Text.Trim().ToUpper() != "NOT DEFINED")
			{
				//save variables
				if (this.m_intCurVariableDefinitionStepCount==uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_BETTER)
				{
					m_oCurVar.m_strBetterExpr[this.m_intCurVar-1]=this.txtExpression.Text;
					//load worse expression
					this.loadvalues_variable(m_intCurVar-1,uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_WORSE);
				}
				else if (this.m_intCurVariableDefinitionStepCount==uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_WORSE)
				{
					m_oCurVar.m_strWorseExpr[this.m_intCurVar-1]=this.txtExpression.Text;
					//load effective expression
					this.loadvalues_variable(m_intCurVar-1,uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_EFFECTIVE);
				}
				
				this.txtExpression.Focus();
				this.txtExpression.Select(this.txtExpression.Text.Length,0);
			}
		}

		private void btnFVSVariablesPrePostExpressionDone_Click(object sender, System.EventArgs e)
		{
			this.ReferenceOptimizerScenarioForm.m_bSave=true;
			if (this.m_intCurVariableDefinitionStepCount==WIZARD_STEP_VARIABLE_BETTER)
				m_oCurVar.m_strBetterExpr[this.m_intCurVar-1]=this.txtExpression.Text;
			else if (this.m_intCurVariableDefinitionStepCount==WIZARD_STEP_VARIABLE_WORSE)
				m_oCurVar.m_strWorseExpr[this.m_intCurVar-1]=this.txtExpression.Text;
			else if (this.m_intCurVariableDefinitionStepCount==WIZARD_STEP_VARIABLE_EFFECTIVE)
				m_oCurVar.m_strEffectiveExpr[this.m_intCurVar-1]=this.txtExpression.Text;
			else if (this.m_intCurVariableDefinitionStepCount==WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
			{
				m_oCurVar.m_strOverallEffectiveExpr=this.txtExpression.Text;
			}

			

			SaveToCurrentVariables();
			
			this.grpboxFVSVariablesPrePostExpression.Hide();
			
			this.ReferenceOptimizerScenarioForm.m_bEnableSelectedTabPage=true;
			ReferenceOptimizerScenarioForm.m_strCurrentEditTabControlName="";
			this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabPageText="";
			this.m_intCurVar=-1;
			EnableTabs(true);
			this.grpboxFVSVariablesPrePost.Show();
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

		private void btnFVSVariablesPrePostExpressionCancel_Click(object sender, System.EventArgs e)
		{
			this.grpboxFVSVariablesPrePostExpression.Hide();
			
			this.ReferenceOptimizerScenarioForm.m_bEnableSelectedTabPage=true;
			ReferenceOptimizerScenarioForm.m_strCurrentEditTabControlName="";
			this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabPageText="";
			this.m_intCurVar=-1;
			EnableTabs(true);
			this.grpboxFVSVariablesPrePost.Show();
		}

		private void btnFVSVariablesPrePostOverallExp_Click(object sender, System.EventArgs e)
		{
			loadvalues_overall();
			

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

		private void grpboxFVSVariablesPrePostVariable3_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void lstFVSVariablesPrePostExpressionSelectedVariables_DoubleClick(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtExpression," " + this.lstFVSVariablesPrePostExpressionSelectedVariables.SelectedItem.ToString().Trim() + " ");
		}

	


		private void btnFVSVariablesPrePostExpressionClear_Click(object sender, System.EventArgs e)
		{
			this.txtExpression.Text = "";
		}

		private void btnFVSVariablesPrePostExpressionDefault_Click(object sender, System.EventArgs e)
		{
			string str="";
			switch (this.m_intCurVariableDefinitionStepCount)
			{
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_BETTER:
					this.txtExpression.Text = "post_variable" + this.m_intCurVar.ToString().Trim() + "_value > pre_variable" + this.m_intCurVar.ToString().Trim() + "_value + 10 AND " + 
						                      "post_variable" + this.m_intCurVar.ToString().Trim() + "_value <> -1 AND pre_variable" + this.m_intCurVar.ToString().Trim() + "_value <> -1";
					break;
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_WORSE:
					this.txtExpression.Text = "post_variable" + this.m_intCurVar.ToString().Trim() + "_value < pre_variable" + this.m_intCurVar.ToString().Trim() + "_value - 10 AND " + 
											  "post_variable" + this.m_intCurVar.ToString().Trim() + "_value <> -1 AND pre_variable" + this.m_intCurVar.ToString().Trim() + "_value <> -1";;
					break;
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_EFFECTIVE:
					if (this.m_oCurVar.m_strBetterExpr[m_intCurVar-1].Trim().Length > 0 &&
						this.m_oCurVar.m_strWorseExpr[m_intCurVar-1].Trim().Length > 0)
					{
						this.txtExpression.Text = "variable" + this.m_intCurVar.ToString().Trim() + "_better_yn='Y' AND " + 
												  "variable" + this.m_intCurVar.ToString().Trim() + "_worse_yn='N'";
					}
					else if (this.m_oCurVar.m_strBetterExpr[m_intCurVar-1].Trim().Length > 0)
					{
						this.txtExpression.Text = "variable" + this.m_intCurVar.ToString().Trim() + "_better_yn='Y'";
					}
					else if (this.m_oCurVar.m_strWorseExpr[m_intCurVar-1].Trim().Length > 0)
					{
						this.txtExpression.Text = "variable" + this.m_intCurVar.ToString().Trim() + "_worse_yn='N'";
					}
					break;
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE:
					if (this.m_oCurVar.m_strEffectiveExpr[0].Trim().Length > 0)
						str="variable1_effective_yn='Y' AND ";
					if (this.m_oCurVar.m_strEffectiveExpr[1].Trim().Length > 0)
					    str=str + "variable2_effective_yn='Y' AND ";
					if (this.m_oCurVar.m_strEffectiveExpr[2].Trim().Length > 0)
						str=str + "variable3_effective_yn='Y' AND ";
					if (this.m_oCurVar.m_strEffectiveExpr[3].Trim().Length > 0)
						str=str + "variable4_effective_yn='Y' AND ";

					//get rid of the and 
					if (str.Trim().Length > 0)
						str = str.Substring(0,str.Length - 5);
	
					this.txtExpression.Text = str;
					break;

					
						
				default:
					break;
			}
		}

		private void btnFVSVariablesPrePostGo_Click(object sender, System.EventArgs e)
		{
			string strFind="";
			if (this.m_intCurVar!=-1)
			{
				//see if we are navigating away from the current variable
				strFind = "VARIABLE " + this.m_intCurVar.ToString().Trim();
				if (this.cmbFVSVariablesPrePost.Text.ToUpper().IndexOf(strFind,0) < 0)
				{
					if (Modified())
					{
						DialogResult result = MessageBox.Show("Save current variable changes(Y/N)?","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
						if (result == DialogResult.Yes)
						{
							this.SaveToCurrentVariables();
						}
					}
				}

			}
			else if (this.m_intCurVariableDefinitionStepCount == uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
			{
				if (Modified())
				{
					DialogResult result = MessageBox.Show("Save current variable changes(Y/N)?","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
					if (result == DialogResult.Yes)
					{
						this.SaveToCurrentVariables();
					}
				}
			}
			else
			{
                if (this.m_oSavVar == null) m_oSavVar = new Variables();

				this.m_oCurVar = new Variables();
				this.m_oCurVar.Copy(this.m_oSavVar,ref m_oCurVar);
			}
			switch (this.cmbFVSVariablesPrePost.Text.Trim().ToUpper())
			{
				case "SELECT VARIABLE 1":
					if (this.m_intCurVar!=1)
					{
					}
					this.loadvalues_variable(0,WIZARD_STEP_VARIABLE_SELECT);
					ShowGroupBox(this.grpboxFVSVariablesPrePostVariable.Name);
					
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 1 POST-TREATMENT IMPROVEMENT (BETTER)":
					if (this.m_intCurVar!=1)
					{
					}
					if (m_oSavVar.m_strPostVarArray[0].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable1 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(0,WIZARD_STEP_VARIABLE_BETTER);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 1 POST-TREATMENT DISIMPROVEMENT (WORSE)":
					if (this.m_intCurVar!=1)
					{
					}
					if (m_oSavVar.m_strPostVarArray[0].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable1 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(0,WIZARD_STEP_VARIABLE_WORSE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 1 EFFECTIVE TREATMENT":
					if (this.m_intCurVar!=1)
					{
					}
					if (m_oSavVar.m_strPostVarArray[0].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable1 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(0,WIZARD_STEP_VARIABLE_EFFECTIVE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "SELECT VARIABLE 2":
					if (this.m_intCurVar!=2)
					{
					}
					this.loadvalues_variable(1,WIZARD_STEP_VARIABLE_SELECT);
					ShowGroupBox(this.grpboxFVSVariablesPrePostVariable.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 2 POST-TREATMENT IMPROVEMENT (BETTER)":
					if (this.m_intCurVar!=2)
					{
					}
					if (m_oSavVar.m_strPostVarArray[1].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable2 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(1,WIZARD_STEP_VARIABLE_BETTER);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
                case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 2 POST-TREATMENT DISIMPROVEMENT (WORSE)":
					if (this.m_intCurVar!=2)
					{
					}
					if (m_oSavVar.m_strPostVarArray[1].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable2 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(1,WIZARD_STEP_VARIABLE_WORSE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 2 EFFECTIVE TREATMENT":
					if (this.m_intCurVar!=2)
					{
					}
					if (m_oSavVar.m_strPostVarArray[1].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable2 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(1,WIZARD_STEP_VARIABLE_EFFECTIVE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "SELECT VARIABLE 3":
					if (this.m_intCurVar!=3)
					{
					}
					this.loadvalues_variable(2,WIZARD_STEP_VARIABLE_SELECT);
					ShowGroupBox(this.grpboxFVSVariablesPrePostVariable.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 3 POST-TREATMENT IMPROVEMENT (BETTER)":
					if (this.m_intCurVar!=3)
					{
					}
					if (m_oSavVar.m_strPostVarArray[2].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable3 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(2,WIZARD_STEP_VARIABLE_BETTER);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
                case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 3 POST-TREATMENT DISIMPROVEMENT (WORSE)":
					if (this.m_intCurVar!=3)
					{
					}
					if (m_oSavVar.m_strPostVarArray[2].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable3 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(2,WIZARD_STEP_VARIABLE_WORSE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 3 EFFECTIVE TREATMENT":
					if (this.m_intCurVar!=3)
					{
					}
					if (m_oSavVar.m_strPostVarArray[2].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable3 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(2,WIZARD_STEP_VARIABLE_EFFECTIVE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;

				case "SELECT VARIABLE 4":
					if (this.m_intCurVar!=4)
					{
					}
					this.loadvalues_variable(3,WIZARD_STEP_VARIABLE_SELECT);
					ShowGroupBox(this.grpboxFVSVariablesPrePostVariable.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 4 POST-TREATMENT IMPROVEMENT (BETTER)":
					if (this.m_intCurVar!=4)
					{
					}
					if (m_oSavVar.m_strPostVarArray[3].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable4 not defined","FIA Biosum");
						return;
					}
     				this.loadvalues_variable(3,WIZARD_STEP_VARIABLE_BETTER);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
                case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 4 POST-TREATMENT DISIMPROVEMENT (WORSE)":
					if (this.m_intCurVar!=4)
					{
					}
					if (m_oSavVar.m_strPostVarArray[3].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable4 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(3,WIZARD_STEP_VARIABLE_WORSE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES VARIABLE 4 EFFECTIVE TREATMENT":
					if (this.m_intCurVar!=4)
					{
					}
					if (m_oSavVar.m_strPostVarArray[3].Trim().ToUpper() == "NOT DEFINED")
					{
						MessageBox.Show("Variable4 not defined","FIA Biosum");
						return;
					}
					this.loadvalues_variable(3,WIZARD_STEP_VARIABLE_EFFECTIVE);
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;
				case "DEFINE EXPRESSION FOR WHAT CONSTITUTES OVERALL EFFECTIVENESS":
					this.loadvalues_overall();
					ShowGroupBox(grpboxFVSVariablesPrePostExpression.Name);
					break;




			}
		    if (this.grpboxFVSVariablesPrePost.Visible==false) EnableTabs(false);
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
		private bool Modified()
		{

			if (this.m_oCurVar!=null)
			{
				if (this.m_intCurVar != -1)
				{
					this.m_oCurVar.m_strPostVarArray[m_intCurVar-1] = lblFVSVariablesPrePostVariablePostSelected.Text.Trim();
					this.m_oCurVar.m_strPreVarArray[m_intCurVar-1] = lblFVSVariablesPrePostVariablePreSelected.Text.Trim();
					switch (this.m_intCurVariableDefinitionStepCount)
					{
						case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_BETTER:
							this.m_oCurVar.m_strBetterExpr[m_intCurVar-1] = this.txtExpression.Text;
							break;
						case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_WORSE:
							this.m_oCurVar.m_strWorseExpr[m_intCurVar-1] = this.txtExpression.Text;
							break;
						case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_EFFECTIVE:
							this.m_oCurVar.m_strEffectiveExpr[m_intCurVar-1] = this.txtExpression.Text;
							break;
						case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE:
							this.m_oCurVar.m_strOverallEffectiveExpr = this.txtExpression.Text;
							break;
					
					}
				
				}
				else if (this.m_intCurVariableDefinitionStepCount==uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE)
				{
					this.m_oCurVar.m_strOverallEffectiveExpr = this.txtExpression.Text;
				}

				if (this.m_oCurVar.Modified(this.m_oSavVar)) return true;
			}

			return false;
            

		}

		private void btnFVSVariablesPrePostValuesButtonsEdit_Click(object sender, System.EventArgs e)
		{
            if (this.lvFVSVariablesPrePostValues.SelectedItems.Count == 0)
            {
                MessageBox.Show("!! No variable selected to edit!", "FIA Biosum", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            this.loadFVSTableAndField();
            this.lblFVSVariablesPrePostVariablePreSelected.Text = this.lvFVSVariablesPrePostValues.SelectedItems[0].SubItems[COLUMN_PREVAR].Text.Trim();
            this.lblFVSVariablesPrePostVariablePostSelected.Text = this.lvFVSVariablesPrePostValues.SelectedItems[0].SubItems[COLUMN_POSTVAR].Text.Trim();
			EditVariable();

		}
		private void EditVariable()
		{
			if (this.lvFVSVariablesPrePostValues.SelectedItems.Count==0) return;

			this.m_oCurVar = new Variables();
			this.m_oCurVar.Copy(this.m_oSavVar,ref m_oCurVar);

			this.loadvalues_variable(this.lvFVSVariablesPrePostValues.SelectedItems[0].Index,1);
			
			this.grpboxFVSVariablesPrePost.Hide();

			this.ReferenceOptimizerScenarioForm.m_bEnableSelectedTabPage=false;
			this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabControlName="tabControlFVSVariables";
			this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabPageText="Effective";
			this.ReferenceOptimizerScenarioForm.m_intEditTabPageIndex=0;
			EnableTabs(false);
			grpboxFVSVariablesPrePostVariable.Show();
		}

		
		private void btnFVSVariablesPrePost2Overall_Click(object sender, System.EventArgs e)
		{
			loadvalues_overall();
		}

		private void btnFVSVariablesPrePostValuesButtonsClear_Click(object sender, System.EventArgs e)
		{
            if (this.lvFVSVariablesPrePostValues.SelectedItems.Count == 0)
            {
                MessageBox.Show("!! No variable selected to clear!", "FIA Biosum", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
			if (this.lvFVSVariablesPrePostValues.SelectedItems[0].SubItems[COLUMN_PREVAR].Text.Trim()=="Not Defined") return;

			DialogResult result = MessageBox.Show("Are you sure you wish to delete variable " + this.lvFVSVariablesPrePostValues.SelectedItems[0].SubItems[COLUMN_VARID].Text + "? (YN)","FIA Biosum",System.Windows.Forms.MessageBoxButtons.YesNo,System.Windows.Forms.MessageBoxIcon.Question);
			if (result == System.Windows.Forms.DialogResult.Yes) 
				RemoveVariable(this.lvFVSVariablesPrePostValues.SelectedItems[0].Index);
		}

		private void RemoveVariable(int p_intIndex)
		{
			this.m_oSavVar.m_strPreVarArray[p_intIndex]="Not Defined";
			this.m_oSavVar.m_strPostVarArray[p_intIndex]="Not Defined";
			this.m_oSavVar.m_strBetterExpr[p_intIndex] = "";
			this.m_oSavVar.m_strWorseExpr[p_intIndex]="";
			this.m_oSavVar.m_strEffectiveExpr[p_intIndex]="";

			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_PREVAR].Text = "Not Defined";
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_POSTVAR].Text = "Not Defined";
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_BETTER].ForeColor=Color.Gray;
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_BETTER].Text = "No";
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_WORSE].ForeColor=Color.Gray;
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_WORSE].Text = "No";
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_EFFECTIVE].ForeColor=Color.Gray;
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].SubItems[COLUMN_EFFECTIVE].Text = "No";
			this.lvFVSVariablesPrePostValues.Items[p_intIndex].Checked=false;

			int x=0;
			//check to see if all the variables are not defined, if so, initialize overall effectiveness.
			for (x=0;x<=NUMBER_OF_VARIABLES-1;x++)
			{
				if (this.m_oSavVar.m_strPostVarArray[x]!="Not Defined") break;
			}
			if (x > NUMBER_OF_VARIABLES-1) 
			{
					this.m_oSavVar.m_strOverallEffectiveExpr="";
				    this.btnFVSVariablesPrePostOverall.Enabled=false;
			}

			if (m_oCurVar != null) this.m_oSavVar.Copy(m_oSavVar,ref m_oCurVar);


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

		private void btnFVSVariablesPrePostOverall_Click(object sender, System.EventArgs e)
		{

			   
			    this.ReferenceOptimizerScenarioForm.m_bEnableSelectedTabPage=false;
			    this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabControlName="tabControlFVSVariables";
			    this.ReferenceOptimizerScenarioForm.m_strCurrentEditTabPageText="Effective";

			    this.ReferenceOptimizerScenarioForm.m_intEditTabPageIndex=0;
			    EnableTabs(false);
				loadvalues_overall();
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
		private void Audit()
        {
			DataMgr oDataMgr = new DataMgr();
			string str = "";
			string strTable = "";
			string strColumn = "";
			//bool bOptimized=false;
			int x, y;
			this.m_intError = 0;
			this.m_strError = "Audit Results \r\n";
			this.m_strError = m_strError + "-------------\r\n\r\n";
			string fvsDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + Tables.FVS.DefaultFVSOutPrePostDbFile;
			using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(fvsDb)))
            {
				conn.Open();

				string weightedDb = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableSqliteDbFile;
				if (!oDataMgr.DatabaseAttached(conn, weightedDb))
				{
					oDataMgr.m_strSQL = "ATTACH DATABASE '" + weightedDb + "' AS weighted";
					oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
				}
				
				if (oDataMgr.m_intError == 0)
                {
					for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
                    {
						if (m_oSavVar.m_strPreVarArray[x].Trim().ToUpper() != "NOT DEFINED")
                        {
							strTable = this.m_oSavVar.TableName(m_oSavVar.m_strPreVarArray[x]);
							if (oDataMgr.TableExist(conn, strTable) || oDataMgr.AttachedTableExist(conn, strTable))
                            {
								strColumn = m_oSavVar.ColumnName(m_oSavVar.m_strPreVarArray[x]);
								if (!oDataMgr.ColumnExist(conn, strTable, strColumn) && !oDataMgr.AttachedColumnExist(conn, strTable, strColumn))
                                {
									m_intError = -1;
									m_strError = m_strError + "Table column " + m_oSavVar.m_strPreVarArray[x] + " does not exist in Db files PREPOST_FVSOUT.db or prepost_fvs_weighted.db\r\n";
								}
							}
							else
                            {
								m_intError = -1;
								m_strError = m_strError + "Table " + strTable + " does not exist in Db files PREPOST_FVSOUT.db or prepost_fvs_weighted.db\r\n";
							}
						}

					}

				}
			}
			for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
            {
				if (m_oSavVar.m_strPreVarArray[x].Trim().ToUpper() == "NOT DEFINED")
				{
					this.AuditSearch("VARIABLE" + Convert.ToString(x + 1).Trim());
				}
				else
				{

					if (this.m_oSavVar.m_strBetterExpr[x].Trim().Length == 0)
					{
						this.AuditSearch("VARIABLE" + Convert.ToString(x + 1).Trim() + "_BETTER_YN");
					}
					if (this.m_oSavVar.m_strWorseExpr[x].Trim().Length == 0)
					{
						this.AuditSearch("VARIABLE" + Convert.ToString(x + 1).Trim() + "_WORSE_YN");
					}
					if (this.m_oSavVar.m_strEffectiveExpr[x].Trim().Length == 0)
					{
						this.AuditSearch("VARIABLE" + Convert.ToString(x + 1).Trim() + "_EFFECTIVE_YN");
					}

					string[] strPieces = m_oSavVar.m_strPreVarArray[x].Trim().ToUpper().Split('.');
					if (strPieces.Length == 2)
					{
						if (strPieces[0].ToUpper().Contains("_WEIGHTED"))
						{
							// Strip PRE_ prefix from the table before sending it to the shared AuditWeightedFvsVariables
							// function so that it can be shared with optimizer and tie break
							char[] preArray = "PRE_".ToCharArray();
							string strTempTableName = strPieces[0].TrimStart(preArray);
							string strWeightedError = this.m_oOptimizerScenarioTools.AuditWeightedFvsVariables(strTempTableName, out m_intError);
							if (m_intError != 0)
							{
								m_strError = m_strError + strWeightedError;
							}
						}
					}
				}
			}
			if (m_oSavVar.m_strOverallEffectiveExpr.Trim().Length == 0)
			{
				m_intError = -1;
				m_strError = m_strError + "Overall effectiveness expression not defined. \r\n";
			}
			//make sure no duplicates exist
			for (x = 0; x <= NUMBER_OF_VARIABLES - 1; x++)
			{
				if (this.m_oSavVar.m_strPreVarArray[x].Trim().ToUpper() != "NOT DEFINED")
				{
					if (str.IndexOf("," + this.m_oSavVar.m_strPreVarArray[x].Trim() + ",", 0) < 0)
					{
						str = str + "," + this.m_oSavVar.m_strPreVarArray[x].Trim() + ",";
						for (y = x + 1; y <= NUMBER_OF_VARIABLES - 1; y++)
						{
							if (this.m_oSavVar.m_strPreVarArray[x].Trim() ==
								this.m_oSavVar.m_strPreVarArray[y].Trim())
							{
								m_intError = -1;
								m_strError = m_strError + m_oSavVar.m_strPreVarArray[x] + " is defined multiple times\r\n";
								break;
							}
						}
					}
				}

			}
			AuditEffectiveExpressionDefined(m_oSavVar.m_strEffectiveExpr);

			this.ReferenceOptimizationUserControl.ReferenceFVSVariables = this.m_oSavVar;


			this.ReferenceOptimizationUserControl.DisplayAuditMessage = false;

			this.ReferenceOptimizationUserControl.Audit();

			this.m_strError = m_strError + ReferenceOptimizationUserControl.m_strError;
			x = ReferenceOptimizationUserControl.m_intError;
			if (x != 0)
				this.m_intError = x;

			if (m_intError == 0) this.m_strError = m_strError + "Passed Audit";
			else m_strError = m_strError + "\r\n\r\n" + "Failed Audit";

			if (this.DisplayAuditMessage)
				MessageBox.Show(m_strError, "FIA Biosum");

		}
        /// <summary>
        /// Check to make sure that at least one effective expression is defined.
        /// </summary>
        /// <param name="p_strExpressions"></param>
        /// <returns></returns>
        private void AuditEffectiveExpressionDefined(string[] p_strExpressions)
        {

            if (p_strExpressions == null)
            {
                m_strError = m_strError + "No effective expression defined. At least 1 effective expression must be defined.\r\n";
                m_intError = -1;
                return;
            }

            for (int x = 0; x <= p_strExpressions.Length - 1; x++)
            {
                if (p_strExpressions[x] != null)
                {
                    if (p_strExpressions[x].Trim().Length > 0)
                    {
                        return;
                    }
                }
            }
            m_strError = m_strError + "No effective expression defined. At least 1 effective expression must be defined.\r\n";
            m_intError = -1;   
        }
        private void AuditSearch(string p_strSearchFor)
		{
			int x;
			for (x=0;x<=NUMBER_OF_VARIABLES-1;x++)
			{
				for (x=0;x<=NUMBER_OF_VARIABLES-1;x++)
				{
					if (m_oSavVar.m_strBetterExpr[x].Trim().ToUpper().IndexOf(p_strSearchFor.ToUpper(),0) >= 0)
					{
						m_intError=-1;
						this.m_strError=this.m_strError + p_strSearchFor + " is not defined and is referenced in variable" + Convert.ToString(x+1).Trim() + " improved (better) expression. \r\n";
					}
					if (m_oSavVar.m_strWorseExpr[x].Trim().ToUpper().IndexOf(p_strSearchFor.ToUpper(),0) >= 0)
					{
						m_intError=-1;
						this.m_strError=this.m_strError + p_strSearchFor + " is not defined and is referenced in variable" + Convert.ToString(x+1).Trim() + " regressed (worse) expression. \r\n";
					}
					if (m_oSavVar.m_strEffectiveExpr[x].Trim().ToUpper().IndexOf(p_strSearchFor.ToUpper(),0) >= 0)
					{
						m_intError=-1;
						this.m_strError=this.m_strError + p_strSearchFor + " is not defined and is referenced in variable" + Convert.ToString(x+1).Trim() + " effective expression. \r\n";
					}
				}
				if (m_oSavVar.m_strOverallEffectiveExpr.Trim().ToUpper().IndexOf(p_strSearchFor.ToUpper(),0) >= 0)
				{
					m_intError=-1;
					this.m_strError=this.m_strError + p_strSearchFor + " is not defined and is referenced in the overall effective expression. \r\n";
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

		private void btnFVSVariablesPrePostExpressionPrev_Click(object sender, System.EventArgs e)
		{
			
			
			string strConn="";
			string strSQL="";

			DialogResult result;
            
			string strScenarioId =  ((frmOptimizerScenario)this.ParentForm).uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioDB = 
				((frmMain)this.ParentForm.ParentForm).frmProject.uc_project1.m_strProjectDirectory + "\\" +
                    Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;

			frmDialog frmPrevExp = new frmDialog();
				
			frmPrevExp.Width = frmPrevExp.uc_previous_expressions1.m_intFullWd;
			frmPrevExp.Height = frmPrevExp.uc_previous_expressions1.m_intFullHt;
			frmPrevExp.Text = "Treatment Optimizer: Previous FVS Variables SQL Expressions";
					
			frmPrevExp.uc_previous_expressions1.Visible=true;

			DataMgr p_dataMgr = new DataMgr();
			strConn = p_dataMgr.GetConnectionString(strScenarioDB);

			switch (this.m_intCurVariableDefinitionStepCount)
			{
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_BETTER:
					strSQL = "SELECT scenario_id,pre_fvs_variable,better_expression,current_yn FROM scenario_fvs_variables";
					frmPrevExp.uc_previous_expressions1.loadvalues(strConn,strSQL,"BETTER_EXPRESSION","BETTER_EXPRESSION", "scenario_fvs_variables");
					break;
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_WORSE:
					strSQL = "SELECT scenario_id,pre_fvs_variable,worse_expression,current_yn FROM scenario_fvs_variables";
					frmPrevExp.uc_previous_expressions1.loadvalues(strConn,strSQL,"WORSE_EXPRESSION","WORSE_EXPRESSION", "scenario_fvs_variables");
					break;
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLE_EFFECTIVE:
					strSQL = "SELECT scenario_id,fvs_variables_list,pre_fvs_variable,effective_expression,current_yn FROM scenario_fvs_variables";
					frmPrevExp.uc_previous_expressions1.loadvalues(strConn,strSQL,"EFFECTIVE_EXPRESSION","EFFECTIVE_EXPRESSION", "scenario_fvs_variables");
					break;
				case uc_optimizer_scenario_fvs_prepost_variables_effective.WIZARD_STEP_VARIABLES_OVERALL_EFFECTIVE:
					strSQL = "SELECT scenario_id,fvs_variables_list,overall_effective_expression,current_yn FROM scenario_fvs_variables_overall_effective";
					frmPrevExp.uc_previous_expressions1.loadvalues(strConn,strSQL,"OVERALL_EFFECTIVE_EXPRESSION","OVERALL_EFFECTIVE_EXPRESSION", "scenario_fvs_variables_overall_effective");
					break;
					
			}
			
			
            
			result = frmPrevExp.ShowDialog(this);
			if (result == DialogResult.OK)
			{
				result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtExpression.Text + "' \n\n\n WITH  \n----------------\n\n'" + frmPrevExp.uc_previous_expressions1.lblSQL.Text + "'", "Previous SQL",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
				if (result == DialogResult.Yes)
				{
					this.txtExpression.Text = frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[frmPrevExp.uc_previous_expressions1.m_intSelectColumn + 1].Text;
					((frmOptimizerScenario)this.ParentForm).m_bSave=true;
				}
			}
			frmPrevExp.Close();
			frmPrevExp = null;
		}

		private void grpboxFVSVariablesPrePost_VisibleChanged(object sender, System.EventArgs e)
		{
			if (grpboxFVSVariablesPrePost.Visible==false) 
			{

				ReferenceOptimizerScenarioForm.m_strAllowLeaveTabPageMsg="Cannot edit Optimization until finished editing FVS Variable";

			}
			else 
			{
				
				ReferenceOptimizerScenarioForm.m_strAllowLeaveTabPageMsg="";

			}
		}


		
		private void EnableTabs(bool p_bEnable)
		{

            int x;
			ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlScenario,"tbdesc,tbnotes,tbdatasources",p_bEnable);
			ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlRules,"tbpsites,tbowners,tbcost,tbtreatmentintensity,tbfilterplots,tbrun",p_bEnable);
			ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlFVSVariables,"tboptimization,tbtiebreaker",p_bEnable);
            ReferenceOptimizerScenarioForm.EnableTabPage(ReferenceOptimizerScenarioForm.tabControlCosts, "tbcosts", p_bEnable);
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

		private void lvFVSVariablesPrePostValues_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			try
			{
				if (e.Button == MouseButtons.Left)
				{
					int intRowHt = lvFVSVariablesPrePostValues.Items[0].Bounds.Height;
					double dblRow = (double)(e.Y / intRowHt);
					this.lvFVSVariablesPrePostValues.Items[lvFVSVariablesPrePostValues.TopItem.Index + (int)dblRow-1].Selected=true;
					
				}
			}
			catch 
			{
			}

		}

		private void lvFVSVariablesPrePostValues_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (this.lvFVSVariablesPrePostValues.SelectedItems.Count > 0)
				this.m_oLvRowColors.DelegateListViewItem(this.lvFVSVariablesPrePostValues.SelectedItems[0]);
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
		public FIA_Biosum_Manager.uc_optimizer_scenario_fvs_prepost_variables_tiebreaker ReferenceTieBreakerUserControl
		{
			get {return _uc_tiebreaker;}
			set {_uc_tiebreaker=value;}
		}

        private void lstFVSTablesList_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstFVSFieldsList.Items.Clear();
            txtEffVarDescr.Text = "";
            this.lblFVSVariablesPrePostVariablePreSelected.Text = "Not Defined";
            this.lblFVSVariablesPrePostVariablePostSelected.Text = "Not Defined";
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
                lblEffVarDescr.Visible = bDisplayWeightedFields;
                txtEffVarDescr.Visible = bDisplayWeightedFields;
            }
            else
            {
                lblEffVarDescr.Visible = false;
                txtEffVarDescr.Visible = false;
            }

        }

        private void lstFVSFieldsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.lblFVSVariablesPrePostVariablePreSelected.Text = "Not Defined";
            this.lblFVSVariablesPrePostVariablePostSelected.Text = "Not Defined";
            this.txtEffVarDescr.Text = "";
            if (this.lstFVSFieldsList.SelectedIndex > -1)
            {
                this.btnFVSVariablesPrePostVariableValue.Enabled = true;
                if (this.lstFVSTablesList.SelectedItems[0].ToString().ToUpper().Contains("_WEIGHTED") == true)
                {
                    foreach (uc_optimizer_scenario_calculated_variables.VariableItem oItem in this.ReferenceOptimizerScenarioForm.m_oWeightedVariableCollection)
                    {
                        if (oItem.strVariableName.Equals(Convert.ToString(this.lstFVSFieldsList.SelectedItem)))
                        {
                            if (!String.IsNullOrEmpty(oItem.strVariableDescr))
                                txtEffVarDescr.Text = oItem.strVariableDescr;
                            break;
                        }
                    }
                }                
            }
            else
            {
                this.btnFVSVariablesPrePostVariableValue.Enabled = false;
            }
        }

        private void grpboxFVSVariablesPrePostExpression_VisibleChanged(object sender, EventArgs e)
        {
            if (grpboxFVSVariablesPrePostExpression.Visible == false)
            {
                m_strHelpChapter = "EFFECTIVE_SETTINGS";
            }
            else
            {
                m_strHelpChapter = "EFFECTIVE_EXPRESSION_BUILDER";
            }
        }

        public string HelpChapter
        {
            get
            {
                return m_strHelpChapter;
            }
        }
       
	
	}
}

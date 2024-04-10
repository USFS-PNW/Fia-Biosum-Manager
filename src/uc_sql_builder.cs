using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_sql_builder.
	/// </summary>
	public class uc_sql_builder : System.Windows.Forms.UserControl
	{
		private env m_oEnv;
		public System.Windows.Forms.TextBox txtSQLCommand;
		private System.Windows.Forms.Label label1;
		public System.Windows.Forms.ListBox lstTables;
		private System.Windows.Forms.Button btnGetTables;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ListBox lstFields;
		private System.Windows.Forms.ListBox lstValues;
		private System.Windows.Forms.Button btnUnion;
		private System.Windows.Forms.Label lblFilterCondition;
		private System.Windows.Forms.Button btnHaving;
		private System.Windows.Forms.Label lblFieldList2;
		private System.Windows.Forms.Button btnOn;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ComboBox cmbJoinType;
		private System.Windows.Forms.Label lblDBList;
		private System.Windows.Forms.Button btnJoin;
		private System.Windows.Forms.Button btnFrom;
		private System.Windows.Forms.Button btnAs;
		private System.Windows.Forms.Button btnAsterick;
		private System.Windows.Forms.Label lblFieldList;
		private System.Windows.Forms.Button btnDistinct;
		private System.Windows.Forms.Button btnSelect;
		private System.Windows.Forms.GroupBox grpboxSelectSyntax;
		private System.Windows.Forms.Button btnAll;
		private System.Windows.Forms.Label lblSqlExp;
		private System.Windows.Forms.Label lblFieldList3;
		private System.Windows.Forms.Button btnOrderBy;
		private System.Windows.Forms.Button btnAsc;
		private System.Windows.Forms.Button btnDesc;
		private System.Windows.Forms.Label lbljoincondexp;
		private System.Windows.Forms.GroupBox grpboxVariableManipulation;
		private System.Windows.Forms.GroupBox grpboxCondExp;
		private System.Windows.Forms.GroupBox grpboxMathFunc;
		private System.Windows.Forms.Button btnAnd;
		private System.Windows.Forms.Button btnOr;
		private System.Windows.Forms.Button btnLike;
		private System.Windows.Forms.Button btnBetween;
		private System.Windows.Forms.Button btnNot;
		private System.Windows.Forms.Button btnIn;
		private System.Windows.Forms.Button btnIIF;
		private System.Windows.Forms.Button btnDoubleQuote;
		private System.Windows.Forms.Button btnSingleQuote;
		private System.Windows.Forms.Button btnParenthLeft;
		private System.Windows.Forms.Button btnParenthRight;
		private System.Windows.Forms.Button btnComma;
		private System.Windows.Forms.Button btnPercent;
		private System.Windows.Forms.Button btnEqual;
		private System.Windows.Forms.Button btnNotEqual;
		private System.Windows.Forms.Button btnMoreThan;
		private System.Windows.Forms.Button btnMoreThanEqualTo;
		private System.Windows.Forms.Button btnLessThanEqualTo;
		private System.Windows.Forms.Button btnLessThan;
		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.Button btnSubtract;
		private System.Windows.Forms.Button btnMultiply;
		private System.Windows.Forms.Button btnDivide;
		private System.Windows.Forms.Button btnExponent;
		private System.Windows.Forms.Button btnAvg;
		private System.Windows.Forms.Button btnCount;
		private System.Windows.Forms.Button btnMin;
		private System.Windows.Forms.Button btnMax;
		private System.Windows.Forms.Button btnSum;
		private System.Windows.Forms.Button btnGroupBy;
		private System.Windows.Forms.Button btnIsNull;
		private System.Windows.Forms.Button btnTrim;
		private System.Windows.Forms.Button btnUCase;
		private System.Windows.Forms.Button btnLCase;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.Button btnClear;
		public System.Windows.Forms.Button btnClipboardCopy;
		private System.Windows.Forms.Button btnClipboardPaste;
		public int m_intFullHt=700;
		public int m_intFullWd=744;
		public System.Data.OleDb.OleDbConnection m_OleDbConnectionScenario;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Button btnExecute;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.Label lblMsg;
		private System.Windows.Forms.Button btnSort;
		private System.Windows.Forms.TextBox txtSearchValues;
		private System.Windows.Forms.Button btnFind;
		private string m_strCurTable;
		private string m_strCurField;
		private System.Windows.Forms.CheckBox chkValues;
		private System.Windows.Forms.Label lblWhere;
		private System.Windows.Forms.Label lblFieldAlias;
		private System.Windows.Forms.Button btnMid;
		private System.Windows.Forms.Button btnPrevSQL;
		public FIA_Biosum_Manager.frmGridView frmGridView1;
		public string[] m_strListedTablesLoadedIntoDatasets;
		public int m_intNumberOfTablesLoadedIntoDatasets;
		public int m_intNumberOfValidTables;
		public string m_strTempMDBFile;
		public System.Data.OleDb.OleDbConnection m_TempMDBFileConn;
		public System.Windows.Forms.ListView lvDataSource;
		public System.Data.DataSet m_dsDataSource;
		const int TABLETYPE = 0;
		const int PATH = 1;
		const int MDBFILE = 2;
		const int FILESTATUS = 3;
		const int TABLE = 4;
		const int TABLESTATUS = 5;
		const int RECORDCOUNT = 6;
		public int m_intError=0;
		private System.Data.OleDb.OleDbDataAdapter m_OleDbDataAdapter;
		private System.Windows.Forms.Button btnExists;
		private FIA_Biosum_Manager.utils m_oUtils;
		private System.Windows.Forms.Button btnN;
		private System.Windows.Forms.Button btnY;

		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public uc_sql_builder()
		{
            
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
            
            this.m_oEnv = new env();
            
            this.m_oUtils = new utils();
            
            this.m_strCurTable = "";
            
            this.m_strCurField="";
            
            this.m_strListedTablesLoadedIntoDatasets = new string[50];
            
            this.frmGridView1 = new frmGridView();
            
            this.frmGridView1.MdiParent = this.ParentForm;
            
            this.m_OleDbDataAdapter = new System.Data.OleDb.OleDbDataAdapter();
            
            this.m_dsDataSource = new DataSet();
            
            this.frmGridView1.Hide();
            // TODO: Add any initialization after the InitializeComponent call
            
        }
		~uc_sql_builder()
		{
			this.m_oUtils = null;
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
		public void SendKeyStrokes(System.Windows.Forms.TextBox p_oTextBox, string strKeyStrokes)
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

		public void SendSingleKeyStrokes(System.Windows.Forms.TextBox p_oTextBox, string strKeyStrokes)
		{
			string strKeyStroke="";
			p_oTextBox.Focus();
			try 
			{
				//MessageBox.Show(strKeyStrokes);
				
				
				//while (p_oTextBox.Focus()==false )
				//{
//
//				}
				for (int x=0;x<=strKeyStrokes.Length-1;x++)
				{
					
					switch (strKeyStrokes.Substring(x,1))
					{
						case ")":
							strKeyStroke = "{)}";
							break;
						case "(":
							strKeyStroke = "{(}";
							break;
						case "%":
							strKeyStroke = "{%}";
							break;
						case "^":
							strKeyStroke = "{^}";
							break;
						case "+":
							strKeyStroke = "{+}";
							break;
						case "~":
							strKeyStroke = "{~}";
							break;
						case "[":
							strKeyStroke = "{[}";
							break;
						case "]":
							strKeyStroke = "{]}";
							break;
						case "{":
							strKeyStroke = "{{}";
							break;
						case "}":
							strKeyStroke = "{}}";
							break;
						default:
							strKeyStroke = strKeyStrokes.Substring(x,1).ToString();
							break;

					}
					//p_oTextBox.Focus();
					System.Windows.Forms.SendKeys.Send(strKeyStroke);
					//p_oTextBox.Focus();
					//p_oTextBox.Refresh();
				}
				p_oTextBox.Refresh();
			}
			catch  (Exception caught)
			{
				MessageBox.Show("SendKeyStrokes Method Failed With This Message:" + caught.Message);
			}

		}


		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnPrevSQL = new System.Windows.Forms.Button();
            this.chkValues = new System.Windows.Forms.CheckBox();
            this.btnFind = new System.Windows.Forms.Button();
            this.txtSearchValues = new System.Windows.Forms.TextBox();
            this.btnSort = new System.Windows.Forms.Button();
            this.lblMsg = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnExecute = new System.Windows.Forms.Button();
            this.btnClipboardPaste = new System.Windows.Forms.Button();
            this.btnClipboardCopy = new System.Windows.Forms.Button();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnTest = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.grpboxMathFunc = new System.Windows.Forms.GroupBox();
            this.btnSum = new System.Windows.Forms.Button();
            this.btnMax = new System.Windows.Forms.Button();
            this.btnMin = new System.Windows.Forms.Button();
            this.btnCount = new System.Windows.Forms.Button();
            this.btnAvg = new System.Windows.Forms.Button();
            this.btnExponent = new System.Windows.Forms.Button();
            this.btnDivide = new System.Windows.Forms.Button();
            this.btnMultiply = new System.Windows.Forms.Button();
            this.btnSubtract = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.grpboxCondExp = new System.Windows.Forms.GroupBox();
            this.btnN = new System.Windows.Forms.Button();
            this.btnY = new System.Windows.Forms.Button();
            this.btnExists = new System.Windows.Forms.Button();
            this.btnIsNull = new System.Windows.Forms.Button();
            this.btnLessThanEqualTo = new System.Windows.Forms.Button();
            this.btnLessThan = new System.Windows.Forms.Button();
            this.btnMoreThanEqualTo = new System.Windows.Forms.Button();
            this.btnMoreThan = new System.Windows.Forms.Button();
            this.btnNotEqual = new System.Windows.Forms.Button();
            this.btnEqual = new System.Windows.Forms.Button();
            this.btnPercent = new System.Windows.Forms.Button();
            this.btnComma = new System.Windows.Forms.Button();
            this.btnParenthRight = new System.Windows.Forms.Button();
            this.btnParenthLeft = new System.Windows.Forms.Button();
            this.btnSingleQuote = new System.Windows.Forms.Button();
            this.btnDoubleQuote = new System.Windows.Forms.Button();
            this.btnIIF = new System.Windows.Forms.Button();
            this.btnIn = new System.Windows.Forms.Button();
            this.btnNot = new System.Windows.Forms.Button();
            this.btnBetween = new System.Windows.Forms.Button();
            this.btnLike = new System.Windows.Forms.Button();
            this.btnOr = new System.Windows.Forms.Button();
            this.btnAnd = new System.Windows.Forms.Button();
            this.grpboxVariableManipulation = new System.Windows.Forms.GroupBox();
            this.btnMid = new System.Windows.Forms.Button();
            this.btnLCase = new System.Windows.Forms.Button();
            this.btnUCase = new System.Windows.Forms.Button();
            this.btnTrim = new System.Windows.Forms.Button();
            this.grpboxSelectSyntax = new System.Windows.Forms.GroupBox();
            this.lblFieldAlias = new System.Windows.Forms.Label();
            this.lbljoincondexp = new System.Windows.Forms.Label();
            this.btnDesc = new System.Windows.Forms.Button();
            this.btnAsc = new System.Windows.Forms.Button();
            this.lblFieldList3 = new System.Windows.Forms.Label();
            this.btnOrderBy = new System.Windows.Forms.Button();
            this.lblSqlExp = new System.Windows.Forms.Label();
            this.btnAll = new System.Windows.Forms.Button();
            this.btnUnion = new System.Windows.Forms.Button();
            this.lblFilterCondition = new System.Windows.Forms.Label();
            this.btnHaving = new System.Windows.Forms.Button();
            this.lblFieldList2 = new System.Windows.Forms.Label();
            this.btnGroupBy = new System.Windows.Forms.Button();
            this.btnOn = new System.Windows.Forms.Button();
            this.lblWhere = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbJoinType = new System.Windows.Forms.ComboBox();
            this.lblDBList = new System.Windows.Forms.Label();
            this.btnJoin = new System.Windows.Forms.Button();
            this.btnFrom = new System.Windows.Forms.Button();
            this.btnAs = new System.Windows.Forms.Button();
            this.btnAsterick = new System.Windows.Forms.Button();
            this.lblFieldList = new System.Windows.Forms.Label();
            this.btnDistinct = new System.Windows.Forms.Button();
            this.btnSelect = new System.Windows.Forms.Button();
            this.lstValues = new System.Windows.Forms.ListBox();
            this.lstFields = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnGetTables = new System.Windows.Forms.Button();
            this.lstTables = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSQLCommand = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.grpboxMathFunc.SuspendLayout();
            this.grpboxCondExp.SuspendLayout();
            this.grpboxVariableManipulation.SuspendLayout();
            this.grpboxSelectSyntax.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnPrevSQL);
            this.groupBox1.Controls.Add(this.chkValues);
            this.groupBox1.Controls.Add(this.btnFind);
            this.groupBox1.Controls.Add(this.txtSearchValues);
            this.groupBox1.Controls.Add(this.btnSort);
            this.groupBox1.Controls.Add(this.lblMsg);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.btnExecute);
            this.groupBox1.Controls.Add(this.btnClipboardPaste);
            this.groupBox1.Controls.Add(this.btnClipboardCopy);
            this.groupBox1.Controls.Add(this.btnClear);
            this.groupBox1.Controls.Add(this.btnTest);
            this.groupBox1.Controls.Add(this.btnOK);
            this.groupBox1.Controls.Add(this.grpboxMathFunc);
            this.groupBox1.Controls.Add(this.grpboxCondExp);
            this.groupBox1.Controls.Add(this.grpboxVariableManipulation);
            this.groupBox1.Controls.Add(this.grpboxSelectSyntax);
            this.groupBox1.Controls.Add(this.lstValues);
            this.groupBox1.Controls.Add(this.lstFields);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnGetTables);
            this.groupBox1.Controls.Add(this.lstTables);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtSQLCommand);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(744, 592);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // btnPrevSQL
            // 
            this.btnPrevSQL.Location = new System.Drawing.Point(190, 520);
            this.btnPrevSQL.Name = "btnPrevSQL";
            this.btnPrevSQL.Size = new System.Drawing.Size(57, 32);
            this.btnPrevSQL.TabIndex = 24;
            this.btnPrevSQL.Text = "Previous SQL";
            this.btnPrevSQL.Click += new System.EventHandler(this.btnPrevSQL_Click);
            // 
            // chkValues
            // 
            this.chkValues.Checked = true;
            this.chkValues.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkValues.Location = new System.Drawing.Point(432, 14);
            this.chkValues.Name = "chkValues";
            this.chkValues.Size = new System.Drawing.Size(120, 16);
            this.chkValues.TabIndex = 23;
            this.chkValues.Text = "Table/Field Values";
            this.chkValues.Click += new System.EventHandler(this.chkValues_Click);
            // 
            // btnFind
            // 
            this.btnFind.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFind.Location = new System.Drawing.Point(608, 12);
            this.btnFind.Name = "btnFind";
            this.btnFind.Size = new System.Drawing.Size(32, 16);
            this.btnFind.TabIndex = 22;
            this.btnFind.Text = "Find";
            this.btnFind.Click += new System.EventHandler(this.btnFind_Click);
            // 
            // txtSearchValues
            // 
            this.txtSearchValues.Location = new System.Drawing.Point(648, 10);
            this.txtSearchValues.Name = "txtSearchValues";
            this.txtSearchValues.Size = new System.Drawing.Size(72, 20);
            this.txtSearchValues.TabIndex = 21;
            // 
            // btnSort
            // 
            this.btnSort.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSort.Location = new System.Drawing.Point(560, 12);
            this.btnSort.Name = "btnSort";
            this.btnSort.Size = new System.Drawing.Size(32, 16);
            this.btnSort.TabIndex = 20;
            this.btnSort.Text = "Sort";
            this.btnSort.Click += new System.EventHandler(this.btnSort_Click);
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.White;
            this.lblMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMsg.ForeColor = System.Drawing.Color.Black;
            this.lblMsg.Location = new System.Drawing.Point(512, 80);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(144, 16);
            this.lblMsg.TabIndex = 19;
            this.lblMsg.Text = "Loading Values...Stand By";
            this.lblMsg.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(302, 520);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(56, 32);
            this.btnCancel.TabIndex = 18;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(22, 520);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(57, 32);
            this.btnExecute.TabIndex = 17;
            this.btnExecute.Text = "Execute SQL";
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // btnClipboardPaste
            // 
            this.btnClipboardPaste.Location = new System.Drawing.Point(191, 552);
            this.btnClipboardPaste.Name = "btnClipboardPaste";
            this.btnClipboardPaste.Size = new System.Drawing.Size(152, 24);
            this.btnClipboardPaste.TabIndex = 16;
            this.btnClipboardPaste.Text = "Paste SQL From Clipboard";
            this.btnClipboardPaste.Click += new System.EventHandler(this.btnClipboardPaste_Click);
            // 
            // btnClipboardCopy
            // 
            this.btnClipboardCopy.Location = new System.Drawing.Point(39, 552);
            this.btnClipboardCopy.Name = "btnClipboardCopy";
            this.btnClipboardCopy.Size = new System.Drawing.Size(152, 24);
            this.btnClipboardCopy.TabIndex = 15;
            this.btnClipboardCopy.Text = "Copy SQL To Clipboard";
            this.btnClipboardCopy.Click += new System.EventHandler(this.btnClipboardCopy_Click);
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(134, 520);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(56, 32);
            this.btnClear.TabIndex = 14;
            this.btnClear.Text = "Clear";
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnTest
            // 
            this.btnTest.Location = new System.Drawing.Point(78, 520);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(57, 32);
            this.btnTest.TabIndex = 13;
            this.btnTest.Text = "Test";
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(246, 520);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(57, 32);
            this.btnOK.TabIndex = 12;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // grpboxMathFunc
            // 
            this.grpboxMathFunc.Controls.Add(this.btnSum);
            this.grpboxMathFunc.Controls.Add(this.btnMax);
            this.grpboxMathFunc.Controls.Add(this.btnMin);
            this.grpboxMathFunc.Controls.Add(this.btnCount);
            this.grpboxMathFunc.Controls.Add(this.btnAvg);
            this.grpboxMathFunc.Controls.Add(this.btnExponent);
            this.grpboxMathFunc.Controls.Add(this.btnDivide);
            this.grpboxMathFunc.Controls.Add(this.btnMultiply);
            this.grpboxMathFunc.Controls.Add(this.btnSubtract);
            this.grpboxMathFunc.Controls.Add(this.btnAdd);
            this.grpboxMathFunc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxMathFunc.Location = new System.Drawing.Point(360, 531);
            this.grpboxMathFunc.Name = "grpboxMathFunc";
            this.grpboxMathFunc.Size = new System.Drawing.Size(368, 40);
            this.grpboxMathFunc.TabIndex = 11;
            this.grpboxMathFunc.TabStop = false;
            this.grpboxMathFunc.Text = "Mathmatical Functions";
            // 
            // btnSum
            // 
            this.btnSum.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSum.Location = new System.Drawing.Point(296, 16);
            this.btnSum.Name = "btnSum";
            this.btnSum.Size = new System.Drawing.Size(40, 16);
            this.btnSum.TabIndex = 40;
            this.btnSum.Text = "SUM";
            this.btnSum.Click += new System.EventHandler(this.btnSum_Click);
            // 
            // btnMax
            // 
            this.btnMax.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMax.Location = new System.Drawing.Point(256, 16);
            this.btnMax.Name = "btnMax";
            this.btnMax.Size = new System.Drawing.Size(40, 16);
            this.btnMax.TabIndex = 39;
            this.btnMax.Text = "MAX";
            this.btnMax.Click += new System.EventHandler(this.btnMax_Click);
            // 
            // btnMin
            // 
            this.btnMin.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMin.Location = new System.Drawing.Point(216, 16);
            this.btnMin.Name = "btnMin";
            this.btnMin.Size = new System.Drawing.Size(40, 16);
            this.btnMin.TabIndex = 38;
            this.btnMin.Text = "MIN";
            this.btnMin.Click += new System.EventHandler(this.btnMin_Click);
            // 
            // btnCount
            // 
            this.btnCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCount.Location = new System.Drawing.Point(168, 16);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(48, 16);
            this.btnCount.TabIndex = 37;
            this.btnCount.Text = "COUNT";
            this.btnCount.Click += new System.EventHandler(this.btnCount_Click);
            // 
            // btnAvg
            // 
            this.btnAvg.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAvg.Location = new System.Drawing.Point(128, 16);
            this.btnAvg.Name = "btnAvg";
            this.btnAvg.Size = new System.Drawing.Size(40, 16);
            this.btnAvg.TabIndex = 36;
            this.btnAvg.Text = "AVG";
            this.btnAvg.Click += new System.EventHandler(this.btnAvg_Click);
            // 
            // btnExponent
            // 
            this.btnExponent.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExponent.Location = new System.Drawing.Point(104, 16);
            this.btnExponent.Name = "btnExponent";
            this.btnExponent.Size = new System.Drawing.Size(24, 16);
            this.btnExponent.TabIndex = 35;
            this.btnExponent.Text = "^";
            this.btnExponent.Click += new System.EventHandler(this.btnExponent_Click);
            // 
            // btnDivide
            // 
            this.btnDivide.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDivide.Location = new System.Drawing.Point(80, 16);
            this.btnDivide.Name = "btnDivide";
            this.btnDivide.Size = new System.Drawing.Size(24, 16);
            this.btnDivide.TabIndex = 34;
            this.btnDivide.Text = "/";
            this.btnDivide.Click += new System.EventHandler(this.btnDivide_Click);
            // 
            // btnMultiply
            // 
            this.btnMultiply.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMultiply.Location = new System.Drawing.Point(56, 16);
            this.btnMultiply.Name = "btnMultiply";
            this.btnMultiply.Size = new System.Drawing.Size(24, 16);
            this.btnMultiply.TabIndex = 33;
            this.btnMultiply.Text = "*";
            this.btnMultiply.Click += new System.EventHandler(this.btnMultiply_Click);
            // 
            // btnSubtract
            // 
            this.btnSubtract.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubtract.ImageAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSubtract.Location = new System.Drawing.Point(32, 16);
            this.btnSubtract.Name = "btnSubtract";
            this.btnSubtract.Size = new System.Drawing.Size(24, 16);
            this.btnSubtract.TabIndex = 32;
            this.btnSubtract.Text = "-";
            this.btnSubtract.Click += new System.EventHandler(this.btnSubtract_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAdd.Location = new System.Drawing.Point(8, 16);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(24, 16);
            this.btnAdd.TabIndex = 31;
            this.btnAdd.Text = "+";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // grpboxCondExp
            // 
            this.grpboxCondExp.Controls.Add(this.btnN);
            this.grpboxCondExp.Controls.Add(this.btnY);
            this.grpboxCondExp.Controls.Add(this.btnExists);
            this.grpboxCondExp.Controls.Add(this.btnIsNull);
            this.grpboxCondExp.Controls.Add(this.btnLessThanEqualTo);
            this.grpboxCondExp.Controls.Add(this.btnLessThan);
            this.grpboxCondExp.Controls.Add(this.btnMoreThanEqualTo);
            this.grpboxCondExp.Controls.Add(this.btnMoreThan);
            this.grpboxCondExp.Controls.Add(this.btnNotEqual);
            this.grpboxCondExp.Controls.Add(this.btnEqual);
            this.grpboxCondExp.Controls.Add(this.btnPercent);
            this.grpboxCondExp.Controls.Add(this.btnComma);
            this.grpboxCondExp.Controls.Add(this.btnParenthRight);
            this.grpboxCondExp.Controls.Add(this.btnParenthLeft);
            this.grpboxCondExp.Controls.Add(this.btnSingleQuote);
            this.grpboxCondExp.Controls.Add(this.btnDoubleQuote);
            this.grpboxCondExp.Controls.Add(this.btnIIF);
            this.grpboxCondExp.Controls.Add(this.btnIn);
            this.grpboxCondExp.Controls.Add(this.btnNot);
            this.grpboxCondExp.Controls.Add(this.btnBetween);
            this.grpboxCondExp.Controls.Add(this.btnLike);
            this.grpboxCondExp.Controls.Add(this.btnOr);
            this.grpboxCondExp.Controls.Add(this.btnAnd);
            this.grpboxCondExp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxCondExp.Location = new System.Drawing.Point(360, 406);
            this.grpboxCondExp.Name = "grpboxCondExp";
            this.grpboxCondExp.Size = new System.Drawing.Size(368, 120);
            this.grpboxCondExp.TabIndex = 10;
            this.grpboxCondExp.TabStop = false;
            this.grpboxCondExp.Text = "Condition Expressions";
            // 
            // btnN
            // 
            this.btnN.BackColor = System.Drawing.SystemColors.Control;
            this.btnN.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnN.Location = new System.Drawing.Point(280, 88);
            this.btnN.Name = "btnN";
            this.btnN.Size = new System.Drawing.Size(32, 24);
            this.btnN.TabIndex = 36;
            this.btnN.Text = "\'N\'";
            this.btnN.UseVisualStyleBackColor = false;
            this.btnN.Click += new System.EventHandler(this.btnN_Click);
            // 
            // btnY
            // 
            this.btnY.BackColor = System.Drawing.SystemColors.Control;
            this.btnY.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnY.Location = new System.Drawing.Point(248, 88);
            this.btnY.Name = "btnY";
            this.btnY.Size = new System.Drawing.Size(32, 24);
            this.btnY.TabIndex = 35;
            this.btnY.Text = "\'Y\'";
            this.btnY.UseVisualStyleBackColor = false;
            this.btnY.Click += new System.EventHandler(this.btnY_Click);
            // 
            // btnExists
            // 
            this.btnExists.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExists.Location = new System.Drawing.Point(112, 64);
            this.btnExists.Name = "btnExists";
            this.btnExists.Size = new System.Drawing.Size(56, 24);
            this.btnExists.TabIndex = 34;
            this.btnExists.Text = "EXISTS";
            this.btnExists.Click += new System.EventHandler(this.btnExists_Click);
            // 
            // btnIsNull
            // 
            this.btnIsNull.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnIsNull.Location = new System.Drawing.Point(296, 40);
            this.btnIsNull.Name = "btnIsNull";
            this.btnIsNull.Size = new System.Drawing.Size(48, 24);
            this.btnIsNull.TabIndex = 33;
            this.btnIsNull.Text = "IS NULL";
            this.btnIsNull.Click += new System.EventHandler(this.btnIsNull_Click);
            // 
            // btnLessThanEqualTo
            // 
            this.btnLessThanEqualTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLessThanEqualTo.Location = new System.Drawing.Point(128, 88);
            this.btnLessThanEqualTo.Name = "btnLessThanEqualTo";
            this.btnLessThanEqualTo.Size = new System.Drawing.Size(120, 24);
            this.btnLessThanEqualTo.TabIndex = 32;
            this.btnLessThanEqualTo.Text = "Less Than Or Equal To";
            this.btnLessThanEqualTo.Click += new System.EventHandler(this.btnLessThanEqualTo_Click);
            // 
            // btnLessThan
            // 
            this.btnLessThan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLessThan.Location = new System.Drawing.Point(240, 64);
            this.btnLessThan.Name = "btnLessThan";
            this.btnLessThan.Size = new System.Drawing.Size(72, 24);
            this.btnLessThan.TabIndex = 31;
            this.btnLessThan.Text = "Less Than";
            this.btnLessThan.Click += new System.EventHandler(this.btnLessThan_Click);
            // 
            // btnMoreThanEqualTo
            // 
            this.btnMoreThanEqualTo.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMoreThanEqualTo.Location = new System.Drawing.Point(8, 88);
            this.btnMoreThanEqualTo.Name = "btnMoreThanEqualTo";
            this.btnMoreThanEqualTo.Size = new System.Drawing.Size(120, 24);
            this.btnMoreThanEqualTo.TabIndex = 30;
            this.btnMoreThanEqualTo.Text = "More Than Or Equal To";
            this.btnMoreThanEqualTo.Click += new System.EventHandler(this.btnMoreThanEqualTo_Click);
            // 
            // btnMoreThan
            // 
            this.btnMoreThan.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMoreThan.Location = new System.Drawing.Point(168, 64);
            this.btnMoreThan.Name = "btnMoreThan";
            this.btnMoreThan.Size = new System.Drawing.Size(72, 24);
            this.btnMoreThan.TabIndex = 29;
            this.btnMoreThan.Text = "More Than";
            this.btnMoreThan.Click += new System.EventHandler(this.btnMoreThan_Click);
            // 
            // btnNotEqual
            // 
            this.btnNotEqual.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNotEqual.Location = new System.Drawing.Point(56, 64);
            this.btnNotEqual.Name = "btnNotEqual";
            this.btnNotEqual.Size = new System.Drawing.Size(56, 24);
            this.btnNotEqual.TabIndex = 28;
            this.btnNotEqual.Text = "Not Equal";
            this.btnNotEqual.Click += new System.EventHandler(this.btnNotEqual_Click);
            // 
            // btnEqual
            // 
            this.btnEqual.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEqual.Location = new System.Drawing.Point(8, 64);
            this.btnEqual.Name = "btnEqual";
            this.btnEqual.Size = new System.Drawing.Size(48, 24);
            this.btnEqual.TabIndex = 26;
            this.btnEqual.Text = "Equal";
            this.btnEqual.Click += new System.EventHandler(this.btnEqual_Click);
            // 
            // btnPercent
            // 
            this.btnPercent.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPercent.Location = new System.Drawing.Point(248, 40);
            this.btnPercent.Name = "btnPercent";
            this.btnPercent.Size = new System.Drawing.Size(48, 24);
            this.btnPercent.TabIndex = 25;
            this.btnPercent.Text = "%";
            this.btnPercent.Click += new System.EventHandler(this.btnPercent_Click);
            // 
            // btnComma
            // 
            this.btnComma.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnComma.Location = new System.Drawing.Point(200, 40);
            this.btnComma.Name = "btnComma";
            this.btnComma.Size = new System.Drawing.Size(48, 24);
            this.btnComma.TabIndex = 24;
            this.btnComma.Text = ",";
            this.btnComma.Click += new System.EventHandler(this.btnComma_Click);
            // 
            // btnParenthRight
            // 
            this.btnParenthRight.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnParenthRight.Location = new System.Drawing.Point(152, 40);
            this.btnParenthRight.Name = "btnParenthRight";
            this.btnParenthRight.Size = new System.Drawing.Size(48, 24);
            this.btnParenthRight.TabIndex = 23;
            this.btnParenthRight.Text = ")";
            this.btnParenthRight.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnParenthRight.Click += new System.EventHandler(this.btnParenthRight_Click);
            // 
            // btnParenthLeft
            // 
            this.btnParenthLeft.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnParenthLeft.Location = new System.Drawing.Point(104, 40);
            this.btnParenthLeft.Name = "btnParenthLeft";
            this.btnParenthLeft.Size = new System.Drawing.Size(48, 24);
            this.btnParenthLeft.TabIndex = 22;
            this.btnParenthLeft.Text = "(";
            this.btnParenthLeft.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnParenthLeft.Click += new System.EventHandler(this.btnParenthLeft_Click);
            // 
            // btnSingleQuote
            // 
            this.btnSingleQuote.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSingleQuote.Location = new System.Drawing.Point(56, 40);
            this.btnSingleQuote.Name = "btnSingleQuote";
            this.btnSingleQuote.Size = new System.Drawing.Size(48, 24);
            this.btnSingleQuote.TabIndex = 21;
            this.btnSingleQuote.Text = "\'";
            this.btnSingleQuote.Click += new System.EventHandler(this.btnSingleQuote_Click);
            // 
            // btnDoubleQuote
            // 
            this.btnDoubleQuote.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDoubleQuote.Location = new System.Drawing.Point(8, 40);
            this.btnDoubleQuote.Name = "btnDoubleQuote";
            this.btnDoubleQuote.Size = new System.Drawing.Size(48, 24);
            this.btnDoubleQuote.TabIndex = 20;
            this.btnDoubleQuote.Text = "\"";
            this.btnDoubleQuote.Click += new System.EventHandler(this.btnDoubleQuote_Click);
            // 
            // btnIIF
            // 
            this.btnIIF.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnIIF.Location = new System.Drawing.Point(312, 24);
            this.btnIIF.Name = "btnIIF";
            this.btnIIF.Size = new System.Drawing.Size(48, 16);
            this.btnIIF.TabIndex = 19;
            this.btnIIF.Text = "IIF";
            this.btnIIF.Click += new System.EventHandler(this.btnIIF_Click);
            // 
            // btnIn
            // 
            this.btnIn.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnIn.Location = new System.Drawing.Point(264, 24);
            this.btnIn.Name = "btnIn";
            this.btnIn.Size = new System.Drawing.Size(48, 16);
            this.btnIn.TabIndex = 18;
            this.btnIn.Text = "IN";
            this.btnIn.Click += new System.EventHandler(this.btnIn_Click);
            // 
            // btnNot
            // 
            this.btnNot.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNot.Location = new System.Drawing.Point(216, 24);
            this.btnNot.Name = "btnNot";
            this.btnNot.Size = new System.Drawing.Size(48, 16);
            this.btnNot.TabIndex = 17;
            this.btnNot.Text = "NOT";
            this.btnNot.Click += new System.EventHandler(this.btnNot_Click);
            // 
            // btnBetween
            // 
            this.btnBetween.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBetween.Location = new System.Drawing.Point(152, 24);
            this.btnBetween.Name = "btnBetween";
            this.btnBetween.Size = new System.Drawing.Size(64, 16);
            this.btnBetween.TabIndex = 16;
            this.btnBetween.Text = "BETWEEN";
            this.btnBetween.Click += new System.EventHandler(this.btnBetween_Click);
            // 
            // btnLike
            // 
            this.btnLike.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLike.Location = new System.Drawing.Point(104, 24);
            this.btnLike.Name = "btnLike";
            this.btnLike.Size = new System.Drawing.Size(48, 16);
            this.btnLike.TabIndex = 15;
            this.btnLike.Text = "LIKE";
            this.btnLike.Click += new System.EventHandler(this.btnLike_Click);
            // 
            // btnOr
            // 
            this.btnOr.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOr.Location = new System.Drawing.Point(56, 24);
            this.btnOr.Name = "btnOr";
            this.btnOr.Size = new System.Drawing.Size(48, 16);
            this.btnOr.TabIndex = 14;
            this.btnOr.Text = "OR";
            this.btnOr.Click += new System.EventHandler(this.btnOr_Click);
            // 
            // btnAnd
            // 
            this.btnAnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnd.Location = new System.Drawing.Point(8, 24);
            this.btnAnd.Name = "btnAnd";
            this.btnAnd.Size = new System.Drawing.Size(48, 16);
            this.btnAnd.TabIndex = 13;
            this.btnAnd.Text = "AND";
            this.btnAnd.Click += new System.EventHandler(this.btnAnd_Click);
            // 
            // grpboxVariableManipulation
            // 
            this.grpboxVariableManipulation.Controls.Add(this.btnMid);
            this.grpboxVariableManipulation.Controls.Add(this.btnLCase);
            this.grpboxVariableManipulation.Controls.Add(this.btnUCase);
            this.grpboxVariableManipulation.Controls.Add(this.btnTrim);
            this.grpboxVariableManipulation.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxVariableManipulation.Location = new System.Drawing.Point(360, 360);
            this.grpboxVariableManipulation.Name = "grpboxVariableManipulation";
            this.grpboxVariableManipulation.Size = new System.Drawing.Size(368, 40);
            this.grpboxVariableManipulation.TabIndex = 9;
            this.grpboxVariableManipulation.TabStop = false;
            this.grpboxVariableManipulation.Text = "Variable Manipulation";
            // 
            // btnMid
            // 
            this.btnMid.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMid.Location = new System.Drawing.Point(152, 16);
            this.btnMid.Name = "btnMid";
            this.btnMid.Size = new System.Drawing.Size(48, 16);
            this.btnMid.TabIndex = 17;
            this.btnMid.Text = "MID";
            this.btnMid.Click += new System.EventHandler(this.btnMid_Click);
            // 
            // btnLCase
            // 
            this.btnLCase.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLCase.Location = new System.Drawing.Point(104, 16);
            this.btnLCase.Name = "btnLCase";
            this.btnLCase.Size = new System.Drawing.Size(48, 16);
            this.btnLCase.TabIndex = 16;
            this.btnLCase.Text = "LCASE";
            this.btnLCase.Click += new System.EventHandler(this.btnLCase_Click);
            // 
            // btnUCase
            // 
            this.btnUCase.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUCase.Location = new System.Drawing.Point(56, 16);
            this.btnUCase.Name = "btnUCase";
            this.btnUCase.Size = new System.Drawing.Size(48, 16);
            this.btnUCase.TabIndex = 15;
            this.btnUCase.Text = "UCASE";
            this.btnUCase.Click += new System.EventHandler(this.btnUCase_Click);
            // 
            // btnTrim
            // 
            this.btnTrim.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTrim.Location = new System.Drawing.Point(8, 16);
            this.btnTrim.Name = "btnTrim";
            this.btnTrim.Size = new System.Drawing.Size(48, 16);
            this.btnTrim.TabIndex = 14;
            this.btnTrim.Text = "TRIM";
            this.btnTrim.Click += new System.EventHandler(this.btnTrim_Click);
            // 
            // grpboxSelectSyntax
            // 
            this.grpboxSelectSyntax.Controls.Add(this.lblFieldAlias);
            this.grpboxSelectSyntax.Controls.Add(this.lbljoincondexp);
            this.grpboxSelectSyntax.Controls.Add(this.btnDesc);
            this.grpboxSelectSyntax.Controls.Add(this.btnAsc);
            this.grpboxSelectSyntax.Controls.Add(this.lblFieldList3);
            this.grpboxSelectSyntax.Controls.Add(this.btnOrderBy);
            this.grpboxSelectSyntax.Controls.Add(this.lblSqlExp);
            this.grpboxSelectSyntax.Controls.Add(this.btnAll);
            this.grpboxSelectSyntax.Controls.Add(this.btnUnion);
            this.grpboxSelectSyntax.Controls.Add(this.lblFilterCondition);
            this.grpboxSelectSyntax.Controls.Add(this.btnHaving);
            this.grpboxSelectSyntax.Controls.Add(this.lblFieldList2);
            this.grpboxSelectSyntax.Controls.Add(this.btnGroupBy);
            this.grpboxSelectSyntax.Controls.Add(this.btnOn);
            this.grpboxSelectSyntax.Controls.Add(this.lblWhere);
            this.grpboxSelectSyntax.Controls.Add(this.button1);
            this.grpboxSelectSyntax.Controls.Add(this.label4);
            this.grpboxSelectSyntax.Controls.Add(this.cmbJoinType);
            this.grpboxSelectSyntax.Controls.Add(this.lblDBList);
            this.grpboxSelectSyntax.Controls.Add(this.btnJoin);
            this.grpboxSelectSyntax.Controls.Add(this.btnFrom);
            this.grpboxSelectSyntax.Controls.Add(this.btnAs);
            this.grpboxSelectSyntax.Controls.Add(this.btnAsterick);
            this.grpboxSelectSyntax.Controls.Add(this.lblFieldList);
            this.grpboxSelectSyntax.Controls.Add(this.btnDistinct);
            this.grpboxSelectSyntax.Controls.Add(this.btnSelect);
            this.grpboxSelectSyntax.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpboxSelectSyntax.Location = new System.Drawing.Point(360, 154);
            this.grpboxSelectSyntax.Name = "grpboxSelectSyntax";
            this.grpboxSelectSyntax.Size = new System.Drawing.Size(368, 200);
            this.grpboxSelectSyntax.TabIndex = 8;
            this.grpboxSelectSyntax.TabStop = false;
            this.grpboxSelectSyntax.Text = "Select Syntax";
            // 
            // lblFieldAlias
            // 
            this.lblFieldAlias.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFieldAlias.Location = new System.Drawing.Point(272, 18);
            this.lblFieldAlias.Name = "lblFieldAlias";
            this.lblFieldAlias.Size = new System.Drawing.Size(72, 16);
            this.lblFieldAlias.TabIndex = 25;
            this.lblFieldAlias.Text = "field alias";
            // 
            // lbljoincondexp
            // 
            this.lbljoincondexp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbljoincondexp.Location = new System.Drawing.Point(160, 88);
            this.lbljoincondexp.Name = "lbljoincondexp";
            this.lbljoincondexp.Size = new System.Drawing.Size(128, 16);
            this.lbljoincondexp.TabIndex = 24;
            this.lbljoincondexp.Text = "join condition expression";
            // 
            // btnDesc
            // 
            this.btnDesc.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDesc.Location = new System.Drawing.Point(208, 177);
            this.btnDesc.Name = "btnDesc";
            this.btnDesc.Size = new System.Drawing.Size(40, 16);
            this.btnDesc.TabIndex = 23;
            this.btnDesc.Text = "DESC";
            this.btnDesc.Click += new System.EventHandler(this.btnDesc_Click);
            // 
            // btnAsc
            // 
            this.btnAsc.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAsc.Location = new System.Drawing.Point(160, 177);
            this.btnAsc.Name = "btnAsc";
            this.btnAsc.Size = new System.Drawing.Size(40, 16);
            this.btnAsc.TabIndex = 22;
            this.btnAsc.Text = "ASC";
            this.btnAsc.Click += new System.EventHandler(this.btnAsc_Click);
            // 
            // lblFieldList3
            // 
            this.lblFieldList3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFieldList3.Location = new System.Drawing.Point(112, 176);
            this.lblFieldList3.Name = "lblFieldList3";
            this.lblFieldList3.Size = new System.Drawing.Size(48, 16);
            this.lblFieldList3.TabIndex = 21;
            this.lblFieldList3.Text = "field list";
            // 
            // btnOrderBy
            // 
            this.btnOrderBy.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOrderBy.Location = new System.Drawing.Point(32, 175);
            this.btnOrderBy.Name = "btnOrderBy";
            this.btnOrderBy.Size = new System.Drawing.Size(64, 16);
            this.btnOrderBy.TabIndex = 20;
            this.btnOrderBy.Text = "ORDER BY";
            this.btnOrderBy.Click += new System.EventHandler(this.btnOrderBy_Click);
            // 
            // lblSqlExp
            // 
            this.lblSqlExp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSqlExp.Location = new System.Drawing.Point(160, 159);
            this.lblSqlExp.Name = "lblSqlExp";
            this.lblSqlExp.Size = new System.Drawing.Size(120, 16);
            this.lblSqlExp.TabIndex = 19;
            this.lblSqlExp.Text = "select-sql expression";
            // 
            // btnAll
            // 
            this.btnAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAll.Location = new System.Drawing.Point(112, 158);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(40, 16);
            this.btnAll.TabIndex = 18;
            this.btnAll.Text = "ALL";
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click);
            // 
            // btnUnion
            // 
            this.btnUnion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUnion.Location = new System.Drawing.Point(32, 158);
            this.btnUnion.Name = "btnUnion";
            this.btnUnion.Size = new System.Drawing.Size(64, 16);
            this.btnUnion.TabIndex = 17;
            this.btnUnion.Text = "UNION";
            this.btnUnion.Click += new System.EventHandler(this.btnUnion_Click);
            // 
            // lblFilterCondition
            // 
            this.lblFilterCondition.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFilterCondition.Location = new System.Drawing.Point(112, 142);
            this.lblFilterCondition.Name = "lblFilterCondition";
            this.lblFilterCondition.Size = new System.Drawing.Size(88, 16);
            this.lblFilterCondition.TabIndex = 16;
            this.lblFilterCondition.Text = "filter condition";
            // 
            // btnHaving
            // 
            this.btnHaving.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnHaving.Location = new System.Drawing.Point(32, 141);
            this.btnHaving.Name = "btnHaving";
            this.btnHaving.Size = new System.Drawing.Size(64, 16);
            this.btnHaving.TabIndex = 15;
            this.btnHaving.Text = "HAVING";
            this.btnHaving.Click += new System.EventHandler(this.btnHaving_Click);
            // 
            // lblFieldList2
            // 
            this.lblFieldList2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFieldList2.Location = new System.Drawing.Point(112, 125);
            this.lblFieldList2.Name = "lblFieldList2";
            this.lblFieldList2.Size = new System.Drawing.Size(56, 16);
            this.lblFieldList2.TabIndex = 14;
            this.lblFieldList2.Text = "field list";
            // 
            // btnGroupBy
            // 
            this.btnGroupBy.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGroupBy.Location = new System.Drawing.Point(32, 124);
            this.btnGroupBy.Name = "btnGroupBy";
            this.btnGroupBy.Size = new System.Drawing.Size(64, 16);
            this.btnGroupBy.TabIndex = 13;
            this.btnGroupBy.Text = "GROUP BY";
            this.btnGroupBy.Click += new System.EventHandler(this.btnGroupBy_Click);
            // 
            // btnOn
            // 
            this.btnOn.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOn.Location = new System.Drawing.Point(104, 86);
            this.btnOn.Name = "btnOn";
            this.btnOn.Size = new System.Drawing.Size(48, 16);
            this.btnOn.TabIndex = 12;
            this.btnOn.Text = "ON";
            this.btnOn.Click += new System.EventHandler(this.btnOn_Click);
            // 
            // lblWhere
            // 
            this.lblWhere.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWhere.Location = new System.Drawing.Point(112, 108);
            this.lblWhere.Name = "lblWhere";
            this.lblWhere.Size = new System.Drawing.Size(144, 16);
            this.lblWhere.TabIndex = 11;
            this.lblWhere.Text = "filter or join expression(s)";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(32, 108);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(64, 16);
            this.button1.TabIndex = 10;
            this.button1.Text = "WHERE";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(200, 59);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(88, 16);
            this.label4.TabIndex = 9;
            this.label4.Text = "database name";
            // 
            // cmbJoinType
            // 
            this.cmbJoinType.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbJoinType.ItemHeight = 13;
            this.cmbJoinType.Items.AddRange(new object[] {
            "LEFT",
            "RIGHT",
            "FULL",
            "OUTER",
            "INNER",
            "LEFT OUTER",
            "RIGHT OUTER",
            "FULL OUTER"});
            this.cmbJoinType.Location = new System.Drawing.Point(32, 59);
            this.cmbJoinType.Name = "cmbJoinType";
            this.cmbJoinType.Size = new System.Drawing.Size(104, 21);
            this.cmbJoinType.TabIndex = 8;
            // 
            // lblDBList
            // 
            this.lblDBList.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDBList.Location = new System.Drawing.Point(88, 37);
            this.lblDBList.Name = "lblDBList";
            this.lblDBList.Size = new System.Drawing.Size(68, 16);
            this.lblDBList.TabIndex = 7;
            this.lblDBList.Text = "database list";
            // 
            // btnJoin
            // 
            this.btnJoin.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnJoin.Location = new System.Drawing.Point(144, 59);
            this.btnJoin.Name = "btnJoin";
            this.btnJoin.Size = new System.Drawing.Size(48, 16);
            this.btnJoin.TabIndex = 6;
            this.btnJoin.Text = "JOIN";
            this.btnJoin.Click += new System.EventHandler(this.btnJoin_Click);
            // 
            // btnFrom
            // 
            this.btnFrom.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFrom.Location = new System.Drawing.Point(32, 36);
            this.btnFrom.Name = "btnFrom";
            this.btnFrom.Size = new System.Drawing.Size(48, 16);
            this.btnFrom.TabIndex = 5;
            this.btnFrom.Text = "FROM";
            this.btnFrom.Click += new System.EventHandler(this.btnFrom_Click);
            // 
            // btnAs
            // 
            this.btnAs.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAs.Location = new System.Drawing.Point(232, 17);
            this.btnAs.Name = "btnAs";
            this.btnAs.Size = new System.Drawing.Size(24, 16);
            this.btnAs.TabIndex = 4;
            this.btnAs.Text = "AS";
            this.btnAs.Click += new System.EventHandler(this.btnAs_Click);
            // 
            // btnAsterick
            // 
            this.btnAsterick.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAsterick.Location = new System.Drawing.Point(208, 17);
            this.btnAsterick.Name = "btnAsterick";
            this.btnAsterick.Size = new System.Drawing.Size(16, 16);
            this.btnAsterick.TabIndex = 3;
            this.btnAsterick.Text = "*";
            this.btnAsterick.Click += new System.EventHandler(this.btnAsterick_Click);
            // 
            // lblFieldList
            // 
            this.lblFieldList.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFieldList.Location = new System.Drawing.Point(152, 17);
            this.lblFieldList.Name = "lblFieldList";
            this.lblFieldList.Size = new System.Drawing.Size(48, 16);
            this.lblFieldList.TabIndex = 2;
            this.lblFieldList.Text = "field list";
            // 
            // btnDistinct
            // 
            this.btnDistinct.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDistinct.Location = new System.Drawing.Point(77, 17);
            this.btnDistinct.Name = "btnDistinct";
            this.btnDistinct.Size = new System.Drawing.Size(67, 16);
            this.btnDistinct.TabIndex = 1;
            this.btnDistinct.Text = "DISTINCT";
            this.btnDistinct.Click += new System.EventHandler(this.btnDistinct_Click);
            // 
            // btnSelect
            // 
            this.btnSelect.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelect.Location = new System.Drawing.Point(16, 17);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(56, 16);
            this.btnSelect.TabIndex = 0;
            this.btnSelect.Text = "SELECT";
            this.btnSelect.Click += new System.EventHandler(this.btnSelect_Click);
            // 
            // lstValues
            // 
            this.lstValues.Location = new System.Drawing.Point(432, 36);
            this.lstValues.Name = "lstValues";
            this.lstValues.Size = new System.Drawing.Size(296, 108);
            this.lstValues.TabIndex = 7;
            this.lstValues.DoubleClick += new System.EventHandler(this.lstValues_DoubleClick);
            // 
            // lstFields
            // 
            this.lstFields.Location = new System.Drawing.Point(216, 35);
            this.lstFields.Name = "lstFields";
            this.lstFields.Size = new System.Drawing.Size(200, 108);
            this.lstFields.TabIndex = 5;
            this.lstFields.SelectedValueChanged += new System.EventHandler(this.lstFields_SelectedValueChanged);
            this.lstFields.DoubleClick += new System.EventHandler(this.lstFields_DoubleClick);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(216, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "Fields";
            // 
            // btnGetTables
            // 
            this.btnGetTables.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGetTables.Location = new System.Drawing.Point(64, 145);
            this.btnGetTables.Name = "btnGetTables";
            this.btnGetTables.Size = new System.Drawing.Size(80, 21);
            this.btnGetTables.TabIndex = 3;
            this.btnGetTables.Text = "Get Table{s}";
            this.btnGetTables.Click += new System.EventHandler(this.btnGetTables_Click);
            // 
            // lstTables
            // 
            this.lstTables.Location = new System.Drawing.Point(24, 34);
            this.lstTables.Name = "lstTables";
            this.lstTables.Size = new System.Drawing.Size(176, 108);
            this.lstTables.TabIndex = 2;
            this.lstTables.SelectedValueChanged += new System.EventHandler(this.lstTables_SelectedValueChanged);
            this.lstTables.DoubleClick += new System.EventHandler(this.lstTables_DoubleClick);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(24, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Table";
            // 
            // txtSQLCommand
            // 
            this.txtSQLCommand.Location = new System.Drawing.Point(24, 171);
            this.txtSQLCommand.Multiline = true;
            this.txtSQLCommand.Name = "txtSQLCommand";
            this.txtSQLCommand.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtSQLCommand.Size = new System.Drawing.Size(328, 349);
            this.txtSQLCommand.TabIndex = 0;
            // 
            // uc_sql_builder
            // 
            this.AutoScroll = true;
            this.AutoSize = true;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_sql_builder";
            this.Size = new System.Drawing.Size(744, 592);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.grpboxMathFunc.ResumeLayout(false);
            this.grpboxCondExp.ResumeLayout(false);
            this.grpboxVariableManipulation.ResumeLayout(false);
            this.grpboxSelectSyntax.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		private void btnLessThan_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " < ");
		}

		private void btnSelect_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, "select ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, "SELECT ");
			}
		}

		private void btnDistinct_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " distinct ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " DISTINCT ");
			}
		}

		private void btnTest_Click(object sender, System.EventArgs e)
		{
			//((frmDialog)this.ParentForm).m_frmScenarioCallingForm.uc_scenario_filter1.m_ado_core_tables.m_dsCoreTables.WriteXmlSchema("c:\\temp\\temp.xsd");
		    //((frmDialog)this.ParentForm).m_frmScenarioCallingForm.uc_scenario_filter1.m_ado_core_tables.m_dsCoreTables.WriteXml("c:\\temp\\temp.xml");
			
			
			string strSQL="";
			//int intArrayCount;
			//int x=0;
			string strConn="";
			//string strCommand="";
			//string str="";
            System.Data.OleDb.OleDbConnection p_conn;
			ado_data_access p_ado = new ado_data_access();
            
			p_conn = new System.Data.OleDb.OleDbConnection();
			strConn = p_ado.getMDBConnString(this.m_strTempMDBFile,"admin",""); //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.m_strTempMDBFile + ";User Id=admin;Password=;";
			p_ado.OpenConnection(strConn, ref p_conn);	
			
			if (p_ado.m_intError != 0)
			{
				p_ado = null;
				return;
			}

			strSQL = this.txtSQLCommand.Text.Trim();

			p_ado.SqlQueryReader(p_conn, strSQL);

			
			if (p_ado.m_intError == 0)
			{
				

				MessageBox.Show("Valid Syntax");
				
				p_ado.m_OleDbDataReader.Close();
				
				p_ado.m_OleDbDataReader = null;
				p_ado.m_OleDbCommand = null;
				
				
			}
		    p_conn.Close();
			p_conn = null;
			p_ado = null;
			
			
		}

		/*******************************************************************
		 ** provide list box for user to select the tables that
		 ** he/she  can select from.
		 *******************************************************************/
		private void btnGetTables_Click(object sender, System.EventArgs e)
		{

			//string strSQL="";

			//string strConn="";
			//string strCommand="";
			//string strLargestString="";
			//string str="";
			//string strScenarioId="";
			//string strScenarioMDB="";
			int x=0;
			DialogResult result;
            

			ado_data_access p_ado = new ado_data_access();

		
				frmDialog frmTemp = new frmDialog();
				frmTemp.Text = "Select Database Tables For SQL";
				frmTemp.uc_select_list_item1.lblMsg.Enabled=true;
				frmTemp.uc_select_list_item1.lblMsg.Text= "Data Sources";
				frmTemp.uc_select_list_item1.lblMsg.Visible = true;
				frmTemp.uc_select_list_item1.Dock = System.Windows.Forms.DockStyle.Fill;
						
				frmTemp.uc_project1.Visible=false;
				frmTemp.uc_select_list_item1.listBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
				frmTemp.uc_select_list_item1.listBox1.Items.Clear();
				this.AddDataSourceTablesIntoListBox(frmTemp.uc_select_list_item1.listBox1);
				if (frmTemp.uc_select_list_item1.listBox1.Items.Count==0)
				{
					MessageBox.Show("!!No Tables Found!!");
				}
				else 
				{
					frmTemp.uc_select_list_item1.Visible=true;
					result = frmTemp.ShowDialog(this);
                        
						
					if (result == DialogResult.OK) 
					{
						this.m_strCurField="";
						this.m_strCurTable="";
						this.lstTables.Items.Clear();
						this.lstFields.Items.Clear();
						this.lstValues.Items.Clear();
						for (x=0; x<=frmTemp.uc_select_list_item1.listBox1.SelectedItems.Count-1;x++)
						{
							this.lstTables.Items.Add(frmTemp.uc_select_list_item1.listBox1.SelectedItems[x]);
						}
						this.LoadTablesIntoDatasets();
					}
					
					
					frmTemp.Close();
					frmTemp = null;
				


			}
			p_ado.m_OleDbDataReader = null;
			p_ado.m_OleDbCommand = null;

			p_ado = null;
		
			

		}

		private void lstTables_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strCurTable == this.lstTables.Text)
				return;
		    
			int x=0;
			string strField="";
			string strTable = this.lstTables.Text;
			this.lstFields.Items.Clear();
            int ColumnCount = this.m_dsDataSource.Tables[strTable].Columns.Count;
			
			for (x=0;x<=ColumnCount-1;x++)
			{
              
			   strField = this.m_dsDataSource.Tables[strTable].Columns[x].ColumnName;
               this.lstFields.Items.Add(strField);
			}
            this.m_strCurTable = this.lstTables.Text;
			this.m_strCurField = "";
    	}

		private void lstFields_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (this.m_strCurField == this.lstFields.Text || this.chkValues.Checked==false)
				return;

			string strField=this.lstFields.Text;
			string strTable = this.lstTables.Text;
			string str = "";
			int count = 0;
			int x=0;
			bool lSingleQuote = false;
			string strBuild;
			//string[] strValues = new string[((frmDialog)this.Parent).m_frmScenarioCallingForm.uc_scenario_filter1.m_ado_core_tables.m_dsCoreTables.Tables[strTable].Rows.Count];
			string[] strValues = new string[this.m_dsDataSource.Tables[strTable].Rows.Count];
			this.lstValues.Items.Clear();
			this.lstValues.Sorted=false;

			//strValues.Initialize();
			this.lblMsg.Text = "Loading values...stand by";
			this.lblMsg.Visible=true;
			this.lblMsg.Refresh();
			
            strBuild="";
             
			switch (this.m_dsDataSource.Tables[strTable].Columns[strField].DataType.ToString().Trim())
			{
			    case "System.Single":
					lSingleQuote = false;
					break;
				case "System.Double":
					lSingleQuote = false;
					break;
				case "System.Decimal":
					lSingleQuote = false;
					break;
				case "System.String":
					lSingleQuote = true;
					break;
                case "System.Int16":
					lSingleQuote = false;
					break;
				case "System.Char":
					lSingleQuote = true;
					break;
				case "System.Int32":
					lSingleQuote = false;
					break;
                case "System.DateTime":
					lSingleQuote = false;
					break;
				case "System.DayOfWeek":
					lSingleQuote = false;
					break; 
				case "System.Int64":
					lSingleQuote = false;
					break;
				case "System.Byte":
					lSingleQuote = false;
					break;
                case "System.Boolean":
					lSingleQuote = false;
					break;
                default:
					lSingleQuote = false;
					MessageBox.Show("unknown datatype for " + strField + " " + this.m_dsDataSource.Tables[strTable].Columns[strField].DataType.ToString());
					break;
			}
			
			//if (((frmDialog)this.Parent).m_frmScenarioCallingForm.uc_scenario_filter1.m_ado_core_tables.m_dsCoreTables.Tables[strTable].Columns[strField].DataType == 
			//	  System.Type.GetTypeData.OleDb.OleDbType.VarChar) MessageBox.Show("found one");

			for (x = 0; x<= this.m_dsDataSource.Tables[strTable].Rows.Count -1; x++)
			{
                str = this.m_dsDataSource.Tables[strTable].Rows[x][strField].ToString().Trim();
				if (strBuild.IndexOf(str,0,strBuild.Length) == -1)
				{
					if (lSingleQuote == true)
					{
						strValues[count] = "'" + str + "'";
					}
					else
					{
						strValues[count] = str;
					}
					
					strBuild = strBuild + "'" + str + "',";
					count++;
				}

			}																																					 
 					
			for (x = 0; x<= count - 1; x++)
				this.lstValues.Items.Add(strValues[x].ToString().Trim());																							 
			
		
			strBuild = null;
			strValues = null;
			this.lblMsg.Visible=false;
			this.lblMsg.Refresh();
           
		    this.m_strCurField = this.lstFields.Text ;
			

		   
		}

		private void btnSort_Click(object sender, System.EventArgs e)
		{
			if (this.lstValues.Sorted == true)
			{
				this.lstValues.Sorted = false;
			}
			else this.lstValues.Sorted = true;
		}

		private void btnFind_Click(object sender, System.EventArgs e)
		{
			int x = lstValues.FindString(this.txtSearchValues.Text,0);

			
			// If no item is found that matches exit.
			if (x != -1)
			{
				// Since the FindString loops infinitely, determine if we found first item again and exit.
				if (this.lstValues.SelectedIndices.Count > 0)
				{
					if(x == this.lstValues.SelectedIndices[0])
						return;
				}
				// Select the item in the ListBox once it is found.
				this.lstValues.SetSelected(x,true);
			}

		}

		private void lstTables_DoubleClick(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " " + this.lstTables.Text + " ");
		}

		private void lstFields_DoubleClick(object sender, System.EventArgs e)
		{
			
			this.SendKeyStrokes(this.txtSQLCommand," " + this.lstTables.Text + "." + this.lstFields.Text + " " );
		}

		private void lstValues_DoubleClick(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand," " + this.lstValues.Text + " " );
		}

		private void btnAsterick_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " * ");
		}

		private void btnAs_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " as ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand," AS ");
			}
		}

		private void btnFrom_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " from ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " FROM ");
			}

		}

		private void btnJoin_Click(object sender, System.EventArgs e)
		{
			if (this.cmbJoinType.Text.Trim().Length == 0)
			{
				this.cmbJoinType.Text = "INNER";
				if (this.m_oUtils.IsCapsLockOn() == true)
				{
					this.SendKeyStrokes(this.txtSQLCommand, " inner join ");
				}
				else
				{
					this.SendKeyStrokes(this.txtSQLCommand, " INNER JOIN ");
				}
			}
			else 
			{
				if (this.m_oUtils.IsCapsLockOn() == true)
				{
					this.SendKeyStrokes(this.txtSQLCommand, this.cmbJoinType.Text.Trim() + " join ");
				}
				else
				{
					this.SendKeyStrokes(this.txtSQLCommand, this.cmbJoinType.Text.Trim() + " JOIN ");
				}
			}
		}

		private void btnOn_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " on ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " ON ");
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " where ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " WHERE ");
			}
		}

		private void btnGroupBy_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " group by ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " GROUP BY ");
			}
		}

		private void btnHaving_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " having ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " HAVING ");
			}
		}

		private void btnUnion_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " union ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " UNION ");
			}
		}

		private void btnAll_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " all ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " ALL ");
			}
		}

		private void btnOrderBy_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " order by ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " ORDER BY ");
			}
		}

		private void btnAsc_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " asc ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " ASC ");
			}
		}

		private void btnDesc_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " desc ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " DESC ");
			}
		}

		private void btnAnd_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " and ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " AND ");
			}
		}

		private void btnOr_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " or ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " OR ");
			}
		}

		private void btnLike_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " like ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " LIKE ");
			}
		}

		private void btnBetween_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " between ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " BETWEEN ");
			}
		}

		private void btnNot_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " not ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " NOT ");
			}
		}

		private void btnIn_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " in ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " IN ");
			}
		}

		private void btnIIF_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " iif{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " IIF{(}");
			}
		}

		private void btnDoubleQuote_Click(object sender, System.EventArgs e)
		{

		}

		private void btnSingleQuote_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, "'");
		}

		private void btnParenthLeft_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " {(}");
		}

		private void btnParenthRight_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, "{)} ");
		}

		private void btnComma_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, ",");
		}

		private void btnPercent_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, "{%}");
		}

		private void btnIsNull_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " is null ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " IS NULL ");
			}
		}

		private void btnEqual_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " = ");
		}

		private void btnExact_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " == ");
		}

		private void btnNotEqual_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " <> ");
		}

		private void btnMoreThan_Click(object sender, System.EventArgs e)
		{
		   this.SendKeyStrokes(this.txtSQLCommand, " > ");
		}

		private void btnMoreThanEqualTo_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " >= ");
		}

		private void btnLessThanEqualTo_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " <= ");
		}

		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " {+} ");
		}

		private void btnSubtract_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " - ");
		}

		private void btnMultiply_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " * ");
		}

		private void btnDivide_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, " / ");
		}

		private void btnExponent_Click(object sender, System.EventArgs e)
		{
			this.SendKeyStrokes(this.txtSQLCommand, "{^} ");
		}

		private void btnAvg_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " avg{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " AVG{(}");
			}
		}

		private void btnCount_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " count{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " COUNT{(}");
			}
		}

		private void btnMin_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " min{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " MIN{(}");
			}
		}

		private void btnMax_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " max{(}");
			}
			else
			{
			    this.SendKeyStrokes(this.txtSQLCommand, " MAX{(}");
			}

		}

		private void btnSum_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " sum{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " SUM{(}");
			}
		}

		private void btnClear_Click(object sender, System.EventArgs e)
		{
			this.txtSQLCommand.Text = "";
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.ParentForm.Close();
		}

		private void chkValues_Click(object sender, System.EventArgs e)
		{
			if (this.chkValues.Checked==false)
				this.lstValues.Items.Clear();
		}

		private void btnClipboardCopy_Click(object sender, System.EventArgs e)
		{
		   System.Windows.Forms.Clipboard.SetDataObject(this.txtSQLCommand.Text);
		}

		private void btnClipboardPaste_Click(object sender, System.EventArgs e)
		{
			System.Windows.Forms.IDataObject iData = Clipboard.GetDataObject();
			if (iData.GetDataPresent(System.Windows.Forms.DataFormats.Text))
				this.txtSQLCommand.Text = (String)iData.GetData(DataFormats.Text);
		}

		private void btnTrim_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " trim{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " TRIM{(}");
			}
		}

		private void btnLCase_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " lcase{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " LCASE{(}");
			}
		}

		private void btnUCase_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " ucase{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " UCASE{(}");
			}
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{

			((frmDialog)this.ParentForm).DialogResult = System.Windows.Forms.DialogResult.OK;
		}

		private void btnMid_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " mid{(}");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " MID{(}");
			}
		}

		private void btnPrevSQL_Click(object sender, System.EventArgs e)
		{
			//string strSQL="";
			string strConn="";
			//string str="";
			//int x=0;

			DialogResult result;
            
  			string strScenarioId =  ((frmDialog)this.ParentForm).m_frmScenarioCallingForm.uc_scenario1.txtScenarioId.Text.Trim().ToLower();
			string strScenarioMDB = ((frmDialog)this.ParentForm).m_frmMain.frmProject.uc_project1.m_strProjectDirectory + 
				"\\core\\db\\scenario_core_rule_definitions.mdb";
			
			frmDialog frmPrevExp = new frmDialog();
				
			frmPrevExp.Width = frmPrevExp.uc_previous_expressions1.m_intFullWd;
			frmPrevExp.Height = frmPrevExp.uc_previous_expressions1.m_intFullHt;
			frmPrevExp.Text = "Previous SQL Expressions";
					
			frmPrevExp.uc_previous_expressions1.Visible=true;
			if (strScenarioMDB.Substring(strScenarioMDB.Trim().Length - 6,6).ToUpper()==".ACCDB")
				strConn= "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + strScenarioMDB + ";User Id=admin;Password=;";
			else
				strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strScenarioMDB + ";User Id=admin;Password=;";

			frmPrevExp.uc_previous_expressions1.loadvalues(strConn,"SELECT * FROM scenario_plot_filter;","SQL_COMMAND","SQL_COMMAND", "scenario_plot_filter");
            
			result = frmPrevExp.ShowDialog(this);
			if (result == DialogResult.OK)
            {
				//result = MessageBox.Show("Replace \n'" + this.txtSQLCommand.Text + "' \n with \n'" + frmPrevExp.uc_previous_expressions1.lblSQL.Text + "'", "Previous SQL",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
				result = MessageBox.Show("REPLACE \n" + "----------------\n\n" + "'" + this.txtSQLCommand.Text  + "' \n\n\n WITH  \n----------------\n\n'" + frmPrevExp.uc_previous_expressions1.lblSQL.Text + "'", "Previous SQL",MessageBoxButtons.YesNo,MessageBoxIcon.Question);
			    if (result == DialogResult.Yes)
			    {
					this.txtSQLCommand.Text = frmPrevExp.uc_previous_expressions1.listView1.SelectedItems[0].SubItems[frmPrevExp.uc_previous_expressions1.m_intSelectColumn + 1].Text;
			    }
		    }
			frmPrevExp.Close();
			frmPrevExp = null;
			
		
		}

		private void btnExecute_Click(object sender, System.EventArgs e)
		{
			
			string strRandomPathAndFile="";
			string strConn="";
			if (((frmDialog)this.ParentForm).strCallingFormType == "CS")  //core scenario client
			{
				strRandomPathAndFile=((frmDialog)this.ParentForm).m_frmScenarioCallingForm.uc_datasource1.CreateMDBAndScenarioTableDataSourceLinks(this.m_oEnv.strTempDir);
				if (strRandomPathAndFile.Trim().Length > 0)
				{
					if (strRandomPathAndFile.Substring(strRandomPathAndFile.Trim().Length - 6,6).ToUpper()==".ACCDB")
						strConn= "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" + strRandomPathAndFile + ";User Id=admin;Password=;";
					else
						strConn= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strRandomPathAndFile + ";User Id=admin;Password=;";
					try
					{
					   this.frmGridView1.LoadDataSet(strConn,this.txtSQLCommand.Text);
					}
					catch
					{
						this.frmGridView1 = new frmGridView();
						this.frmGridView1.MdiParent = this.ParentForm.ParentForm;
						this.frmGridView1.Visible=false;
						this.frmGridView1.LoadDataSet(strConn,this.txtSQLCommand.Text);
						
					}
					
					if (this.frmGridView1.Visible==false)
					{
						//((frmDialog)this.ParentForm).m_frmScenarioCallingForm.MdiParent = ((frmDialog)this.ParentForm).m_frmMain;
						this.frmGridView1.Text =  "SQL Builder Data Set Browser";
						try
						{
							this.frmGridView1.Show();
						}
						catch
						{
							this.frmGridView1 = new frmGridView();
							this.frmGridView1.MdiParent = this.ParentForm.ParentForm;
							this.frmGridView1.Visible=false;
							this.frmGridView1.LoadDataSet(strConn,this.txtSQLCommand.Text);
							this.frmGridView1.Text =  "SQL Builder Data Set Browser";
							this.frmGridView1.Show();

						}
					}
					if (this.frmGridView1.WindowState == System.Windows.Forms.FormWindowState.Minimized)
						 this.frmGridView1.WindowState = System.Windows.Forms.FormWindowState.Normal;
					this.frmGridView1.Focus();

				}
			}
		}
		public void InitializeDataSource(System.Windows.Forms.ListView lv)
		{
			this.lvDataSource = lv;
            dao_data_access p_dao = new dao_data_access();
			utils p_utils = new utils();
            this.getNumberOfValidTables();
			for (int x=0; x <= this.m_intNumberOfValidTables - 1; x++)
			{
				this.m_strListedTablesLoadedIntoDatasets[x] = "";
			}
			this.m_intNumberOfTablesLoadedIntoDatasets = 0;

			/*************************
			 **get temporary mdb file
			 *************************/
			this.m_strTempMDBFile =  p_utils.getRandomFile(this.m_oEnv.strTempDir,"accdb");
			p_utils = null;

			/*****************************************************
			 **create a temporary mdb that will contain all 
			 **the links to the tables that 
			 *****************************************************/
			p_dao.CreateMDB(this.m_strTempMDBFile);

			/************************************************
			 **load up the tables into datasets
			 ************************************************/
			this.LoadTablesIntoDatasets();




		}
		private void LoadTablesIntoDatasets()
		{
			//string strScenarioMDB="";
			string strSQL="";
			string strFullPathMDB="";
			string strConn="";
			
			int x=0;
			int y=0;
			bool lLoaded=false;
			
			this.m_intError=0;
              
			
		
			//ado specific routines class
			ado_data_access p_ado = new ado_data_access();

			//go through each of the items in the listbox
			for (y = 0; y <= this.lstTables.Items.Count - 1; y++)
			{
				lLoaded=false;
				//see if the listbox item is already loaded into a dataset table and the linked mdb table
				if (this.m_intNumberOfTablesLoadedIntoDatasets != 0)
				{   
					for (x=0;x<=this.m_intNumberOfValidTables - 1;x++)
					{
						if (this.m_strListedTablesLoadedIntoDatasets[x].Trim().Length > 0)
						{
							if (this.lstTables.Items[y].ToString().Trim().ToLower() == 
								this.m_strListedTablesLoadedIntoDatasets[x].Trim().ToLower())
							{
								lLoaded=true;
								break;
							}
						}
					}
				}
				if (lLoaded==false)
				{
					for (x=0; x<=this.lvDataSource.Items.Count-1;x++)
					{
						//look to make sure we have the correct record
						if (this.lstTables.Items[y].ToString().Trim().ToUpper() 
							==
							this.lvDataSource.Items[x].SubItems[TABLE].Text.Trim().ToUpper())
						{
							strFullPathMDB = this.lvDataSource.Items[x].SubItems[PATH].Text.Trim() + 
								"\\" + this.lvDataSource.Items[x].SubItems[MDBFILE].Text.Trim();
							
							//used to create a link to the table
							dao_data_access p_dao = new dao_data_access();

							//create a link to the table in an mdb file
							p_dao.CreateTableLink(this.m_strTempMDBFile,this.lstTables.Items[y].ToString().Trim(),
								strFullPathMDB,this.lstTables.Items[y].ToString().Trim());
							p_dao = null;

							//connect to mdb file that will be used as the master table link file
							this.m_TempMDBFileConn = new System.Data.OleDb.OleDbConnection();
							strConn = p_ado.getMDBConnString(this.m_strTempMDBFile,"admin",""); //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.m_strTempMDBFile + ";User Id=admin;Password=;";
							p_ado.OpenConnection(strConn, ref this.m_TempMDBFileConn);
	
							//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\\temp\\temp.mdb;User Id=admin;Password=;";
							//strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFullPathMDB + ";User Id=admin;Password=;";
							strSQL = "SELECT * FROM " + this.lvDataSource.Items[x].SubItems[TABLE].Text.Trim();
							this.m_OleDbDataAdapter.SelectCommand = new System.Data.OleDb.OleDbCommand(strSQL,this.m_TempMDBFileConn);
							try 
							{

								this.m_OleDbDataAdapter.Fill(this.m_dsDataSource,this.lvDataSource.Items[x].SubItems[TABLE].Text.Trim());
							}
							catch (Exception e)
							{
								MessageBox.Show(e.Message,"Table",MessageBoxButtons.OK,MessageBoxIcon.Error);
								this.m_intError=-1;
							}
							this.m_strListedTablesLoadedIntoDatasets[this.m_intNumberOfTablesLoadedIntoDatasets] = this.lvDataSource.Items[x].SubItems[TABLE].Text.Trim();
							this.m_intNumberOfTablesLoadedIntoDatasets++;

							this.m_TempMDBFileConn.Close();
							this.m_TempMDBFileConn= null;
								
						}
					}
				}
			}
			p_ado = null;

			
		}

		public void getNumberOfValidTables()
		{
			int x;
			this.m_intNumberOfValidTables=0;
			for (x=0; x <= this.lvDataSource.Items.Count - 1; x++)
			{
				if (this.lvDataSource.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lvDataSource.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					this.m_intNumberOfValidTables++;
				}
			}

		}

		public void AddDataSourceTablesIntoListBox(System.Windows.Forms.ListBox p_lst)
		{
			int x;
			for (x=0; x <= this.lvDataSource.Items.Count - 1; x++)
			{
				if (this.lvDataSource.Items[x].SubItems[TABLESTATUS].Text.Trim().ToUpper()=="FOUND" &&
					this.lvDataSource.Items[x].SubItems[FILESTATUS].Text.Trim().ToUpper()=="FOUND")
				{
					p_lst.Items.Add(this.lvDataSource.Items[x].SubItems[TABLE].Text.Trim());
				}
			}
		}
		public void TablesFromListBoxToDatasets()
		{
			

		}
		/********************************************************
		 ** return the row associated with the table type
		 ********************************************************/
		private int getDataSourceTableNameRow(string pcTableId)
		{
			int x;
			for (x=0; x<= this.lvDataSource.Items.Count-1;x++)
			{
				if (pcTableId.Trim().ToUpper() == 
					this.lvDataSource.Items[x].SubItems[TABLETYPE].Text.Trim().ToUpper())
				{
					return x;
				}
			}
			return -1;

		}

		private void btnExists_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, " exists ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, " EXISTS ");
			}
		}

		private void btnY_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, "'y' ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, "'Y' ");
			}
		}

		private void btnN_Click(object sender, System.EventArgs e)
		{
			if (this.m_oUtils.IsCapsLockOn() == true)
			{
				this.SendKeyStrokes(this.txtSQLCommand, "'n' ");
			}
			else
			{
				this.SendKeyStrokes(this.txtSQLCommand, "'N' ");
			}
		}
		
	}
}

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario.
	/// </summary>
	public class uc_scenario : System.Windows.Forms.UserControl
	{
		public System.Windows.Forms.Button btnOpen;
		public System.Windows.Forms.TextBox txtDescription;
		private System.Data.DataSet dataSet1;
		private System.Data.OleDb.OleDbCommand oleDbSelectCommand1;
		private System.Data.OleDb.OleDbCommand oleDbInsertCommand1;
		private System.Data.OleDb.OleDbCommand oleDbUpdateCommand1;
		private System.Data.OleDb.OleDbCommand oleDbDeleteCommand1;
		private System.Data.OleDb.OleDbDataAdapter oleDbDataAdapter1;
		private System.Data.OleDb.OleDbCommand oleDbCommand1;
		private System.Data.OleDb.OleDbConnection oleDbConnection1;
		private System.ComponentModel.Container components = null;
		public int intError;
		public System.Windows.Forms.Button btnCancel;
		public string strError;
		public System.Windows.Forms.Label lblScenarioDescription;
		public System.Windows.Forms.Label lblScenarioPath;
		public System.Windows.Forms.TextBox txtScenarioPath;
		public System.Windows.Forms.Label lblNewScenario;
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Button btnClose;
		public System.Windows.Forms.Label lblTitle;
		public System.Windows.Forms.TextBox txtScenarioId;
		public int m_intFullHt=416;
		public int m_intFullWd=632;
		public int m_intError=0;
		public string m_strError="";
		private FIA_Biosum_Manager.frmOptimizerScenario _frmOptimizerScenario;
		private FIA_Biosum_Manager.frmProcessorScenario _frmProcessorScenario;
		private string _strScenarioType="optimizer";
		
		// public FIA_Biosum_Manager.frmScenario frmscenario1;
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		

		public uc_scenario()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();
			this.txtScenarioPath.Enabled=false;
			


			this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
			this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
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
            this.btnOpen = new System.Windows.Forms.Button();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.lblScenarioDescription = new System.Windows.Forms.Label();
            this.lblScenarioPath = new System.Windows.Forms.Label();
            this.txtScenarioPath = new System.Windows.Forms.TextBox();
            this.dataSet1 = new System.Data.DataSet();
            this.oleDbSelectCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbInsertCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbUpdateCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDeleteCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbDataAdapter1 = new System.Data.OleDb.OleDbDataAdapter();
            this.oleDbCommand1 = new System.Data.OleDb.OleDbCommand();
            this.oleDbConnection1 = new System.Data.OleDb.OleDbConnection();
            this.lblNewScenario = new System.Windows.Forms.Label();
            this.txtScenarioId = new System.Windows.Forms.TextBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.BackColor = System.Drawing.SystemColors.Control;
            this.btnOpen.Location = new System.Drawing.Point(208, 392);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(96, 32);
            this.btnOpen.TabIndex = 1;
            this.btnOpen.Text = "OK";
            this.btnOpen.UseVisualStyleBackColor = false;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // txtDescription
            // 
            this.txtDescription.Enabled = false;
            this.txtDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDescription.Location = new System.Drawing.Point(16, 193);
            this.txtDescription.Multiline = true;
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(592, 152);
            this.txtDescription.TabIndex = 2;
            this.txtDescription.TextChanged += new System.EventHandler(this.txtDescription_TextChanged);
            this.txtDescription.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDescription_KeyPress);
            // 
            // lblScenarioDescription
            // 
            this.lblScenarioDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioDescription.Location = new System.Drawing.Point(16, 168);
            this.lblScenarioDescription.Name = "lblScenarioDescription";
            this.lblScenarioDescription.Size = new System.Drawing.Size(160, 16);
            this.lblScenarioDescription.TabIndex = 5;
            this.lblScenarioDescription.Text = "Scenario Description";
            // 
            // lblScenarioPath
            // 
            this.lblScenarioPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioPath.Location = new System.Drawing.Point(16, 112);
            this.lblScenarioPath.Name = "lblScenarioPath";
            this.lblScenarioPath.Size = new System.Drawing.Size(136, 15);
            this.lblScenarioPath.TabIndex = 6;
            this.lblScenarioPath.Text = "Scenario Directory Path";
            // 
            // txtScenarioPath
            // 
            this.txtScenarioPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtScenarioPath.Location = new System.Drawing.Point(16, 129);
            this.txtScenarioPath.Name = "txtScenarioPath";
            this.txtScenarioPath.Size = new System.Drawing.Size(592, 30);
            this.txtScenarioPath.TabIndex = 7;
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            this.dataSet1.Locale = new System.Globalization.CultureInfo("en-US");
            // 
            // oleDbDataAdapter1
            // 
            this.oleDbDataAdapter1.DeleteCommand = this.oleDbDeleteCommand1;
            this.oleDbDataAdapter1.InsertCommand = this.oleDbInsertCommand1;
            this.oleDbDataAdapter1.SelectCommand = this.oleDbSelectCommand1;
            this.oleDbDataAdapter1.UpdateCommand = this.oleDbUpdateCommand1;
            // 
            // lblNewScenario
            // 
            this.lblNewScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNewScenario.Location = new System.Drawing.Point(16, 53);
            this.lblNewScenario.Name = "lblNewScenario";
            this.lblNewScenario.Size = new System.Drawing.Size(128, 15);
            this.lblNewScenario.TabIndex = 9;
            this.lblNewScenario.Text = "Scenario Id";
            // 
            // txtScenarioId
            // 
            this.txtScenarioId.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtScenarioId.Location = new System.Drawing.Point(16, 75);
            this.txtScenarioId.MaxLength = 20;
            this.txtScenarioId.Name = "txtScenarioId";
            this.txtScenarioId.Size = new System.Drawing.Size(120, 30);
            this.txtScenarioId.TabIndex = 10;
            this.txtScenarioId.Leave += new System.EventHandler(this.txtScenarioId_Leave);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.Control;
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(336, 392);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(96, 32);
            this.btnCancel.TabIndex = 14;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.lblNewScenario);
            this.groupBox1.Controls.Add(this.txtScenarioId);
            this.groupBox1.Controls.Add(this.lblScenarioPath);
            this.groupBox1.Controls.Add(this.txtScenarioPath);
            this.groupBox1.Controls.Add(this.lblScenarioDescription);
            this.groupBox1.Controls.Add(this.txtDescription);
            this.groupBox1.Controls.Add(this.btnOpen);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(632, 480);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(626, 32);
            this.lblTitle.TabIndex = 25;
            this.lblTitle.Text = "New Scenario";
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.BackColor = System.Drawing.SystemColors.Control;
            this.btnClose.Location = new System.Drawing.Point(528, 440);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(96, 32);
            this.btnClose.TabIndex = 15;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // uc_scenario
            // 
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_scenario";
            this.Size = new System.Drawing.Size(632, 480);
            this.Load += new System.EventHandler(this.uc_scenario_Load);
            this.MouseDown += new System.Windows.Forms.MouseEventHandler(this.uc_scenario_MouseDown);
            this.Resize += new System.EventHandler(this.uc_scenario_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_scenario_Load(object sender, System.EventArgs e)
		{
		
		}

		private void uc_scenario_Resize(object sender, System.EventArgs e)
		{
			resize_uc_scenario();
		}
		public void resize_uc_scenario()
		{
			try
			{
				this.btnClose.Top = this.groupBox1.Height - this.btnClose.Height - 5;
				this.btnClose.Left = this.groupBox1.Width - this.btnClose.Width - 5;
				this.txtDescription.Width = this.Width - (this.txtDescription.Left * 2);
			}
			catch
			{
			}
		}

		public void NewScenario()
		{
			
			this.btnCancel.Enabled = true;
			this.btnOpen.Enabled = true;
			this.lblNewScenario.Visible = true;
			this.txtScenarioId.Visible = true;
			//intscenario = getScenarioCount() + 1;
            string strScenario = getNextScenarioId();
            this.txtScenarioId.Text = strScenario;
			this.txtScenarioPath.Text = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\" + this.txtScenarioId.Text;
			this.txtScenarioId.Focus();
			this.txtDescription.Enabled=true;
			this.txtDescription.Text = "";

		}
        private string getNextScenarioId()
        {
            string strProjDir = frmMain.g_oFrmMain.getProjectDirectory();
            string strScenarioDir = strProjDir + "\\" + ScenarioType + "\\db";
            System.Collections.Generic.IList<string> lstExistingScenarios = new System.Collections.Generic.List<string>();
            if (ReferenceProcessorScenarioForm != null &&
                ReferenceProcessorScenarioForm.m_bUsingSqlite == true)
            {
                lstExistingScenarios = QueryScenarioNamesSqlite(strScenarioDir);
            }
            else
            {
                lstExistingScenarios = QueryScenarioNames(strScenarioDir);
            }
            

            int i = 1;
            string strTestName;
            // keep incrementing the scenario name until we find one that doesn't exist
            while (i < (lstExistingScenarios.Count + 1) )
            {
                strTestName = "scenario" + Convert.ToString(i);
                if (! lstExistingScenarios.Contains(strTestName))
                    break;
                i++;
            }

            strTestName = "scenario" + Convert.ToString(i);
            return strTestName;
        }

        private System.Collections.Generic.IList<string> QueryScenarioNames(string strScenarioDir)
        {
            ado_data_access oAdo = new ado_data_access();
            string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
            string strFullPath = strScenarioDir + "\\" + strFile;
            string strConn = oAdo.getMDBConnString(strFullPath, "", "");
            System.Collections.Generic.IList<string> lstExistingScenarios = new System.Collections.Generic.List<string>();
            using (var conn = new System.Data.OleDb.OleDbConnection(strConn))
            {
                conn.Open();
                string strSQL = "SELECT scenario_id from " + Tables.Scenario.DefaultScenarioTableName;
                oAdo.SqlQueryReader(conn, strSQL);
                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    // Load all of the existing scenarios into a list we can query
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strScenario = Convert.ToString(oAdo.m_OleDbDataReader["scenario_id"]).Trim();
                        if (!String.IsNullOrEmpty(strScenario))
                        {
                            lstExistingScenarios.Add(strScenario);
                        }
                    }
                }
            }
            return lstExistingScenarios;
        }
        private System.Collections.Generic.IList<string> QueryScenarioNamesSqlite(string strScenarioDir)
        {
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strFile = "scenario_" + ScenarioType + "_rule_definitions.db";
            string strFullPath = strScenarioDir + "\\" + strFile;
            string strConn = dataMgr.GetConnectionString(strFullPath);
            System.Collections.Generic.IList<string> lstExistingScenarios = new System.Collections.Generic.List<string>();
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                string strSQL = "SELECT scenario_id from " + Tables.Scenario.DefaultScenarioTableName;
                dataMgr.SqlQueryReader(conn, strSQL);
                if (dataMgr.m_DataReader.HasRows)
                {
                    // Load all of the existing scenarios into a list we can query
                    while (dataMgr.m_DataReader.Read())
                    {
                        string strScenario = Convert.ToString(dataMgr.m_DataReader["scenario_id"]).Trim();
                        if (!String.IsNullOrEmpty(strScenario))
                        {
                            lstExistingScenarios.Add(strScenario);
                        }
                    }
                }
                dataMgr.m_DataReader.Close();
            }
            return lstExistingScenarios;
        }
        private void btnFolder_Click(object sender, System.EventArgs e)
		{
		
            
			DialogResult result = ((frmMain)this.Parent.Parent).folderBrowserDialog1.ShowDialog();
			//the variable myPic contains the string of the full File Name,it includes the full path. 
			if (result == DialogResult.OK) 
			{
				string strTemp = ((frmMain)this.Parent.Parent).folderBrowserDialog1.SelectedPath;
			
				if (strTemp.Length > 0) 
				{
					this.txtScenarioPath.Text = strTemp;
				}
			}
		}

		public void SaveScenarioPropertiesSqlite()
		{
			bool bOptimizer;
			string strDesc="";
			string strSQL="";
			m_intError=0;
			//validate the input
			//
			//Optimization id
			//
			if (this.txtScenarioId.Text.Length == 0 ) 
			{
                MessageBox.Show("Enter a unique " + ScenarioType + " scenario id");
					this.txtScenarioId.Focus();
					m_intError=-1;
					return;
			}
            else
            {
                System.Text.RegularExpressions.Regex rx =
                    new System.Text.RegularExpressions.Regex("^[a-zA-Z_][a-zA-Z0-9_]*$");

                System.Text.RegularExpressions.MatchCollection matches = rx.Matches(txtScenarioId.Text);
                if (matches.Count < 1)
                {
                    MessageBox.Show("The scenario id contains an invalid character. Only letters, numbers and underscores are permitted!!", "FIA Biosum");
                    this.txtScenarioId.Focus();
                    m_intError=-1;
                    return;
                }
            }
            //
            //check for duplicate scenario id
            //

            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strProjDir = frmMain.g_oFrmMain.getProjectDirectory();
            string strScenarioDir = strProjDir + "\\" + ScenarioType + "\\db";
            string strFile = "scenario_" + ScenarioType + "_rule_definitions.db";
            string strConn = dataMgr.GetConnectionString(strScenarioDir + "\\" + strFile);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                strSQL = "select scenario_id from " + Tables.Scenario.DefaultScenarioTableName;
                dataMgr.SqlQueryReader(conn, strSQL);
                if (dataMgr.m_DataReader.HasRows)
                {
                    // Load all of the existing scenarios into a list we can query
                    while (dataMgr.m_DataReader.Read())
                    {
                        string strScenario = Convert.ToString(dataMgr.m_DataReader["scenario_id"]).Trim();
                        if (this.txtScenarioId.Text.Trim().ToUpper() == strScenario.ToUpper())
                        {
                            this.m_intError = -1;
                            MessageBox.Show("Cannot have a duplicate scenario id");
                            this.txtScenarioId.Focus();
                            dataMgr.m_DataReader.Close();
                            return;
                        }
                    }
                }
                dataMgr.m_DataReader.Close();
            }

            //
            //create the scenario path if it does not exist and
            //copy the scenario_results.mdb to it
            //
            try
			{
				if (!System.IO.Directory.Exists(this.txtScenarioPath.Text)) 
				{
					System.IO.Directory.CreateDirectory(this.txtScenarioPath.Text);
					System.IO.Directory.CreateDirectory(this.txtScenarioPath.Text.ToString() + "\\db");		
				}
                //copy default processor scenario_results database to the new project directory
                if (this.ScenarioType == "processor")
                {
                    string strDestFile = this.txtScenarioPath.Text + "\\" + Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile;
                    if (!System.IO.File.Exists(strDestFile))
                    {
                        dataMgr.CreateDbFile(strDestFile);
                        string strScenarioResultsConn = dataMgr.GetConnectionString(strDestFile);
                        using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strScenarioResultsConn))
                        {
                            conn.Open();
                            frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(dataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                            frmMain.g_oTables.m_oProcessor.CreateSqliteTreeVolValSpeciesDiamGroupsTable(dataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);
                        }
                    }
                }
            }
			catch 
			{
				MessageBox.Show("Error Creating Empty scenario_results.db");
				m_intError=-1;
				return;
			}
            //
            //copy the project data source values to the scenario data source
            //
            string strScenarioDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
            string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.db";
            StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
            strScenarioFullPath.Append("\\");
            strScenarioFullPath.Append(strScenarioFile);
            string strScenarioConn = dataMgr.GetConnectionString(strScenarioFullPath.ToString());

            ado_data_access oAdo = new ado_data_access();
            string strProjDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db";
			string strProjFile = "project.mdb";
			StringBuilder strProjFullPath = new StringBuilder(strProjDBDir);
			strProjFullPath.Append("\\");
			strProjFullPath.Append(strProjFile);
			string strProjConn=oAdo.getMDBConnString(strProjFullPath.ToString(),"admin","");
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strScenarioConn))
            {
                conn.Open();
                using (var p_OleDbProjConn = new System.Data.OleDb.OleDbConnection(strProjConn))
                {
                    p_OleDbProjConn.Open();
                    if (oAdo.m_intError == 0)
                    {
                        if (this.txtDescription.Text.Trim().Length > 0)
                            strDesc = dataMgr.FixString(this.txtDescription.Text.Trim(), "'", "''");
                        strSQL = "INSERT INTO scenario (scenario_id,description,Path,File) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                            "'" + strDesc + "'," +
                            "'" + this.txtScenarioPath.Text.Trim() + "','scenario_" + ScenarioType + "_rule_definitions.db');";
                        dataMgr.SqlNonQuery(conn, strSQL);

                        oAdo.SqlQueryReader(p_OleDbProjConn, "select * from datasource");
                        if (oAdo.m_intError == 0)
                        {
                            try
                            {
                                while (oAdo.m_OleDbDataReader.Read())
                                {
                                    bOptimizer = false;
                                    switch (oAdo.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
                                    {
                                        case "PLOT":
                                            bOptimizer = true;
                                            break;
                                        case "CONDITION":
                                            bOptimizer = true;
                                            break;
                                        case "ADDITIONAL HARVEST COSTS":
                                            bOptimizer = true;
                                            break;
                                        case "TREATMENT PRESCRIPTIONS":
                                            bOptimizer = true;
                                            break;
                                        case "TRAVEL TIMES":
                                            bOptimizer = true;
                                            break;
                                        case "PROCESSING SITES":
                                            bOptimizer = true;
                                            break;
                                        case "TREE":
                                            if (ScenarioType == "processor") bOptimizer = true;
                                            break;
                                        case "HARVEST METHODS":
                                            bOptimizer = true;
                                            break;
                                        case "TREATMENT PACKAGES":
                                            bOptimizer = true;
                                            break;
                                        case "TREE SPECIES":
                                            if (ScenarioType == "processor") bOptimizer = true;
                                            break;
                                        case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                            bOptimizer = true;
                                            break;
                                        case "FIA TREE SPECIES REFERENCE":
                                            if (ScenarioType == "processor") bOptimizer = true;
                                            break;

                                        default:
                                            break;
                                    }
                                    if (bOptimizer == true)
                                    {
                                        strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                                            "'" + oAdo.m_OleDbDataReader["table_type"].ToString().Trim() + "'," +
                                            "'" + oAdo.m_OleDbDataReader["path"].ToString().Trim() + "'," +
                                            "'" + oAdo.m_OleDbDataReader["file"].ToString().Trim() + "'," +
                                            "'" + oAdo.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
                                        dataMgr.SqlNonQuery(conn, strSQL);
                                    }
                                }
                            }
                            catch (Exception caught)
                            {
                                m_intError = -1;
                                m_strError = caught.Message;
                                MessageBox.Show(strError);
                            }
                            oAdo.m_OleDbDataReader.Close();
                        }
                    }
                }
                if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
                {
                    string strTemp = dataMgr.FixString("SELECT @@PlotTable@@.* FROM @@PlotTable@@ ", "'", "''");
                    strSQL = "INSERT INTO scenario_plot_filter (scenario_id,sql_command,current_yn) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                        "'" + strTemp + "'," +
                            "'Y');";
                    dataMgr.SqlNonQuery(conn, strSQL);

                    strTemp = dataMgr.FixString("SELECT @@CondTable@@.* FROM @@CondTable@@", "'", "''");
                    strSQL = "INSERT INTO scenario_cond_filter (scenario_id,sql_command,current_yn) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                        "'" + strTemp + "'," +
                        "'Y');";
                    dataMgr.SqlNonQuery(conn, strSQL);
                }
            }
		}

        public void SaveScenarioProperties()
        {
            bool bOptimizer;
            string strDesc = "";
            string strSQL = "";
            System.Text.StringBuilder strFullPath;
            m_intError = 0;
            //validate the input
            //
            //Optimization id
            //
            if (this.txtScenarioId.Text.Length == 0)
            {
                MessageBox.Show("Enter a unique " + ScenarioType + " scenario id");
                this.txtScenarioId.Focus();
                m_intError = -1;
                return;
            }
            else
            {
                System.Text.RegularExpressions.Regex rx =
                    new System.Text.RegularExpressions.Regex("^[a-zA-Z_][a-zA-Z0-9_]*$");

                System.Text.RegularExpressions.MatchCollection matches = rx.Matches(txtScenarioId.Text);
                if (matches.Count < 1)
                {
                    MessageBox.Show("The scenario id contains an invalid character. Only letters, numbers and underscores are permitted!!", "FIA Biosum");
                    this.txtScenarioId.Focus();
                    m_intError = -1;
                    return;
                }
            }
            //
            //check for duplicate scenario id
            //
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            string strProjDir = frmMain.g_oFrmMain.getProjectDirectory();
            string strScenarioDir = strProjDir + "\\" + ScenarioType + "\\db";
            string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
            strFullPath = new System.Text.StringBuilder(strScenarioDir);
            strFullPath.Append("\\");
            strFullPath.Append(strFile);
            ado_data_access oAdo = new ado_data_access();
            string strConn = oAdo.getMDBConnString(strFullPath.ToString(), "admin", "");
            oAdo.SqlQueryReader(strConn, "select scenario_id from scenario");
            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    if (oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                    {
                        if (this.txtScenarioId.Text.Trim().ToUpper() ==
                            Convert.ToString(oAdo.m_OleDbDataReader["scenario_id"]).Trim().ToUpper())
                        {
                            this.m_intError = -1;
                            MessageBox.Show("Cannot have a duplicate Optimization scenario id");
                            oAdo.m_OleDbDataReader.Close();
                            oAdo.m_OleDbDataReader = null;
                            oAdo = null;
                            this.txtScenarioId.Focus();
                            return;
                        }
                    }
                }
            }
            else
            {
            }
            oAdo.m_OleDbDataReader.Close();
            oAdo.m_OleDbDataReader = null;
            //
            //create the scenario path if it does not exist and
            //copy the scenario_results.mdb to it
            //
            try
            {
                if (!System.IO.Directory.Exists(this.txtScenarioPath.Text))
                {
                    System.IO.Directory.CreateDirectory(this.txtScenarioPath.Text);
                    System.IO.Directory.CreateDirectory(this.txtScenarioPath.Text.ToString() + "\\db");
                }
                //copy default processor scenario_results database to the new project directory
                if (this.ScenarioType == "processor")
                {
                    string strDestFile = this.txtScenarioPath.Text + "\\" + Tables.ProcessorScenarioRun.DefaultHarvestCostsTableDbFile;
                    if (!System.IO.File.Exists(strDestFile))
                    {
                        dao_data_access oDao = new dao_data_access();
                        oDao.CreateMDB(strDestFile);
                        oDao.m_DaoWorkspace.Close();
                        oDao = null;
                        string strScenarioResultsConn = oAdo.getMDBConnString(strDestFile, "", "");
                        System.Data.OleDb.OleDbConnection OleDbScenarioResultsConn = new System.Data.OleDb.OleDbConnection();
                        oAdo.OpenConnection(strScenarioResultsConn, ref OleDbScenarioResultsConn);
                        frmMain.g_oTables.m_oProcessor.CreateHarvestCostsTable(
                            oAdo,
                            OleDbScenarioResultsConn,
                            Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);

                        frmMain.g_oTables.m_oProcessor.CreateTreeVolValSpeciesDiamGroupsTable(
                            oAdo,
                            OleDbScenarioResultsConn,
                            Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName);

                        OleDbScenarioResultsConn.Close();
                        OleDbScenarioResultsConn.Dispose();
                    }
                }
            }
            catch
            {
                MessageBox.Show("Error Creating Folder");
                m_intError = -1;
                return;
            }
            //
            //copy the project data source values to the scenario data source
            //
            string strProjDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\db";
            string strProjFile = "project.mdb";
            StringBuilder strProjFullPath = new StringBuilder(strProjDBDir);
            strProjFullPath.Append("\\");
            strProjFullPath.Append(strProjFile);
            string strProjConn = oAdo.getMDBConnString(strProjFullPath.ToString(), "admin", "");
            System.Data.OleDb.OleDbConnection p_OleDbProjConn = new System.Data.OleDb.OleDbConnection();
            oAdo.OpenConnection(strProjConn, ref p_OleDbProjConn);

            string strScenarioDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
            string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
            StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
            strScenarioFullPath.Append("\\");
            strScenarioFullPath.Append(strScenarioFile);
            string strScenarioConn = oAdo.getMDBConnString(strScenarioFullPath.ToString(), "admin", "");
            oAdo.OpenConnection(strScenarioConn);


            if (oAdo.m_intError == 0)
            {

                if (this.txtDescription.Text.Trim().Length > 0)
                    strDesc = oAdo.FixString(this.txtDescription.Text.Trim(), "'", "''");
                strSQL = "INSERT INTO scenario (scenario_id,description,Path,File) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                    "'" + strDesc + "'," +
                    "'" + this.txtScenarioPath.Text.Trim() + "','scenario_" + ScenarioType + "_rule_definitions.mdb');";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

                oAdo.SqlQueryReader(p_OleDbProjConn, "select * from datasource");
                if (oAdo.m_intError == 0)
                {
                    try
                    {


                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            bOptimizer = false;
                            switch (oAdo.m_OleDbDataReader["table_type"].ToString().Trim().ToUpper())
                            {
                                case "PLOT":
                                    bOptimizer = true;
                                    break;
                                case "CONDITION":
                                    bOptimizer = true;
                                    break;
                                //case "FIRE AND FUEL EFFECTS":
                                //	bCore = true;
                                //	break;
                                //case "HARVEST COSTS":
                                //	bCore = true;
                                //	break;
                                case "ADDITIONAL HARVEST COSTS":
                                    bOptimizer = true;
                                    break;
                                case "TREATMENT PRESCRIPTIONS":
                                    bOptimizer = true;
                                    break;
                                //case "TREE VOLUMES AND VALUES BY SPECIES AND DIAMETER GROUPS":
                                //	bCore = true;
                                //	break;
                                case "TRAVEL TIMES":
                                    bOptimizer = true;
                                    break;
                                case "PROCESSING SITES":
                                    bOptimizer = true;
                                    break;
                                //case "TREE SPECIES AND DIAMETER GROUPS DOLLAR VALUES":
                                //	bCore = true;
                                //	break;
                                case "TREE":
                                    if (ScenarioType == "processor") bOptimizer = true;
                                    break;
                                case "HARVEST METHODS":
                                    bOptimizer = true;
                                    break;
                                case "TREATMENT PACKAGES":
                                    bOptimizer = true;
                                    break;
                                //case "FVS TREE LIST FOR PROCESSOR":
                                //	if (ScenarioType=="processor") bCore=true;
                                //	break;
                                case "TREE SPECIES":
                                    if (ScenarioType == "processor") bOptimizer = true;
                                    break;
                                case "TREATMENT PRESCRIPTIONS HARVEST COST COLUMNS":
                                    bOptimizer = true;
                                    break;
                                case "FIA TREE SPECIES REFERENCE":
                                    if (ScenarioType == "processor") bOptimizer = true;
                                    break;

                                default:
                                    break;
                            }
                            if (bOptimizer == true)
                            {
                                strSQL = "INSERT INTO scenario_datasource (scenario_id,table_type,Path,file,table_name) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                                    "'" + oAdo.m_OleDbDataReader["table_type"].ToString().Trim() + "'," +
                                    "'" + oAdo.m_OleDbDataReader["path"].ToString().Trim() + "'," +
                                    "'" + oAdo.m_OleDbDataReader["file"].ToString().Trim() + "'," +
                                    "'" + oAdo.m_OleDbDataReader["table_name"].ToString().Trim() + "');";
                                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                            }

                        }
                    }
                    catch (Exception caught)
                    {
                        m_intError = -1;
                        m_strError = caught.Message;
                        MessageBox.Show(strError);
                    }
                    oAdo.m_OleDbDataReader.Close();
                    oAdo.m_OleDbDataReader = null;
                    oAdo.m_OleDbCommand = null;
                    p_OleDbProjConn.Close();
                    p_OleDbProjConn = null;
                }
                if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
                {
                    string strTemp = oAdo.FixString("SELECT @@PlotTable@@.* FROM @@PlotTable@@ ", "'", "''");
                    strSQL = "INSERT INTO scenario_plot_filter (scenario_id,sql_command,current_yn) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                        "'" + strTemp + "'," +
                        "'Y');";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);

                    strTemp = oAdo.FixString("SELECT @@CondTable@@.* FROM @@CondTable@@", "'", "''");
                    strSQL = "INSERT INTO scenario_cond_filter (scenario_id,sql_command,current_yn) VALUES " + "('" + this.txtScenarioId.Text.Trim() + "'," +
                        "'" + strTemp + "'," +
                        "'Y');";
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, strSQL);
                }
            }
            oAdo.m_OleDbConnection.Close();
            oAdo.m_OleDbConnection = null;
            oAdo = null;


        }
        public void UpdateDescription()
		{
			string strDesc="";
            string strScenarioDBDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
            if ((ReferenceProcessorScenarioForm != null && !ReferenceProcessorScenarioForm.m_bUsingSqlite)
                 || ScenarioType.Trim().ToUpper() == "OPTIMIZER")
            {
                ado_data_access oAdo = new ado_data_access();
                string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
                StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
                strScenarioFullPath.Append("\\");
                strScenarioFullPath.Append(strScenarioFile);
                string strScenarioConn = oAdo.getMDBConnString(strScenarioFullPath.ToString(), "admin", "");
                oAdo.OpenConnection(strScenarioConn);
                if (this.txtDescription.Text.Trim().Length > 0) strDesc = oAdo.FixString(this.txtDescription.Text.Trim(), "'", "''");
                oAdo.m_strSQL = "UPDATE scenario SET description='" + strDesc.Trim() + "' WHERE TRIM(scenario_id)='" + this.txtScenarioId.Text.Trim() + "'";
                oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                oAdo.m_OleDbConnection.Close();
                while (oAdo.m_OleDbConnection.State != System.Data.ConnectionState.Closed)
                {
                    oAdo.m_OleDbConnection.Close();
                    System.Threading.Thread.Sleep(1000);
                }
                oAdo = null;
            }
            else
            {
                SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.db";
                StringBuilder strScenarioFullPath = new StringBuilder(strScenarioDBDir);
                strScenarioFullPath.Append("\\");
                strScenarioFullPath.Append(strScenarioFile);
                string strConn = oDataMgr.GetConnectionString(strScenarioFullPath.ToString());
                if (this.txtDescription.Text.Trim().Length > 0) strDesc = oDataMgr.FixString(this.txtDescription.Text.Trim(), "'", "''");
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = "UPDATE scenario SET description='" + strDesc.Trim() + "' WHERE TRIM(scenario_id)='" + this.txtScenarioId.Text.Trim() + "'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                oDataMgr = null;
            }
		}
		public string strScenarioId
		{
			get {return this.txtScenarioId.Text;}
			set {this.txtScenarioId.Text = value;}
		}
		public string strScenarioPath
		{
			get {return this.txtScenarioPath.Text;}
			set {this.txtScenarioPath.Text = value;}
		}
		public string strScenarioDescription
		{
			get {return this.txtDescription.Text;}
			set {this.txtDescription.Text = value;}
		}
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			if (this.ScenarioType.Trim().ToUpper()=="OPTIMIZER")
			{
				if (((frmOptimizerScenario)this.ParentForm).m_bScenarioOpen == false) 
				{
					((frmOptimizerScenario)this.ParentForm).Close();
				}
				else 
				{
					this.lblTitle.Text = "";
					((frmOptimizerScenario)this.ParentForm).SetMenu("scenario");
					this.Visible =false;
				}
			}
			else
			{
				if (this.ReferenceProcessorScenarioForm.m_bScenarioOpen==false)
				{
					this.ReferenceProcessorScenarioForm.Close();
				}
				else
				{
					this.lblTitle.Text = "";
					this.Visible=false;
				}

			}

		
		}
		private void lstScenario_Click(object sender, System.EventArgs e)
		{
			
		   
		
		}

		private void txtDescription_TextChanged(object sender, System.EventArgs e)
		{
			
			
 

		}
        public bool DeleteScenario(string p_strScenarioId)
        {

            StringBuilder strFullPath;
            string strSQL = "Delete Scenario '" + p_strScenarioId + "' (Y/N)?";
            DialogResult result = MessageBox.Show(strSQL, "Delete Scenario", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            switch (result)
            {
                case DialogResult.Yes:
                    break;
                case DialogResult.No:
                    return false;
            }


            string strScenarioFile = "scenario_" + ScenarioType + "_rule_definitions.mdb"; //this.txtScenarioMDBFile.Text;
            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";

            string strFile = "scenario_" + ScenarioType + "_rule_definitions.mdb";
            strFullPath = new System.Text.StringBuilder(strScenarioDir);
            strFullPath.Append("\\");
            strFullPath.Append(strFile);
            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            ado_data_access p_ado = new ado_data_access();
            string strConn = p_ado.getMDBConnString(strFullPath.ToString(), "admin", "");
            p_ado.OpenConnection(strConn);
            string strScenarioPath = Convert.ToString(p_ado.getSingleStringValueFromSQLQuery(p_ado.m_OleDbConnection, "SELECT path FROM scenario WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'", "scenario")).Trim();

            strSQL = "DELETE * FROM scenario WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
            p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
            if (p_ado.m_intError == 0)
            {
                strSQL = "DELETE * FROM scenario_datasource WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
            }
            {
                strSQL = "DELETE * FROM scenario_datasource WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
            }
            if (p_ado.m_intError == 0)
            {
                // This table exists in both PROCESSOR and CORE scenario rule definitions
                strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
            }
            if (ScenarioType.Equals("processor"))
            {
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultRxHarvestMethodTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
            }
            else
            {
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
                if (p_ado.m_intError == 0)
                {
                    strSQL = "DELETE * FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    p_ado.SqlNonQuery(p_ado.m_OleDbConnection, strSQL);
                }
            }
            if (p_ado.m_intError == 0)
            {
                try
                {
                    // Delete scenario output folder
                    System.IO.Directory.Delete(strScenarioPath, true);
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }

                strSQL = p_strScenarioId + " was successfully deleted.";
                result = MessageBox.Show(strSQL, "Delete Scenario", MessageBoxButtons.OK);
            }
            p_ado.CloseConnection(p_ado.m_OleDbConnection);
            p_ado = null;
            return true;
        }

        public bool DeleteScenarioSqlite(string p_strScenarioId)
        {

            StringBuilder strFullPath;
            string strSQL = "Delete Scenario '" + p_strScenarioId + "' (Y/N)?";
            DialogResult result = MessageBox.Show(strSQL, "Delete Scenario", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            switch (result)
            {
                case DialogResult.Yes:
                    break;
                case DialogResult.No:
                    return false;
            }

            string strScenarioDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + ScenarioType + "\\db";
            string strFile = "scenario_" + ScenarioType + "_rule_definitions.db";
            string strFolderToDelete = "";
            strFullPath = new StringBuilder(strScenarioDir);
            strFullPath.Append("\\");
            strFullPath.Append(strFile);
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            string strConn = oDataMgr.GetConnectionString(strFullPath.ToString());
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();
                string strScenarioPath = Convert.ToString(oDataMgr.getSingleStringValueFromSQLQuery(conn, 
                    "SELECT path FROM scenario WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'", "scenario")).Trim();

                strSQL = "DELETE FROM scenario WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                oDataMgr.SqlNonQuery(conn, strSQL);
                if (oDataMgr.m_intError == 0)
                {
                    strSQL = "DELETE FROM scenario_datasource WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    oDataMgr.SqlNonQuery(conn, strSQL);
                }
                {
                    strSQL = "DELETE FROM scenario_datasource WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    oDataMgr.SqlNonQuery(conn, strSQL);
                }
                if (oDataMgr.m_intError == 0)
                {
                    // This table exists in both PROCESSOR and CORE scenario rule definitions
                    strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestCostColumnsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                    oDataMgr.SqlNonQuery(conn, strSQL);
                }
                if (ScenarioType.Equals("processor"))
                {
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultRxHarvestMethodTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                }
                else
                {
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterMiscTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                    if (oDataMgr.m_intError == 0)
                    {
                        strSQL = "DELETE FROM " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName + " WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'";
                        oDataMgr.SqlNonQuery(conn, strSQL);
                    }
                }
                strFolderToDelete = Convert.ToString(oDataMgr.getSingleStringValueFromSQLQuery(conn, "SELECT path FROM scenario WHERE scenario_id =  " + "'" + p_strScenarioId.Trim() + "'", "scenario")).Trim();
            }

            if (oDataMgr.m_intError == 0)
            {
                try
                {
                    // Delete scenario output folder
                    if (!String.IsNullOrEmpty(strFolderToDelete))
                    {
                        System.IO.Directory.Delete(strFolderToDelete, true);
                    }                    
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }

                strSQL = p_strScenarioId + " was successfully deleted.";
                result = MessageBox.Show(strSQL, "Delete Scenario", MessageBoxButtons.OK);
            }
            return true;
        }

        private void RefreshForm()
		{
			
		}
		private void txtDescription_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
            if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
            {
            }
            else
                this.ReferenceProcessorScenarioForm.m_bSave = true;
		}

		

		private void uc_scenario_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
            if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				((frmOptimizerScenario)this.ParentForm).m_bPopup = false;
			else
				this.ReferenceProcessorScenarioForm.m_bPopup=false;

		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
            if (this.ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				((frmOptimizerScenario)this.ParentForm).Close();
			else
				this.ReferenceProcessorScenarioForm.Close();
		}

		private void txtScenarioId_Leave(object sender, System.EventArgs e)
		{
			if (this.txtScenarioId.Text.Length > 0) 
			{
				this.txtScenarioId.Text = this.txtScenarioId.Text.Trim();
				//replace spaces with underscores
				this.txtScenarioId.Text = this.txtScenarioId.Text.Replace(" ","_");
				int intLastDir = this.txtScenarioPath.Text.LastIndexOf("\\");
				this.txtScenarioPath.Text = this.txtScenarioPath.Text.Substring(0,intLastDir) + "\\" + this.txtScenarioId.Text;
			}
		}

		private void btnOpen_Click(object sender, System.EventArgs e)
		{
			//lets see if this scenario is already open
			try
			{
				utils p_oUtils = new utils();
				p_oUtils.m_intLevel=1;
                if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
				{
					if (p_oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Treatment Optimizer: Optimization Scenario (" + this.txtScenarioId.Text.Trim() + ")","*",true,false) > 0)
					{
						MessageBox.Show("!!Scenario Already Open!!","Scenario Open",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						return;
					}
				}
				else
				{
					if (p_oUtils.FindWindowLike(frmMain.g_oFrmMain.Handle, "Processor: Scenario (" + this.txtScenarioId.Text.Trim() + ")","*",true,false) > 0)
					{
						MessageBox.Show("!!Scenario Already Open!!","Scenario Open",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						return;
					}
				}
                if (ReferenceProcessorScenarioForm != null &&
                    ReferenceProcessorScenarioForm.m_bUsingSqlite)
                {
                    this.SaveScenarioPropertiesSqlite();
                }
                else
                {
                    this.SaveScenarioProperties();
                }


				if (this.m_intError==0)
				{
					this.btnOpen.DialogResult=DialogResult.OK;
                    if (ScenarioType.Trim().ToUpper() == "OPTIMIZER")
					{
					
						((frmOptimizerScenario)this.ParentForm).DialogResult=DialogResult.OK;
						((frmOptimizerScenario)this.ParentForm).Close();
					}
					else
					{
						this.ReferenceProcessorScenarioForm.DialogResult=DialogResult.OK;
						this.ReferenceProcessorScenarioForm.Close();
					}
					
				}
			}
			catch (Exception err)
			{
				MessageBox.Show("!!Error!! \n" + 
					"Module - uc_scenario:btnOpen_Click  \n" + 
					"Err Msg - " + err.Message.ToString().Trim(),
					"Open Scenario",System.Windows.Forms.MessageBoxButtons.OK,
					System.Windows.Forms.MessageBoxIcon.Exclamation);
			
			}
		}
		public FIA_Biosum_Manager.frmProcessorScenario ReferenceProcessorScenarioForm
		{
			get {return _frmProcessorScenario;}
			set {_frmProcessorScenario=value;}
		}
		public FIA_Biosum_Manager.frmOptimizerScenario ReferenceOptimizerScenarioForm
		{
			get {return _frmOptimizerScenario;}
			set {_frmOptimizerScenario=value;}
		}
		public string ScenarioType
		{
			get {return this._strScenarioType;}
			set {this._strScenarioType=value;}
		}
		
	}
}

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;
using System.Threading;
using System.Data.SQLite;
using System.Data.OleDb;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for uc_scenario_tree_groupings.
	/// </summary>
	public class uc_optimizer_sqlite_export : System.Windows.Forms.UserControl
    {
		private System.Windows.Forms.GroupBox groupBox1;
		public System.Windows.Forms.Label lblTitle;
        private Button BtnExport;
        private FIA_Biosum_Manager.frmMain m_frmMain;
        private frmTherm m_frmTherm;
        private int m_intError;
        private dao_data_access m_oDao;
        private string m_strDebugFile = "";
        private string m_strOptimizerScenario = "";
        private string m_strContextAccdbPath = "";
        private string m_strResultsDbPath = "";
        private string m_strResultsAccdbPath = "";
        private string m_strFvsContextAccdbPath = "";
        private System.Collections.Generic.IDictionary<string, string> m_dictScenarios = null;


        private int m_intDatabaseCount;
        public ListBox lstScenario;
        public Label lblScenarioId;
        public Label lblScenarioDescription;
        public TextBox txtDescription;
        private CheckBox chkFvsContext;
        private CheckBox chkContext;
        private CheckBox chkResults;
        private ComboBox cboResultsDb;
        private Button BtnTestConnection;
        private Button btnHelp;
        private env m_oEnv;
        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultTreatmentOptimizerFile;


        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public uc_optimizer_sqlite_export(FIA_Biosum_Manager.frmMain p_frmMain)
        {
            // This call is required by the Windows.Forms Form Designer.
            InitializeComponent();
            this.m_frmMain = p_frmMain;

            // TODO: Add any initialization after the InitializeComponent call
            this.m_oEnv = new env();
            load_values();
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
            this.BtnTestConnection = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.cboResultsDb = new System.Windows.Forms.ComboBox();
            this.chkFvsContext = new System.Windows.Forms.CheckBox();
            this.chkContext = new System.Windows.Forms.CheckBox();
            this.chkResults = new System.Windows.Forms.CheckBox();
            this.lblScenarioDescription = new System.Windows.Forms.Label();
            this.txtDescription = new System.Windows.Forms.TextBox();
            this.lblScenarioId = new System.Windows.Forms.Label();
            this.lstScenario = new System.Windows.Forms.ListBox();
            this.BtnExport = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnTestConnection);
            this.groupBox1.Controls.Add(this.btnHelp);
            this.groupBox1.Controls.Add(this.cboResultsDb);
            this.groupBox1.Controls.Add(this.chkFvsContext);
            this.groupBox1.Controls.Add(this.chkContext);
            this.groupBox1.Controls.Add(this.chkResults);
            this.groupBox1.Controls.Add(this.lblScenarioDescription);
            this.groupBox1.Controls.Add(this.txtDescription);
            this.groupBox1.Controls.Add(this.lblScenarioId);
            this.groupBox1.Controls.Add(this.lstScenario);
            this.groupBox1.Controls.Add(this.BtnExport);
            this.groupBox1.Controls.Add(this.lblTitle);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(664, 457);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // BtnTestConnection
            // 
            this.BtnTestConnection.Enabled = false;
            this.BtnTestConnection.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnTestConnection.Location = new System.Drawing.Point(235, 369);
            this.BtnTestConnection.Name = "BtnTestConnection";
            this.BtnTestConnection.Size = new System.Drawing.Size(175, 33);
            this.BtnTestConnection.TabIndex = 36;
            this.BtnTestConnection.Text = "Test Connection";
            this.BtnTestConnection.UseVisualStyleBackColor = true;
            this.BtnTestConnection.Click += new System.EventHandler(this.BtnTestConnection_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.btnHelp.Location = new System.Drawing.Point(9, 407);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(115, 37);
            this.btnHelp.TabIndex = 49;
            this.btnHelp.Text = "Help";
            this.btnHelp.Click += new System.EventHandler(this.btnHelp_Click);
            // 
            // cboResultsDb
            // 
            this.cboResultsDb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboResultsDb.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboResultsDb.FormattingEnabled = true;
            this.cboResultsDb.Location = new System.Drawing.Point(235, 255);
            this.cboResultsDb.Name = "cboResultsDb";
            this.cboResultsDb.Size = new System.Drawing.Size(406, 26);
            this.cboResultsDb.TabIndex = 35;
            this.cboResultsDb.SelectedIndexChanged += new System.EventHandler(this.cboResultsDb_SelectedIndexChanged);
            // 
            // chkFvsContext
            // 
            this.chkFvsContext.AutoSize = true;
            this.chkFvsContext.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkFvsContext.Location = new System.Drawing.Point(219, 307);
            this.chkFvsContext.Name = "chkFvsContext";
            this.chkFvsContext.Size = new System.Drawing.Size(149, 22);
            this.chkFvsContext.TabIndex = 34;
            this.chkFvsContext.Text = "fvs_context.accdb";
            this.chkFvsContext.UseVisualStyleBackColor = true;
            this.chkFvsContext.Visible = false;
            // 
            // chkContext
            // 
            this.chkContext.AutoSize = true;
            this.chkContext.Checked = true;
            this.chkContext.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkContext.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkContext.Location = new System.Drawing.Point(219, 282);
            this.chkContext.Name = "chkContext";
            this.chkContext.Size = new System.Drawing.Size(122, 22);
            this.chkContext.TabIndex = 33;
            this.chkContext.Text = "context.accdb";
            this.chkContext.UseVisualStyleBackColor = true;
            this.chkContext.Visible = false;
            // 
            // chkResults
            // 
            this.chkResults.AutoSize = true;
            this.chkResults.Checked = true;
            this.chkResults.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkResults.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkResults.Location = new System.Drawing.Point(219, 258);
            this.chkResults.Name = "chkResults";
            this.chkResults.Size = new System.Drawing.Size(18, 17);
            this.chkResults.TabIndex = 32;
            this.chkResults.UseVisualStyleBackColor = true;
            this.chkResults.Visible = false;
            // 
            // lblScenarioDescription
            // 
            this.lblScenarioDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioDescription.Location = new System.Drawing.Point(216, 50);
            this.lblScenarioDescription.Name = "lblScenarioDescription";
            this.lblScenarioDescription.Size = new System.Drawing.Size(165, 16);
            this.lblScenarioDescription.TabIndex = 31;
            this.lblScenarioDescription.Text = "Scenario Description";
            // 
            // txtDescription
            // 
            this.txtDescription.Enabled = false;
            this.txtDescription.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDescription.Location = new System.Drawing.Point(216, 77);
            this.txtDescription.Multiline = true;
            this.txtDescription.Name = "txtDescription";
            this.txtDescription.Size = new System.Drawing.Size(425, 152);
            this.txtDescription.TabIndex = 30;
            // 
            // lblScenarioId
            // 
            this.lblScenarioId.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblScenarioId.Location = new System.Drawing.Point(6, 50);
            this.lblScenarioId.Name = "lblScenarioId";
            this.lblScenarioId.Size = new System.Drawing.Size(120, 23);
            this.lblScenarioId.TabIndex = 29;
            this.lblScenarioId.Text = "Scenario List";
            // 
            // lstScenario
            // 
            this.lstScenario.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstScenario.ItemHeight = 25;
            this.lstScenario.Location = new System.Drawing.Point(9, 76);
            this.lstScenario.Name = "lstScenario";
            this.lstScenario.Size = new System.Drawing.Size(181, 154);
            this.lstScenario.TabIndex = 28;
            this.lstScenario.SelectedIndexChanged += new System.EventHandler(this.lstScenario_SelectedIndexChanged);
            // 
            // BtnExport
            // 
            this.BtnExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnExport.Location = new System.Drawing.Point(466, 369);
            this.BtnExport.Name = "BtnExport";
            this.BtnExport.Size = new System.Drawing.Size(175, 33);
            this.BtnExport.TabIndex = 27;
            this.BtnExport.Text = "Export";
            this.BtnExport.UseVisualStyleBackColor = true;
            this.BtnExport.Click += new System.EventHandler(this.BtnExport_Click);
            // 
            // lblTitle
            // 
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.ForeColor = System.Drawing.Color.Green;
            this.lblTitle.Location = new System.Drawing.Point(3, 18);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(658, 32);
            this.lblTitle.TabIndex = 26;
            this.lblTitle.Text = "Export to SQLITE";
            // 
            // uc_optimizer_sqlite_export
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_optimizer_sqlite_export";
            this.Size = new System.Drawing.Size(664, 457);
            this.Resize += new System.EventHandler(this.uc_optimizer_sqlite_export_Resize);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

		}
		#endregion

		private void uc_optimizer_sqlite_export_Resize(object sender, System.EventArgs e)
		{
			
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			this.Visible=false;
		}

        public frmDialog ReferenceFormDialog { set; get; }

        private void SetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName, int p_intValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)p_oPb, p_strPropertyName,
                (int)p_intValue);
        }

        private int GetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName)
        {
            return (int)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control)p_oPb,
                p_strPropertyName, false);
        }

        private void SetLabelValue(System.Windows.Forms.Label p_oLabel, string p_strPropertyName, string p_strValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label)p_oLabel, p_strPropertyName,
                p_strValue);
        }

        private void UpdateProgressBar1(string label, int value)
        {
            SetLabelValue(m_frmTherm.lblMsg, "Text", label);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)m_frmTherm.progressBar1,
                "Value", value);
        }
        
        private void BtnExport_Click(object sender, EventArgs e)
        {
            m_intDatabaseCount = 0;
            m_strDebugFile = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\db\sqlite_log.txt";
            bool bDbExists = false;
            if (chkResults.Checked)
            {
                m_intDatabaseCount++;
                if (System.IO.File.Exists(m_strResultsDbPath))
                {
                    bDbExists = true;
                }
            }
            if (chkContext.Checked)
            {
                m_intDatabaseCount++;
                if (System.IO.File.Exists(m_strResultsDbPath))
                {
                    bDbExists = true;
                }
            }
            if (chkFvsContext.Checked)
            {
                m_intDatabaseCount++;
                if (System.IO.File.Exists(m_strResultsDbPath))
                {
                    bDbExists = true;
                }
            }

            if (bDbExists == true)
            {
                DialogResult res = MessageBox.Show("The output SQLITE database already exists!! Do you want to replace it?", "FIA Biosum",
                    MessageBoxButtons.YesNo);
                if (res != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    // Delete existing db
                    if (System.IO.File.Exists(m_strResultsDbPath))
                        System.IO.File.Delete(m_strResultsDbPath);
                }
            }

            if (m_intDatabaseCount > 0)
            {
                CreateSqliteDatabases_Start();
            }
            else
            {
                MessageBox.Show("At least 1 database must be selected to export!!", "FIA Biosum");
            }

        }

        private void CreateSqliteDatabases_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            StartTherm("Create SQLITE Optimizer database");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(CreateSqliteDatabases_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }
        
        private void CreateSqliteDatabases_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            m_intError = 0;
            
            try
            {
                // Delete old log so we start fresh
                if (System.IO.File.Exists(m_strDebugFile))
                    System.IO.File.Delete(m_strDebugFile);

                m_oDao = new dao_data_access();
                
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: SQLITE export Log " + System.DateTime.Now.ToString() + "\r\n\r\n");
                if (frmMain.g_intDebugLevel > 1)
                {

                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Project: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtProjectId.Text + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Project Directory: " + frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Scenario Directory: " + m_strOptimizerScenario + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "---------------------------------------------------------------\r\n");
                }

                //progress bar: single process
                SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                bool bCreateResults = (bool) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox) chkResults, "Checked", false);
                bool bCreateContext = (bool) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox) chkContext, "Checked", false);
                bool bCreateFvsContext = (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)chkFvsContext, "Checked", false);

                SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                oDataMgr.CreateDbFile(m_strResultsDbPath);    // create new, blank database
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created the SQLITE database to hold all tables \r\n");
                }
                
                if (bCreateResults)
                {
                    m_intError = CreateResultsSqliteDb();
                }

                System.Threading.Thread.Sleep(5000);

                if (m_intError == 0 && bCreateContext)
                {
                    m_intError = CreateContextSqliteDb();
                }

                if (m_intError == 0 && bCreateFvsContext)
                {
                    m_intError = CreateFvsContextSqliteDb();
                }

                if (m_intError == 0)
                {
                    UpdateProgressBar1("All databases complete!!", 100);
                }
                else
                {
                    UpdateProgressBar1("An error occurred. Export failed!!", 100);
                }

                if (m_oDao != null)
                {
                    m_oDao.m_DaoWorkspace.Close();
                    m_oDao.m_DaoWorkspace = null;
                    m_oDao = null;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled",
                    true);
                CreateSqliteDatabases_Finish();

                if (m_intError == 0)
                {
                    MessageBox.Show("SQLite database successfully exported!!", "FIA Biosum");
                }
                else
                {
                    MessageBox.Show("An error occurred while exporting the SQLite database. Check the " + m_strDebugFile + " file in the Optimizer scenario/db directory!!", "FIA Biosum");
                }
                
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                ThreadCleanUp();
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
        }

        public int CreateContextSqliteDb()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            ado_data_access oAdo = new ado_data_access();
            
            string strConnection = "data source=" + m_strResultsDbPath;
            string strAccdbConnection = oAdo.getMDBConnString(m_strContextAccdbPath, "", "");
            string strTable = "";
            System.Collections.Generic.IList<string> lstTables = new System.Collections.Generic.List<string>();
            try
            {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsDiameterSpeciesGroupRefTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteDiameterSpeciesGroupRefTable(oDataMgr, con,
                    strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsWeightedVariablesRefTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteFvsWeightedVariableRefTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconWeightedVariablesRefTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEconWeightedVariableRefTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsHarvestMethodRefTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteHarvestMethodRefTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                System.Collections.Generic.IList<string> lstTableNames = 
                    new System.Collections.Generic.List<string>();
                    lstTableNames.Add(Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + "_C");
                    lstTableNames.Add(Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + "_C");
                    System.Collections.Generic.IList<string[]> lstColumnNames =
                    new System.Collections.Generic.List<string[]>();
                System.Collections.Generic.IList<string[]> lstDataTypes =
                    new System.Collections.Generic.List<string[]>();
                // Store the column names and data types in lists so we don't have to close/reopen
                // the ado connection
                using (OleDbConnection oAccessConn = new OleDbConnection(strAccdbConnection))
                {
                    oAccessConn.Open();
                    foreach (var strName in lstTableNames)
                    {
                        string[] strColumnNamesArray = new string[0];
                        string[] strDataTypesArray = new string[0];

                        string strColumnNamesList = "";
                        string strDataTypesList = "";
                        oAdo.getFieldNamesAndDataTypes(oAccessConn, "SELECT * FROM " + strName,
                            ref strColumnNamesList, ref strDataTypesList);
                        if (!String.IsNullOrEmpty(strColumnNamesList))
                        {
                            strColumnNamesArray = strColumnNamesList.Split(",".ToCharArray());
                            strDataTypesArray = strDataTypesList.Split(",".ToCharArray());
                            lstColumnNames.Add(strColumnNamesArray);
                            lstDataTypes.Add(strDataTypesArray);
                        }
                    } 
                }

                for (int i = 0; i < lstTableNames.Count; i++)
                {
                    if (AddDynamicColumnsToTable(strConnection, lstTableNames[i], lstColumnNames[i], lstDataTypes[i]))
                    {
                        lstTables.Add(lstTableNames[i]);
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                        }
                    }
                    else
                    {
                        MessageBox.Show("An error ocurred while creating the " + strTable + " table", "FIA Biosum");
                        return -1;
                    }
                }
               
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsRxPackageRefTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteRxPackageRefTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.FVS.DefaultRxHarvestCostColumnsTableName + "_C";
                frmMain.g_oTables.m_oFvs.CreateSqliteRxHarvestCostColumnTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName + "_C";
                if (this.CreateScenarioAdditionalHarvestCostsTable(strConnection, strTable) == true)
                {
                    lstTables.Add(strTable);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                    }
                }
                else
                {
                    MessageBox.Show("An error occurred while creating the additional harvest costs table!!", "FIA Biosum");
                    return -1;
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsSpeciesGroupRefTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteSpeciesGroupRefTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created target tables \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                }

                using (OleDbConnection oAccessConn = new OleDbConnection(strAccdbConnection))
                {
                    oAccessConn.Open();
                    var counter = 0;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "About to populate context tables \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    }
                    foreach (string strTableName in lstTables)
                    {
                        if (oAdo.TableExist(oAccessConn, strTableName))
                        {
                            counter += 1;
                            string strMessage = "Writing rows to " + strTableName + " in " + System.IO.Path.GetFileName(m_strResultsDbPath);
                            UpdateProgressBar1(strMessage, counter + (100 / (m_intDatabaseCount * 10)));
                            string strSql = "select * from " + strTableName;
                            oAdo.CreateDataTable(oAccessConn, strSql, strTableName, false);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "About to populate context table " + strTableName + " \r\n");
                            }

                            using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSql, con))
                            {
                                using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                {
                                        using (var transaction = con.BeginTransaction())
                                        {
                                            da.InsertCommand = cb.GetInsertCommand();
                                            int rows = da.Update(oAdo.m_DataTable);
                                            transaction.Commit();
                                        }
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    {
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated context table " + strTableName + " \r\n");
                                    }
                                }
                            }
                        }
                    }

                        // Copy POP tables from sqlite master.db
                        string[] arrPopTables = { frmMain.g_oTables.m_oFIAPlot.DefaultPopEvalTableName,
                                                  frmMain.g_oTables.m_oFIAPlot.DefaultPopEstnUnitTableName,
                                                  frmMain.g_oTables.m_oFIAPlot.DefaultPopStratumTableName,
                                                  frmMain.g_oTables.m_oFIAPlot.DefaultPopPlotStratumAssgnTableName};
                        string strMasterDbPath = m_frmMain.getProjectDirectory() + @"\" + frmMain.g_oTables.m_oFIAPlot.DefaultPopTableDbFile;
                        oDataMgr.m_strSQL = "ATTACH DATABASE '" + strMasterDbPath + "' AS MASTER";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile,
                                oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                        foreach(var nextTable in arrPopTables)
                        {
                            // Use the schema from the original tables in master.db to create tables in the context.db
                            oDataMgr.m_strSQL = "SELECT sql FROM master.sqlite_master WHERE type = 'table' " +
                                "AND name = '" + nextTable + "'";
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.m_strSQL = oDataMgr.getSingleStringValueFromSQLQuery(con, oDataMgr.m_strSQL, "sqlite_master");
                            if (!String.IsNullOrEmpty(oDataMgr.m_strSQL))
                            {
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, oDataMgr.m_strSQL + "\r\n");
                                oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);
                            }

                            // Populate the data in the context.db
                            oDataMgr.m_strSQL = "INSERT INTO " + nextTable +
                                " SELECT * FROM master." + nextTable;
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                                frmMain.g_oUtils.WriteText(m_strDebugFile, oDataMgr.m_strSQL + "\r\n");
                            oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);
                        }

                        //detach FIADB database
                        oDataMgr.m_strSQL = "DETACH DATABASE MASTER";
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile,
                                oDataMgr.m_strSQL + "\r\n");
                        oDataMgr.SqlNonQuery(con, oDataMgr.m_strSQL);

                    }
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Finished populating context tables! \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                }
            }
            return 0;
            }
            catch (Exception e)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "CreateContextSqliteDb ERROR!! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                return -1;
            }
            finally
            {
                if (oAdo != null)
                {
                    if (oAdo.m_DataSet != null)
                    {
                        oAdo.m_DataSet.Clear();
                        oAdo.m_DataSet.Dispose();
                    }
                    oAdo = null;
                }

                if (oDataMgr != null)
                {
                    if (oDataMgr.m_DataTable != null)
                    {
                        oAdo.m_DataTable.Clear();
                        oAdo.m_DataTable.Dispose();
                    }
                    if (oDataMgr.m_Connection != null)
                    {
                        oDataMgr.m_Connection.Close();
                        oDataMgr.m_Connection.Dispose();
                    }
                    oDataMgr = null;
                }
            }
        }

        public int CreateFvsContextSqliteDb()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            ado_data_access oAdo = new ado_data_access();

            string[] strTableNames = new string[0];
            System.Collections.Generic.IList<string> lstWeightedTableNames = new System.Collections.Generic.List<string>();
            int tableCount = m_oDao.getTableNames(m_strFvsContextAccdbPath, ref strTableNames);
            string strConnection = "data source=" + (m_strResultsDbPath);
            string strAccdbConnection = oAdo.getMDBConnString(m_strFvsContextAccdbPath, "", "");
            int counter = 1;

            try
            {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                using (OleDbConnection oAccessConn = new OleDbConnection(strAccdbConnection))
                {
                    oAccessConn.Open();
                    foreach(string strTable in strTableNames)
                    {
                        if (! String.IsNullOrEmpty(strTable) && strTable.ToUpper().Contains("WEIGHTED"))
                        {
                            // We will copy this table later. It has a slightly different schema
                            lstWeightedTableNames.Add(strTable);
                        }
                        else if (! String.IsNullOrEmpty(strTable))
                        {
                            // Create the scaffold table
                            Tables.OptimizerScenarioResults.CreateSqliteFvsPrePostTable(oDataMgr, con, strTable, false);
                            string strColumnNamesList = "";
                            string strDataTypesList = "";
                            oAdo.getFieldNamesAndDataTypes(oAccessConn, "SELECT * FROM " + strTable,
                                ref strColumnNamesList, ref strDataTypesList);
                            string[] strColumnNamesArray = new string[0];
                            string[] strDataTypesArray = new string[0];
                            if (!String.IsNullOrEmpty(strColumnNamesList))
                            {
                                strColumnNamesArray = strColumnNamesList.Split(",".ToCharArray());
                                strDataTypesArray = strDataTypesList.Split(",".ToCharArray());
                            }
                            int i = 0;
                            foreach (string strColumn in strColumnNamesArray)
                            {
                                if (! string.IsNullOrEmpty(strColumn))
                                {
                                    if (!oDataMgr.ColumnExist(con, strTable, strColumn))
                                    {
                                        string strDataType = strDataTypesArray[i];
                                        switch (strDataType.ToUpper())
                                        {
                                            case "SYSTEM.STRING":
                                                oDataMgr.AddColumn(con, strTable, strColumn, "TEXT", "");
                                                break;
                                            case "SYSTEM.INT32":
                                                oDataMgr.AddColumn(con, strTable, strColumn, "INTEGER", "");
                                                break;
                                            case "SYSTEM.DOUBLE":
                                                oDataMgr.AddColumn(con, strTable, strColumn, "DOUBLE", "");
                                                break;
                                            default:
                                                MessageBox.Show("Not found!");
                                                break;
                                        }
                                    }
                                }
                                i++;
                            }
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + "table \r\n");
                            }
                            string strMessage = "Writing rows to " + strTable + " in " + System.IO.Path.GetFileName(m_strResultsDbPath);
                            UpdateProgressBar1(strMessage, counter + (100 / (m_intDatabaseCount * 10)));
                            counter++;

                            string strSql = "select * from " + strTable;
                            oAdo.CreateDataTable(oAccessConn, strSql, strTable, false);

                            using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSql, con))
                            {
                                using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                {
                                        using (var transaction = con.BeginTransaction())
                                        {
                                            da.InsertCommand = cb.GetInsertCommand();
                                            int rows = da.Update(oAdo.m_DataTable);
                                            transaction.Commit();
                                        }
                                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    {
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated fvs context table " + strTable + " \r\n");
                                    }
                                }
                            }
                        }
                    }
                    foreach (string strTable in lstWeightedTableNames)
                    {
                        // Create the scaffold table
                        Tables.OptimizerScenarioResults.CreateSqliteFvsPrePostTable(oDataMgr, con, strTable, true);
                        string strColumnNamesList = oAdo.getFieldNames(oAccessConn, "SELECT * FROM " + strTable);
                        string[] strColumnNamesArray = new string[0];
                        if (!String.IsNullOrEmpty(strColumnNamesList))
                        {
                            strColumnNamesArray = strColumnNamesList.Split(",".ToCharArray());
                        }
                        int i = 0;
                        foreach (string strColumn in strColumnNamesArray)
                        {
                            if (!string.IsNullOrEmpty(strColumn))
                            {
                                if (!oDataMgr.ColumnExist(con, strTable, strColumn))
                                {
                                    // Fields for weighted variables are always double/REAL
                                    oDataMgr.AddColumn(con, strTable, strColumn, "REAL", "");
                                }
                            }
                            i++;
                        }
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + "table \r\n");
                        }
                        string strMessage = "Writing rows to " + strTable + " in " + System.IO.Path.GetFileName(m_strResultsDbPath);
                        UpdateProgressBar1(strMessage, counter + (100 / (m_intDatabaseCount * 10)));
                        counter++;

                        string strSql = "select * from " + strTable;
                        oAdo.CreateDataTable(oAccessConn, strSql, strTable, false);

                        using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSql, con))
                        {
                            using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                            {
                                    using (var transaction = con.BeginTransaction())
                                    {
                                        da.InsertCommand = cb.GetInsertCommand();
                                        int rows = da.Update(oAdo.m_DataTable);
                                        transaction.Commit();
                                    }
                                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                {
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated fvs context table " + strTable + " \r\n");
                                }
                            }
                        }
                    }
                }
            }
            return 0;
            }
            catch (Exception e)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "CreateFvsContextSqliteDb ERROR!! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                return -1;
            }
            finally
            {
                if (oAdo != null)
                {
                    if (oAdo.m_DataSet != null)
                    {
                        oAdo.m_DataSet.Clear();
                        oAdo.m_DataSet.Dispose();
                    }
                    oAdo = null;
                }

                if (oDataMgr != null)
                {
                    if (oDataMgr.m_DataTable != null)
                    {
                        oAdo.m_DataTable.Clear();
                        oAdo.m_DataTable.Dispose();
                    }
                    if (oDataMgr.m_Connection != null)
                    {
                        oDataMgr.m_Connection.Close();
                        oDataMgr.m_Connection.Dispose();
                    }
                    oDataMgr = null;
                }
            }
        }

        public void load_values()
        {
            string strScenarioId = "";
            string strDescription = "";
            m_dictScenarios = new System.Collections.Generic.Dictionary<string, string>();    //scenarioName and description

            System.Data.OleDb.OleDbConnection oConn = new System.Data.OleDb.OleDbConnection();
            string strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

            // load scenario list
            ado_data_access p_ado = new ado_data_access();
            string strConn = p_ado.getMDBConnString(strProjDir + "\\"  + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableDbFile, "admin", "");
            try 
            { 
                using (OleDbConnection oAccessConn = new OleDbConnection(strConn))
                {
                    oAccessConn.Open();
                    p_ado.SqlQueryReader(oAccessConn, "select * from scenario");
                    if (p_ado.m_intError == 0)
                    {
                        this.lstScenario.Items.Clear();
                        while (p_ado.m_OleDbDataReader.Read())
                        {
                            strScenarioId = p_ado.m_OleDbDataReader["scenario_id"].ToString().Trim();
                            strDescription = p_ado.m_OleDbDataReader["description"].ToString();
                            this.lstScenario.Items.Add(strScenarioId);
                            m_dictScenarios.Add(strScenarioId, strDescription);
                        }
                        this.lstScenario.SelectedIndex = this.lstScenario.Items.Count - 1;
                    }
                }
            }
            catch (Exception caught)
            {
                MessageBox.Show(caught.Message);
            }
            finally
            {
                if (p_ado.m_OleDbDataReader != null)
                {
                    p_ado.m_OleDbDataReader.Close();
                    p_ado.m_OleDbDataReader = null;
                    p_ado.m_OleDbCommand = null;
                }
                p_ado = null;
            }

        }

        public void update_checkboxes()
        {
            // determine which databases are available to export
            int posY = 258;
            if (cboResultsDb.Items.Count > 0)
            {
                chkResults.Visible = true;
                chkResults.Checked = true;
                cboResultsDb.Visible = true;
                posY = posY + 25;
            }
            else
            {
                chkResults.Visible = false;
                chkResults.Checked = false;
                cboResultsDb.Visible = false;
            }
            if (System.IO.File.Exists(m_strContextAccdbPath))
            {
                chkContext.Visible = true;
                chkContext.Checked = true;
                chkContext.Location = new Point(chkContext.Location.X, posY);
                posY = posY + 25;
            }
            else
            {
                chkContext.Visible = false;
                chkContext.Checked = false;
            }
            chkFvsContext.Checked = false;  // Not checked by default; Export takes a long time
            if (System.IO.File.Exists(m_strFvsContextAccdbPath))
            {
                chkFvsContext.Visible = true;                
                chkFvsContext.Location = new Point(chkFvsContext.Location.X, posY);
            }
            else
            {
                chkFvsContext.Visible = false;
            }
            if (chkResults.Visible == false && chkContext.Visible == false 
                && chkFvsContext.Visible == false)
            {
                BtnExport.Enabled = false;
            }
            else
            {
                BtnExport.Enabled = true;
            }
        }

        public int CreateResultsSqliteDb()
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            ado_data_access oAdo = new ado_data_access();

            string strConnection = "data source=" + m_strResultsDbPath;
            string strAccdbConnection = oAdo.getMDBConnString(m_strResultsAccdbPath, "", "");
            string strTablePrefix = "all_cycles";
            string strTable = "";
            System.Collections.Generic.IList<string> lstTables = new System.Collections.Generic.List<string>();

            try
            {            
            if (!m_oDao.TableExists(m_strResultsAccdbPath, strTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix))
            {
                strTablePrefix = "cycle_1";
            }

            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                strTable = strTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsBestRxSummaryTableSuffix;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteBestRxSummaryTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = strTable + "_before_tiebreaks";
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteBestRxSummaryTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteValidComboTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsValidCombosFVSPrePostTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteValidComboFVSPrePostTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsTieBreakerTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteTieBreakerTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = strTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsOptimizationTableSuffix;
                if (CreateScenarioTableWithDynamicColumns(strConnection,strTable))
                {
                    lstTables.Add(strTable);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                    }
                }
                else
                {
                    MessageBox.Show("An error ocurred while creating the " + strTable + " table", "FIA Biosum");
                    return -1;
                }
                strTable = strTablePrefix + Tables.OptimizerScenarioResults.DefaultScenarioResultsEffectiveTableSuffix;
                if (CreateScenarioTableWithDynamicColumns(strConnection, strTable))
                {
                    lstTables.Add(strTable);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                    }
                }
                else
                {
                    MessageBox.Show("An error ocurred while creating the " + strTable + " table", "FIA Biosum");
                    return -1;
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxCycleTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteProductYieldsTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsEconByRxUtilSumTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEconByRxUtilSumTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsPostEconomicWeightedTableName;
                if (CreateScenarioTableWithDynamicColumns(strConnection, strTable))
                {
                    lstTables.Add(strTable);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                    }
                }
                else
                {
                    MessageBox.Show("An error ocurred while creating the " + strTable + " table", "FIA Biosum");
                    return -1;
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteHaulCostTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }
                strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsCondPsiteTableName;
                frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteCondPsiteTable(oDataMgr, con, strTable);
                lstTables.Add(strTable);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                }

                    System.Collections.Generic.IList<string> lstTableNames = new System.Collections.Generic.List<string>();
                    System.Collections.Generic.IList<string> lstDatabases = new System.Collections.Generic.List<string>();
                    System.Collections.Generic.IDictionary<string, string> dictCustomDb = 
                        new System.Collections.Generic.Dictionary<string, string>();
                    Datasource p_datasource = new Datasource(frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim(),
                                                             m_strOptimizerScenario);
                    int intCondTableType = p_datasource.getDataSourceTableNameRow(Datasource.TableTypes.Condition);
                    if (p_datasource.DataSourceTableExist(intCondTableType))
                    {
                        lstTableNames.Add(p_datasource.m_strDataSource[intCondTableType, Datasource.TABLE].Trim());
                        lstDatabases.Add(p_datasource.m_strDataSource[intCondTableType, Datasource.PATH].Trim() + "\\" +
                                                p_datasource.m_strDataSource[intCondTableType, Datasource.MDBFILE].Trim());
                        dictCustomDb.Add(p_datasource.m_strDataSource[intCondTableType, Datasource.TABLE].Trim(),
                                                p_datasource.m_strDataSource[intCondTableType, Datasource.PATH].Trim() + "\\" +
                                                p_datasource.m_strDataSource[intCondTableType, Datasource.MDBFILE].Trim());
                    }
                    int intPlotTableType = p_datasource.getDataSourceTableNameRow(Datasource.TableTypes.Plot);
                    if (p_datasource.DataSourceTableExist(intPlotTableType))
                    {
                        lstTableNames.Add(p_datasource.m_strDataSource[intPlotTableType, Datasource.TABLE].Trim());
                        lstDatabases.Add(p_datasource.m_strDataSource[intPlotTableType, Datasource.PATH].Trim() + "\\" +
                                                p_datasource.m_strDataSource[intPlotTableType, Datasource.MDBFILE].Trim());
                        dictCustomDb.Add(p_datasource.m_strDataSource[intPlotTableType, Datasource.TABLE].Trim(),
                            p_datasource.m_strDataSource[intPlotTableType, Datasource.PATH].Trim() + "\\" +
                            p_datasource.m_strDataSource[intPlotTableType, Datasource.MDBFILE].Trim());
                    }
                    int intPSiteType = p_datasource.getDataSourceTableNameRow(Datasource.TableTypes.ProcessingSites);
                    if (p_datasource.DataSourceTableExist(intPSiteType))
                    {
                        lstTableNames.Add(p_datasource.m_strDataSource[intPSiteType, Datasource.TABLE].Trim());
                        lstDatabases.Add(p_datasource.m_strDataSource[intPSiteType, Datasource.PATH].Trim() + "\\" +
                                                p_datasource.m_strDataSource[intPSiteType, Datasource.MDBFILE].Trim());
                        dictCustomDb.Add(p_datasource.m_strDataSource[intPSiteType, Datasource.TABLE].Trim(),
                            p_datasource.m_strDataSource[intPSiteType, Datasource.PATH].Trim() + "\\" +
                            p_datasource.m_strDataSource[intPSiteType, Datasource.MDBFILE].Trim());
                    }

                    System.Collections.Generic.IList<string[]> lstColumnNames = new System.Collections.Generic.List<string[]>();
                    System.Collections.Generic.IList<string[]> lstDataTypes = new System.Collections.Generic.List<string[]>();
                    // Store the column names and data types in lists so we don't have to close/reopen
                    // the ado connection
                    int j = 0;
                    using (OleDbConnection oAccessConn = new OleDbConnection())
                    {                        
                        foreach (var strName in lstTableNames)
                        {
                            string strConn = oAdo.getMDBConnString(lstDatabases[j], "", "");
                            if (oAccessConn.State == ConnectionState.Closed)
                            {
                                oAccessConn.ConnectionString = strConn;
                                oAccessConn.Open();
                            }
                            else if (oAccessConn.ConnectionString != strConn)
                            {
                                oAccessConn.Close();
                                oAccessConn.ConnectionString = strConn;
                                oAccessConn.Open();
                            }
                            string[] strColumnNamesArray = new string[0];
                            string[] strDataTypesArray = new string[0];

                            string strColumnNamesList = "";
                            string strDataTypesList = "";
                            oAdo.getFieldNamesAndDataTypes(oAccessConn, "SELECT * FROM " + strName,
                                ref strColumnNamesList, ref strDataTypesList);
                            if (!String.IsNullOrEmpty(strColumnNamesList))
                            {
                                strColumnNamesArray = strColumnNamesList.Split(",".ToCharArray());
                                strDataTypesArray = strDataTypesList.Split(",".ToCharArray());
                                lstColumnNames.Add(strColumnNamesArray);
                                lstDataTypes.Add(strDataTypesArray);
                            }
                            j++;
                        }
                    }

                    for (int i = 0; i < lstTableNames.Count; i++)
                    {
                        if (AddDynamicColumnsToTable(strConnection, lstTableNames[i], lstColumnNames[i], lstDataTypes[i]))
                        {
                            lstTables.Add(lstTableNames[i]);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "Created " + strTable + " table \r\n");
                            }
                        }
                        else
                        {
                            MessageBox.Show("An error ocurred while creating the " + strTable + " table", "FIA Biosum");
                            return -1;
                        }
                    }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Created target tables \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                }

                using (OleDbConnection oAccessConn = new OleDbConnection(strAccdbConnection))
                {
                    oAccessConn.Open();
                    var counter = 0;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    {
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "About to populate results tables \r\n");
                        frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                    }
                    foreach (string strTableName in lstTables)
                    {
                            string strTempConn = strAccdbConnection;
                            // Check for connection strings other than results.accdb (cond, plot, etc)
                            if (dictCustomDb.ContainsKey(strTableName))
                            {
                                strTempConn = oAdo.getMDBConnString(dictCustomDb[strTableName], "", "");
                            }
                            // Close and reset connection if its different than preceding table
                            if (oAccessConn.ConnectionString != strTempConn)
                            {
                                oAccessConn.Close();
                                oAccessConn.ConnectionString = strTempConn;
                                oAccessConn.Open();
                            }
                            if (oAdo.TableExist(oAccessConn, strTableName))
                            {
                            counter += 1;
                            string strMessage = "Writing rows to " + strTableName + " in " + System.IO.Path.GetFileName(m_strResultsDbPath);
                            UpdateProgressBar1(strMessage, counter + (100 / (m_intDatabaseCount * 10)));
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "About to populate results table " + strTableName + " \r\n");
                            }
                            string strSql = "select * from " + strTableName;
                            oAdo.CreateDataTable(oAccessConn, strSql, strTableName, false);

                            using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSql, con))
                            {
                                using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                {
                                        using (var transaction = con.BeginTransaction())
                                        {
                                            da.InsertCommand = cb.GetInsertCommand();
                                            int rows = da.Update(oAdo.m_DataTable);
                                            transaction.Commit();
                                        }
                                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                    {
                                        frmMain.g_oUtils.WriteText(m_strDebugFile, "Populated results table " + strTableName + " \r\n");
                                    }
                                }
                            }
                        }
                    }
                }
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Finished populating results tables! \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                }
            }
            return 0;
            }
            catch (Exception e)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "CreateResultsSqliteDb ERROR!! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                return -1;
            }
            finally
            {
                if (oAdo != null)
                {
                    if (oAdo.m_DataSet != null)
                    {
                        oAdo.m_DataSet.Clear();
                        oAdo.m_DataSet.Dispose();
                    }
                    oAdo = null;
                }

                if (oDataMgr != null)
                {
                    if (oDataMgr.m_DataTable != null)
                    {
                        oAdo.m_DataTable.Clear();
                        oAdo.m_DataTable.Dispose();
                    }
                    if (oDataMgr.m_Connection != null)
                    {
                        oDataMgr.m_Connection.Close();
                        oDataMgr.m_Connection.Dispose();
                    }
                    oDataMgr = null;
                }
            }
        }

        private void CreateSqliteDatabases_Finish()
        {
            if (m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                m_frmTherm = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            ((frmDialog)ParentForm).MinimizeMainForm = false;
        }

        private void StartTherm(string p_strTitle)
        {
            m_frmTherm = new frmTherm((frmDialog)ParentForm, p_strTitle);

            m_frmTherm.Text = p_strTitle;
            m_frmTherm.lblMsg.Text = "";
            m_frmTherm.lblMsg2.Text = "";
            m_frmTherm.Visible = false;
            m_frmTherm.btnCancel.Visible = false;
            m_frmTherm.btnCancel.Enabled = false;
            m_frmTherm.lblMsg.Visible = true;
            m_frmTherm.progressBar1.Minimum = 0;
            m_frmTherm.progressBar1.Visible = true;
            m_frmTherm.progressBar1.Maximum = 10;
            m_frmTherm.AbortProcess = false;
            m_frmTherm.Refresh();
            m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            ((frmDialog)ParentForm).Enabled = false;
            m_frmTherm.Visible = true;
        }

        private void ThreadCleanUp()
        {
            try
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled",
                    true);

                if (m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form)m_frmTherm, "Dispose");

                    m_frmTherm = null;
                }
            }
            catch
            {
            }
        }
        
        private void CodeExamples()
        {
            SQLite.ADO.DataMgr dataMgr = new SQLite.ADO.DataMgr();
            string strConn = @"Data Source=C:\sqlite\db\chinook.db;Version=3;";
            try
            {
                // opening a connection and reading records
                //dataMgr.OpenConnection(strConn);
                //dataMgr.m_strSQL = "select title from albums";
                //dataMgr.SqlQueryReader(dataMgr.m_Connection, dataMgr.m_strSQL);
                //int rowCount = -1;
                //if (dataMgr.m_DataReader.HasRows == true)
                //{
                //    int i = 0;
                //    while(dataMgr.m_DataReader.Read())
                //    {
                //        System.Diagnostics.Debug.WriteLine(i + " " + dataMgr.m_DataReader["title"].ToString().Trim());
                //        i++;
                //    }
                //    rowCount = i;
                //}
                //dataMgr.CloseConnection(dataMgr.m_Connection);

                // sample of copying a SQLITE table link into an MS Access table
                ado_data_access oAdo = new ado_data_access();
                //string strConnection = oAdo.getMDBConnString(@"C:\Docs\Biosum\Blue11_demo_588\optimizer\weighted_ptmod\db\optimizer_results.accdb", "", "");
                //using (OleDbConnection oConn = new OleDbConnection(strConnection))
                //{
                //    oConn.Open();
                //    if (oAdo.TableExist(oConn, "albums"))
                //    {
                //        string albums2 = "albums2";
                //        if (oAdo.TableExist(oConn, albums2))
                //        {
                //            oAdo.m_strSQL = "drop table " + albums2;
                //            oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                //        }
                //        oAdo.m_strSQL = "select * into " + albums2 + " from albums";
                //        oAdo.SqlNonQuery(oConn, oAdo.m_strSQL);
                //        MessageBox.Show("Table created!!");
                //    }
                //}

                //Create database and create table using our MS Access SQL
                //System.Data.SQLite.SQLiteConnection.CreateFile(@"C:\Docs\fia_biosum\Docs\Reporting\attach.db");
                //using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(@"data source=C:\Docs\fia_biosum\Docs\Reporting\attach.db"))
                //{
                //    using (System.Data.SQLite.SQLiteCommand com = new System.Data.SQLite.SQLiteCommand(con))
                //    {
                //        con.Open();
                //        com.CommandText = Tables.OptimizerScenarioResults.CreateHaulCostTableSQL(Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName);
                //        com.ExecuteNonQuery();
                //    }
                //}

                //string strHaulCosts = oAdo.getMDBConnString(@"C:\Docs\Biosum\Blue11_demo_588\optimizer\weighted_ptmod\db\optimizer_results.accdb", "", "");
                //string strSQL = "Select * from " + Tables.OptimizerScenarioResults.DefaultScenarioResultsHaulCostsTableName;
                DataTable dt = new DataTable();
                //using (OleDbDataAdapter da = new OleDbDataAdapter(strSQL, strHaulCosts))
                //{
                //    //tell it not to set rows to Unchanged
                //    da.AcceptChangesDuringFill = false;
                //    da.Fill(dt);
                //}

                //string LiteConnStr = @"data source=C:\Docs\fia_biosum\Docs\Reporting\attach.db";
                //using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(strSQL, LiteConnStr))
                //{
                //    System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da);
                //    da.InsertCommand = cb.GetInsertCommand();
                //    int rows = da.Update(dt);
                //}

                string conFirstDb = @"data source=C:\Docs\fia_biosum\Docs\Reporting\ok_test.db;Version=3;";
                string attachSQL = @"attach 'C:\Docs\fia_biosum\Docs\Reporting\attach.db' as db1";
                string sqlQuery = "select a.biosum_plot_id, b.complete_haul_cost_dpgt from plot as a " +
                                  "inner join db1.haul_costs as b on a.biosum_plot_id = b.biosum_plot_id " +
                                  "where trim(b.materialcd) = 'M'";
                using (SQLiteConnection singleConnectionFor2DBFiles = new SQLiteConnection(conFirstDb))
                {
                    singleConnectionFor2DBFiles.Open();
                    dataMgr.SqlNonQuery(singleConnectionFor2DBFiles, attachSQL);
                    dataMgr.SqlQueryReader(singleConnectionFor2DBFiles, sqlQuery);
                    int rowCount = 0;
                    if (dataMgr.m_DataReader.HasRows == true)
                    {
                        int i = 0;
                        while (dataMgr.m_DataReader.Read())
                        {
                            i++;
                        }
                        rowCount = i;
                    }
                    System.Diagnostics.Debug.WriteLine("number of rows retrieved: " + rowCount);
                }
            }
            catch (Exception e2)
            {
                MessageBox.Show(e2.Message);
            }
        }

        private bool CreateScenarioAdditionalHarvestCostsTable(string strSqliteConn, string strTableName)
        {
            try
            {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strSqliteConn))
            {
                SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                con.Open();
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateSqliteScenarioAdditionalHarvestCostsTable(oDataMgr, con, strTableName);
                string[] strSourceColumnsArray = new string[0];
                m_oDao.getFieldNames(m_strContextAccdbPath, strTableName, ref strSourceColumnsArray);
                foreach (string strColumn in strSourceColumnsArray)
                {
                    if (!oDataMgr.ColumnExists(con, strTableName, strColumn))
                    {
                        oDataMgr.AddColumn(con, strTableName, strColumn, "DOUBLE", "");
                    }
                }
            }
            return true;
            }
            catch (Exception e)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "CreateScenarioAdditionalHarvestCostsTable ERROR!! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                return false;
            }
        }

        private bool CreateScenarioTableWithDynamicColumns(string strSqliteConn, string strTableName)
        {
            try
            {
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strSqliteConn))
                {
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    con.Open();
                    if (strTableName.Contains("optimization"))
                    {
                        frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteOptimizationTable(oDataMgr, con, strTableName);
                    }
                    else if (strTableName.Contains("weighted"))
                    {
                        frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqlitePostEconomicWeightedTable(oDataMgr, con, strTableName);
                    }
                    else
                    {
                        frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteEffectiveTable(oDataMgr, con, strTableName);
                    }
                
                    // This code adds any filter fields that change depending on the scenario configuration
                    string[] strSourceColumnsArray = new string[0];
                    m_oDao.getFieldNames(m_strResultsAccdbPath, strTableName, ref strSourceColumnsArray);
                    foreach (string strColumn in strSourceColumnsArray)
                    {
                        if (!oDataMgr.ColumnExists(con, strTableName, strColumn))
                        {
                            oDataMgr.AddColumn(con, strTableName, strColumn, "DOUBLE", "");
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "CreateScenarioTableWithDynamicColumns ERROR!! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                return false;
            }
        }

        private bool AddDynamicColumnsToTable(string strSqliteConn, string strTableName, string[] strColumnNamesArray,
            string[] strDataTypesArray)
        {
            try
            {
                using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strSqliteConn))
                {
                    SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
                    con.Open();
                    if (strTableName.Equals(Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName + "_C"))
                    {
                        frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(oDataMgr, con, strTableName);
                    }
                    else if (strTableName.Equals(Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName + "_C"))
                    {
                        frmMain.g_oTables.m_oProcessor.CreateSqliteTreeVolValSpeciesDiamGroupsTable(oDataMgr, con, strTableName, true);
                    }
                    else if (strTableName.Equals(frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqliteConditionTable(oDataMgr, con, strTableName);

                    }
                    else if (strTableName.Equals(frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName))
                    {
                        frmMain.g_oTables.m_oFIAPlot.CreateSqlitePlotTable(oDataMgr, con, strTableName);

                    }
                    else if (strTableName.Equals(Tables.TravelTime.DefaultProcessingSiteTableName))
                    {
                        frmMain.g_oTables.m_oTravelTime.CreateSqliteProcessingSiteTable(oDataMgr, con, strTableName);

                    }

                    // This code adds any custom fields the user may have added
                    int i = 0;
                    foreach (string strColumn in strColumnNamesArray)
                    {
                        if (!string.IsNullOrEmpty(strColumn))
                        {
                            if (!oDataMgr.ColumnExist(con, strTableName, strColumn))
                            {
                                string strDataType = strDataTypesArray[i];
                                // Note: Made decision to use SQLite data types here as we don't have a way to get the
                                // precise Access column definition. We would have to write something new. These data types
                                // will only be used for ad-hoc columns that aren't in the standard BioSum tables
                                switch (strDataType.ToUpper())
                                {
                                    case "SYSTEM.STRING":
                                        oDataMgr.AddColumn(con, strTableName, strColumn, "TEXT", "");
                                        break;
                                    case "SYSTEM.INT32":
                                        oDataMgr.AddColumn(con, strTableName, strColumn, "INTEGER", "");
                                        break;
                                    case "SYSTEM.DOUBLE":
                                        oDataMgr.AddColumn(con, strTableName, strColumn, "DOUBLE", "");
                                        break;
                                    case "SYSTEM.SINGLE":
                                        oDataMgr.AddColumn(con, strTableName, strColumn, "SINGLE", "");
                                        break;
                                    default:
                                        MessageBox.Show("Not found!");
                                        break;
                                }
                            }
                        }
                        i++;
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "AddDynamicColumnsToTable ERROR!! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, e.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                return false;
            }
        }

        private void lstScenario_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strKey = "";
            if (this.lstScenario.SelectedIndex >= 0)
            {
                string strValue = "";
                strKey = Convert.ToString(lstScenario.SelectedItem);
                if (m_dictScenarios.TryGetValue(strKey, out strValue) == true)
                {
                    txtDescription.Text = strValue;
                    m_strOptimizerScenario = strKey;
                }
                else
                {
                    txtDescription.Text = "";
                    m_strOptimizerScenario = "";
                }
            }
            else
            {
                txtDescription.Text = "";
                m_strOptimizerScenario = "";
            }
            m_strContextAccdbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsContextDbFile;
            m_strResultsDbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsSqliteResultsDbFile;
            m_strFvsContextAccdbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\" + Tables.OptimizerScenarioResults.DefaultScenarioResultsFvsContextDbFile;
            BtnExport.Enabled = !String.IsNullOrEmpty(strKey);
            cboResultsDb.Items.Clear();
            chkResults.Enabled = false;
            if (!String.IsNullOrEmpty(m_strOptimizerScenario))
            {
                string strDirectory = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\db";
                string[] arrFilePaths = System.IO.Directory.GetFiles(strDirectory, "*.accdb");
                foreach (string strFilePath in arrFilePaths)
                {
                    string strFileName = System.IO.Path.GetFileName(strFilePath);
                    if (strFileName.ToLower().Contains("optimizer_results"))
                    {
                        cboResultsDb.Items.Add(strFileName);
                    }
                }
                cboResultsDb.SelectedIndex = cboResultsDb.FindStringExact(System.IO.Path.GetFileName(Tables.OptimizerScenarioResults.DefaultScenarioResultsDbFile));
                if (cboResultsDb.Items.Count > 0)
                {
                    chkResults.Enabled = true;
                }
                BtnTestConnection.Enabled = true;
            }
            else
            {
                BtnTestConnection.Enabled = false;
            }
            update_checkboxes();
        }

        private void cboResultsDb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboResultsDb.SelectedIndex > -1)
            {
                string selectedDb = cboResultsDb.GetItemText(cboResultsDb.SelectedItem);
                m_strResultsAccdbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\db\" + selectedDb;
            }
        }

        private void BtnTestConnection_Click(object sender, EventArgs e)
        {
            SQLite.ADO.DataMgr oDataMgr = new SQLite.ADO.DataMgr();
            m_strDebugFile = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\db\sqlite_log.txt";
            // Delete old log so we start fresh
            if (System.IO.File.Exists(m_strDebugFile))
                System.IO.File.Delete(m_strDebugFile);

            frmMain.g_oUtils.WriteText(m_strDebugFile, "START: SQLITE test Log " + System.DateTime.Now.ToString() + "\r\n\r\n");

            string strTestDbPath = m_frmMain.getProjectDirectory() + @"\optimizer\" + m_strOptimizerScenario + @"\TestSQLite.db3";
            string strConnection = "data source=" + strTestDbPath;
            if (System.IO.File.Exists(strTestDbPath) == true)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Test database found: " + strTestDbPath + ". Attempt to delete \r\n");
                System.IO.File.Delete(strTestDbPath);
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Test database deleted! \r\n");
            }
            try
            {
                oDataMgr.CreateDbFile(strTestDbPath);
                frmMain.g_oUtils.WriteText(m_strDebugFile, "Test SQLite database successfully created \r\n");
                using (SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
                {
                    con.Open();
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Connection made to test SQLite database \r\n");
                    string strTable = Tables.OptimizerScenarioResults.DefaultScenarioResultsTieBreakerTableName;
                    frmMain.g_oTables.m_oOptimizerScenarioResults.CreateSqliteTieBreakerTable(oDataMgr, con, strTable);
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Sample table successfully created in test SQLite database \r\n");
                }
                if (System.IO.File.Exists(strTestDbPath) == true)
                {
                    System.IO.File.Delete(strTestDbPath);
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Test database deleted after testing! \r\n");
                }
                MessageBox.Show("Passed all tests for using SQLite!!", "FIA Biosum");
            }
            catch (Exception ex)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "An ERROR occurred while trying to connect with the sqlite database! \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, ex.Message + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, ex.StackTrace + " \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n");
                MessageBox.Show("Unable to connect with an SQLite database. Check the log file at " + strTestDbPath + "!!", "FIA Biosum");
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "TREATMENT_OPTIMIZER", "EXPORT_SQLITE" });
        }

    }


}

using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;
using SQLite.ADO;
using System.Data.OleDb;

namespace FIA_Biosum_Manager
{


    public class uc_fvs_create_mdbs : System.Windows.Forms.UserControl
    {
        private GroupBox groupBox1;
        private TextBox textBox1;
        private ToolTip toolTip1;
        private IContainer components;
        private Button button1;
        private string m_strProjDir;
        private Queries m_oQueries;
        private Tables m_oTables;
        public int m_intError;
        private ado_data_access m_ado;
        private dao_data_access m_dao;
        private string m_strTempMDBFileConnectionString;
        private env m_oEnv;
        private utils m_oUtils;
        private bool m_bDebug;
        private Datasource m_oDatasource = null;


        public uc_fvs_create_mdbs(string p_strProjDir)
        {
            InitializeComponent();
            InitializeDatasource();
            this.m_strProjDir = p_strProjDir;
            this.m_oQueries = new Queries();
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oFIAPlot.LoadDatasource = true;
            m_oQueries.LoadDatasources(true);

            if (m_oQueries.m_oFvs.m_strFvsTreeTable.Trim().Length == 0)
            {
                m_oQueries.m_oFvs.m_strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
            }

            this.m_ado = new ado_data_access();
            this.m_dao = new dao_data_access();
            this.m_strTempMDBFileConnectionString = this.m_ado.getMDBConnString(this.m_oQueries.m_strTempDbFile, "", "");
            this.m_ado.OpenConnection(this.m_strTempMDBFileConnectionString);
            if (this.m_ado.m_intError != 0)
            {
                this.m_intError = this.m_ado.m_intError;
                this.m_ado = null;
                return;

            }
            this.m_oEnv = new env();
            this.m_bDebug = frmMain.g_bDebug;
        }

        private void InitializeDatasource()
        {
            string strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

            m_oDatasource = new Datasource();
            m_oDatasource.LoadTableColumnNamesAndDataTypes = false;
            m_oDatasource.LoadTableRecordCount = false;
            m_oDatasource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
            m_oDatasource.m_strDataSourceTableName = "datasource";
            m_oDatasource.m_strScenarioId = "";
            m_oDatasource.populate_datasource_array();

            ////get table names
            //this.m_strPlotTable = m_oDatasource.getValidDataSourceTableName("PLOT");
            //this.m_strCondTable = m_oDatasource.getValidDataSourceTableName("CONDITION");
            //this.m_strTreeTable = m_oDatasource.getValidDataSourceTableName("TREE");
            //this.m_strSiteTreeTable = m_oDatasource.getValidDataSourceTableName("SITE TREE");
            //this.m_strTreeRegionalBiomassTable = m_oDatasource.getValidDataSourceTableName("TREE REGIONAL BIOMASS");
            //this.m_strPpsaTable = m_oDatasource.getValidDataSourceTableName("POPULATION PLOT STRATUM ASSIGNMENT");
            //this.m_strPopEstUnitTable = m_oDatasource.getValidDataSourceTableName("POPULATION ESTIMATION UNIT");
            //this.m_strPopStratumTable = m_oDatasource.getValidDataSourceTableName("POPULATION STRATUM");
            //this.m_strPopEvalTable = m_oDatasource.getValidDataSourceTableName("POPULATION EVALUATION");
            //this.m_strBiosumPopStratumAdjustmentFactorsTable = m_oDatasource.getValidDataSourceTableName("BIOSUM POP STRATUM ADJUSTMENT FACTORS");
            //this.m_strTreeMacroPlotBreakPointDiaTable = m_oDatasource.getValidDataSourceTableName("FIA TREE MACRO PLOT BREAKPOINT DIAMETER");
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(32, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(684, 377);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(20, 336);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(126, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Create MDBs";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(20, 21);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(638, 285);
            this.textBox1.TabIndex = 1;
            // 
            // uc_fvs_create_mdbs
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_create_mdbs";
            this.Size = new System.Drawing.Size(746, 421);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        private void CreateMDBs_Main()
        {
            // When you need to see how to update progress bars:
            // RunAppend_Main in uc_fvs_output.cs! (also good for seeing how to interact with m_intError;

            //get the fiadb table structures
            DataMgr oDataMgr = new DataMgr();
            var dbFileName = "fvsout.db";
            var dbPath = "\\fvs\\data\\" + dbFileName;
            string strSQLiteConnection = oDataMgr.GetConnectionString(m_strProjDir + dbPath);
            var oUtils = new utils(); // Cargo cult!(?) For doing SQLite to Acces oledb reader conversions
            var oEnv = new env(); // Cargo cult!(?) For getting stuff about the temp directory and app directory
            dao_data_access oDao = new dao_data_access();
            var fileNamesList = new List<List<string>>();
            fileNamesList = getRunTitleFilenames();
            // GetFVSVariantRxPackageSQL
            var tableQueriesList = getTableCreateQueries(strSQLiteConnection, dbFileName, oDataMgr);
            foreach (var file in fileNamesList)
            {
                // Create file in root\fvs\data\<variant>\<filename>
                var strDbPathFile = m_strProjDir + "\\fvs\\data\\"+file[1]+"\\"+file[0];

                if (System.IO.File.Exists(strDbPathFile))
                {
                    System.IO.File.Delete(strDbPathFile);
                    // TODO: DEBUG LOG HERE.
                    appendStringToDebugTextbox($@"File exists. Deleting: {strDbPathFile}");
                }
                oDao.CreateMDB(strDbPathFile);
                // Open a connection to new file 
                using (var accessConn = new OleDbConnection(m_ado.getMDBConnString(strDbPathFile, "", "")))
                {
                    // Populate tables via these queries
                    // Log to files in utils.cs WriteText method
                    appendStringToDebugTextbox($@"Connecting to: {file[0]}");
                    accessConn.Open();
                    // TODO: Map goofy table names to proper ones. Either via a mapping, regex, or a simple string compare. FVS_Summary2 -> FVS_Summary
                    var validTables = Tables.FVS.g_strFVSOutTablesArray;
                    createMDBTablesfromSQLite(strDbPathFile, oDataMgr, oDao, strSQLiteConnection);
                    // Populate new tables from SQLite
                        // Code written for #223 does something similar
                    // Get answers for what analysts would prefer to do for setting base year?
                        // Make text box as prototype.
                    // Diff new and old access DBs if possible.
                        // Idea: Use Tyler's access macro on the new and old, and compare outputs.
                    // Progress indicators? Instantiate thermometer? Calculate max number of steps and add increments. Use delegate to update bar from background thread.
                        // Update BioSum ready/working indicator in lower right corner of frmMain.
                    // Cancel button
                    // Debug log with log levels.
                }



                // Add index (if needed)
                //var strTempIndex = column + "_delete_idx";
                //if (!m_dao.IndexExists(strDbPathFile, table, strTempIndex))
                //{
                //    m_ado.AddIndex(conn, table, strTempIndex, column);
                //}
            }
            // TODO: PotFireBaseYr (sp?) special case handling. Each DB has two tables, FVS_Cases and FVS_PotFire.
            // Only one PotFireBaseYr db per variant. Probably.

            appendStringToDebugTextbox("Done.");
            return;
        }

        private void appendStringToDebugTextbox(string text)
        {
            var textBoxValue = frmMain.g_oDelegate.GetControlPropertyValue(this.textBox1, "Text", false);
            frmMain.g_oDelegate.SetControlPropertyValue(this.textBox1, "Text", textBoxValue += text +System.Environment.NewLine);

        }
        private string dataTypeConvert(string dataTypeFromDB)
        {
            var convertedType = "";
            switch (dataTypeFromDB)
            {
                case "SYSTEM.INT32":
                    convertedType = "LONG";
                    break;
                case "SYSTEM.DOUBLE":
                    convertedType = "DOUBLE";
                    break;
                case "SYSTEM.STRING":
                    convertedType = "CHAR(255)";
                    break;
                default:
                    convertedType = "UNRECOGNIZED";
                    break;
            }
            // SYSTEM.INT32 > "LONG"
            // SYSTEM.DOUBLE > "DOUBLE"
            // SYSTEM.STRING > "CHAR(255)"
            // ??? > "SINGLE"
            return convertedType;
        }

        private List<string> getTableCreateQueries(string strConnection, string fileName, DataMgr oDataMgr)
        {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                //getTableNames
                var tableNames = oDataMgr.getTableNames(con);
                //build field list string to insert sql by matching 
                //up the column names in the biosum plot table and the fiadb plot table

                var strFields = "";
                var createTableQueriesList = new List<string>();

                // Run this loop for each database we need to make.
                foreach (var tblName in tableNames)
                {
                    DataTable dtSourceSchema = oDataMgr.getTableSchema(con, $"select * from {tblName}");
                    var sb = new System.Text.StringBuilder();
                    var strCol = "";
                    sb.Append($@"CREATE TABLE {tblName} (");

                    // TODO Make CaseID Unique? Make it the index (via ado_data_access index creation method?)
                    for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                    {
                        strCol = dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper() + " " + dataTypeConvert(dtSourceSchema.Rows[y]["datatype"].ToString().ToUpper());
                        if (strFields.Trim().Length == 0)
                        {
                            strFields = strCol;
                        }
                        else
                        {
                            strFields += "," + strCol;
                        }
                    }
                    sb.Append(strFields + ") ");
                    createTableQueriesList.Add(sb.ToString());
                }
                return createTableQueriesList;

            }

        }

        private void createMDBTablesfromSQLite(string strMDBPathAndFile, DataMgr oDataMgr, dao_data_access oDao, string strConnection)
        {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                foreach (var tblName in oDataMgr.getTableNames(con))
                {
                    DataTable dtSourceSchema = oDataMgr.getTableSchema(con, $"select * from {tblName}");
                    // TODO: Map goofy table names to proper ones. Either via a mapping, regex, or a simple string compare. FVS_Summary2 -> FVS_Summary
                    var validTables = Tables.FVS.g_strFVSOutTablesArray;
                    var validTablesList = new List<string>(validTables);
                    var tablesSet = new HashSet<string>(validTablesList);
                    var newTblName = tblName;
                    if (!tablesSet.Contains(tblName))
                    {
                        if (tblName.ToUpper().Contains("SUMMARY"))
                        {
                            newTblName = "FVS_Summary";
                        }
                    }
                    oDao.CreateMDBTableFromDataSetTable(strMDBPathAndFile, newTblName, dtSourceSchema, true);
                }
            }

        }

        private List<List<string>> getRunTitleFilenames()
        {
            // List of three items; databasename, variant, runtitle.
            var fileNames = new List<List<string>>();
            var strTempMDB = m_oDatasource.CreateMDBAndTableDataSourceLinks();
            // Similar example: uc_delete_conditions.cs line 504 (Execute
            var fileNameList = new List<string>();
            using (System.Data.OleDb.OleDbConnection con = new System.Data.OleDb.OleDbConnection(m_ado.getMDBConnString(strTempMDB, "", "")))
            {
                con.Open();
                this.m_ado.SqlQueryReader(con, Queries.FVS.GetFVSVariantRxPackageSQL(m_oQueries.m_oFIAPlot.m_strPlotTable, m_oQueries.m_oFvs.m_strRxPackageTable));

                while (this.m_ado.m_OleDbDataReader.Read())
                {
                    //rxpackage,simyear1_rx,simyear2_rx,simyear3_rx,simyear4_rx,rxcycle_length
                    var strVariant = this.m_ado.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                    var strPackage = this.m_ado.m_OleDbDataReader["rxpackage"].ToString().Trim();
                    var strRx1 = this.m_ado.m_OleDbDataReader["simyear1_rx"].ToString().Trim();
                    strRx1 = String.IsNullOrEmpty(strRx1) ? "000" : strRx1;
                    var strRx2 = this.m_ado.m_OleDbDataReader["simyear2_rx"].ToString().Trim();
                    strRx2 = String.IsNullOrEmpty(strRx2) ? "000" : strRx2;
                    var strRx3 = this.m_ado.m_OleDbDataReader["simyear3_rx"].ToString().Trim();
                    strRx3 = String.IsNullOrEmpty(strRx3) ? "000" : strRx3;
                    var strRx4 = this.m_ado.m_OleDbDataReader["simyear4_rx"].ToString().Trim();
                    strRx4 = String.IsNullOrEmpty(strRx4) ? "000" : strRx4;
                    // Example: FVSOUT_CA_P001-001-000-000-000.MDB

                    // Examples:
                    //FVSOUT_CA_P003 - 003 - 000 - 000 - 000
                    //FVSOUT_CA_POTFIRE_BaseYr
                    //FVSOUT_CA_P001 - 001 - 000 - 000 - 000
                    //FVSOUT_CA_P002 - 002 - 000 - 000 - 000
                    //FVSOUT_CA_P005 - 000 - 000 - 000 - 000
                    var runTitle = $@"FVSOUT_{strVariant}_P{strPackage}-{strRx1}-{strRx2}-{strRx3}-{strRx4}";
                    var fileName = $@"FVSOUT_{strVariant}_P{strPackage}-{strRx1}-{strRx2}-{strRx3}-{strRx4}.MDB";
                    fileNameList.Add(fileName);
                    fileNames.Add(new List<string>() { fileName, strVariant, runTitle });
                }
            }
            return fileNames;
        }
        private void button1_Click(object sender, EventArgs e)
        {

            //if (this.lstFvsOutput.CheckedItems.Count == 0)
            //{
            //    MessageBox.Show("No Boxes Are Checked", "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
            //    return;
            //}

            //this.m_frmTherm = new frmTherm(((frmDialog)ParentForm), "FVS OUT DATA",
            //    "FVS Output", "2");
            //m_frmTherm.Visible = false;
            //this.m_frmTherm.lblMsg.Text = "";
            //this.m_frmTherm.TopMost = true;

            //this.cmbStep.Enabled = false;
            //this.btnExecute.Enabled = false;
            //this.btnChkAll.Enabled = false;
            //this.btnClearAll.Enabled = false;
            //this.btnRefresh.Enabled = false;
            //this.btnClose.Enabled = false;
            //this.btnHelp.Enabled = false;
            //this.btnCancel.Visible = false;
            //this.btnViewLogFile.Enabled = false;
            //this.btnViewPostLogFile.Enabled = false;
            //this.btnAuditDb.Enabled = false;
            //this.btnPostAppendAuditDb.Enabled = false;

            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Maximum", 100);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Minimum", 0);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.progressBar2, "Value", 0);
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg2, "Text", "Overall Progress");
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Text", "");
            //frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control)this.m_frmTherm.lblMsg, "Visible", true);
            //m_frmTherm.Show((frmDialog)ParentForm);

            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(CreateMDBs_Main));
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        

        internal void uc_fvs_create_mdbs_Resize()
        {
            return;
        }
    }
}

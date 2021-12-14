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
using System.Linq;

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
        private Dictionary<string, List<Tuple<string, utils.DataType>>> m_listDictFVSOutputTablesColumnsDefinitions;
        private Dictionary<string, string> m_dictCreateTableQueries;

        // Mapping of sqlite to Access tables names. Add new entries here.
        public static Dictionary<string, string> sqliteToAccessTblNames = new Dictionary<string, string>
        {
            { "FVS_SUMMARY2", "FVS_SUMMARY" }
        };
        public static Dictionary<string, string> AccessToSqliteTblNames = sqliteToAccessTblNames.ToDictionary((i) => i.Value, (i) => i.Key);


        public uc_fvs_create_mdbs(string p_strProjDir)
        {
            InitializeComponent();
            InitializeDatasource();
            this.m_listDictFVSOutputTablesColumnsDefinitions = new Dictionary<string, List<Tuple<string, utils.DataType>>>();
            this.m_dictCreateTableQueries = new Dictionary<string, string>();
            this.m_strProjDir = p_strProjDir;
            this.m_oQueries = new Queries();
            m_oQueries.m_oFvs.LoadDatasource = true;
            m_oQueries.m_oFIAPlot.LoadDatasource = true;
            m_oQueries.LoadDatasources(true);

            if (m_oQueries.m_oFvs.m_strFvsTreeTable.Trim().Length == 0)
            {
                m_oQueries.m_oFvs.m_strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
            }

            textBox1.TextChanged += (sender, e) =>
            {
                if (textBox1.Visible)
                {
                    textBox1.SelectionStart = textBox1.TextLength;
                    textBox1.ScrollToCaret();
                }
            };
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
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Location = new System.Drawing.Point(32, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(875, 435);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // textBox1
            // 
            this.textBox1.AccessibleDescription = "This textbox outputs logs on the Create MDB process.";
            this.textBox1.AccessibleName = "Output Textbox";
            this.textBox1.Location = new System.Drawing.Point(20, 21);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(849, 379);
            this.textBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(20, 406);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(126, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Create MDBs";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // uc_fvs_create_mdbs
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_create_mdbs";
            this.Size = new System.Drawing.Size(937, 470);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
        }

        private void CreateMDBs_Main()
        {
            // When you need to see how to update progress bars:
            // RunAppend_Main in uc_fvs_output.cs! (also good for seeing how to interact with m_intError;

            //get the fiadb table structures
            DataMgr sqliteDataMgr = new DataMgr();
            var dbFileName = "fvsout.db";
            var dbPath = "\\fvs\\data\\" + dbFileName;
            string strSQLiteConnection = sqliteDataMgr.GetConnectionString(m_strProjDir + dbPath);
            var oUtils = new utils(); // Cargo cult!(?) For doing SQLite to Acces oledb reader conversions
            var oEnv = new env(); // Cargo cult!(?) For getting stuff about the temp directory and app directory
            dao_data_access oDao = new dao_data_access();
            var fileNamesList = new List<List<string>>();
            fileNamesList = getRunTitleFilenames();
            // GetFVSVariantRxPackageSQL
            populateTableQueryDictionaries(strSQLiteConnection,sqliteDataMgr);

            foreach (var file in fileNamesList)
            {
                // Create file in root\fvs\data\<variant>\<filename>
                var strDbPathFile = m_strProjDir + "\\fvs\\data\\"+file[1]+"\\"+file[0];

                var fi = new System.IO.FileInfo(strDbPathFile);
                if (fi.Exists)
                {
                    // Unhandled exception if file is open in Access; catch and prompt user to close and retry?
                    fi.Delete();
                    fi.Refresh();
                    while (fi.Exists)
                    {
                        Thread.Sleep(100);
                        fi.Refresh();
                    }
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
                    executeSQLListOnAccessConnection(m_dictCreateTableQueries, accessConn, m_ado);
                    // Populate new tables from SQLite
                    using (var sqliteConn = new System.Data.SQLite.SQLiteConnection(strSQLiteConnection))
                    {
                        sqliteConn.Open();
                        foreach (var tblName in m_listDictFVSOutputTablesColumnsDefinitions.Keys)
                        {
                            var accessTblName = convertAccessTblNameToSqliteTblName(tblName);
                            var cols = m_listDictFVSOutputTablesColumnsDefinitions[tblName];
                            var strColumns = string.Join(",", m_listDictFVSOutputTablesColumnsDefinitions[tblName].Select(item => wrapInBackTick(item.Item1)));
                            if (tblName.ToUpper().Contains("SUMMARY"))
                            {
                                var problem = true;
                            }
                            sqliteDataMgr.SqlQueryReader(sqliteConn, generateRuntitleSubsetQuery(accessTblName, file[2]));
                            appendStringToDebugTextbox(generateRuntitleSubsetQuery(tblName, file[2]));


                            // TODO: Make a dictionary mapping Access -> SQLite table names to SQLite -> Access
                            if (sqliteDataMgr.m_DataReader.HasRows)
                            {
                                System.Data.OleDb.OleDbTransaction transaction;
                                System.Data.OleDb.OleDbCommand command = accessConn.CreateCommand();
                                // Start a local transaction
                                transaction = accessConn.BeginTransaction(IsolationLevel.ReadCommitted);
                                // Assign transaction object for a pending local transaction
                                command.Transaction = transaction;
                                try
                                {
                                    var recordCount = 0;
                                    while (sqliteDataMgr.m_DataReader.Read())
                                    {
                                        if (sqliteDataMgr.m_DataReader["CASEID"] != DBNull.Value && Convert.ToString(sqliteDataMgr.m_DataReader["CASEID"]).Trim().Length > 0)
                                        {
                                            // Can't use year without backtick, can't use backticks

                                            var strValues = utils.GetParsedInsertValues(sqliteDataMgr.m_DataReader, m_listDictFVSOutputTablesColumnsDefinitions[tblName]);
                                            command.CommandText = $"INSERT INTO {tblName} ({strColumns}) VALUES ({strValues})";
                                            command.ExecuteNonQuery();
                                            recordCount++;
                                        }
                                    }
                                    transaction.Commit();
                                    appendStringToDebugTextbox($@"Inserted {recordCount} records into {tblName}");

                                }
                                catch (Exception err)
                                {
                                    m_intError = -1;
                                    appendStringToDebugTextbox(err.Message);
                                    transaction.Rollback();
                                }
                                transaction.Dispose();
                            }
                            sqliteDataMgr.m_DataReader.Dispose();
                            
                        }
                    }

                    // Code written for #223 does something similar
                    // Get answers for what analysts would prefer to do for setting base year?
                    // Make text box as prototype.
                    // Diff new and old access DBs if possible.utputs.
                    // Progress indicators? Instantiate thermometer? Calculate max number o
                    // Idea: Use Tyler's access macro on the new and old, and compare onf steps and add increments. Use delegate to update bar from background thread.
                    // Update BioSum ready/working indicator in lower right corner of frmMai.
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
            m_dictCreateTableQueries.Clear();
            m_listDictFVSOutputTablesColumnsDefinitions.Clear();
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
                    convertedType = "TEXT(255)";
                    break;
                default:
                    convertedType = "UNRECOGNIZED";
                    break;
            }
            // ??? > "SINGLE" ?
            return convertedType;
        }

        private string generateRuntitleSubsetQuery(string strTableName, string strRunTitle)
        {
            return !strTableName.ToUpper().Contains("CASES") ? $"SELECT f.* FROM FVS_CASES c INNER JOIN {strTableName} f ON c.CaseID=f.CaseID WHERE c.RunTitle='{strRunTitle}'" : $"SELECT c.* FROM {strTableName} c WHERE c.RunTitle='{strRunTitle}'";
        }

        private utils.DataType getDataTypeEnumValueFromString(string dataTypeFromDB)
        {
            utils.DataType convertedType;
            switch (dataTypeFromDB)
            {
                case "SYSTEM.INT32":
                    convertedType = utils.DataType.INTEGER;
                    break;
                case "SYSTEM.DOUBLE":
                    convertedType = utils.DataType.DOUBLE;
                    break;
                case "SYSTEM.STRING":
                    convertedType = utils.DataType.STRING;
                    break;
                    // Byte case?
                case "SYSTEM.BYTE":
                    convertedType = utils.DataType.BYTE;
                    break;
                default:
                    convertedType = utils.DataType.STRING;
                    break;
            }
            // ??? > "SINGLE" ?
            return convertedType;
        }
        // TODO: Make this a void function
        private void populateTableQueryDictionaries(string strConnection, DataMgr oDataMgr)
        {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                //getTableNames
                var tableNames = oDataMgr.getTableNames(con);
                //build field list string to insert sql by matching 
                //up the column names in the biosum plot table and the fiadb plot table


                // Run this loop for each database we need to make.
                foreach (var tblName in tableNames)
                {
                    DataTable dtSourceSchema = oDataMgr.getTableSchema(con, $"select * from {tblName}");
                    var sb = new System.Text.StringBuilder();
                    var strCol = "";
                    sb.Append($@"CREATE TABLE {convertSqliteTblNameToBiosumTblName(tblName)} (");
                    var strFields = "";
                    var listColDataTypes = new List<Tuple<string, utils.DataType>>();
                    // TODO Make CaseID Unique? Make it the index (via ado_data_access index creation method?)
                    for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                    {
                        var colName = translateColumn(dtSourceSchema.Rows[y]["columnname"].ToString()).ToUpper();
                        //if (colName.Contains("YEAR"))
                        //{
                        //    continue;
                        //}

                        var dataType = dtSourceSchema.Rows[y]["datatype"].ToString().ToUpper();

                        listColDataTypes.Add(Tuple.Create(colName, getDataTypeEnumValueFromString(dataType)));

                        strCol = wrapInBackTick(colName) + " " + dataTypeConvert(dataType);
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
                    m_dictCreateTableQueries[convertSqliteTblNameToBiosumTblName(tblName)] =sb.ToString();
                    m_listDictFVSOutputTablesColumnsDefinitions[convertSqliteTblNameToBiosumTblName(tblName)] = listColDataTypes;
                }
            }
        }

        private void executeSQLListOnAccessConnection(Dictionary<string,string> queryDict, OleDbConnection accessConn, ado_data_access oAdo)
        {
            foreach (var tblName in queryDict.Keys)
            {
                appendStringToDebugTextbox($@"Creating table: {tblName}");
                oAdo.SqlNonQuery(accessConn, queryDict[tblName]);
                //oAdo.AddIndex(accessConn, tblName, tblName+"_CaseId_Idx", "CaseID");
            }
        }

        private string translateColumn (string strToCheck)
        {
            var translatedStr = strToCheck.ToUpper();
            // Map SPECIESFIA to SPECIES. TODO: Make prettier?
            if (translatedStr.Contains("SPECIESFIA"))
            {
                translatedStr = translatedStr.Replace("SPECIESFIA", "SPECIES");
                appendStringToDebugTextbox($@"Mapping SpeciesFIA to Species for col: {strToCheck}. New value: {translatedStr}");

            }
            return translatedStr;
        }

        // Disused due to datatype incompatibilities
        //private void createMDBTablesfromSQLite(string strMDBPathAndFile, DataMgr oDataMgr, dao_data_access oDao, string strConnection)
        //{
        //    using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
        //    {
        //        con.Open();
        //        foreach (var tblName in oDataMgr.getTableNames(con))
        //        {
        //            DataTable dtSourceSchema = oDataMgr.getTableSchema(con, $"select * from {tblName}");
        //            var newTblName = convertTblNameToBiosumTblName(tblName);
        //            oDao.CreateMDBTableFromDataSetTable(strMDBPathAndFile, newTblName, dtSourceSchema, true);
        //        }
        //    }
        //}

        private string convertSqliteTblNameToBiosumTblName(string tblName)
        {
            // TODO: Map changed table names to proper ones. Either via a mapping, regex, or a simple string compare. FVS_Summary2 -> FVS_Summary
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
            return newTblName;
        }

        private string convertAccessTblNameToSqliteTblName(string tblName)
        {
            // TODO: Map changed table names to proper ones. Either via a mapping, regex, or a simple string compare. FVS_Summary2 -> FVS_Summary
            var validTables = Tables.FVS.g_strFVSOutTablesArray;
            var validTablesList = new List<string>(validTables);
            var tablesSet = new HashSet<string>(validTablesList);
            var newTblName = tblName;
            if (!tablesSet.Contains(tblName))
            {
                if (tblName.ToUpper().Contains("SUMMARY"))
                {
                    newTblName = "FVS_Summary2";
                }
            }
            return newTblName;
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

        private string wrapInBackTick(string str)
        {
            return $@"`{str}`";
        }

        internal void uc_fvs_create_mdbs_Resize()
        {
            return;
        }
    }
}

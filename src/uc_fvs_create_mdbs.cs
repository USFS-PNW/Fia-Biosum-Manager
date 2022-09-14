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
using System.IO;

namespace FIA_Biosum_Manager
{
    public class uc_fvs_create_mdbs : System.Windows.Forms.UserControl
    {
        private GroupBox groupBox1;
        private TextBox textBox1;
        private IContainer components;
        private Button btnCreateMdbs;
        private string m_strProjDir;
        public int m_intError;
        private ado_data_access m_ado;
        private string m_strTempMDBFileConnectionString;
        private env m_oEnv;
        private bool m_bDebug;
        private Dictionary<string, List<Tuple<string, utils.DataType>>> m_listDictFVSOutputTablesColumnsDefinitions;
        private Dictionary<string, string> m_dictCreateTableQueries;
        Dictionary<string, List<string>> m_dictRunTitleTables;

        // Mapping of sqlite column names to biosum column names. Add new entries here.
        public static Dictionary<string, string> sqliteToAccessColNames = new Dictionary<string, string>
        {
            { "SPECIESFIA", "SPECIES" }
        };

        // Only tables we want to create from POTFIRE BaseYr runtitle
        public static List<string> lstPotfireTables = new List<string> { "FVS_Cases", "FVS_PotFire", "FVS_Fuels", "FVS_BurnReport",
            "FVS_Consumption", "FVS_Mortality", "FVS_SnagSum", "FVS_SnagDet", "FVS_Carbon", "FVS_Hrv_Carbon", "FVS_Down_Wood_Cov",
            "FVS_Down_Wood_Vol"};
        public static string runTitle = "RUNTITLE";
        private Button btnCancel;
        private Button btnExportLog;
        private ToolTip createMdbsTooltip;
        private ToolTip cancelTooltip;
        private ToolTip exportLogTooltip;
        private Button btnClose;

        public uc_fvs_create_mdbs(string p_strProjDir)
        {
            InitializeComponent();
            this.m_listDictFVSOutputTablesColumnsDefinitions = new Dictionary<string, List<Tuple<string, utils.DataType>>>();
            this.m_dictCreateTableQueries = new Dictionary<string, string>();
            this.m_strProjDir = p_strProjDir;
            Queries oQueries = new Queries();
            oQueries.m_oFvs.LoadDatasource = true;
            oQueries.m_oFIAPlot.LoadDatasource = true;
            oQueries.LoadDatasources(true);

            if (oQueries.m_oFvs.m_strFvsTreeTable.Trim().Length == 0)
            {
                oQueries.m_oFvs.m_strFvsTreeTable = Tables.FVS.DefaultFVSTreeTableName;
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
            this.m_strTempMDBFileConnectionString = this.m_ado.getMDBConnString(oQueries.m_strTempDbFile, "", "");
            this.m_ado.OpenConnection(this.m_strTempMDBFileConnectionString);
            if (this.m_ado.m_intError != 0)
            {
                this.m_ado.CloseAndDisposeConnection(this.m_ado.m_OleDbConnection, true);
                this.m_intError = this.m_ado.m_intError;
                this.m_ado = null;
                return;
            }
            this.m_oEnv = new env();
            this.m_bDebug = frmMain.g_bDebug;
        }

        private void InitializeDatasource(Datasource p_dataSource)
        {
            string strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            p_dataSource.LoadTableColumnNamesAndDataTypes = false;
            p_dataSource.LoadTableRecordCount = false;
            p_dataSource.m_strDataSourceMDBFile = strProjDir.Trim() + "\\db\\project.mdb";
            p_dataSource.m_strDataSourceTableName = "datasource";
            p_dataSource.m_strScenarioId = "";
            p_dataSource.populate_datasource_array();
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnExportLog = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnCreateMdbs = new System.Windows.Forms.Button();
            this.createMdbsTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.cancelTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.exportLogTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnExportLog);
            this.groupBox1.Controls.Add(this.btnCancel);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.btnCreateMdbs);
            this.groupBox1.Location = new System.Drawing.Point(32, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(940, 465);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Create MDBs Output";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(809, 434);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(125, 25);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnExportLog
            // 
            this.btnExportLog.Enabled = false;
            this.btnExportLog.Location = new System.Drawing.Point(276, 434);
            this.btnExportLog.Name = "btnExportLog";
            this.btnExportLog.Size = new System.Drawing.Size(125, 25);
            this.btnExportLog.TabIndex = 3;
            this.btnExportLog.Text = "Export log as .txt";
            this.btnExportLog.UseVisualStyleBackColor = true;
            this.btnExportLog.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Enabled = false;
            this.btnCancel.Location = new System.Drawing.Point(145, 434);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(125, 25);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // textBox1
            // 
            this.textBox1.AccessibleDescription = "This textbox outputs logs on the Create MDB process.";
            this.textBox1.AccessibleName = "Output Textbox";
            this.textBox1.Location = new System.Drawing.Point(6, 21);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(928, 407);
            this.textBox1.TabIndex = 1;
            // 
            // btnCreateMdbs
            // 
            this.btnCreateMdbs.Location = new System.Drawing.Point(13, 434);
            this.btnCreateMdbs.Name = "btnCreateMdbs";
            this.btnCreateMdbs.Size = new System.Drawing.Size(125, 25);
            this.btnCreateMdbs.TabIndex = 0;
            this.btnCreateMdbs.Text = "Create MDBs";
            this.btnCreateMdbs.UseVisualStyleBackColor = true;
            this.btnCreateMdbs.Click += new System.EventHandler(this.createMdbsMain_Click);
            // 
            // createMdbsTooltip
            // 
            this.createMdbsTooltip.ToolTipTitle = "Create MDBs Button Tooltip";
            // 
            // exportLogTooltip
            // 
            this.exportLogTooltip.ToolTipTitle = "Export Logs Button Tooltip";
            // 
            // uc_fvs_create_mdbs
            // 
            this.Controls.Add(this.groupBox1);
            this.Name = "uc_fvs_create_mdbs";
            this.Size = new System.Drawing.Size(1000, 500);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        /// <summary>Main function for the "create MDBS" process. Executes the process of creating the Access files, subsetting the SQLite DB,
        /// and inserting those rows into the new Access DBs. The result should be compatible with BioSum.</summary>
        private void CreateMDBs_Main()
        {
            frmMain.g_oDelegate.SetControlPropertyValue(this.btnExportLog, "Enabled", false);
            frmMain.g_oDelegate.SetControlPropertyValue(this.btnCancel, "Enabled", true);
            // Get the fiadb table structures
            appendStringToDebugTextbox("Generating MDBs on this date and time:" + DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString().ToString());
            DataMgr sqliteDataMgr = new DataMgr();
            var dbFileName = "fvsout.db";
            var dbPath = "\\fvs\\data\\" + dbFileName;
            string strSQLiteConnection = sqliteDataMgr.GetConnectionString(m_strProjDir + dbPath);
            // Initialize utils, environment, data access objects
            var oUtils = new utils();
            var oEnv = new env();
            dao_data_access oDao = new dao_data_access();
            var fileNamesList = new List<List<string>>();
            fileNamesList = getRunTitleFilenames();
            populateTableQueryDictionaries(strSQLiteConnection, sqliteDataMgr);

            // Iterate over file name list to create new access files and populate them.
            foreach (var file in fileNamesList)
            {
                // Create file in root\fvs\data\<variant>\<filename>
                var strDbPathFile = m_strProjDir + "\\fvs\\data\\" + file[1] + "\\" + file[0];
                var strRunTitle = System.IO.Path.GetFileNameWithoutExtension(file[0]);

                var fi = new FileInfo(strDbPathFile);
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
                    appendStringToDebugTextbox($@"File exists. Deleting: {strDbPathFile}");
                }
                oDao.CreateMDB(strDbPathFile);

                // Open a connection to new file 
                using (var accessConn = new OleDbConnection(m_ado.getMDBConnString(strDbPathFile, "", "")))
                {
                    // Populate tables via these queries
                    appendStringToDebugTextbox($@"Connecting to: {file[0]}");
                    accessConn.Open();
                    executeSQLListOnAccessConnection(m_dictCreateTableQueries, accessConn, m_ado, strRunTitle);
                    // Populate new tables from SQLite
                    using (var sqliteConn = new System.Data.SQLite.SQLiteConnection(strSQLiteConnection))
                    {
                        sqliteConn.Open();
                        foreach (var tblName in m_listDictFVSOutputTablesColumnsDefinitions.Keys)
                        {
                            if (m_ado.TableExist(accessConn, tblName))
                                // Add this to avoid exceptions and to support POTFIRE (only has 2 tables)
                            {
                                var cols = m_listDictFVSOutputTablesColumnsDefinitions[tblName];
                                // Generate comma-seperated column string for insert statements. Wrap in back-tick to prevent "reserved word" errors.
                                var strColumns = string.Join(",", m_listDictFVSOutputTablesColumnsDefinitions[tblName].Select(item => wrapInBackTick(translateColumn(item.Item1))));

                                // Open up a data manager for the subsetted query for the run title we're on.  
                                sqliteDataMgr.SqlQueryReader(sqliteConn, generateRuntitleSubsetQuery(tblName, file[2]));
                                appendStringToDebugTextbox(generateRuntitleSubsetQuery(tblName, file[2]));
                                if (sqliteDataMgr.m_DataReader.HasRows)
                                {
                                    OleDbTransaction transaction;
                                    OleDbCommand command = accessConn.CreateCommand();
                                    // Start a local transaction
                                    transaction = accessConn.BeginTransaction(IsolationLevel.ReadCommitted);
                                    // Assign transaction object for a pending local transaction
                                    command.Transaction = transaction;
                                    try
                                    {
                                        // Iterate over reader, create, and execute insert statements. Count them for display to user.
                                        var recordCount = 0;
                                        while (sqliteDataMgr.m_DataReader.Read())
                                        {
                                            if (sqliteDataMgr.m_DataReader["CASEID"] != DBNull.Value && Convert.ToString(sqliteDataMgr.m_DataReader["CASEID"]).Trim().Length > 0)
                                            {
                                                var strValues = utils.GetParsedInsertValues(sqliteDataMgr.m_DataReader, m_listDictFVSOutputTablesColumnsDefinitions[tblName]);
                                                // Insert statements
                                                command.CommandText = $"INSERT INTO {tblName} ({strColumns}) VALUES ({strValues})";
                                                command.ExecuteNonQuery();
                                                recordCount++;
                                            }
                                        }
                                        transaction.Commit();
                                        // Update runtitle as needed
                                        command.CommandText = $"UPDATE {tblName} SET {runTitle} = IIF(IsNull({runTitle}),'{file[2]}',{runTitle})";
                                        command.ExecuteNonQuery();
                                        appendStringToDebugTextbox($@"Inserted {recordCount} records into {tblName}");
                                    }
                                    catch (Exception err)
                                    {
                                        m_intError = -1;
                                        appendStringToDebugTextbox(err.Message);
                                        transaction.Rollback();
                                    }
                                    transaction.Dispose();
                                    command.Dispose();
                                }
                                sqliteDataMgr.m_DataReader.Dispose();
                                sqliteDataMgr.m_DataReader = null;
                            }
                        }
                        sqliteConn.Close();
                    }
                    accessConn.Close();
                    // Potential future TODOs:
                    // Get answers for what analysts would prefer to do for setting base year?
                    // Diff new and old access DBs if possible.
                    // Progress indicators? Instantiate thermometer? Calculate max number of steps and show progress?
                    // Idea: Use Tyler's access macro on the new and old, and compare of steps and add increments. Use delegate to update bar from background thread.
                    // Update BioSum ready/working indicator in lower right corner of frmMain.
                }
            }
            // TODO: PotFireBaseYr (sp?) special case handling. Each DB has two tables, FVS_Cases and FVS_PotFire.
            // Only one PotFireBaseYr db per variant. Probably.
            appendStringToDebugTextbox("Done.");

            // Lesley added this
            frmMain.g_oDelegate.CurrentThreadProcessDone = true;
            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

            CleanupThread();
            frmMain.g_oDelegate.SetControlPropertyValue(this.btnExportLog, "Enabled", true);
            frmMain.g_oDelegate.SetControlPropertyValue(this.btnCancel, "Enabled", false);
            return;
        }

        /// <summary>
        /// Closes the form.
        /// </summary>
        /// <param name="sender"> Sender object (unused).</param>
        /// <param name="e">Event arguments (unused).</param>
        private void btnClose_Click(object sender, System.EventArgs e)
        {
            this.ParentForm.Close();
        }

        /// <summary>Appends the supplied string to the main output textbox. Adds a newline, so just include the message you want to output. </summary>
        /// <param name="text">The string to append to the textbox.</param>
        private void appendStringToDebugTextbox(string text)
        {
            var textBoxValue = frmMain.g_oDelegate.GetControlPropertyValue(this.textBox1, "Text", false);
            frmMain.g_oDelegate.SetControlPropertyValue(this.textBox1, "Text", textBoxValue += text + System.Environment.NewLine);
        }

        /// <summary>
        /// Event handler for the export button. 
        /// Takes the text from the textbox and supplies it to the user to download as a .txt file. 
        /// Only enabled when the process is complete.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.FileName = "CreatMDBSExportLog-" + DateTime.Now.ToShortDateString().Replace("/", "-") + ".txt";
            save.Filter = "Text File | *.txt";
            if (save.ShowDialog() == DialogResult.OK)
            {
                StreamWriter writer = new StreamWriter(save.OpenFile());
                for (int i = 0; i < textBox1.Lines.Count(); i++)
                {
                    writer.WriteLine(textBox1.Lines[i].ToString());
                }
                writer.Dispose();
                writer.Close();
            }
        }

        /// <summary>Converts the string representation of an SQLite data type to the equivalent Access data type strings. </summary>
        /// <param name="dataTypeFromDB">The string representation of an SQLite data type.</param>
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
                    convertedType = "VARCHAR(255)";
                    break;
                default:
                    convertedType = "UNRECOGNIZED";
                    break;
            }
            return convertedType;
        }

        /// <summary>Generates a select query to subset the supplied table by the given run title. Used to populate access dbs from the SQLite tables. </summary>
        /// <param name="strTableName">The SQLite table name to subset.</param>
        ///  <param name="strRunTitle">The run title to subset the table by.</param>
        private string generateRuntitleSubsetQuery(string strTableName, string strRunTitle)
        {
            return !strTableName.ToUpper().Contains("CASES") ? $"SELECT f.* FROM FVS_CASES c INNER JOIN {strTableName} f ON c.CaseID=f.CaseID WHERE c.RunTitle='{strRunTitle}'" : $"SELECT c.* FROM {strTableName} c WHERE c.RunTitle='{strRunTitle}'";
        }

        /// <summary>Converts the string representation of an SQLite data type to the equivalent utils.DataType enum value. </summary>
        /// <param name="dataTypeFromDB">The string representation of an SQLite data type.</param>
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
                case "SYSTEM.BYTE":
                    convertedType = utils.DataType.BYTE;
                    break;
                default:
                    convertedType = utils.DataType.STRING;
                    break;
            }
            return convertedType;
        }

        /// <summary>This method populates dictionaries storing table creation scripts and 
        /// for the new Access MDBs, and a dictionary of column definition information for each type of table
        /// within those Access MDBs. </summary>
        /// <param name="strConnection">The connection string for the source SQLite table.</param>
        /// <param name="oDataMgr">An instance of the DataMgr class.</param>
        private void populateTableQueryDictionaries(string strConnection, DataMgr oDataMgr)
        {
            using (System.Data.SQLite.SQLiteConnection con = new System.Data.SQLite.SQLiteConnection(strConnection))
            {
                con.Open();
                //getTableNames
                var tableNames = oDataMgr.getTableNames(con);
                //Build field list string to insert sql by matching up the column names in the biosum plot table and the fiadb plot table

                // Run this loop for each database we need to make.
                m_dictRunTitleTables = new Dictionary<string, List<string>>();
                foreach (var tblName in tableNames)
                {
                    // We only want tables that start with FVS_
                    if (tblName.ToUpper().Contains("FVS_"))
                    {
                        DataTable dtSourceSchema = oDataMgr.getTableSchema(con, $"select * from {tblName}");
                        var sb = new System.Text.StringBuilder();
                        var strCol = "";
                        sb.Append($@"CREATE TABLE {convertSqliteTblNameToBiosumTblName(tblName)} (");
                        var strFields = "";
                        var listColDataTypes = new List<Tuple<string, utils.DataType>>();
                        // Iterate over table schema defined in Data Table format. Check debugger to see columns.
                        var hasRunTitleField = false;   // only add runtitle field if table doesn't have it
                        for (int y = 0; y <= dtSourceSchema.Rows.Count - 1; y++)
                        {
                            var colName = dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper();
                            // This maps SPECIESFIA to SPECIES currently (and in the future should map any other column name differences between SQLite and target access DB
                            var convertedColName = translateColumn(dtSourceSchema.Rows[y]["columnname"].ToString().ToUpper());
                            if (convertedColName.ToUpper() == runTitle)
                            {
                                hasRunTitleField = true;
                            }
                            var dataType = dtSourceSchema.Rows[y]["datatype"].ToString().ToUpper();

                            listColDataTypes.Add(Tuple.Create(colName, getDataTypeEnumValueFromString(dataType)));

                            // Use converted name here. We want SPECIES for the access creation, and SPECIESFIA for the selects.
                            strCol = wrapInBackTick(convertedColName) + " " + dataTypeConvert(dataType);
                            if (strFields.Trim().Length == 0)
                            {
                                strFields = strCol;
                            }
                            else
                            {
                                strFields += "," + strCol;
                            }
                        }
                        if (!hasRunTitleField)
                        {
                            // Add runTitle field to tables without it
                            strCol = wrapInBackTick(runTitle) + " " + dataTypeConvert("SYSTEM.STRING");
                            strFields += "," + strCol;
                        }

                        sb.Append(strFields + ") ");
                        // Populate the table create queries dict and the dictionary of output table columns and their access datatype.
                        m_dictCreateTableQueries[convertSqliteTblNameToBiosumTblName(tblName)] = sb.ToString();
                        m_listDictFVSOutputTablesColumnsDefinitions[convertSqliteTblNameToBiosumTblName(tblName)] = listColDataTypes;

                        oDataMgr.SqlQueryReader(con, $@"SELECT distinct(c.RunTitle) FROM FVS_CASES c 
                                                        INNER JOIN {tblName} f ON c.CaseID=f.CaseID ");
                        List<string> lstTables = null;
                        if (oDataMgr.m_DataReader.HasRows)
                        {
                            while (oDataMgr.m_DataReader.Read())
                            {
                                string strRunTitle = Convert.ToString(oDataMgr.m_DataReader["RunTitle"]);
                                // dictionary key: runTitle                                 
                                if (! m_dictRunTitleTables.Keys.Contains(strRunTitle))
                                {
                                    lstTables = new List<string>();
                                }
                                else
                                {
                                    lstTables = m_dictRunTitleTables[strRunTitle];
                                }
                                lstTables.Add(tblName);
                                m_dictRunTitleTables[strRunTitle] = lstTables;
                            }

                        }
                        oDataMgr.m_DataReader.Close();
                    }
                }
                // Add the FVS_CutList if it's missing so that FVS_CUTLIST_PREPOST_SEQNUM_MATRIX is handled
                // correctly in succeeding steps
                foreach (var key in m_dictRunTitleTables.Keys)
                {
                    List<string> lstTables = m_dictRunTitleTables[key];
                    if (!lstTables.Contains("FVS_CutList") && !(key.IndexOf("BaseYr") > -1))
                    {
                        lstTables.Add("FVS_CutList");
                    }
                }
            }
        }

        /// <summary>Executes the supplied dictionary of queries on the table name used as the key, using the connection suppled.
        /// Used to insert values into a new Access DB.</summary>
        /// <param name="queryDict">A <string,string> dict where the key is a table and the value is a query to run.</param>
        /// <param name="accessConn">An OleDbConnection for the access DB to run against.</param>
        /// <param name="oAdo">An ado_data_access instance.</param>
        private void executeSQLListOnAccessConnection(Dictionary<string, string> queryDict, OleDbConnection accessConn, ado_data_access oAdo,
            string strRunTitle)
        {
            List<string> lstTables = new List<string>();
            if (m_dictRunTitleTables.Keys.Contains(strRunTitle))
            {
                lstTables = m_dictRunTitleTables[strRunTitle];
            }
            foreach (var tblName in queryDict.Keys)
            {
                if (lstTables.Contains(tblName))
                {
                    string strDataSource = accessConn.DataSource.ToUpper();
                    if (strDataSource.IndexOf("POTFIRE_BASEYR") > -1) // This is a baseyr database; Only create 2 tables
                    {
                        if (lstPotfireTables.Contains(tblName))
                        {
                            appendStringToDebugTextbox($@"Creating table: {tblName}");
                            oAdo.SqlNonQuery(accessConn, queryDict[tblName]);
                        }
                    }
                    else
                    {
                        appendStringToDebugTextbox($@"Creating table: {tblName}");
                        oAdo.SqlNonQuery(accessConn, queryDict[tblName]);
                    }
                }
            }
        }

        /// <summary>Translates "FVSOut.db" column naming conventions into BioSum ones, e.g. SPECIESFIA -> SPECIES.</summary>
        /// <param name="strToCheck">String to check. If it matches, map it. Otherwise return the same string.</param>
        private string translateColumn(string strToCheck)
        {
            var translatedStr = strToCheck.ToUpper();
            // Map SPECIESFIA to SPECIES. In the future, add other column mappings (e.g. stuff that's different in FVSOUT.db from the target Access Mdbs) here.
            if (sqliteToAccessColNames.Keys.Contains(translatedStr))
            {
                translatedStr = sqliteToAccessColNames[translatedStr];
                appendStringToDebugTextbox($@"Mapping column {strToCheck} to {translatedStr}");
            }
            return translatedStr;
        }

        /// <summary>Translates "FVSOut.db" table naming conventions into BioSum ones, e.g. Any table with "Summary" in the name -> FVS_Summary.</summary>
        /// <param name="tblName">String to check. If it matches, map it. Otherwise return the same string.</param>
        private string convertSqliteTblNameToBiosumTblName(string tblName)
        {
            var validTables = Tables.FVS.g_strFVSOutTablesArray;
            var validTablesList = new List<string>(validTables);
            var tablesSet = new HashSet<string>(validTablesList);
            var newTblName = tblName;
            if (!tablesSet.Contains(tblName))
            {
                //07-JUN-2022: There is also an fvs_summary2 table created that we do not want
                // Only publish fvs_summary as named
                //if (tblName.ToUpper().Contains("SUMMARY"))
                //{
                //    newTblName = "FVS_Summary";
                //}
            }
            return newTblName;
        }

        /// <summary>This function generates the list of Access DB file names we'll be populating.
        /// It returns a List of Lists of strings. Each member of this list contains 3 strings: fileName, strVariant, and runTitle, in that order.
        /// </summary>
        private List<List<string>> getRunTitleFilenames()
        {
            // List of three items; databasename, variant, runtitle.
            var fileNames = new List<List<string>>();
            Datasource oDatasource = new Datasource();
            InitializeDatasource(oDatasource);
            var strTempMDB = oDatasource.CreateMDBAndTableDataSourceLinks();
            string strPlotTable = oDatasource.getValidDataSourceTableName("PLOT");
            string strRxPackageTable = oDatasource.getValidDataSourceTableName("TREATMENT PACKAGES");
            var variantList = new List<string>();
            using (OleDbConnection con = new OleDbConnection(m_ado.getMDBConnString(strTempMDB, "", "")))
            {
                con.Open();
                this.m_ado.SqlQueryReader(con, Queries.FVS.GetFVSVariantRxPackageSQL(strPlotTable, strRxPackageTable));

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
                    // maintain variant list so we can add the base year files
                    if (!variantList.Contains(strVariant))
                    {
                        variantList.Add(strVariant);
                    }
                    var runTitle = $@"FVSOUT_{strVariant}_P{strPackage}-{strRx1}-{strRx2}-{strRx3}-{strRx4}";
                    var fileName = $@"FVSOUT_{strVariant}_P{strPackage}-{strRx1}-{strRx2}-{strRx3}-{strRx4}.MDB";
                    fileNames.Add(new List<string>() { fileName, strVariant, runTitle });
                }
                this.m_ado.m_OleDbDataReader.Close();
                // Add POTFIRE Base year mdbs
                foreach (var strVariant in variantList)
                {
                    var runTitle = $@"FVSOUT_{strVariant}_POTFIRE_BaseYr";
                    var fileName = $@"FVSOUT_{strVariant}_POTFIRE_BaseYr.MDB";
                    fileNames.Add(new List<string>() { fileName, strVariant, runTitle });
                }
            }
            File.Delete(strTempMDB);
            return fileNames;
        }
        /// <summary>Event handler for the Create MDBs button. Spins up the CreateMDBS_Main thread. </summary>
        private void createMdbsMain_Click(object sender, EventArgs e)
        {
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            frmMain.g_oDelegate.m_oThread = new System.Threading.Thread(new System.Threading.ThreadStart(CreateMDBs_Main));
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.m_oThread.Start();
        }
        /// <summary>
        /// Attempts to cancel the thread.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            CancelThread();
        }
        /// <summary>
        /// The thread cancel function. Aborts the thread and calls CleanupThread.
        /// </summary>
        private void CancelThread()
        {
            if (frmMain.g_oDelegate.m_oThread.IsAlive)
            {
                bool bAbort = frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel processing (Y/N)?");
                if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
                {
                    frmMain.g_oDelegate.StopThread();
                    CleanupThread();
                }
            }
        }
        /// <summary>
        /// Clean's up the current thread. Clears the create table and column definition dictionaries (so they can be repopulated on new runs).
        /// Enables the parent form.
        /// </summary>
        private void CleanupThread()
        {
            m_dictCreateTableQueries.Clear();
            m_listDictFVSOutputTablesColumnsDefinitions.Clear();
            this.ParentForm.Enabled = true;
        }
        /// <summary>
        /// Wraps the supplied string in ` characters. 
        /// This is necessary for certain Access column names, so we just do it for everything.
        /// </summary>
        /// <param name="str">The string to wrap in `.</param>
        /// <returns></returns>
        private string wrapInBackTick(string str)
        {
            return $@"`{str}`";
        }

        /// <summary>
        /// Resize function. Currently does nothing. I don't remember making this; perhaps we could resize the textbox?
        /// </summary>
        internal void uc_fvs_create_mdbs_Resize()
        {
            return;
        }
    }
}

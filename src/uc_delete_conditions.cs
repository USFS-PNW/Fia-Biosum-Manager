using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_delete_conditions : UserControl
    {
        public int m_DialogHt;
        public int m_DialogWd;
        private Datasource m_oDatasource;
        private string m_strPlotTable;
        private string m_strCondTable;
        private string m_strTreeTable;
        private string m_strSiteTreeTable;
        private string m_strPpsaTable;
        private string m_strPopEstUnitTable;
        private string m_strPopStratumTable;
        private string m_strPopEvalTable;
        private string m_strBiosumPopStratumAdjustmentFactorsTable;
        private string m_strTreeMacroPlotBreakPointDiaTable;

        private HashSet<string> setBiosumCondCNs = new HashSet<string>();
        private HashSet<string> setBiosumCondIds = new HashSet<string>();
        private HashSet<string> setBiosumPlotIds = new HashSet<string>();
        private HashSet<string> setPlotCNs = new HashSet<string>();
        private HashSet<string> setTreeCNs = new HashSet<string>();
        private string m_strCondCNs;
        Dictionary<string, HashSet<string>> m_dictIdentityColumnsToValues;
        Dictionary<string, Dictionary<string, int>> m_dictDeletedRowCountsByDatabaseAndTable =
            new Dictionary<string, Dictionary<string, int>>();

        //TODO: Help files
        private env m_oEnv;

        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
        private int m_intError;
        private string m_strStateCountyPlotSQL;
        private string m_strStateCountySQL;
        private string m_strCurrentProcess;
        private frmTherm m_frmTherm;
        private ado_data_access m_ado;
        private dao_data_access m_dao;
        private OleDbConnection m_connTempMDBFile;
        private string m_strTempMDBFileConn;
        private string m_strTempMDBFile;
        private string m_strSQL;
        private string m_strProjDir;
        private string m_strMessage = "";
        private string m_strDebugFile = "";

        public uc_delete_conditions()
        {
            ReferenceFormDialog = null;
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            Width = 700;
            m_DialogWd = Width + 10;
            m_DialogHt = groupBox1.Top + grpboxFilter.Top + grpboxFilter.Height + 100;

            rdoFilterByMenu.Enabled = false;
            rdoDeleteAllConds.Enabled = false;

            m_oEnv = new env();
            m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_delete_conditions_debug" +
                             String.Format("{0:yyyyMMddhhmm}", DateTime.Now) + ".txt";
        }

        private void InitializeDatasource()
        {
            m_strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();

            m_oDatasource = new Datasource();
            m_oDatasource.LoadTableColumnNamesAndDataTypes = false;
            m_oDatasource.LoadTableRecordCount = false;
            m_oDatasource.m_strDataSourceMDBFile = m_strProjDir + "\\db\\project.mdb";
            m_oDatasource.m_strDataSourceTableName = "datasource";
            m_oDatasource.m_strScenarioId = "";
            m_oDatasource.populate_datasource_array();

            m_strPlotTable = m_oDatasource.getValidDataSourceTableName("PLOT");
            m_strCondTable = m_oDatasource.getValidDataSourceTableName("CONDITION");
            m_strTreeTable = m_oDatasource.getValidDataSourceTableName("TREE");
        }

        private void rdoFilterByFile_Click(object sender, System.EventArgs e)
        {
            this.btnFilterFinish.Enabled = false;
            this.btnFilterNext.Enabled = false;
            this.txtFilterByFile.Enabled = true;
            this.btnFilterByFileBrowse.Enabled = true;
            if (!string.IsNullOrWhiteSpace(txtFilterByFile.Text))
            {
                this.btnFilterFinish.Enabled = true;
            }
        }

        private void rdoFilterByMenu_Click(object sender, System.EventArgs e)
        {
            this.btnFilterFinish.Enabled = false;
            this.btnFilterNext.Enabled = true;
            this.txtFilterByFile.Enabled = false;
            this.btnFilterByFileBrowse.Enabled = false;
        }

        private void rdoDeleteAllConds_Click(object sender, System.EventArgs e)
        {
            this.btnFilterFinish.Enabled = true;
            this.btnFilterNext.Enabled = true;
            this.txtFilterByFile.Enabled = false;
            this.btnFilterByFileBrowse.Enabled = false;
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            ((frmDialog) this.ParentForm).Close();
        }

        private void btnFilterFinish_Click(object sender, System.EventArgs e)
        {
            this.m_strStateCountyPlotSQL = "";
            this.m_strStateCountySQL = "";
            this.m_intError = 0;
            InitializeDatasource();
            if (m_intError == 0)
            {
                //No specific conditions to remove. Delete all of them.
                if (this.rdoDeleteAllConds.Checked)
                {
                    /*TODO: either set m_strCondCNs to all possible cond.cn values 
                     * or modify the SQL to skip the where filters 
                     * (caution: this approach deletes tables that don't have 
                     * plot/cond/tree/stand identifiers at all)
                     */
                    DeleteCondsFromBiosumProject_Start();
                }
                else if (this.rdoFilterByMenu.Checked)
                {
                    //TODO: set m_strCondCNs via the GUI
                    DeleteCondsFromBiosumProject_Start();
                }
                else if (this.rdoFilterByFile.Checked)
                {
                    if (System.IO.File.Exists(this.txtFilterByFile.Text.Trim()) == true)
                    {
                        this.m_strCondCNs =
                            this.CreateDelimitedStringList(this.txtFilterByFile.Text.Trim(), ",", ",", false);
                        if (this.m_intError == 0)
                        {
                            DeleteCondsFromBiosumProject_Start();
                        }
                    }
                    else
                    {
                        MessageBox.Show("!!" + this.txtFilterByFile.Text.Trim() + " could not be found!!",
                            "Delete Conditions", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        //throw new NotImplementedException("Delete conds, with or without filters, using m_strCondCNs built with either a text file or GUI menu. use it to collect biosum_cond_ids, plot.cn, tree.cn");
        private void DeleteCondsFromBiosumProject_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            this.m_strCurrentProcess = "DeleteConditions";
            this.StartTherm("2", "Delete Biosum Condition Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(DeleteCondsFromBiosumProject_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void DeleteCondsFromBiosumProject_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            this.m_intError = 0;
            try
            {
                this.m_ado = new ado_data_access();
                this.m_dao = new dao_data_access();

                //progress bar 1: single process
                this.SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "Visible", true);
                //progress bar 2: overall progress
                this.SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                this.SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                this.SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                this.SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "Visible", true);


                this.m_strTempMDBFile = m_oDatasource.CreateMDBAndTableDataSourceLinks();
                //get an ado connection string for the temp mdb file
                this.m_strTempMDBFileConn = this.m_ado.getMDBConnString(this.m_strTempMDBFile, "", "");
                //create a new connection to the temp MDB file
                this.m_connTempMDBFile = new System.Data.OleDb.OleDbConnection();
                //open the connection to the temp mdb file 
                this.m_ado.OpenConnection(this.m_strTempMDBFileConn, ref this.m_connTempMDBFile);

                UpdateProgressBar2(10);
                SetConditionLookupAndIDStrings();

                //Need to access plot.fvs_variants before deleting records
                string[] strVariants = m_oDatasource.getVariants();

                //ProjectRoot\gis Section
                UpdateProgressBar2(20);
                // Travel times is now in gis_travel_times.accdb
                ConnectToDatabasesInPathAndExecuteDeletes(Directory
                    .GetFiles(m_strProjDir + "\\gis\\db\\", "*.*", SearchOption.AllDirectories)
                    .Where(s => s.ToLower().EndsWith(".mdb") || s.ToLower().EndsWith(".accdb")).ToArray());


                //ProjectRoot\processor Section
                UpdateProgressBar2(30);
                foreach (string subfolder in Directory.GetDirectories(m_strProjDir + "\\processor\\"))
                {
                    if (Path.GetFileName(subfolder) == "db")
                        ConnectToDatabasesInPathAndExecuteDeletes(Directory.GetFiles(subfolder, "*.mdb"));
                    else 
                        if (Directory.Exists(subfolder + "\\db\\"))
                            ConnectToDatabasesInPathAndExecuteDeletes(Directory.GetFiles(subfolder + "\\db\\", "*.mdb"));
                }

                //FVS Section
                UpdateProgressBar2(40);
                if (strVariants != null && strVariants.Length > 0)
                {
                    string strFvsDataDir = m_strProjDir + "\\fvs\\data\\";
                    int num_variants_in_fvs_data = 0;
                    foreach (string variant in strVariants)
                        if (Directory.Exists(strFvsDataDir + variant + "\\"))
                            num_variants_in_fvs_data++;
                    int progressbar2_value = 40;
                    foreach (string variant in strVariants)
                    {
                        //Collect pathfiles of databases to delete from in FVS Data directory
                        string strVariantPath = strFvsDataDir + variant + "\\";
                        if (Directory.Exists(strVariantPath))
                        {
                            progressbar2_value += (30 / (num_variants_in_fvs_data + 1));
                            UpdateProgressBar2(progressbar2_value);
                            ConnectToDatabasesInPathAndExecuteDeletes(
                                Directory.GetFiles(strVariantPath, "*.*", SearchOption.AllDirectories)
                                    .Where(s => s.ToLower().EndsWith(".mdb") || s.ToLower().EndsWith(".accdb"))
                                    .ToArray(),
                                strTableExceptions: new string[] {"FVS_GroupAddFilesAndKeywords"});
                        }
                    }
                }

                //PREPOST Databases
                UpdateProgressBar2(70);
                ConnectToDatabasesInPathAndExecuteDeletes(
                    Directory.GetFiles(m_strProjDir + "\\fvs\\db\\", "PREPOST*.accdb"),
                    strTableExceptions: new string[] {"FVS_TREE"});

                //OPTIMIZER weighted variables
                UpdateProgressBar2(80);
                ConnectToDatabasesInPathAndExecuteDeletes(
                    Directory.GetFiles(m_strProjDir + "\\optimizer\\db\\", "prepost*.accdb"));

                //ProjectRoot\db Section
                UpdateProgressBar2(90);
                ConnectToDatabasesInPathAndExecuteDeletes(Directory
                    .GetFiles(m_strProjDir + "\\db\\", "*.*", SearchOption.AllDirectories)
                    .Where(s => s.ToLower().EndsWith(".mdb") || s.ToLower().EndsWith(".accdb")).ToArray());

                //Done
                UpdateProgressBar2(100);

                if (!Checked(chkDeletesDisabled))
                {
                    MessageBox.Show(
                        String.Format("Successfully deleted data associated with {0} Conditions!",
                            ((HashSet<string>)m_dictIdentityColumnsToValues["biosum_cond_id"]).Count),
                        "Delete Conditions Results");
                }

                //Cleanup section, assuming no exceptions were thrown
                this.m_connTempMDBFile.Close();
                while (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    System.Threading.Thread.Sleep(1000);
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                if (m_dao != null)
                {
                    this.m_dao.m_DaoWorkspace.Close();
                    this.m_dao.m_DaoWorkspace = null;
                    this.m_dao = null;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                    true);
                
                DeleteCondsFromBiosumProject_Finish();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (this.m_connTempMDBFile != null)
                {
                    if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                    {
                        m_ado.CloseConnection(m_connTempMDBFile);
                    }
                    m_connTempMDBFile = null;
                }
                if (m_ado != null)
                {
                    if (m_ado.m_DataSet != null)
                    {
                        this.m_ado.m_DataSet.Clear();
                        this.m_ado.m_DataSet.Dispose();
                    }
                    this.m_ado = null;
                }
                this.ThreadCleanUp();
                this.CleanupThread();
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_delete_conditions:DeleteCondsFromBiosumProject_Process  \n" +
                                "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                this.m_intError = -1;
            }
            finally
            {
                string strLogFileName = String.Concat(frmMain.g_oFrmMain.getProjectDirectory(),
                    "\\db\\biosum_deleted_records_",
                    String.Format("{0:yyyyMMddhhmm}", DateTime.Now),
                    Checked(chkDeletesDisabled) ? "_no_deletes" : "",
                    ".txt");
                if (Checked(chkCreateLog) || Checked(chkDeletesDisabled))
                {
                    CreateLogFile(strLogFileName);
                    if (Checked(chkDeletesDisabled))
                    {
                        MessageBox.Show("Record count complete! The results were recorded in " + strLogFileName,
                        "Delete Conditions Results");
                    }
                }
                m_dictDeletedRowCountsByDatabaseAndTable = new Dictionary<string, Dictionary<string, int>>();
                m_dictIdentityColumnsToValues = new Dictionary<string, HashSet<string>>();
            }

            if (this.m_connTempMDBFile != null)
            {
                if (m_connTempMDBFile.State != System.Data.ConnectionState.Closed)
                {
                    m_ado.CloseConnection(m_connTempMDBFile);
                }
                m_connTempMDBFile = null;
            }
            if (m_ado != null)
            {
                if (m_ado.m_DataSet != null)
                {
                    this.m_ado.m_DataSet.Clear();
                    this.m_ado.m_DataSet.Dispose();
                }
                this.m_ado = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                true);
            if (this.m_frmTherm != null)
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "Visible", false);


            DeleteCondsFromBiosumProject_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }


        private void UpdateProgressBar1(string label, int value)
        {
            this.SetLabelValue(m_frmTherm.lblMsg, "Text", label);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) this.m_frmTherm.progressBar1,
                "Value", value);
        }

        private void UpdateProgressBar2(int value)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) this.m_frmTherm.progressBar2,
                "Value", value);
        }

        private void ConnectToDatabasesInPathAndExecuteDeletes(string[] strDatabaseNames, string[] strTargetTables = null, string[] strTableExceptions = null)
        {
            int counter = 0;

            foreach (string db in strDatabaseNames)
            {
                var message = (Checked(chkDeletesDisabled) ? "Checking " : "Deleting from ") + Path.GetFileName(db);
                UpdateProgressBar1(message, (int) Math.Floor(100 * ((double) (counter + 1) / (strDatabaseNames.Length + 1))));
                counter += 1;
                ExecuteDeleteOnTables(db, tables: strTargetTables, exceptions: strTableExceptions);
            }
            UpdateProgressBar1("", 0);
        }

        private void SetConditionLookupAndIDStrings()
        {
            string[] condCNs = m_strCondCNs.Split(',');

            setBiosumCondCNs = new HashSet<string>();
            setBiosumCondIds = new HashSet<string>();
            setBiosumPlotIds = new HashSet<string>();
            setPlotCNs = new HashSet<string>();
            setTreeCNs = new HashSet<string>();

            if (m_intError == 0 && !GetBooleanValue(m_frmTherm, "AbortProcess"))
            {
                SetThermValue(m_frmTherm.progressBar1, "Value", 20);

                //Create table of condition CNs to join Plot/Cond/Tree
                m_ado.SqlNonQuery(m_connTempMDBFile, "CREATE TABLE conds_to_delete (cn TEXT(34))");
				m_ado.AddPrimaryKey(m_connTempMDBFile,"conds_to_delete","conds_to_delete_pk","cn");
                foreach (string condCN in condCNs)
                {
                    m_ado.SqlNonQuery(m_connTempMDBFile, "INSERT INTO conds_to_delete (cn) VALUES (" + condCN + ");");
                }

                m_ado.CreateDataSet(m_connTempMDBFile,
                    string.Format(
                        "SELECT c.cn, c.biosum_cond_id, t.cn " +
                        "FROM (({0} c INNER JOIN conds_to_delete ON c.cn=conds_to_delete.cn) " +
                        "INNER JOIN {1} p ON c.biosum_plot_id=p.biosum_plot_id) " +
                        "LEFT JOIN {2} t on c.biosum_cond_id=t.biosum_cond_id;",
                        m_strCondTable, m_strPlotTable, m_strTreeTable), "identifiers");
                foreach (DataRow row in m_ado.m_DataSet.Tables["identifiers"].Rows)
                {
                    setBiosumCondCNs.Add(String.Format("'{0}'", row[0]));
                    setBiosumCondIds.Add(String.Format("'{0}'", row[1]));
                    setTreeCNs.Add(String.Format("'{0}'", row[2]));
                }
                m_ado.CreateDataSet(m_connTempMDBFile,
                    String.Format(
                        "SELECT allconds.biosum_plot_id, allconds.cn " + //, allconds.cntConds, someconds.cntConds " +
                        "FROM (SELECT p.biosum_plot_id, p.cn, count(*) as cntConds FROM ({0} c INNER JOIN conds_to_delete ON c.cn=conds_to_delete.cn) INNER JOIN {1} p ON c.biosum_plot_id = p.biosum_plot_id WHERE c.cn IN ({2}) GROUP BY p.biosum_plot_id, p.cn) someconds " +
                        "RIGHT JOIN (SELECT p.biosum_plot_id, p.cn, count(*) as cntConds FROM {0} c INNER JOIN {1} p ON c.biosum_plot_id = p.biosum_plot_id GROUP BY p.biosum_plot_id, p.cn) allconds " +
                        "ON allconds.biosum_plot_id=someconds.biosum_plot_id WHERE allconds.cntConds=someconds.cntConds",
                        m_strCondTable, m_strPlotTable, m_strCondCNs), "plots_with_all_conds_deleted");
                foreach (DataRow row in m_ado.m_DataSet.Tables["plots_with_all_conds_deleted"].Rows)
                {
                    setBiosumPlotIds.Add(String.Format("'{0}'", row[0]));
                    setPlotCNs.Add(String.Format("'{0}'", row[1]));
                }
                m_intError = m_ado.m_intError;
            }

            m_dictIdentityColumnsToValues = new Dictionary<string, HashSet<string>>
            {
                {"biosum_cond_id", setBiosumCondIds},
                {"StandID", setBiosumCondIds},
                {"Stand_ID", setBiosumCondIds},
                {"biosum_plot_id", setBiosumPlotIds},
                {"cnd_cn", setBiosumCondCNs},
                {"plt_cn", setPlotCNs},
                {"tre_cn", setTreeCNs},
            };
        }


        private void ExecuteDeleteOnTables(string strDbPathFile, string[] tables = null, string[] exceptions = null)
        {
            using (var conn = new OleDbConnection(m_ado.getMDBConnString(strDbPathFile, "", "")))
            {
                conn.Open();

                string[] strTables = tables;
                if (tables == null || tables.Length == 0)
                {
                    strTables = m_ado.getTableNamesOfSpecificTypes(conn).Where(s => !(s.Contains("~") || s.Contains(" "))).ToArray();
                    //In case none of the tables are valid
                    if (strTables.Length == 1 && strTables[0] == "")
                        return;
                }

                string column;
                foreach (string table in strTables)
                {
                    if (exceptions != null && exceptions.Contains(table))
                    {
                        continue;
                    }

                    column = null;
                    foreach (string col in m_dictIdentityColumnsToValues.Keys)
                    {
                        if (m_ado.ColumnExist(conn, table, col))
                        {
                            column = col;
                            break;
                        }
                    }

                    if (!String.IsNullOrEmpty(column) && m_dictIdentityColumnsToValues[column].Count > 0)
                    {
                        var strTempIndex = column + "_delete_idx";
                        if (!m_dao.IndexExists(strDbPathFile, table, strTempIndex))
                        {
                            m_ado.AddIndex(conn, table, strTempIndex, column);
                        }

                        if (frmMain.g_bDebug)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, Checked(chkDeletesDisabled)
                                ? "\r\nCounting records to delete from " + strDbPathFile + " " + table + " using " + column + "\r\n"
                                : "\r\nDeleting from " + strDbPathFile + " " + table + " using " + column + "\r\n");
                        }

                        int deletedRecords = BuildAndExecuteDeleteSQLStmts(conn, table, column);
                        if (deletedRecords > 0) AddDeletedCountToDictionary(strDbPathFile, table, deletedRecords);
                        m_ado.SqlNonQuery(conn, String.Format("DROP INDEX {0} ON {1}", strTempIndex, table));
                    }
                }
            }
            if (Checked(chkCompactMDB)) m_dao.CompactMDB(strDbPathFile);
        }

        private void AddDeletedCountToDictionary(string strDbPathFile, string table, int deletedRecords)
        {
            if (!m_dictDeletedRowCountsByDatabaseAndTable.Keys.Contains(strDbPathFile))
            {
                m_dictDeletedRowCountsByDatabaseAndTable.Add(strDbPathFile,
                    new Dictionary<string, int>());
            }
            m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Add(table, deletedRecords);
        }

        private int BuildAndExecuteDeleteSQLStmts(OleDbConnection oConn, string table, string column)
        {
            int matchedRecordCount = 0;
            string columnKey = column;
            string[] edgeCaseTables = {"fcs_biosum_volumes_input"};

            if (edgeCaseTables.Contains(table))
            {
                //FCS schema uses cnd_cn/plt_cn, but we set these values to biosum_cond_id/biosum_plot_id in FVSOUT stage
                if (table.ToLower() == "fcs_biosum_volumes_input" && column.ToLower() == "cnd_cn")
                    columnKey = "biosum_cond_id";
            }

            var batches = m_dictIdentityColumnsToValues[columnKey]
                .Select((id, i) => new {id, i})
                .GroupBy(p => p.i / 100, p => p.id);
            foreach (var batch in batches)
            {
                var values = string.Join(",", batch);
                matchedRecordCount += (int) m_ado.getSingleDoubleValueFromSQLQuery(oConn,
                    String.Format("SELECT COUNT(*) FROM {0} WHERE {1} IN ({2});", table, column, values), table);
                var deleteSQL = String.Format("DELETE FROM {0} WHERE {1} IN ({2});", table, column, values);
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                {
                    var message = Checked(chkDeletesDisabled)
                        ? "Generated but not executed: " + deleteSQL + "\r\n"
                        : "Attempting to execute: " + deleteSQL + "\r\n";
                    frmMain.g_oUtils.WriteText(m_strDebugFile, message);
                }

                if (!Checked(chkDeletesDisabled))
                    m_ado.SqlNonQuery(oConn, deleteSQL);
            }

            return matchedRecordCount;
        }

        private void CreateLogFile(string strLogFileName)
        {
            if (m_dictDeletedRowCountsByDatabaseAndTable.Keys.Count > 0)
            {
                var strMessage = Checked(chkDeletesDisabled)
                    ? "records would be deleted from"
                    : "records were deleted from";
                foreach (var strDbPathFile in m_dictDeletedRowCountsByDatabaseAndTable.Keys)
                {
                    if (m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Keys.Count > 0)
                    {
                        frmMain.g_oUtils.WriteText(strLogFileName, String.Concat(strDbPathFile, Environment.NewLine));
                        foreach (var strTable in m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Keys)
                        {
                            frmMain.g_oUtils.WriteText(strLogFileName,
                                String.Format("  {0} {1} {2}{3}",
                                    m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile][strTable], strMessage,
                                    strTable, Environment.NewLine));
                        }
                    }
                }
            }
            else
            {
                frmMain.g_oUtils.WriteText(strLogFileName, "No records were deleted." + Environment.NewLine);
            }
        }

        private void CleanupThread()
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                true);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                true);
        }

        private void ThreadCleanUp()
        {
            try
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) ReferenceFormDialog, "Enabled",
                    true);

                if (this.m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form) m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form) m_frmTherm, "Dispose");

                    this.m_frmTherm = null;
                }
            }
            catch
            {
            }
        }

        private void DeleteCondsFromBiosumProject_Finish()
        {
            if (this.m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                this.m_frmTherm = null;
            }
            this.m_strCurrentProcess = "";
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            ((frmDialog) ParentForm).MinimizeMainForm = false;
        }


        private void StartTherm(string p_strNumberOfTherms, string p_strTitle)
        {
            this.m_frmTherm = new frmTherm((frmDialog) this.ParentForm, p_strTitle);

            this.m_frmTherm.Text = p_strTitle;
            this.m_frmTherm.lblMsg.Text = "";
            this.m_frmTherm.lblMsg2.Text = "";
            this.m_frmTherm.Visible = false;
            this.m_frmTherm.btnCancel.Visible = false; 
            this.m_frmTherm.btnCancel.Enabled = false;
            this.m_frmTherm.lblMsg.Visible = true;
            this.m_frmTherm.progressBar1.Minimum = 0;
            this.m_frmTherm.progressBar1.Visible = true;
            this.m_frmTherm.progressBar1.Maximum = 10;

            if (p_strNumberOfTherms == "2")
            {
                this.m_frmTherm.progressBar2.Size = this.m_frmTherm.progressBar1.Size;
                this.m_frmTherm.progressBar2.Left = this.m_frmTherm.progressBar1.Left;
                this.m_frmTherm.progressBar2.Top =
                    Convert.ToInt32(this.m_frmTherm.progressBar1.Top + (this.m_frmTherm.progressBar1.Height * 3));
                this.m_frmTherm.lblMsg2.Top =
                    this.m_frmTherm.progressBar2.Top + this.m_frmTherm.progressBar2.Height + 5;
                this.m_frmTherm.Height = this.m_frmTherm.lblMsg2.Top + this.m_frmTherm.lblMsg2.Height +
                                         this.m_frmTherm.btnCancel.Height + 50;
                this.m_frmTherm.btnCancel.Top =
                    this.m_frmTherm.ClientSize.Height - this.m_frmTherm.btnCancel.Height - 5;
                this.m_frmTherm.lblMsg2.Show();
                this.m_frmTherm.progressBar2.Visible = true;
            }
            this.m_frmTherm.AbortProcess = false;
            this.m_frmTherm.Refresh();
            this.m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            ((frmDialog)this.ParentForm).Enabled=false;
            this.m_frmTherm.Visible = true;
        }

        public void StopThread()
        {
            string strMsg = "";

            frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel adding plot data?");

            if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "AbortProcess",
                    true);
//                this.CancelThreadCleanup();
//                this.ThreadCleanUp();
                throw new NotImplementedException();
            }
        }

        private bool Checked(System.Windows.Forms.RadioButton p_rdoButton)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.RadioButton)p_rdoButton, "Checked", false);
        }
        private bool Checked(System.Windows.Forms.CheckBox p_chkBox)
        {
            return (bool)frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.CheckBox)p_chkBox, "Checked", false);
        }

        private void SetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName, int p_intValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) p_oPb, p_strPropertyName,
                (int) p_intValue);
        }

        private int GetThermValue(System.Windows.Forms.ProgressBar p_oPb, string p_strPropertyName)
        {
            return (int) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control) p_oPb,
                p_strPropertyName, false);
        }

        private bool GetBooleanValue(System.Windows.Forms.Control p_oControl, string p_strPropertyName)
        {
            return (bool) frmMain.g_oDelegate.GetControlPropertyValue((System.Windows.Forms.Control) p_oControl,
                p_strPropertyName, false);
        }


        private void SetLabelValue(System.Windows.Forms.Label p_oLabel, string p_strPropertyName, string p_strValue)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Label) p_oLabel, p_strPropertyName,
                p_strValue);
        }

        public frmDialog ReferenceFormDialog { set; get; }


        private void btnFilterByFileBrowse_click(object sender, EventArgs e)
        {
            var OpenFileDialog1 = new OpenFileDialog();
            OpenFileDialog1.Title = "Text File With PLOT_CN data";
            OpenFileDialog1.Filter = "Text File (*.TXT;*.DAT) |*.txt;*.dat";
            var result = OpenFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (OpenFileDialog1.FileName.Trim().Length > 0)
                    txtFilterByFile.Text = OpenFileDialog1.FileName.Trim();
                OpenFileDialog1 = null;
            }
        }

        /// <summary>
        /// create a delimited string list from a text file
        /// that has a single column of data with multiple rows
        /// </summary>
        /// <param name="p_strTxtFile">text file containing the column of data</param>
        /// <param name="p_strTxtFileDelimiter">specified character between list items</param>
        /// <param name="p_strListDelimiter">specified character between list items</param>
        /// <param name="p_bNumericDataType">specifies if the column data to retrieve in the text file is numeric</param>
        /// <returns></returns>
        private string CreateDelimitedStringList(string p_strTxtFile, string p_strTxtFileDelimiter,
            string p_strListDelimiter, bool p_bNumericDataType)
        {
            //The DataSet to Return
            //DataSet result = new DataSet();
            this.m_intError = 0;
            string strList = "";
            string str = "";
            try
            {
                //Open the file in a stream reader.
                System.IO.StreamReader s = new System.IO.StreamReader(p_strTxtFile);
                //Read the rest of the data in the file.        
                string AllData = s.ReadToEnd();

                //Split off each row at the Carriage Return/Line Feed
                //Default line ending in most <A class=iAs style="FONT-WEIGHT: normal; FONT-SIZE: 100%; PADDING-BOTTOM: 1px; COLOR: darkgreen; BORDER-BOTTOM: darkgreen 0.07em solid; BACKGROUND-COLOR: transparent; TEXT-DECORATION: underline" href="#" target=_blank itxtdid="2592535">windows</A> exports.  
                string[] rows = AllData.Split("\r\n".ToCharArray());

                //Now add each row to the DataSet        
                foreach (string r in rows)
                {
                    //Split the row at the delimiter.
                    string[] items = r.Split(p_strTxtFileDelimiter.ToCharArray());
                    str = items[0].Trim(); //plot_cn in first column
                    str = str.Replace("\"", ""); //remove any quotations
                    if (str.Trim().Length > 0)
                    {
                        if (strList.Trim().Length == 0)
                        {
                            if (p_bNumericDataType == true)
                            {
                                strList = str.Trim();
                            }
                            else
                            {
                                strList = "'" + str.Trim() + "'";
                            }
                        }
                        else
                        {
                            if (p_bNumericDataType == true)
                            {
                                strList = strList + p_strListDelimiter.Trim() + str.Trim();
                            }
                            else
                            {
                                strList = strList + p_strListDelimiter.Trim() + "'" + str.Trim() + "'";
                            }
                        }
                    }
                }
            }
            catch (Exception caught)
            {
                this.m_intError = -1;
                MessageBox.Show("!!Error: CreateDelimitedStringList() Routine Error Msg:" + caught.Message);
            }
            return strList;
        }

        private void txtFilterByFile_TextChanged(object sender, EventArgs e)
        {
            this.btnFilterFinish.Enabled = true;
        }

        private void btnFilterHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "DELETE_CONDITIONS" });
        }

        private void chkDeletesDisabled_CheckedChanged(object sender, EventArgs e)
        {
            // Don't allow chkCompactMDB to be checked if chkDeletesDisabled is checked
            if (Checked(chkDeletesDisabled))
            {
                chkCompactMDB.Checked = false;
            }
        }

        private void chkCompactMDB_CheckedChanged(object sender, EventArgs e)
        {
            // Don't allow chkCompactMDB to be checked if chkDeletesDisabled is checked
            if (Checked(chkCompactMDB))
            {
                chkDeletesDisabled.Checked = false;
            }
        }
    }
}

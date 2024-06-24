using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class frmScanAndSynchronizeProjectRootFolderTool : Form
    {
        const int COLUMN_NULL = 0;
        const int COLUMN_DATASOURCE = 1;
        const int COLUMN_SCENARIO = 2;
        const int COLUMN_TABLETYPE = 3;
        const int COLUMN_PATH = 4;
        const int COLUMN_PATHFOUND = 5;
        const int COLUMN_SYNCD = 6;

        private string m_strRandomPathAndFile = "";
        private ODBCMgr m_odbcMgr  = new ODBCMgr();

        bool m_bSyncd = false;

        public frmScanAndSynchronizeProjectRootFolderTool()
        {
            InitializeComponent();
            this.lblCurrentProjectRootFolder.Text = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
            CreateTableLinks();
            loadvalues();
            
        }
        public void loadvalues()
        {
            m_bSyncd = false;
            ado_data_access oAdo = new ado_data_access();
           
            int x=0;
            int intPathNF = 0;
            int intRootNF = 0;
           

            this.lvDatasources.Clear();

            ado_data_access p_ado = new ado_data_access();
            this.lvDatasources.Columns.Add(" ", 2, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("DataSource", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("Scenario", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("TableType", 50, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("Path", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("PathFound", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("Synchronized", 100, HorizontalAlignment.Left);

            oAdo.OpenConnection(oAdo.getMDBConnString(m_strRandomPathAndFile,"",""));



            oAdo.m_strSQL =
            "SELECT 'Project' AS DataSource, 'NA' AS Scenario,table_type AS TableType,path FROM project_datasource ";

            if (oAdo.TableExist(oAdo.m_OleDbConnection, "optimizer_scenario"))
                oAdo.m_strSQL = oAdo.m_strSQL +
                    "UNION " +
                    "SELECT 'TreatmentOptimizer' AS DataSource, Scenario_Id AS Scenario,'NA' AS TableType,path FROM optimizer_scenario ";
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "optimizer_scenario_datasource"))
                oAdo.m_strSQL = oAdo.m_strSQL +
                    "UNION " +
                    "SELECT 'TreatmentOptimizer' AS DataSource, Scenario_Id AS Scenario,table_type AS TableType, path FROM optimizer_scenario_datasource ";
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "processor_scenario"))
                oAdo.m_strSQL = oAdo.m_strSQL +
                    "UNION " +
                    "SELECT 'Processor' AS DataSource, Scenario_Id AS Scenario,'NA' AS TableType,path FROM processor_scenario ";
            if (oAdo.TableExist(oAdo.m_OleDbConnection, "processor_scenario_datasource"))
                oAdo.m_strSQL = oAdo.m_strSQL +
                    "UNION " +
                    "SELECT 'Processor' AS DataSource, Scenario_Id AS Scenario,table_type AS TableType, path FROM processor_scenario_datasource ";

            oAdo.SqlQueryReader(oAdo.m_OleDbConnection, oAdo.m_strSQL);

            if (oAdo.m_OleDbDataReader.HasRows)
            {
                while (oAdo.m_OleDbDataReader.Read())
                {
                    // Don't add appData data sources to the grid
                    if (oAdo.m_OleDbDataReader["Path"].ToString().IndexOf("@@appdata@@") == -1) 
                    {
                        System.Windows.Forms.ListViewItem entryListItem =
                                this.lvDatasources.Items.Add(" ");
                        entryListItem.UseItemStyleForSubItems = false;
                        this.lvDatasources.Items[x].SubItems.Add(oAdo.m_OleDbDataReader["DataSource"].ToString());
                        this.lvDatasources.Items[x].SubItems.Add(oAdo.m_OleDbDataReader["Scenario"].ToString());
                        this.lvDatasources.Items[x].SubItems.Add(oAdo.m_OleDbDataReader["TableType"].ToString());
                        this.lvDatasources.Items[x].SubItems.Add(oAdo.m_OleDbDataReader["Path"].ToString());
                        if (System.IO.Directory.Exists(oAdo.m_OleDbDataReader["Path"].ToString().Trim()))
                        {
                            ListViewItem.ListViewSubItem FileStatusSubItem =
                                   entryListItem.SubItems.Add("Yes");
                            FileStatusSubItem.ForeColor = System.Drawing.Color.White;
                            FileStatusSubItem.BackColor = System.Drawing.Color.Green;
                        }
                        else
                        {
                            ListViewItem.ListViewSubItem FileStatusSubItem =
                                    entryListItem.SubItems.Add("No");
                            FileStatusSubItem.ForeColor = System.Drawing.Color.White;
                            FileStatusSubItem.BackColor = System.Drawing.Color.Red;
                            intPathNF++;
                        }
                        if (oAdo.m_OleDbDataReader["Path"].ToString().ToUpper().Contains(lblCurrentProjectRootFolder.Text.Trim().ToUpper()))
                        {
                            ListViewItem.ListViewSubItem SyncStatusSubItem =
                                  entryListItem.SubItems.Add("Yes");
                            SyncStatusSubItem.ForeColor = System.Drawing.Color.White;
                            SyncStatusSubItem.BackColor = System.Drawing.Color.Green;
                        }
                        else
                        {
                            ListViewItem.ListViewSubItem SyncStatusSubItem =
                                  entryListItem.SubItems.Add("No");
                            SyncStatusSubItem.ForeColor = System.Drawing.Color.White;
                            SyncStatusSubItem.BackColor = System.Drawing.Color.Red;
                            intRootNF++;
                        }
                        x++;
                    }
                }
                lblFolderPaths.Text = intPathNF.ToString().Trim();
                lblProjectRootFolderNotFound.Text = intRootNF.ToString().Trim();
                oAdo.m_OleDbDataReader.Close();
            }
            if (intRootNF > 0) btnAnalyze.Enabled = true; else btnAnalyze.Enabled = false;
            oAdo.m_OleDbConnection.Close();
            oAdo = null;


        }
        private void CreateTableLinks()
        {
            m_strRandomPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");

            dao_data_access oDao = new dao_data_access();

            oDao.CreateMDB(m_strRandomPathAndFile);
            //
            //PROJECT DATA SOURCE
            //
            string strFullPath = 
                this.lblCurrentProjectRootFolder.Text.Trim() + "\\db\\" + frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectFile;


            oDao.CreateTableLink(
                m_strRandomPathAndFile, "project_datasource", strFullPath, "datasource");
            //
            //CORE ANALYSIS DATA SOURCE
            //
            strFullPath =
                this.lblCurrentProjectRootFolder.Text.Trim() + "\\" +
                Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;

             if (System.IO.File.Exists(strFullPath))
             {
                // Create ODBC entry for the  file scenario_optimizer_rule_definitions.db
                if (m_odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName))
                {
                    m_odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName);
                }
                m_odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, strFullPath);

                oDao.CreateSQLiteTableLink(m_strRandomPathAndFile, "scenario", "optimizer_scenario",
                    ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, strFullPath, true);
                oDao.CreateSQLiteTableLink(m_strRandomPathAndFile, "scenario_datasource", "optimizer_scenario_datasource",
                    ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, strFullPath, true);
             }
             //
            //PROCESSOR SCENARIO DATA SOURCE
            //
            strFullPath = this.lblCurrentProjectRootFolder.Text.Trim() + 
                "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            if (System.IO.File.Exists(strFullPath))
            {
                // Create ODBC entry for the  file processor_optimizer_rule_definitions.db
                if (m_odbcMgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    m_odbcMgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                m_odbcMgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strFullPath);

                oDao.CreateSQLiteTableLink(m_strRandomPathAndFile, "scenario", "processor_scenario",
                    ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strFullPath, true);
                oDao.CreateSQLiteTableLink(m_strRandomPathAndFile, "scenario_datasource", "processor_scenario_datasource",
                    ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, strFullPath, true);
            }

            oDao.m_DaoWorkspace.Close();
            oDao = null;

        }

        private void frmScanAndSynchronizeProjectRootFolderTool_Resize(object sender, EventArgs e)
        {
            ResizeForm();
        }
        private void ResizeForm()
        {
            grpStatus.Left = this.ClientSize.Width - (int)(grpStatus.Width + 10);
            lvDatasources.Width = grpStatus.Left - lvDatasources.Left - 5;
            lvDatasources.Height = this.ClientSize.Height - lvDatasources.Top - 5;
        }

        private void toolBar1_ButtonClick(object sender, ToolBarButtonClickEventArgs e)
        {
            switch (e.Button.Text.Trim().ToUpper())
            {
                case "SAVE":
                    Save();
                    break;
                case "REFRESH":
                    if (m_bSyncd)
                    {
                        DialogResult result = MessageBox.Show("A refresh reloads the values that are in the database table and any changes that have not been saved will be lost. Do you wish to refresh? (Y/N)", "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes) loadvalues();
                    }
                    break;
                case "SYNC":
                    Sync();
                    break;
                case "HELP":
                    break;
                case "CLOSE":
                    if (m_bSyncd)
                    {
                        DialogResult result = MessageBox.Show("Do you wish to close form without saving changes? (Y/N)", "FIA Biosum", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes) this.Close();
                    }
                    else this.Close();
                    break;
            }
        }
        private void Save()
        {
            string strPath = "";
            string strDatasource = "";
            string strTableType = "";
            string strScenario = "";
            ado_data_access oAdo = new ado_data_access();

            oAdo.OpenConnection(oAdo.getMDBConnString(m_strRandomPathAndFile, "", ""));

            // Create primary keys so update command will work on SQLite tables
            oAdo.AddPrimaryKey(oAdo.m_OleDbConnection, "processor_scenario_datasource", "scenario_datasource_pk", "scenario_id, table_type");
            oAdo.AddPrimaryKey(oAdo.m_OleDbConnection, "optimizer_scenario_datasource", "scenario_datasource_pk", "scenario_id, table_type");

            for (int x = 0; x <= lvDatasources.Items.Count - 1; x++)
            {
                strDatasource = lvDatasources.Items[x].SubItems[COLUMN_DATASOURCE].Text.Trim();
                strTableType = lvDatasources.Items[x].SubItems[COLUMN_TABLETYPE].Text.Trim();
                strPath = lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim();
                strScenario = lvDatasources.Items[x].SubItems[COLUMN_SCENARIO].Text.Trim();
                oAdo.m_strSQL = "";
                if (strDatasource == "Project")
                {
                    oAdo.m_strSQL = "UPDATE project_datasource SET path = '" + strPath + "' WHERE TRIM(Table_Type) = '" + strTableType + "'";
                }
                else if (strDatasource == "TreatmentOptimizer" && strTableType == "NA")
                {
                    oAdo.m_strSQL = "UPDATE optimizer_scenario SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "'";
                }
                else if (strDatasource == "TreatmentOptimizer")
                {
                    oAdo.m_strSQL = "UPDATE optimizer_scenario_datasource SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "' AND TRIM(Table_Type) = '" + strTableType + "'";
                }
                else if (strDatasource == "Processor" && strTableType=="NA")
                {
                    oAdo.m_strSQL = "UPDATE processor_scenario SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "'";
                }
                else if (strDatasource == "Processor")
                {
                    oAdo.m_strSQL = "UPDATE processor_scenario_datasource SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "' AND TRIM(Table_Type) = '" + strTableType + "'";
                }
                if (oAdo.m_strSQL.Trim().Length > 0)
                {
                    oAdo.SqlNonQuery(oAdo.m_OleDbConnection, oAdo.m_strSQL);
                }
            }
            oAdo.CloseConnection(oAdo.m_OleDbConnection);
            MessageBox.Show("Done", "FIA Biosum");
            m_bSyncd = false;
        }
        private void Sync()
        {
            if (txtReplace.Text.Trim().Length == 0) return;
            m_bSyncd = true;
           
            int intPathNF = 0;
            int intRootNF = 0;
            string strSyncd = "";
            string strReplace = this.txtReplace.Text.Trim();
            if (strReplace.Substring(strReplace.Length - 1, 1) == "\\")
                strReplace = strReplace.Substring(0, strReplace.Length - 1);

            for (int x = 0; x<= lvDatasources.Items.Count - 1; x++)
            {
                strSyncd = lvDatasources.Items[x].SubItems[COLUMN_SYNCD].Text.Trim();
                if (strSyncd == "No")
                {
                    int intIndex = lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim().ToUpper().IndexOf(strReplace.Trim().ToUpper());
                    if (intIndex == 0)
                    {

                        string strRemainder =
                            lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim().Substring(strReplace.Trim().Length + 1, lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim().Length - strReplace.Trim().Length - 1);

                        string strNewString = lblCurrentProjectRootFolder.Text.Trim() + "\\" + strRemainder;
                        lvDatasources.Items[x].SubItems[COLUMN_PATH].Text = strNewString;

                    }
                    if (System.IO.Directory.Exists(lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim()))
                    {
                        lvDatasources.Items[x].SubItems[COLUMN_PATHFOUND].Text = "Yes";
                        lvDatasources.Items[x].SubItems[COLUMN_PATHFOUND].BackColor = Color.Green;
                    }
                    else
                    {
                        lvDatasources.Items[x].SubItems[COLUMN_PATHFOUND].Text = "No";
                        lvDatasources.Items[x].SubItems[COLUMN_PATHFOUND].BackColor = Color.Red;
                        intPathNF++;

                    }
                    if (lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim().ToUpper().Contains(lblCurrentProjectRootFolder.Text.Trim().ToUpper()))
                    {
                        lvDatasources.Items[x].SubItems[COLUMN_SYNCD].Text = "Yes";
                        lvDatasources.Items[x].SubItems[COLUMN_SYNCD].BackColor = Color.Green;

                    }
                    else
                    {
                        lvDatasources.Items[x].SubItems[COLUMN_SYNCD].Text = "No";
                        lvDatasources.Items[x].SubItems[COLUMN_SYNCD].BackColor = Color.Red;
                        intRootNF++;
                    }
                }
            }
            lblFolderPaths.Text = intPathNF.ToString().Trim();
            lblProjectRootFolderNotFound.Text = intRootNF.ToString().Trim();
            if (intRootNF > 0) btnAnalyze.Enabled = true; else btnAnalyze.Enabled = false;

        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            List<string> oProjectRootFolders = new List<string>();
            for (int x = 0; x <= lvDatasources.Items.Count - 1; x++)
            {
                
                
                string strSyncd = lvDatasources.Items[x].SubItems[COLUMN_SYNCD].Text.Trim();
                if (strSyncd == "No")
                {
                    int intIndex = -1;
                    string strProjectRootFolder = "";
                    string strDatasource = lvDatasources.Items[x].SubItems[COLUMN_DATASOURCE].Text.Trim();
                    string strPath = lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim().ToUpper();
                    if (strDatasource == "Project")
                    {
                        intIndex = strPath.IndexOf(@"\DB", 0);
                        if (intIndex > 0)
                        {
                            strProjectRootFolder = strPath.Substring(0, intIndex + 1);
                        }
                    }
                    else if (strDatasource == "TreatmentOptimizer")
                    {
                        // THIS CONDITION WILL BE MET BY THE 'NA' ROWS THAT ARE LISTED FOR EACH SCENARIO GENERATED FROM THE CORE scenario_core_rule_definitions.mdb\scenario table
                        intIndex = strPath.IndexOf(@"\OPTIMIZER\", 0);
                        if (intIndex > 0)
                        {
                            strProjectRootFolder = strPath.Substring(0, intIndex + 1);
                        }
                    }
                    else if (strDatasource == "Processor")
                    {
                        // THIS CONDITION WILL BE MET BY THE 'NA' ROWS THAT ARE LISTED FOR EACH SCENARIO GENERATED FROM THE PROCESSOR scenario_core_rule_definitions.mdb\scenario table
                        intIndex = strPath.IndexOf(@"\PROCESSOR\", 0);
                        if (intIndex > 0)
                        {
                            strProjectRootFolder = strPath.Substring(0, intIndex + 1);
                        }
                    }
                    if (strProjectRootFolder.Trim().Length > 0)
                    {
                        if ((int)oProjectRootFolders.Where(p => p == strProjectRootFolder).Count() == 0)
                            oProjectRootFolders.Add(strProjectRootFolder);
                    }
                }
            }
            if (oProjectRootFolders != null && oProjectRootFolders.Count > 0 && oProjectRootFolders[0] != null)
            {
                FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();
                oDlg.Text = "FIA Biosum: Scan and Analyze";
                oDlg.uc_select_list_item1.lblTitle.Text = "Suggested Out-of-Sync Project Root Folder(s)";
                oDlg.uc_select_list_item1.listBox1.Sorted = false;
                oDlg.uc_select_list_item1.lblMsg.Hide();

                oDlg.uc_select_list_item1.loadvalues(oProjectRootFolders);
                oDlg.uc_select_list_item1.Show();

                DialogResult result = oDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    if (oDlg.uc_select_list_item1.listBox1.SelectedItems.Count > 0)
                    {
                        this.txtReplace.Text = oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim();
                    }
                }
                oDlg.Dispose();
                oDlg = null;
            }
        }
    }
}

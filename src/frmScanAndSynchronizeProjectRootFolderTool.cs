using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SQLite.ADO;

namespace FIA_Biosum_Manager
{
    public partial class frmScanAndSynchronizeProjectFolderTool : Form
    {
        const int COLUMN_NULL = 0;
        const int COLUMN_DATASOURCE = 1;
        const int COLUMN_SCENARIO = 2;
        const int COLUMN_TABLETYPE = 3;
        const int COLUMN_PATH = 4;
        const int COLUMN_PATHFOUND = 5;
        const int COLUMN_SYNCD = 6;

        private string m_strRandomPathAndFile = "";

        bool m_bSyncd = false;

        public frmScanAndSynchronizeProjectFolderTool()
        {
            InitializeComponent();
            this.lblCurrentProjectFolder.Text = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory.Trim();
            loadvalues();
            
        }
        
        public void loadvalues()
        {
            m_bSyncd = false;
            DataMgr p_dataMgr = new DataMgr();

            int x = 0;
            int intPathNF = 0;
            int intRootNF = 0;


            this.lvDatasources.Clear();

            this.lvDatasources.Columns.Add(" ", 2, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("DataSource", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("Scenario", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("TableType", 50, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("Path", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("PathFound", 100, HorizontalAlignment.Left);
            this.lvDatasources.Columns.Add("Synchronized", 100, HorizontalAlignment.Left);

            m_strRandomPathAndFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");

            p_dataMgr.CreateDbFile(m_strRandomPathAndFile);
            //
            //PROJECT DATA SOURCES
            //
            string strProjectDbPath = this.lblCurrentProjectFolder.Text.Trim() + "\\db\\" + frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectFile;
            string strOptimizerDbPath = this.lblCurrentProjectFolder.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
            string strProcessorDbPath = this.lblCurrentProjectFolder.Text.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile;
            bool bOptimizerExists = false;
            bool bProcessorExists = false;

            string strConn = p_dataMgr.GetConnectionString(m_strRandomPathAndFile);
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strConn))
            {
                conn.Open();

                p_dataMgr.m_strSQL = "ATTACH DATABASE '" + strProjectDbPath + "' AS project";
                p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);

                if (System.IO.File.Exists(strOptimizerDbPath))
                {
                    bOptimizerExists = true;
                    p_dataMgr.m_strSQL = "ATTACH DATABASE '" + strOptimizerDbPath + "' AS optimizer";
                    p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);
                }
                
                if (System.IO.File.Exists(strProcessorDbPath))
                {
                    bProcessorExists = true;
                    p_dataMgr.m_strSQL = "ATTACH DATABASE '" + strProcessorDbPath + "' AS processor";
                    p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);
                }
                
                p_dataMgr.m_strSQL = "SELECT 'Project' AS DataSource, 'NA' AS Scenario, table_type AS TableType, path FROM project.datasource";

                if (bOptimizerExists)
                {
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + " UNION SELECT 'TreatmentOptimizer' AS DataSource, Scenario_Id AS Scenario, " +
                        "'NA' AS TableType, path FROM optimizer.scenario";
                }

                if (bOptimizerExists)
                {
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + " UNION SELECT 'TreatmentOptimizer' AS DataSource, Scenario_Id AS Scenario, " +
                        "table_type AS TableType, path FROM optimizer.scenario_datasource";
                }

                if (bProcessorExists)
                {
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + " UNION SELECT 'Processor' AS DataSource, Scenario_Id AS Scenario, " +
                        "'NA' AS TableType, path FROM processor.scenario";
                }

                if (bProcessorExists)
                {
                    p_dataMgr.m_strSQL = p_dataMgr.m_strSQL + " UNION SELECT 'Processor' AS DataSource, Scenario_Id AS Scenario, " +
                        "table_type AS TableType, path FROM processor.scenario_datasource";
                }

                p_dataMgr.SqlQueryReader(conn, p_dataMgr.m_strSQL);

                if (p_dataMgr.m_DataReader.HasRows)
                {
                    while (p_dataMgr.m_DataReader.Read())
                    {
                        // Don't add appData data sources to the grid
                        if (p_dataMgr.m_DataReader["Path"].ToString().IndexOf("@@appdata@@") == -1)
                        {
                            System.Windows.Forms.ListViewItem entryListItem =
                                this.lvDatasources.Items.Add(" ");
                            entryListItem.UseItemStyleForSubItems = false;
                            this.lvDatasources.Items[x].SubItems.Add(p_dataMgr.m_DataReader["DataSource"].ToString());
                            this.lvDatasources.Items[x].SubItems.Add(p_dataMgr.m_DataReader["Scenario"].ToString());
                            this.lvDatasources.Items[x].SubItems.Add(p_dataMgr.m_DataReader["TableType"].ToString());
                            this.lvDatasources.Items[x].SubItems.Add(p_dataMgr.m_DataReader["Path"].ToString());
                            if (System.IO.Directory.Exists(p_dataMgr.m_DataReader["Path"].ToString().Trim()))
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
                            if (p_dataMgr.m_DataReader["Path"].ToString().ToUpper().Contains(lblCurrentProjectFolder.Text.Trim().ToUpper()))
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
                    lblProjectFolderNotFound.Text = intRootNF.ToString().Trim();
                    p_dataMgr.m_DataReader.Close();
                }
                
            }
            if (intRootNF > 0) btnAnalyze.Enabled = true; else btnAnalyze.Enabled = false;
            p_dataMgr = null;
        }
       
        private void frmScanAndSynchronizeProjectFolderTool_Resize(object sender, EventArgs e)
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
            
            DataMgr p_dataMgr = new DataMgr();

            string strProjectDbPath = this.lblCurrentProjectFolder.Text.Trim() + "\\db\\" + frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectFile;
            string strOptimizerDbPath = this.lblCurrentProjectFolder.Text.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
            string strProcessorDbPath = this.lblCurrentProjectFolder.Text.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile;

            for (int x = 0; x <= lvDatasources.Items.Count - 1; x++)
            {
                strDatasource = lvDatasources.Items[x].SubItems[COLUMN_DATASOURCE].Text.Trim();
                strTableType = lvDatasources.Items[x].SubItems[COLUMN_TABLETYPE].Text.Trim();
                strPath = lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim();
                strScenario = lvDatasources.Items[x].SubItems[COLUMN_SCENARIO].Text.Trim();
                p_dataMgr.m_strSQL = "";

                if (strDatasource == "Project")
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(strProjectDbPath)))
                    {
                        conn.Open();

                        p_dataMgr.m_strSQL = "UPDATE datasource SET path = '" + strPath + "' WHERE TRIM(table_type) = '" + strTableType + "'";
                        p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);
                    }
                }
                if (strDatasource == "TreatmentOptimizer")
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(strOptimizerDbPath)))
                    {
                        conn.Open();

                        if (strTableType == "NA")
                        {
                            p_dataMgr.m_strSQL = "UPDATE scenario SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "'";
                        }
                        else
                        {
                            p_dataMgr.m_strSQL = "UPDATE scenario_datasource SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "' AND TRIM(table_type) = '" + strTableType + "'";
                        }
                        p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);
                    }
                }
                if (strDatasource == "Processor")
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(p_dataMgr.GetConnectionString(strProcessorDbPath)))
                    {
                        conn.Open();

                        if (strTableType == "NA")
                        {
                            p_dataMgr.m_strSQL = "UPDATE scenario SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "'";
                        }
                        else
                        {
                            p_dataMgr.m_strSQL = "UPDATE scenario_datasource SET path = '" + strPath + "' WHERE TRIM(scenario_id) = '" + strScenario + "' AND TRIM(table_type) = '" + strTableType + "'";
                        }
                        p_dataMgr.SqlNonQuery(conn, p_dataMgr.m_strSQL);
                    }
                }
            }
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

                        string strNewString = lblCurrentProjectFolder.Text.Trim() + "\\" + strRemainder;
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
                    if (lvDatasources.Items[x].SubItems[COLUMN_PATH].Text.Trim().ToUpper().Contains(lblCurrentProjectFolder.Text.Trim().ToUpper()))
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
            lblProjectFolderNotFound.Text = intRootNF.ToString().Trim();
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
                    if (strDatasource == "TreatmentOptimizer")
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
                    intIndex = strPath.IndexOf(@"\DB", 0);
                    if (intIndex > 0)
                    {
                        strProjectRootFolder = strPath.Substring(0, intIndex + 1);
                    }
                    intIndex = -1;
                    intIndex = strPath.IndexOf(@"\GIS", 0);
                    if (intIndex > 0)
                    {
                        strProjectRootFolder = strPath.Substring(0, intIndex + 1);
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
                oDlg.uc_select_list_item1.lblTitle.Text = "Suggested Out-of-Sync Project Folder(s)";
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using SQLite.ADO;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace FIA_Biosum_Manager
{
    public partial class uc_delete_packages : UserControl
    {
        public int m_DialogHt;
        public int m_DialogWd;

        private string m_strPackageNumbers;
        private string[] m_packageNumbers;
        private Dictionary<string, string> m_dictDeletedDatabasesAndFileSizes = new Dictionary<string, string>();
        Dictionary<string, Dictionary<string, int>> m_dictDeletedRowCountsByDatabaseAndTable =
            new Dictionary<string, Dictionary<string, int>>();

        //TODO: Help files
        private env m_oEnv;

        private Help m_oHelp;
        private string m_xpsFile = Help.DefaultDatabaseXPSFile;
        private int m_intError;
        private string m_strCurrentProcess;
        private frmTherm m_frmTherm;
        private DataMgr m_dataMgr;
        private string m_strProjDir;
        private string m_strMessage = "";
        private string m_strDebugFile = "";

        public uc_delete_packages()
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


            m_oEnv = new env();
            m_strDebugFile = frmMain.g_oEnv.strTempDir + "\\biosum_delete_packages_debug" +
                             String.Format("{0:yyyyMMddhhmm}", DateTime.Now) + ".txt";
            m_strProjDir = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim();
        }

        private void rdoFilterByFile_Click(object sender, System.EventArgs e)
        {
            btnFilterFinish.Enabled = false;
            btnFilterNext.Enabled = false;
            txtFilterByFile.Enabled = true;
            btnFilterByFileBrowse.Enabled = true;
            if (!string.IsNullOrWhiteSpace(txtFilterByFile.Text))
            {
                btnFilterFinish.Enabled = true;
            }
        }

        private void rdoFilterByMenu_Click(object sender, System.EventArgs e)
        {
            btnFilterFinish.Enabled = false;
            btnFilterNext.Enabled = true;
            txtFilterByFile.Enabled = false;
            btnFilterByFileBrowse.Enabled = false;
        }

        private void rdoDeleteAllPkgs_Click(object sender, System.EventArgs e)
        {
            btnFilterFinish.Enabled = true;
            btnFilterNext.Enabled = true;
            txtFilterByFile.Enabled = false;
            btnFilterByFileBrowse.Enabled = false;
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            ((frmDialog) ParentForm).Close();
        }

        private void btnFilterFinish_Click(object sender, System.EventArgs e)
        {
            m_intError = 0;
            if (m_intError == 0)
            {
                if (rdoFilterByFile.Checked)
                {
                    if (System.IO.File.Exists(txtFilterByFile.Text.Trim()) == true)
                    {
                        m_strPackageNumbers = CreateDelimitedStringList(txtFilterByFile.Text.Trim(), ",", ",", false);
                        m_packageNumbers = m_strPackageNumbers.Replace("'", string.Empty).Split(',').Distinct().ToArray();
                        if (m_intError == 0) { DeletePkgsFromBiosumProject_Start(); }
                    }
                    else
                    {
                        MessageBox.Show("!!" + txtFilterByFile.Text.Trim() + " could not be found!!",
                            "Delete Packages", System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Exclamation);
                    }
                }
            }
        }

        private void DeletePkgsFromBiosumProject_Start()
        {
            frmMain.g_oDelegate.InitializeThreadEvents();
            frmMain.g_oDelegate.m_oEventStopThread.Reset();
            frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
            frmMain.g_oDelegate.CurrentThreadProcessAborted = false;
            frmMain.g_oDelegate.CurrentThreadProcessDone = false;
            frmMain.g_oDelegate.CurrentThreadProcessStarted = false;
            m_strCurrentProcess = "DeletePackages";
            StartTherm("2", "Delete Biosum Package Data");
            frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(DeletePkgsFromBiosumProject_Process));
            frmMain.g_oDelegate.m_oThread.IsBackground = true;
            frmMain.g_oDelegate.CurrentThreadProcessIdle = false;
            frmMain.g_oDelegate.m_oThread.Start();
        }

        private void DeletePkgsFromBiosumProject_Process()
        {
            frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
            m_intError = 0;
            try
            {
                m_dataMgr = new DataMgr();

                //progress bar 1: single process
                SetThermValue(m_frmTherm.progressBar1, "Maximum", 100);
                SetThermValue(m_frmTherm.progressBar1, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar1, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg, "Text", "");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);
                //progress bar 2: overall progress
                SetThermValue(m_frmTherm.progressBar2, "Maximum", 100);
                SetThermValue(m_frmTherm.progressBar2, "Minimum", 0);
                SetThermValue(m_frmTherm.progressBar2, "Value", 0);
                SetLabelValue(m_frmTherm.lblMsg2, "Text", "Overall Progress");
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", true);

                UpdateProgressBar2(0);

                List<string> projectFiles = new List<string>();
                if (Directory.Exists(m_strProjDir + "\\db\\"))
                {
                    projectFiles.AddRange(Directory
                        .GetFiles(m_strProjDir + "\\db\\", "*.*", SearchOption.AllDirectories)
                        .Where(s => !s.ToLower().Contains("fvsmaster.db")));
                }

                if (Directory.Exists(m_strProjDir + "\\fvs\\"))
                {
                    projectFiles.AddRange(Directory.GetFiles(m_strProjDir + "\\fvs\\", "*.*",
                        SearchOption.AllDirectories));
                }

                if (Directory.Exists(m_strProjDir + "\\processor\\"))
                {
                    projectFiles.AddRange(Directory.GetFiles(m_strProjDir + "\\processor\\", "*.*",
                        SearchOption.AllDirectories));
                }

                if (Directory.Exists(m_strProjDir + "\\opcost\\"))
                {
                    projectFiles.AddRange(Directory.GetFiles(m_strProjDir + "\\opcost\\", "*.*",
                        SearchOption.AllDirectories));
                }

                //OPTIMIZER weighted variables
                if (Directory.Exists(m_strProjDir + "\\optimizer\\db\\"))
                {
                    projectFiles.AddRange(Directory.GetFiles(m_strProjDir + "\\optimizer\\db\\", "prepost*.db",
                        SearchOption.AllDirectories));
                }

                string[] allProjectFiles = projectFiles.ToArray();
                string[] databases = allProjectFiles.Where(s => s.ToLower().EndsWith(".db")).ToArray();
                Regex packagePattern = new Regex(".*P\\d{3}.*");
                string[] nonPackageDatabases = databases.Where(s => !packagePattern.Match(s).Success).ToArray();
                string[] packageFiles = allProjectFiles.Where(s => packagePattern.Match(s).Success
                                                              && !s.ToLower().EndsWith(".kcp")
                                                              && !s.ToLower().EndsWith(".kcp.template")).ToArray();

                //Delete databases entirely if their names match one of the package numbers supplied by the user
                var counter = 0;
                foreach (var pathFile in packageFiles)
                {
                    foreach (var packageNumber in m_packageNumbers)
                    {
                        if (pathFile.Contains(packageNumber))
                        {
                            string fileSize = string.Format("{0:0.00}", new FileInfo(pathFile).Length / 1048576.0);
                            m_dictDeletedDatabasesAndFileSizes.Add(pathFile, fileSize);
                            string message = (!Checked(chkDeletesDisabled) ? "Would delete : " : "Attempting to delete: ") + pathFile + "\r\n";
                            if (!Checked(chkDeletesDisabled)) File.Delete(pathFile);
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2) frmMain.g_oUtils.WriteText(m_strDebugFile, message);
                            break;
                        }
                    }

                    counter += 1;
                    UpdateProgressBar1(Path.GetFileName(pathFile), (int)Math.Floor(100 * ((double)(counter + 1) / (packageFiles.Length + 1))));
                    UpdateProgressBar2((int)Math.Floor(50 * ((double)(counter + 1) / (packageFiles.Length + 1))));
                }

                UpdateProgressBar1("", 0);
                UpdateProgressBar2(51);

                //Look at every access database without PXXX in filename in project and perform delete queries
                ConnectToDatabasesInPathAndExecuteDeletes(nonPackageDatabases);

                UpdateProgressBar2(100);

                if (!Checked(chkDeletesDisabled))
                {
                    MessageBox.Show(
                        String.Format("Successfully deleted data associated with {0} packages!",
                            m_packageNumbers.Length), "Delete Packages Results");
                }

                if (m_dataMgr != null)
                {
                    if (m_dataMgr.m_DataSet != null)
                    {
                        m_dataMgr.m_DataSet.Clear();
                        m_dataMgr.m_DataSet.Dispose();
                    }
                    m_dataMgr = null;
                }

                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Visible",
                    true);
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled",
                    true);

                DeletePkgsFromBiosumProject_Finish();
            }
            catch (System.Threading.ThreadInterruptedException err)
            {
                MessageBox.Show("Threading Interruption Error " + err.Message.ToString());
            }
            catch (System.Threading.ThreadAbortException err)
            {
                if (m_dataMgr != null)
                {
                    if (m_dataMgr.m_DataSet != null)
                    {
                        m_dataMgr.m_DataSet.Clear();
                        m_dataMgr.m_DataSet.Dispose();
                    }
                    m_dataMgr = null;
                }
                ThreadCleanUp();
                CleanupThread();
            }
            catch (Exception err)
            {
                MessageBox.Show("!!Error!! \n" +
                                "Module - uc_delete_packages:DeletePackagesFromBiosumProject_Process  \n" +
                                "Err Msg - " + err.Message.ToString().Trim(),
                    "FVS Biosum", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Exclamation);
                m_intError = -1;
            }
            finally
            {
                string strLogFileName = String.Concat(frmMain.g_oFrmMain.getProjectDirectory(),
                        "\\db\\biosum_deleted_packages_",
                        String.Format("{0:yyyyMMddhhmm}", DateTime.Now),
                        Checked(chkDeletesDisabled) ? "_no_deletes" : "",
                        ".txt");
                if (Checked(chkCreateLog) || Checked(chkDeletesDisabled))
                {
                    CreateLogFile(strLogFileName);
                    if (Checked(chkDeletesDisabled))
                    {
                        MessageBox.Show("Record count complete! The results were recorded in " + strLogFileName,
                        "Delete Packages Results");
                    }
                }

                m_dictDeletedDatabasesAndFileSizes = new Dictionary<string, string>();
                m_dictDeletedRowCountsByDatabaseAndTable = new Dictionary<string, Dictionary<string, int>>();
            }

            if (m_dataMgr != null)
            {
                if (m_dataMgr.m_DataSet != null)
                {
                    m_dataMgr.m_DataSet.Clear();
                    m_dataMgr.m_DataSet.Dispose();
                }
                m_dataMgr = null;
            }
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)ReferenceFormDialog, "Enabled",
                true);
            if (m_frmTherm != null)
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form)m_frmTherm, "Visible", false);


            DeletePkgsFromBiosumProject_Finish();

            CleanupThread();

            frmMain.g_oDelegate.m_oEventThreadStopped.Set();
            Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);
        }
        private void UpdateProgressBar1(string label, int value)
        {
            SetLabelValue(m_frmTherm.lblMsg, "Text", label);
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) m_frmTherm.progressBar1,
                "Value", value);
        }

        private void UpdateProgressBar2(int value)
        {
            frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Control) m_frmTherm.progressBar2,
                "Value", value);
        }

        private void ConnectToDatabasesInPathAndExecuteDeletes(string[] strDatabaseNames, string[] strTargetTables = null, string[] strTableExceptions = null)
        {
            int counter = 0;

            foreach (string db in strDatabaseNames)
            {
                var message = (Checked(chkDeletesDisabled) ? "Checking " : "Deleting from ") + Path.GetFileName(db);
                UpdateProgressBar1(message, (int)Math.Floor(100 * ((double)(counter + 1) / (strDatabaseNames.Length + 1))));
                UpdateProgressBar2(50 + (int)Math.Floor(50 * ((double)(counter + 1) / (strDatabaseNames.Length + 1))));
                counter += 1;
                ExecuteDeleteOnTables(db, tables: strTargetTables, exceptions: strTableExceptions);
            }
            UpdateProgressBar1("", 0);
        }


        private void ExecuteDeleteOnTables(string strDbPathFile, string[] tables = null, string[] exceptions = null)
        {
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(m_dataMgr.GetConnectionString(strDbPathFile)))
            {
                conn.Open();

                string[] strTables = tables;
                if (tables == null || tables.Length == 0)
                {
                    strTables = m_dataMgr.getTableNames(conn).Where(s => !(s.Contains("~") || s.Contains(" "))).ToArray();
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
                    foreach (string col in new string[] { "rxpackage" })
                    {
                        if (m_dataMgr.ColumnExist(conn, table, col))
                        {
                            column = col;
                            break;
                        }
                    }

                    if (!String.IsNullOrEmpty(column))
                    {
                        string strTempIndex = column + "_delete_idx";
                        if (!m_dataMgr.IndexExist(conn, strTempIndex))
                        {
                            m_dataMgr.AddIndex(conn, table, strTempIndex, column);
                        }

                        if (frmMain.g_bDebug)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, Checked(chkDeletesDisabled)
                                ? "\r\nCounting records to delete from " + strDbPathFile + " " + table + " using " + column + "\r\n"
                                : "\r\nDeleting from " + strDbPathFile + " " + table + " using " + column + "\r\n");
                        }

                        int deletedRecords = BuildAndExecuteDeleteSQLStmts(conn, table, column);
                        if (deletedRecords > 0) AddDeletedCountToDictionary(strDbPathFile, table, deletedRecords);
                        m_dataMgr.m_strSQL = "DROP INDEX " + strTempIndex;
                        m_dataMgr.SqlNonQuery(conn, m_dataMgr.m_strSQL);
                    }
                }
            }
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


        private int BuildAndExecuteDeleteSQLStmts(System.Data.SQLite.SQLiteConnection oConn, string table, string column)
        {
            int matchedRecordCount = (int)m_dataMgr.getSingleDoubleValueFromSQLQuery(oConn,
                "SELECT COUNT(*) FROM " + table + " WHERE " + column + " IN (" + m_strPackageNumbers + ")", table);
            if (matchedRecordCount == 0) return 0;

            string deleteSQL = "DELETE FROM " + table + " WHERE " + column + " IN (" + m_strPackageNumbers + ")";
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
            {
                var message = Checked(chkDeletesDisabled)
                    ? "Generated but not executed: " + deleteSQL + "\r\n"
                    : "Attempting to execute: " + deleteSQL + "\r\n";
                frmMain.g_oUtils.WriteText(m_strDebugFile, message);
            }
            if (!Checked(chkDeletesDisabled)) m_dataMgr.SqlNonQuery(oConn, deleteSQL);
            return matchedRecordCount;
        }
        private void CreateLogFile(string strLogFileName)
        {
            string strMessage;

            if (m_dictDeletedDatabasesAndFileSizes.Keys.Count > 0)
            {
                frmMain.g_oUtils.WriteText(strLogFileName,
                    String.Format("Databases or other files associated with these packages: {0}{1}",
                        String.Join(",", m_packageNumbers), Environment.NewLine));
                strMessage = Checked(chkDeletesDisabled) ? "would be deleted" : "was deleted";
                foreach (var pathFile in m_dictDeletedDatabasesAndFileSizes.Keys)
                {
                    frmMain.g_oUtils.WriteText(strLogFileName,
                        String.Format("  {0} ({1} MB) {2}{3}", pathFile,
                            m_dictDeletedDatabasesAndFileSizes[pathFile], strMessage, Environment.NewLine));
                }
            }
            else
            {
                frmMain.g_oUtils.WriteText(strLogFileName, "No matching package-specific files were found." + Environment.NewLine);
            }

            if (m_dictDeletedRowCountsByDatabaseAndTable.Keys.Count > 0)
            {
                frmMain.g_oUtils.WriteText(strLogFileName,
                    String.Format("{1}Deleting records from databases where rxpackage in ({0}):{1}",
                        String.Join(",", m_strPackageNumbers), Environment.NewLine));
                strMessage = Checked(chkDeletesDisabled)
                    ? "records would be deleted from"
                    : "records were deleted from";

                foreach (var strDbPathFile in m_dictDeletedRowCountsByDatabaseAndTable.Keys)
                {
                    if (m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Keys.Count > 0)
                    {
                        frmMain.g_oUtils.WriteText(strLogFileName,
                            String.Concat("  ", strDbPathFile, Environment.NewLine));
                    }

                    foreach (var strTable in m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile].Keys)
                    {
                        var deletedRecordsCount = m_dictDeletedRowCountsByDatabaseAndTable[strDbPathFile][strTable];
                        if (deletedRecordsCount > 0)
                        {
                            frmMain.g_oUtils.WriteText(strLogFileName,
                                String.Format("    {0} {1} {2}{3}", deletedRecordsCount, strMessage, strTable,
                                    Environment.NewLine));
                        }
                    }
                }
            }
            else
            {
                frmMain.g_oUtils.WriteText(strLogFileName, "No matching package data was deleted from tables." + Environment.NewLine);
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

                if (m_frmTherm != null)
                {
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form) m_frmTherm, "Close");
                    frmMain.g_oDelegate.ExecuteControlMethod((System.Windows.Forms.Form) m_frmTherm, "Dispose");

                    m_frmTherm = null;
                }
            }
            catch
            {
            }
        }

        private void DeletePkgsFromBiosumProject_Finish()
        {
            if (m_frmTherm != null)
            {
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Close");
                frmMain.g_oDelegate.ExecuteControlMethod(m_frmTherm, "Dispose");
                m_frmTherm = null;
            }
            m_strCurrentProcess = "";
            frmMain.g_oDelegate.SetControlPropertyValue(this, "Enabled", true);
            ((frmDialog) ParentForm).MinimizeMainForm = false;
        }


        private void StartTherm(string p_strNumberOfTherms, string p_strTitle)
        {
            m_frmTherm = new frmTherm((frmDialog) ParentForm, p_strTitle);

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

            if (p_strNumberOfTherms == "2")
            {
                m_frmTherm.progressBar2.Size = m_frmTherm.progressBar1.Size;
                m_frmTherm.progressBar2.Left = m_frmTherm.progressBar1.Left;
                m_frmTherm.progressBar2.Top =
                    Convert.ToInt32(m_frmTherm.progressBar1.Top + (m_frmTherm.progressBar1.Height * 3));
                m_frmTherm.lblMsg2.Top =
                    m_frmTherm.progressBar2.Top + m_frmTherm.progressBar2.Height + 5;
                m_frmTherm.Height = m_frmTherm.lblMsg2.Top + m_frmTherm.lblMsg2.Height +
                                         m_frmTherm.btnCancel.Height + 50;
                m_frmTherm.btnCancel.Top =
                    m_frmTherm.ClientSize.Height - m_frmTherm.btnCancel.Height - 5;
                m_frmTherm.lblMsg2.Show();
                m_frmTherm.progressBar2.Visible = true;
            }
            m_frmTherm.AbortProcess = false;
            m_frmTherm.Refresh();
            m_frmTherm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            ((frmDialog)ParentForm).Enabled=false;
            m_frmTherm.Visible = true;
        }

        public void StopThread()
        {
            string strMsg = "";

            frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Do you wish to cancel adding plot data?");

            if (frmMain.g_oDelegate.CurrentThreadProcessAborted)
            {
                frmMain.g_oDelegate.SetControlPropertyValue((System.Windows.Forms.Form) m_frmTherm, "AbortProcess",
                    true);
//                CancelThreadCleanup();
//                ThreadCleanUp();
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
            m_intError = 0;
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
                m_intError = -1;
                MessageBox.Show("!!Error: CreateDelimitedStringList() Routine Error Msg:" + caught.Message);
            }
            return strList;
        }

        private void txtFilterByFile_TextChanged(object sender, EventArgs e)
        {
            btnFilterFinish.Enabled = true;
        }

        private void btnFilterHelp_Click(object sender, EventArgs e)
        {
            if (m_oHelp == null)
            {
                m_oHelp = new Help(m_xpsFile, m_oEnv);
            }
            m_oHelp.ShowHelp(new string[] { "DATABASE", "DELETE_PACKAGES" });
        }

    }
}

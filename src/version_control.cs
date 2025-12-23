using SQLite.ADO;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;

namespace FIA_Biosum_Manager
{
	/// <summary>
	/// Summary description for version_control.
	/// </summary>
	public class version_control
	{
		private ado_data_access m_oAdo = new ado_data_access();
		const int APP_VERSION_MAJOR=0;
		const int APP_VERSION_MINOR1=1;
		const int APP_VERSION_MINOR2=2;		
		private string[] m_strAppVerArray=null;
		private string m_strProjectVersion="1.0.0";
		private string[] m_strProjectVersionArray=null;
		private string _strProjDir="";
		public version_control()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		/// <summary>
		/// Check the project's application version and update to the current version
		/// if different.
		/// </summary>
		public void PerformVersionCheck()
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//version_control.PerformVersionCheck \r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                frmMain.g_oFrmMain.WindowState,
                frmMain.g_oFrmMain.Left,
                frmMain.g_oFrmMain.Height,
                frmMain.g_oFrmMain.Width,
                frmMain.g_oFrmMain.Top);
            bool bPerformCheck = true;
            string strProjVersionFile = this.ReferenceProjectDirectory.Trim() + "\\application.version";

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: strProjVersionFile=" + strProjVersionFile + "\r\n");

            m_strAppVerArray = frmMain.g_oUtils.ConvertListToArray(frmMain.g_strAppVer, ".");
            if (System.IO.File.Exists(strProjVersionFile))
            {

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: open application version file\r\n");
                try
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: instantiate streamreader and open file\r\n");
                    //Open the file in a stream reader.
                    System.IO.StreamReader s = new System.IO.StreamReader(strProjVersionFile);
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  application version file opened with no errors\r\n");

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  streamreader.ReadLine\r\n");
                    //Split the first line into the columns       
                    string strProjVersion = s.ReadLine();
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  streamreader.ReadLine successful\r\n");
                    s.Close();
                    s.Dispose();
                    s = null;
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  streamreader close and dispose successful\r\n");

                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  strProjVersion=" + strProjVersion + "\r\n");
                    if (strProjVersion.Trim() == frmMain.g_strAppVer.Trim())
                    {
                        bPerformCheck = false;
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  bPerformCheck=false\r\n");
                    }
                    else
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck:  bPerformCheck=true\r\n");

                        if (strProjVersion.Trim().Length > 0)
                        {
                            this.m_strProjectVersion = strProjVersion.Trim();
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Convert " + m_strProjectVersion + " to an array\r\n");
                            m_strProjectVersionArray = frmMain.g_oUtils.ConvertListToArray(m_strProjectVersion, ".");
                            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                            {
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Conversion to array completed\r\n");
                                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: m_strProjectVersionArray[APP_VERSION_MAJOR]=" + m_strProjectVersionArray[APP_VERSION_MAJOR] + " m_strProjectVersionArray[APP_VERSION_MINOR1]=" + m_strProjectVersionArray[APP_VERSION_MINOR1] + " m_strProjectVersionArray[APP_VERSION_MINOR2]=" + m_strProjectVersionArray[APP_VERSION_MINOR2] + "\r\n");
                            }

                        }
                    }
                }
                catch (Exception err)
                {
                    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                        frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: !!Error opening Application.Version File!! ERROR=" + err.Message + "r\n");
                }
            }
            else
            {
                m_strProjectVersionArray = frmMain.g_oUtils.ConvertListToArray(m_strProjectVersion, ".");
            }

            //check for partial update
            if (bPerformCheck)
            {
                if (m_strProjectVersion.Trim().Length > 0)
                {
                    // Upgrade to 5.10.1 to 5.11.1 (Sequence numbers, Optimizer, Processor to SQLite, new field on Psites table)
                    if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 1) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 10 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 1))
                    {
                        UpdateDatasources_5_11_0();
                        UpdateDatasources_5_11_1();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgrade from 5.11.0 to 5.11.1 (new field on Psites table)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 1) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) > 10 ))
                    {
                        UpdateDatasources_5_11_1();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgrade from 5.11.1 to 5.11.2 (rx, rxPackage, harvest_methods, FVSIn move to SQLite)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 2) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 1))
                    {
                        UpdateDatasources_5_11_2();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                    // Upgraded from 5.11.2 to 5.12.0 (master.mdb to SQLite)
                    else if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 12 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 0) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 11 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 2))
                    {
                        UpdateDatasources_5_12_0();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                }
            }
            frmMain.g_oFrmMain.DeactivateStandByAnimation();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "version_control.PerformVersionCheck: Leaving\r\n");
        }
        private void UpdateProjectVersionFile(string p_strProjectVersionFile)
        {
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//version_control.UpdateProjectVersionFile \r\n");
                frmMain.g_oUtils.WriteText(frmMain.g_oFrmMain.frmProject.uc_project1.m_strDebugFile, "//\r\n");
            }
            if (System.IO.File.Exists(p_strProjectVersionFile))
                System.IO.File.Delete(p_strProjectVersionFile);
            frmMain.g_oUtils.WriteText(p_strProjectVersionFile, frmMain.g_strAppVer);            
        }
        public void UpdateDatasources_5_11_0()
        {
            DataMgr oDataMgr = new DataMgr();
            ado_data_access oAdo = new ado_data_access();
            dao_data_access oDao = new dao_data_access();
            ODBCMgr odbcmgr = new ODBCMgr();
            utils oUtils = new utils();
            env oEnv = new env();
            Datasource oProjectDs = new Datasource();

            // MIGRATING SEQUENCE NUMBER SETTINGS TO fvs_master.db
            string strPrePostSeqNumLink = $@"{Tables.FVS.DefaultFVSPrePostSeqNumTable}_1";
            string strRxPackageAssignLink = $@"{Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable}_1";

            oProjectDs.m_strDataSourceMDBFile = ReferenceProjectDirectory.Trim() + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            int intSeqNumDefs = oProjectDs.getValidTableNameRow(Datasource.TableTypes.SeqNumDefinitions);
            int intSeqNumRxPkgAssign = oProjectDs.getValidTableNameRow(Datasource.TableTypes.SeqNumRxPackageAssign);
            if (intSeqNumDefs > -1 && intSeqNumRxPkgAssign > -1)
            {
                if (!System.IO.File.Exists($@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}"))
                {
                    oDataMgr.CreateDbFile($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");
                }
                string dbConn = oDataMgr.GetConnectionString($@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(dbConn))
                {
                    conn.Open();
                    if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateFVSOutputSQLitePrePostSeqNumTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                    }
                    if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable))
                    {
                        frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSQLiteSeqNumRxPackageAssgnTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                    }
                }
                // Create ODBC entry for the new SQLite fvs_master.db file
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName, $@"{ReferenceProjectDirectory.Trim()}\{Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile}");

                string strSeqNumDefsTable = oProjectDs.m_strDataSource[intSeqNumDefs, Datasource.TABLE].Trim();
                string strSeqNumRxPkgAssignTable = oProjectDs.m_strDataSource[intSeqNumRxPkgAssign, Datasource.TABLE].Trim();
                oDao.CreateSQLiteTableLink(oProjectDs.getFullPathAndFile(Datasource.TableTypes.SeqNumDefinitions), Tables.FVS.DefaultFVSPrePostSeqNumTable,
                    strPrePostSeqNumLink, ODBCMgr.DSN_KEYS.FvsMasterDbDsnName, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile, true);
                oDao.CreateSQLiteTableLink(oProjectDs.getFullPathAndFile(Datasource.TableTypes.SeqNumDefinitions), Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable,
                    strRxPackageAssignLink, ODBCMgr.DSN_KEYS.FvsMasterDbDsnName, ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile, true);
                string strCopyConn = oAdo.getMDBConnString(oAdo.getMDBConnString(oProjectDs.getFullPathAndFile(Datasource.TableTypes.SeqNumDefinitions), "", ""), "", "");
                int i = 0;
                int intError = 0;
                using (var oCopyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    oCopyConn.Open();
                    do
                    {
                        // break out of loop if it runs too long
                        if (i > 20)
                        {
                            System.Windows.Forms.MessageBox.Show("An error occurred while trying to migrate sequence number settings! ", "FIA Biosum");
                            break;
                        }
                        System.Threading.Thread.Sleep(1000);
                        i++;
                    }
                    while (!oAdo.TableExist(oCopyConn, strRxPackageAssignLink));

                    oAdo.m_strSQL = $@"INSERT INTO {strPrePostSeqNumLink} SELECT * FROM {strSeqNumDefsTable}";
                    oAdo.SqlNonQuery(oCopyConn, oAdo.m_strSQL);
                    if (oAdo.m_intError == 0)
                    {
                        oAdo.m_strSQL = $@"INSERT INTO {strRxPackageAssignLink} SELECT * FROM {strSeqNumRxPkgAssignTable}";
                        oAdo.SqlNonQuery(oCopyConn, oAdo.m_strSQL);
                        intError = oAdo.m_intError;
                    }
                    else
                    {
                        intError = oAdo.m_intError;
                    }

                    if (oAdo.TableExist(oCopyConn, strPrePostSeqNumLink))
                    {
                        oAdo.SqlNonQuery(oCopyConn, "DROP TABLE " + strPrePostSeqNumLink);
                    }
                    if (oAdo.TableExist(oCopyConn, strRxPackageAssignLink))
                    {
                        oAdo.SqlNonQuery(oCopyConn, "DROP TABLE " + strRxPackageAssignLink);
                    }
                }
                if (intError == 0)
                {
                    // Update entries in project data sources table
                    string strMasterPath = $@"{ReferenceProjectDirectory.Trim()}\{System.IO.Path.GetDirectoryName(Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile)}";
                    string strFvsMasterDb = System.IO.Path.GetFileName(Tables.FVS.DefaultFVSPrePostSeqNumTableDbFile);
                    oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumDefinitions, strMasterPath, strFvsMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                    oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumRxPackageAssign, strMasterPath, strFvsMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                }
            }

            // MIGRATING SETTINGS FROM scenario_processor_rule_definitions.mdb TO scenario_processor_rule_definitions.db
            string targetDbFile = ReferenceProjectDirectory.Trim() +
                @"\processor\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            string sourceDbFile = ReferenceProjectDirectory.Trim() +
                @"\processor\" + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodDbFile;
            if (System.IO.File.Exists(targetDbFile) == false)
            {
                frmMain.g_oFrmMain.frmProject.uc_project1.CreateProcessorScenarioRuleDefinitionDbAndTables(targetDbFile);
            }

            try
            {
                string[] arrTargetTables = { };
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(targetDbFile)))
                {
                    conn.Open();
                    arrTargetTables = oDataMgr.getTableNames(conn);
                    if (arrTargetTables.Length < 1)
                    {
                        System.Windows.Forms.MessageBox.Show("Target SQLite tables could not be created. Migration stopped!!", "FIA Biosum");
                        return;
                    }

                    // custom processing for scenario_additional_harvest_costs
                    string[] strSourceColumnsArray = new string[0];
                    string strAddCpaTableName = Tables.ProcessorScenarioRuleDefinitions.DefaultAdditionalHarvestCostsTableName;
                    oDao.getFieldNames(sourceDbFile, strAddCpaTableName, ref strSourceColumnsArray);
                    foreach (string strColumn in strSourceColumnsArray)
                    {
                        if (!oDataMgr.ColumnExists(conn, strAddCpaTableName, strColumn))
                        {
                            oDataMgr.AddColumn(conn, strAddCpaTableName, strColumn, "DOUBLE", "");
                        }
                    }
                }

                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, targetDbFile);

                // Create temporary database
                string strTempAccdb = oUtils.getRandomFile(oEnv.strTempDir, "accdb");
                oDao.CreateMDB(strTempAccdb);

                // Link all the target tables to the database
                for (int i = 0; i < arrTargetTables.Length; i++)
                {
                    oDao.CreateSQLiteTableLink(strTempAccdb, arrTargetTables[i], arrTargetTables[i] + "_1",
                        ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName, targetDbFile);
                }
                oDao.CreateTableLinks(strTempAccdb, sourceDbFile);  // Link all the source tables to the database

                List<string> lstScenarioPaths = new List<string>();
                string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                using (var copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    foreach (var strTable in arrTargetTables)
                    {
                        oAdo.m_strSQL = "INSERT INTO " + strTable + "_1" +
                        " SELECT * FROM " + strTable;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }

                    if (oAdo.m_intError == 0)
                    {
                        // Set file (database) field to new Sqlite DB
                        string newDbFile = System.IO.Path.GetFileName(Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile);
                        oAdo.m_strSQL = "UPDATE scenario_1 set file = '" +
                            newDbFile + "'";
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }

                    //retrieve paths for all scenarios in the project and put them in list
                    oAdo.m_strSQL = "SELECT scenario_id from scenario";
                    oAdo.SqlQueryReader(copyConn, oAdo.m_strSQL);
                    if (oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (oAdo.m_OleDbDataReader.Read())
                        {
                            string strScenarioId = "";
                            if (oAdo.m_OleDbDataReader["scenario_id"] != System.DBNull.Value)
                                strScenarioId = oAdo.m_OleDbDataReader["scenario_id"].ToString().Trim();
                            if (!String.IsNullOrEmpty(strScenarioId))
                            {
                                //Check to see if the .mdb exists before adding it to the list
                                string strPath = $@"{ReferenceProjectDirectory.Trim()}\processor\{strScenarioId}";
                                string strPathToMdb = strPath + "\\db\\scenario_results.mdb";
                                //sample path: C:\\workspace\\BioSum\\biosum_data\\bluemountains\\processor\\scenario1\\db\\scenario_results.mdb
                                if (System.IO.File.Exists(strPathToMdb))
                                    lstScenarioPaths.Add(strPath);
                            }
                        }
                        oAdo.m_OleDbDataReader.Close();
                    }
                }

                // Create tables in scenario_results.db if missing
                foreach (var sPath in lstScenarioPaths)
                {
                    string strScenarioDbPath = $@"{sPath}\{Tables.ProcessorScenarioRun.DefaultScenarioResultsTableDbFile}";
                    if (!System.IO.File.Exists(strScenarioDbPath))
                    {
                        oDataMgr.CreateDbFile(strScenarioDbPath);
                    }
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strScenarioDbPath)))
                    {
                        conn.Open();
                        if (!oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName))
                        {
                            frmMain.g_oTables.m_oProcessor.CreateSqliteHarvestCostsTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                        }
                        if (!oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName))
                        {
                            frmMain.g_oTables.m_oProcessor.CreateSqliteTreeVolValSpeciesDiamGroupsTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);
                        }
                        if (!oDataMgr.TableExist(conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName))
                        {
                            frmMain.g_oTables.m_oProcessorScenarioRun.CreateSqliteAdditionalKcpCpaTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, false);
                        }
                    }
                }

                // Add SQLite OpCost config file to db directory
                if (System.IO.File.Exists(frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile))
                {
                    if (!System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile))
                    {
                        System.IO.File.Copy(frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile,
                            ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile);
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show($@"The OpCost configuration file is missing from the AppData directory: {frmMain.g_oEnv.strAppDir + "\\" + Tables.Reference.DefaultOpCostReferenceDbFile}");
                }

                // PUT OPTIMIZER AND GIS DATA MIGRATION CODE HERE. IF YOU NEED DATA THAT ISN'T IN THIS METHOD, WE CAN PASS IT
                // VIA THE METHOD CALL FOR NOW

                // MIGRATE CALCULATED VARIABLES
                string strCalculatedVariablesPathAndDbFile = this.ReferenceProjectDirectory + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
                if (!System.IO.File.Exists(strCalculatedVariablesPathAndDbFile))
                {
                    // Create SQLite copy of optimizer_definitions database
                    string variablesSourceFile = frmMain.g_oEnv.strAppDir + "\\db\\" +
                        System.IO.Path.GetFileName(Tables.OptimizerDefinitions.DefaultSqliteDbFile);
                    System.IO.File.Copy(variablesSourceFile, strCalculatedVariablesPathAndDbFile);

                    // Check to see if the input SQLite DSN exists for optimizer_definitions and if so, delete so we can add
                    if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName))
                    {
                        odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName);
                    }
                    odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);

                    // Create temporary database for optimizer_definitions
                    strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                    oDao.CreateMDB(strTempAccdb);

                    // Create table links for transferring data for optimizer_definitions
                    string targetEcon = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName + "_1";
                    string targetFvs = Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName + "_1";
                    string targetVariables = Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName + "_1";

                    string strCalculatedVariablesPathAndAccdbFile = frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
                    // Link to all tables in source database for optimizer_definitons
                    oDao.CreateTableLinks(strTempAccdb, strCalculatedVariablesPathAndAccdbFile);
                    oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName, targetEcon,
                        ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);
                    oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName, targetFvs,
                        ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);
                    oDao.CreateSQLiteTableLink(strTempAccdb, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, targetVariables,
                        ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName, strCalculatedVariablesPathAndDbFile);

                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strCalculatedVariablesPathAndDbFile)))
                    {
                        conn.Open();
                        // Delete any existing data from SQLite tables
                        string defaultEcon = Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                        string defaultVariables = Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName;

                        oDataMgr.m_strSQL = "DELETE FROM " + defaultEcon;
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                        oDataMgr.m_strSQL = "DELETE FROM " + defaultVariables;
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    }

                    strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                    using (var copyConn = new OleDbConnection(strCopyConn))
                    {
                        copyConn.Open();

                        oAdo.m_strSQL = "INSERT INTO " + targetEcon +
                                          " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                        oAdo.m_strSQL = "INSERT INTO " + targetFvs +
                          " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedFVSVariablesTableName;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                        oAdo.m_strSQL = "INSERT INTO " + targetVariables +
                            " SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName;
                        oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    }
                }

                    // MIGRATE SCENARIO_OPTIMIZER_RULES_DEFINITIONS DATABASE
                    string scenarioAccessFile = this.ReferenceProjectDirectory + "\\" +
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;

                    // Create SQLite copy of scenario_optimizer_rule_definitions database
                    string scenarioSqliteFile = frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory + "\\" +
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                    //@ToDo: Don't have code
                    frmMain.g_oFrmMain.frmProject.uc_project1.CreateOptimizerScenarioRuleDefinitionSqliteDbAndTables(scenarioSqliteFile);

                    // Check to see if the input SQLite DSN exists and if so, delete so we can add
                    if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName))
                    {
                        odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName);
                    }
                    odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, scenarioSqliteFile);

                // Set new temporary database
                strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
                    oDao.CreateMDB(strTempAccdb);

                    // Create table links for transferring data
                    string[] sourceTables = {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName,
                        Tables.Scenario.DefaultScenarioDatasourceTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioHarvestCostColumnsTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLandOwnerGroupsTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPlotFilterTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioPSitesTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioLastTieBreakRankTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOverallEffectiveTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesOptimizationTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioFvsVariablesTieBreakerTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterMiscTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCondFilterTableName,
                        Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioProcessorScenarioSelectTableName};
                    string[] targetTables = new string[15];
                    foreach (string sourceTableName in sourceTables)
                    {
                        targetTables[Array.IndexOf(sourceTables, sourceTableName)] = sourceTableName + "_1";
                    }
                    // Link to all tables in source database
                    oDao.CreateTableLinks(strTempAccdb, scenarioAccessFile);
                    foreach (string targetTableName in targetTables)
                    {
                        oDao.CreateSQLiteTableLink(strTempAccdb, sourceTables[Array.IndexOf(targetTables, targetTableName)], targetTableName,
                            ODBCMgr.DSN_KEYS.OptimizerRuleDefinitionsDsnName, scenarioSqliteFile);
                    }

                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(scenarioSqliteFile)))
                    {
                        conn.Open();
                        // Delete any existing data from SQLite tables
                        foreach (string sourceTableName in sourceTables)
                        {
                            oDataMgr.m_strSQL = "DELETE FROM " + sourceTableName;
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        }
                    }

                    strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                    using (var copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                    {
                        copyConn.Open();

                        foreach (string targetTableName in targetTables)
                        {
                            string sourceTableName = sourceTables[Array.IndexOf(targetTables, targetTableName)];
                            oAdo.m_strSQL = "INSERT INTO " + targetTableName +
                                " SELECT * FROM " + sourceTableName;
                            oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                        }
                    }

                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(scenarioSqliteFile)))
                    {
                        conn.Open();
                        oDataMgr.m_strSQL = "UPDATE scenario SET file = 'scenario_optimizer_rule_definitions.db'";
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    }

                // MIGRATE GIS DATA
                // Check if Processor parameters in SQLite
                string strTest = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
                if (!System.IO.File.Exists(strTest))
                {
                    System.Windows.Forms.MessageBox.Show("Processor parameters have not been migrated to SQLite. SQLite GIS data cannot be loaded!", "FIA Biosum");
                    return;
                }
                // Check if Optimizer parameters in SQLite
                strTest = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}";
                if (!System.IO.File.Exists(strTest))
                {
                    System.Windows.Forms.MessageBox.Show("Optimizer parameters have not been migrated to SQLite. SQLite GIS data cannot be loaded!", "FIA Biosum");
                    return;
                }

                string gisPathAndDbFile = this.ReferenceProjectDirectory + "\\" + Tables.TravelTime.DefaultTravelTimePathAndDbFile;
                if (!System.IO.File.Exists(gisPathAndDbFile))
                {
                    oDataMgr.CreateDbFile(gisPathAndDbFile);
                }
                // Create audit db
                string strAuditDBPath = $@"{frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory}\{Tables.TravelTime.DefaultGisAuditPathAndDbFile}";
                if (!System.IO.File.Exists(strAuditDBPath))
                {
                    oDataMgr.CreateDbFile(strAuditDBPath);
                }

                // Create target tables in new database
                strCopyConn = oDataMgr.GetConnectionString(gisPathAndDbFile);
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strCopyConn))
                {
                    conn.Open();
                    frmMain.g_oTables.m_oTravelTime.CreateSqliteProcessingSiteTable(oDataMgr, conn, Tables.TravelTime.DefaultProcessingSiteTableName);
                    frmMain.g_oTables.m_oTravelTime.CreateSqliteTravelTimeTable(oDataMgr, conn, Tables.TravelTime.DefaultTravelTimeTableName);
                }
                // Find path to existing tables
                oProjectDs = new Datasource();
                oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
                oProjectDs.m_strDataSourceTableName = "datasource";
                oProjectDs.m_strScenarioId = "";
                oProjectDs.LoadTableColumnNamesAndDataTypes = false;
                oProjectDs.LoadTableRecordCount = false;
                oProjectDs.populate_datasource_array();

                // Travel times
                int intTravelTable = oProjectDs.getTableNameRow(Datasource.TableTypes.TravelTimes);
                int intPSitesTable = oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
                string strDirectoryPath = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
                string strFileName = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
                //(‘F’ = FILE FOUND, ‘NF’ = NOT FOUND)
                string strTableName = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                string strTableStatus = oProjectDs.m_strDataSource[intTravelTable, FIA_Biosum_Manager.Datasource.TABLESTATUS].Trim();
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(strCopyConn))
                {
                    conn.Open();
                    string[] arrUpdateTableType = new string[2];
                    string[] arrUpdateTableName = new string[2];
                    if (oProjectDs.DataSourceTableExist(intTravelTable))
                    {
                        string strDbName = System.IO.Path.GetFileName(Tables.TravelTime.DefaultTravelTimePathAndDbFile);
                        string strNewDirectoryPath = System.IO.Path.GetDirectoryName(gisPathAndDbFile);
                        string strConn = oAdo.getMDBConnString(oProjectDs.getFullPathAndFile(Datasource.TableTypes.TravelTimes), "", "");
                        using (var pConn = new System.Data.OleDb.OleDbConnection(strConn))
                        {
                            pConn.Open();
                            oAdo.m_strSQL = $@"SELECT TRAVELTIME_ID, PSITE_ID, BIOSUM_PLOT_ID, COLLECTOR_ID,RAILHEAD_ID,
                                TRAVEL_MODE, ONE_WAY_HOURS, PLOT, STATECD FROM {strTableName}";
                            oAdo.CreateDataTable(pConn, oAdo.m_strSQL, strTableName, false);
                            using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(oAdo.m_strSQL, conn))
                            {
                                using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                {
                                    using (var transaction = conn.BeginTransaction())
                                    {
                                        da.InsertCommand = cb.GetInsertCommand();
                                        int rows = da.Update(oAdo.m_DataTable);
                                        transaction.Commit();
                                    }
                                }
                            }
                            oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.TravelTimes, strNewDirectoryPath, strDbName, strTableName);
                            arrUpdateTableType[0] = Datasource.TableTypes.TravelTimes;
                            arrUpdateTableName[0] = strTableName;
                            if (oProjectDs.DataSourceTableExist(intPSitesTable))
                            {
                                strTableName = oProjectDs.m_strDataSource[intPSitesTable, Datasource.TABLE].Trim();
                                oAdo.m_strSQL = $@"SELECT PSITE_ID,NAME,TRANCD,TRANCD_DEF,BIOCD,BIOCD_DEF,EXISTS_YN,LAT,LON,
                                            STATE,CITY,MILL_TYPE,COUNTY,STATUS,NOTES FROM {strTableName}";
                                oAdo.CreateDataTable(pConn, oAdo.m_strSQL, strTableName, false);
                                using (System.Data.SQLite.SQLiteDataAdapter da = new System.Data.SQLite.SQLiteDataAdapter(oAdo.m_strSQL, conn))
                                {
                                    using (System.Data.SQLite.SQLiteCommandBuilder cb = new System.Data.SQLite.SQLiteCommandBuilder(da))
                                    {
                                        using (var transaction = conn.BeginTransaction())
                                        {
                                            da.InsertCommand = cb.GetInsertCommand();
                                            int rows = da.Update(oAdo.m_DataTable);
                                            transaction.Commit();
                                        }
                                    }
                                }
                                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.ProcessingSites, strNewDirectoryPath, strDbName, strTableName);
                                arrUpdateTableType[1] = Datasource.TableTypes.ProcessingSites;
                                arrUpdateTableName[1] = strTableName;
                            }
                        }
                        strConn = oDataMgr.GetConnectionString($@"{this.ReferenceProjectDirectory}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}");
                        using (System.Data.SQLite.SQLiteConnection scenarioConn = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            scenarioConn.Open();
                            for (int i = 0; i < arrUpdateTableType.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(arrUpdateTableType[i]))
                                {
                                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                    sb.Append($@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET ");
                                    sb.Append($@"PATH = '{strNewDirectoryPath}', file='{strDbName}', table_name = '{arrUpdateTableName[i]}' ");
                                    sb.Append($@"WHERE TRIM(table_type) = '{arrUpdateTableType[i]}'");
                                    oDataMgr.SqlNonQuery(scenarioConn, sb.ToString());
                                }
                            }
                        }
                        strConn = oDataMgr.GetConnectionString($@"{frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}");
                        using (System.Data.SQLite.SQLiteConnection scenarioConn = new System.Data.SQLite.SQLiteConnection(strConn))
                        {
                            scenarioConn.Open();
                            for (int i = 0; i < arrUpdateTableType.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(arrUpdateTableType[i]))
                                {
                                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                                    sb.Append($@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET ");
                                    sb.Append($@"PATH = '{strNewDirectoryPath}', file='{strDbName}', table_name = '{arrUpdateTableName[i]}' ");
                                    sb.Append($@"WHERE TRIM(table_type) = '{arrUpdateTableType[i]}'");
                                    oDataMgr.SqlNonQuery(scenarioConn, sb.ToString());
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                // Clean-up
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorRuleDefinitionsDsnName);
                }
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.FvsMasterDbDsnName);
                }
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.OptimizerCalcVariableDsnName);
                }
            }


            if (oAdo != null && oAdo.m_OleDbConnection != null)
            {
                oAdo.CloseConnection(oAdo.m_OleDbConnection);
                oAdo = null;
            }

            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }
        }

        public void UpdateDatasources_5_11_1()
        {
            DataMgr oDataMgr = new DataMgr();
            Datasource oProjectDs = new Datasource();

            // Find path to existing tables
            oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();

            // gis_travel_times.processing_site
            int intPSitesTable = oProjectDs.getTableNameRow(Datasource.TableTypes.ProcessingSites);
            string strDirectoryPath = oProjectDs.m_strDataSource[intPSitesTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oProjectDs.m_strDataSource[intPSitesTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            string strTableName = oProjectDs.m_strDataSource[intPSitesTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            
            if (oProjectDs.DataSourceTableExist(intPSitesTable))
            {
                string strGisConn = oDataMgr.GetConnectionString(strDirectoryPath + "\\" + strFileName);
                using (System.Data.SQLite.SQLiteConnection gisConn = new System.Data.SQLite.SQLiteConnection(strGisConn))
                {
                    gisConn.Open();

                    if (!oDataMgr.FieldExist(gisConn, "SELECT * FROM " + strTableName, "PSITE_CN"))
                    {
                        string strSQL = "ALTER TABLE " + strTableName + " ADD COLUMN PSITE_CN CHAR(12)";
                        oDataMgr.SqlNonQuery(gisConn, strSQL);
                    }
                }
            }

            oDataMgr = null;
            oProjectDs = null;
        }

        public void UpdateDatasources_5_11_2()
        {
            DataMgr oDataMgr = new DataMgr();
            ODBCMgr odbcmgr = new ODBCMgr();
            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();
            utils oUtils = new utils();
            //
            // Create master_aux.db and migrate values from master_aux.accdb
            //
            frmMain.g_sbpInfo.Text = "Version Update: Migrating DWM tables ...Stand by";

            string strSourceFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDbFile;
            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultDWMSqliteDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                oDataMgr.CreateDbFile(strDestFile);
            }
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();

                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMCoarseWoodyDebrisTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMDuffLitterFuelTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMFineWoodyDebrisTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteDWMTransectSegmentTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
                }
                else
                {
                    oDataMgr.SqlNonQuery(conn, "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName);
                }
            }

            // create DSN if needed
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterAuxDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName, strSourceFile);

            // Set new temporary database
            string strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
            oDao.CreateMDB(strTempAccdb);

            //link access tables to temporary database
            oDao.CreateTableLinks(strTempAccdb, strSourceFile);

            //link sqlite tables to temporary database
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);
            oDao.CreateSQLiteTableLink(strTempAccdb, frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName,
                frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName + "_1", ODBCMgr.DSN_KEYS.MasterAuxDsnName, strDestFile);

            //insert data into sqlite tables
            string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
            using (System.Data.OleDb.OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
            {
                copyConn.Open();

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMCoarseWoodyDebrisName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMDuffLitterFuelName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMFineWoodyDebrisName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);

                oAdo.m_strSQL = "INSERT INTO " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName +
                    "_1 SELECT * FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultDWMTransectSegmentName;
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
            }

            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterAuxDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterAuxDsnName);
            }

            // Move sequence number tables from fvs_master.db to master.db
            frmMain.g_sbpInfo.Text = "Version Update: Move sequence number tables ...Stand by";
            strDestFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.FVS.DefaultRxPackageDbFile;
            if (! System.IO.File.Exists(strDestFile))
            {
                oDataMgr.CreateDbFile(strDestFile);
            }
            // Create sequence number tables if they don't exist
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumTable))
                {
                    frmMain.g_oTables.m_oFvs.CreateFVSOutputSQLitePrePostSeqNumTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultFVSPrePostSeqNumTable}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable))
                {
                    frmMain.g_oTables.m_oFvs.CreateFVSOutputPrePostSQLiteSeqNumRxPackageAssgnTable(oDataMgr, conn, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultRxPackageTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateSQLiteRxPackageTable(oDataMgr, conn, Tables.FVS.DefaultRxPackageTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultRxPackageTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultRxTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateSQLiteRxTable(oDataMgr, conn, Tables.FVS.DefaultRxTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultRxTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName))
                {
                    frmMain.g_oTables.m_oFvs.CreateSqliteRxHarvestCostColumnTable(oDataMgr, conn, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = $@"DELETE FROM {Tables.FVS.DefaultRxHarvestCostColumnsTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }

            Datasource oProjectDs = new Datasource();
            // Find path to existing tables
            oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            // FVS PRE-POST SeqNum Definitions. Assuming that all the sequence number tables will be in the same db
            int intSeqNumTable = oProjectDs.getTableNameRow(Datasource.TableTypes.SeqNumDefinitions);
            // Again, assuming that the rx tables are all in the same database
            int intRxPackageTable = oProjectDs.getTableNameRow(Datasource.TableTypes.RxPackage);

            // create DSN if needed
            frmMain.g_sbpInfo.Text = "Version Update: Migrate tables from fvs_master.mdb ...Stand by";
            string fvsMasterDs = "MIGRATE_FVS_MASTER";
            if (odbcmgr.CurrentUserDSNKeyExist(fvsMasterDs))
            {
                odbcmgr.RemoveUserDSN(fvsMasterDs);
            }
            odbcmgr.CreateUserSQLiteDSN(fvsMasterDs, strDestFile);
            string[] arrTargetTables = { Tables.FVS.DefaultRxPackageTableName, Tables.FVS.DefaultRxTableName, Tables.FVS.DefaultRxHarvestCostColumnsTableName };
            strSourceFile = $@"{oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.PATH].Trim()}\{oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim()}";
            for (int i = 0; i < arrTargetTables.Length; i++)
            {
                oDao.CreateSQLiteTableLink(strSourceFile, arrTargetTables[i],
                    arrTargetTables[i] + "_1", fvsMasterDs, strDestFile);
            }
            System.Threading.Thread.Sleep(4000);

            string strDirectoryPath = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            string strTableName = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();            
            if (oProjectDs.DataSourceTableExist(intSeqNumTable))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strDirectoryPath}\{strFileName}' AS source";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = $@"INSERT INTO {Tables.FVS.DefaultFVSPrePostSeqNumTable} SELECT * FROM source.{strTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    intSeqNumTable = oProjectDs.getTableNameRow(Datasource.TableTypes.SeqNumRxPackageAssign);
                    strTableName = oProjectDs.m_strDataSource[intSeqNumTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
                    oDataMgr.m_strSQL = $@"INSERT INTO {Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable} SELECT * FROM source.{strTableName}";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                // Update entries in project data sources table
                string strMasterDb = System.IO.Path.GetFileName(Tables.FVS.DefaultRxPackageDbFile);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumDefinitions, $@"{ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumTable);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SeqNumRxPackageAssign, $@"{ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultFVSPrePostSeqNumRxPackageAssgnTable);
            }

            strDirectoryPath = oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            strTableName = oProjectDs.m_strDataSource[intRxPackageTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();

            if (oProjectDs.DataSourceTableExist(intRxPackageTable))
            {
                strCopyConn = oAdo.getMDBConnString(strSourceFile, "", "");
                using (OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    oAdo.m_strSQL = $@"INSERT INTO {arrTargetTables[0]}_1 SELECT RXPACKAGE, DESCRIPTION, rxcycle_length,
                        simyear1_rx, simyear1_fvscycle, simyear2_rx, simyear2_fvscycle, simyear3_rx, simyear3_fvscycle,
                        simyear4_rx, simyear4_fvscycle FROM {strTableName}";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"DROP TABLE {arrTargetTables[0]}_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"INSERT INTO {arrTargetTables[1]}_1 SELECT RX, DESCRIPTION, HarvestMethodLowSlope, HarvestMethodSteepSlope
                        FROM {arrTargetTables[1]}";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"DROP TABLE {arrTargetTables[1]}_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[2] + "_1" +
                        " SELECT * FROM " + arrTargetTables[2];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = $@"DROP TABLE {arrTargetTables[2]}_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                }

                string strMasterDb = System.IO.Path.GetFileName(Tables.FVS.DefaultRxPackageDbFile);
                // Update project data sources
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.RxPackage, $@"{ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultRxPackageTableName);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Rx, $@"{ ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultRxTableName);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.RxHarvestCostColumns, $@"{ ReferenceProjectDirectory.Trim()}\db", strMasterDb, Tables.FVS.DefaultRxHarvestCostColumnsTableName);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.HarvestMethods, "@@appdata@@\\fiabiosum", Tables.Reference.DefaultBiosumReferenceSqliteFile, Tables.Reference.DefaultHarvestMethodsTableName);

                // Update Optimizer data sources
                strDestFile = $@"{ReferenceProjectDirectory.Trim()}\{Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile}";
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET file = '{strMasterDb}' 
                        where table_type in ('Treatment Prescriptions','Treatment Prescriptions Harvest Cost Columns','{Datasource.TableTypes.RxPackage}')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioDatasourceTableName} SET path = '@@appdata@@\fiabiosum', file = '{Tables.Reference.DefaultBiosumReferenceSqliteFile}' 
                        where table_type = '{Datasource.TableTypes.HarvestMethods}' ";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                // Update Processor data sources
                strDestFile = $@"{ReferenceProjectDirectory.Trim()}\processor\{Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile}";
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.Scenario.DefaultScenarioDatasourceTableName} SET file = '{strMasterDb}' 
                        where table_type in ('Treatment Prescriptions','Treatment Prescriptions Harvest Cost Columns','{Datasource.TableTypes.RxPackage}')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = $@"UPDATE {Tables.Scenario.DefaultScenarioDatasourceTableName} SET path = '@@appdata@@\fiabiosum', file = '{Tables.Reference.DefaultBiosumReferenceSqliteFile}' 
                        where table_type = '{Datasource.TableTypes.HarvestMethods}' ";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                // Remove obsolete data source definitions
                using (OleDbConnection deleteConn = new System.Data.OleDb.OleDbConnection(oAdo.getMDBConnString(oProjectDs.m_strDataSourceMDBFile, "", "")))
                {
                    deleteConn.Open();
                    oAdo.m_strSQL = $@"DELETE FROM {oProjectDs.m_strDataSourceTableName} WHERE TABLE_TYPE IN 
                        ('Treatment Prescriptions Assigned FVS Commands', 'Treatment Prescription Categories', 'Treatment Prescription Subcategories',
                         'Treatment Package Assigned FVS Commands', 'Treatment Package FVS Commands Order', 'Treatment Package Members', 
                         'FVS Western Tree Species Translator', 'FVS Eastern Tree Species Translator')";
                    oAdo.SqlNonQuery(deleteConn, oAdo.m_strSQL);
                }
            }
            if (odbcmgr.CurrentUserDSNKeyExist(fvsMasterDs))
            {
                odbcmgr.RemoveUserDSN(fvsMasterDs);
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }

            // Migrate any existing variant specific FVSIn.db files to one combined FVSIn.db
            frmMain.g_sbpInfo.Text = "Version Update: Migrate FVSIn.db files from [variant]/FVSIn.db to combined FVSIn.db ...Stand by";
            string strFVSInFile = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\" + Tables.FIA2FVS.DefaultFvsInputFile;

            // Get list of FVS Variants from Plot table
            List<String> lstVar = new List<String>();
            int intPlotTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Plot);
            strDirectoryPath = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            strFileName = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            strTableName = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.TABLE].Trim();
            string strVariantCon = oAdo.getMDBConnString(strDirectoryPath + "\\" + strFileName, "", "");
            using (OleDbConnection variantCon = new OleDbConnection(strVariantCon))
            {
                variantCon.Open();

                oAdo.m_strSQL = Queries.FVS.GetFVSVariantSQL_access(strTableName);
                oAdo.SqlQueryReader(variantCon, oAdo.m_strSQL);

                if (oAdo.m_OleDbDataReader.HasRows)
                {
                    while (oAdo.m_OleDbDataReader.Read())
                    {
                        string strVariant = oAdo.m_OleDbDataReader["fvs_variant"].ToString().Trim();
                        lstVar.Add(strVariant);
                    }
                }
            }

            // Cycle through list. If variant FVSIn exists, add it to the combined FVSIn.db. Create the combined FVSIn.db if needed
            string strFVSConn = oDataMgr.GetConnectionString(strFVSInFile);
            foreach (string variant in lstVar)
            {
                string strVarFvsIn = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\" + variant + "\\" + Tables.FIA2FVS.DefaultFvsInputFile;
                if (System.IO.File.Exists(strVarFvsIn))
                {
                    if (!System.IO.File.Exists(strFVSInFile))
                    {
                        System.IO.File.Copy(strVarFvsIn, strFVSInFile, true);

                        using (System.Data.SQLite.SQLiteConnection fvsConn = new System.Data.SQLite.SQLiteConnection(strFVSConn))
                        {
                            fvsConn.Open();

                            if (!oDataMgr.ColumnExist(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT"))
                            {
                                oDataMgr.AddColumn(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT", "CHAR", "2");
                            }

                            oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                " SET VARIANT = '" + variant + "' WHERE VARIANT IS NULL";
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);
                        }
                    }
                    else
                    {
                        using (System.Data.SQLite.SQLiteConnection fvsConn = new System.Data.SQLite.SQLiteConnection(strFVSConn))
                        {
                            fvsConn.Open();

                            oDataMgr.m_strSQL = "ATTACH DATABASE '" + strVarFvsIn + "' AS variant";
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);

                            string strStandFields = oDataMgr.getFieldNames(fvsConn, "SELECT * FROM variant." + Tables.FIA2FVS.DefaultFvsInputStandTableName);
                            oDataMgr.m_strSQL = "INSERT INTO " + Tables.FIA2FVS.DefaultFvsInputStandTableName +
                                " (" + strStandFields + ") SELECT " + strStandFields +
                                " FROM variant." + Tables.FIA2FVS.DefaultFvsInputStandTableName;
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);

                            string strTreeFields = oDataMgr.getFieldNames(fvsConn, "SELECT * FROM variant." + Tables.FIA2FVS.DefaultFvsInputTreeTableName);
                            oDataMgr.m_strSQL = "INSERT INTO " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                " (" + strTreeFields + ") SELECT " + strTreeFields +
                                " FROM variant." + Tables.FIA2FVS.DefaultFvsInputTreeTableName;
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);

                            if (!oDataMgr.ColumnExist(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT"))
                            {
                                oDataMgr.AddColumn(fvsConn, Tables.FIA2FVS.DefaultFvsInputTreeTableName, "VARIANT", "CHAR", "2");
                            }

                            oDataMgr.m_strSQL = "UPDATE " + Tables.FIA2FVS.DefaultFvsInputTreeTableName +
                                " SET VARIANT = '" + variant + "' WHERE VARIANT IS NULL";
                            oDataMgr.SqlNonQuery(fvsConn, oDataMgr.m_strSQL);
                        }
                    }
                }
            }

            // update optimizer calculated variables db to add null threshold table
            // and use negative column to variables table
            frmMain.g_sbpInfo.Text = "Version Update: Updated Optimizer Calculated Variables ...Stand by";
            string strCalculatedVariablesDb = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strCalculatedVariablesDb)))
            {
                conn.Open();
                if (!oDataMgr.ColumnExist(conn, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, "NEGATIVES_YN"))
                {
                    oDataMgr.AddColumn(conn, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, "NEGATIVES_YN", "CHAR", "1");

                    oDataMgr.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " SET NEGATIVES_YN = 'N'";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName))
                {
                    oDataMgr.m_strSQL = "CREATE TABLE " + Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName + " (fvs_null_threshold INTEGER)";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultOptimizerProjectConfigTableName + " VALUES (4)";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }
        }

        public void UpdateDatasources_5_12_0()
        {
            DataMgr oDataMgr = new DataMgr();
            ODBCMgr odbcmgr = new ODBCMgr();
            dao_data_access oDao = new dao_data_access();
            ado_data_access oAdo = new ado_data_access();
            utils oUtils = new utils();

            // Migrate plot, cond, tree, and sitetree tables from master.mdb to master.db
            frmMain.g_sbpInfo.Text = "Version Update: Moving plot, cond, tree, and sitetree tables ...Stand by";

            string strDestFile = ReferenceProjectDirectory.Trim() + "\\" + frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableSqliteDbFile;
            if (!System.IO.File.Exists(strDestFile))
            {
                oDataMgr.CreateDbFile(strDestFile);
            }

            // Create tables if they don't exist
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();

                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqlitePlotTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteConditionTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteTreeTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
                if (!oDataMgr.TableExist(conn, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName))
                {
                    frmMain.g_oTables.m_oFIAPlot.CreateSqliteSiteTreeTable(oDataMgr, conn, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName);
                }
                else
                {
                    oDataMgr.m_strSQL = "DELETE FROM " + frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }

            Datasource oProjectDs = new Datasource();
            // Find path to existing tables
            oProjectDs.m_strDataSourceMDBFile = this.ReferenceProjectDirectory + "\\db\\project.mdb";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            // plot
            int intPlotTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Plot);
            // cond
            int intCondTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Condition);
            // tree
            int intTreeTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Tree);
            // sitetree
            int intSiteTreeTable = oProjectDs.getTableNameRow(Datasource.TableTypes.SiteTree);

            // Create DSN if needed
            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterDsnName);
            }
            odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.MasterDsnName, strDestFile);

            string[] arrTargetTables = { frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName, frmMain.g_oTables.m_oFIAPlot.DefaultConditionTableName,
                frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName, frmMain.g_oTables.m_oFIAPlot.DefaultSiteTreeTableName };
            string strSourceFile = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim() + "\\" + oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();
            string[] strSourceTables = { oProjectDs.m_strDataSource[intPlotTable, Datasource.TABLE].Trim(), oProjectDs.m_strDataSource[intCondTable, Datasource.TABLE].Trim(),
                oProjectDs.m_strDataSource[intTreeTable, Datasource.TABLE].Trim(), oProjectDs.m_strDataSource[intSiteTreeTable, Datasource.TABLE].Trim() };
            string strMasterDb = System.IO.Path.GetFileName(frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableSqliteDbFile);

            // Set new temporary database
            string strTempAccdb = oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
            oDao.CreateMDB(strTempAccdb);

            //link access tables to temporary database
            foreach (string table in strSourceTables)
            {
                oDao.CreateTableLink(strTempAccdb, table, strSourceFile, table);
            }

            for (int i = 0; i < arrTargetTables.Length; i++)
            {
                oDao.CreateSQLiteTableLink(strTempAccdb, arrTargetTables[i],
                    arrTargetTables[i] + "_1", ODBCMgr.DSN_KEYS.MasterDsnName, strDestFile);
            }
            System.Threading.Thread.Sleep(4000);

            if (oProjectDs.DataSourceTableExist(intPlotTable))
            {
                string strCopyConn = oAdo.getMDBConnString(strTempAccdb, "", "");
                using (OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strCopyConn))
                {
                    copyConn.Open();
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[0] + "_1 " +
                        "SELECT biosum_plot_id, statecd, invyr, unitcd, countycd, " + arrTargetTables[0] + ", " +
                        "measyear, elev, fvs_variant, fvsloccode AS fvs_loc_cd, half_state, gis_yard_dist_ft, " +
                        "num_cond, one_cond_yn, lat, lon, macro_breakpoint_dia, precipitation, " +
                        "ecosubcd, biosum_status_cd, cn FROM " + arrTargetTables[0];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[1] + "_1 " +
                        "SELECT biosum_cond_id, biosum_plot_id, invyr, condid, condprop, cond_status_cd, " +
                        "fortypcd, owncd, owngrpcd, reservcd, siteclcd, sibase, sicond, sisp, " +
                        "slope, aspect, stdage, stdszcd, habtypcd1, adforcd, qmd_all_inch, " +
                        "qmd_hwd_inch, qmd_swd_inch, acres, unitcd, vol_loc_grp, tpacurr, " +
                        "hwd_tpacurr, swd_tpacurr, ba_ft2_ac, hwd_ba_ft2_ac, swd_ba_ft2_ac, " +
                        "vol_ac_grs_ft3, hwd_vol_ac_grs_ft3, swd_vol_ac_grs_ft3, " +
                        "volcsgrs, hwd_volcsgrs, swd_volcsgrs, gsstkcd, alstkcd, " +
                        "condprop_unadj, micrprop_unadj, subpprop_unadj, macrprop_unadj, " +
                        "cn, biosum_status_cd, dwm_fuelbed_typcd, balive, stdorgcd FROM " + arrTargetTables[1];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[2] + "_1 " +
                        "SELECT biosum_cond_id, invyr, statecd, unitcd, countycd, subp, " +
                        "tree, condid, statuscd, spcd, spgrpcd, dia, diahtcd, ht, htcd, actualht, " +
                        "formcl, treeclcd, cr, cclcd, cull, roughcull, decaycd, stocking, tpacurr, wdldstem, " +
                        "volcfnet, volcfgrs, volcsnet, volcsgrs, volbfnet, volbfgrs, voltsgrs, " +
                        "drybiot, drybiom, bhage, cullbf, cullcf, totage, mist_cl_cd, agentcd, " +
                        "damtyp1, damsev1, damtyp2, damsev2, tpa_unadj, condprop_specific, sitree, " +
                        "upper_dia, upper_dia_ht, centroid_dia, centroid_dia_ht_actual, sawht, htdmp, " +
                        "boleht, cull_fld, culldead, cullform, cullmstop, cfsnd, bfsnd, standing_dead_cd, " +
                        "volcfsnd, drybio_bole, drybio_top, drybio_sapling, drybio_wdld_spp, " +
                        "fvs_tree_id, cn, biosum_status_cd FROM " + arrTargetTables[2];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "INSERT INTO " + arrTargetTables[3] + "_1 " +
                        "SELECT biosum_plot_id, invyr, condid, tree, spcd, dia, ht, agedia, " +
                        "spgrpcd, sitree, sibase, subp, method, validcd, condlist, biosum_status_cd FROM " + arrTargetTables[3];
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[0] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[1] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[2] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                    oAdo.m_strSQL = "DROP TABLE " + arrTargetTables[3] + "_1";
                    oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
                }

                // Update project datasources
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Plot, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[0]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Condition, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[1]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.Tree, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[2]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.SiteTree, ReferenceProjectDirectory + "\\db", strMasterDb, arrTargetTables[3]);
                oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.PopStratumAdjFactors, ReferenceProjectDirectory + "\\db", strMasterDb,
                    frmMain.g_oTables.m_oFIAPlot.DefaultBiosumPopStratumAdjustmentFactorsTableName);

                // Update processor datasources
                strDestFile = ReferenceProjectDirectory.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();

                    oDataMgr.m_strSQL = "UPDATE " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                        " SET file = '" + strMasterDb + "' WHERE table_type IN ('" +
                        Datasource.TableTypes.Plot + "', '" +
                        Datasource.TableTypes.Condition + "', '" +
                        Datasource.TableTypes.Tree + "')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                // Update optimizer datasources
                strDestFile = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableSqliteDbFile;
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();

                    oDataMgr.m_strSQL = "UPDATE " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                        " SET file = '" + strMasterDb + "' WHERE table_type IN ('" +
                        Datasource.TableTypes.Plot + "', '" +
                        Datasource.TableTypes.Condition + "')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }

            // Update project datasources; tree_species, fvs_tree_species, and fiadb_fvs_variant have been eliminated; fia_tree_species_ref has moved
            string strDsConn = oAdo.getMDBConnString(oProjectDs.m_strDataSourceMDBFile, "", "");
            using (OleDbConnection copyConn = new System.Data.OleDb.OleDbConnection(strDsConn))
            {
                copyConn.Open();
                oAdo.m_strSQL = $@"DELETE FROM {oProjectDs.m_strDataSourceTableName} WHERE TABLE_TYPE IN ('{Datasource.TableTypes.TreeSpecies}','{Datasource.TableTypes.FvsTreeSpecies}', '{Datasource.TableTypes.FVSVariant}',
                    '{Datasource.TableTypes.OwnerGroups}')";
                oAdo.SqlNonQuery(copyConn, oAdo.m_strSQL);
            }

            // Update project datasources
            oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.FiaTreeMacroPlotBreakpointDia, "@@appdata@@\\fiabiosum",
                Tables.Reference.DefaultTreeMacroPlotBreakPointDiaTableDbFile, Tables.Reference.DefaultTreeMacroPlotBreakPointDiaTableName);
            oProjectDs.UpdateDataSourcePath(Datasource.TableTypes.FiaTreeSpeciesReference, "@@appdata@@\\fiabiosum",
                Tables.Reference.DefaultBiosumReferenceSqliteFile, Tables.Reference.DefaultFIATreeSpeciesTableName);

            // Update processor datasource
            strDestFile = ReferenceProjectDirectory.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
            {
                conn.Open();
                oDataMgr.m_strSQL = "DELETE FROM " + Tables.Scenario.DefaultScenarioDatasourceTableName +
                    " WHERE table_type in ('" + Datasource.TableTypes.TreeSpecies + "','" + Datasource.TableTypes.ProcessingSites + "')";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                oDataMgr.m_strSQL = $@"UPDATE {Tables.Scenario.DefaultScenarioDatasourceTableName} SET FILE = '{Tables.Reference.DefaultBiosumReferenceSqliteFile}' 
                    WHERE table_type = '{Datasource.TableTypes.FiaTreeSpeciesReference}'";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
            }

            if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.MasterDsnName))
            {
                odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.MasterDsnName);
            }

            // Add BioSum generated site index flag column to FVS_STANDINIT_COND table in FVSIn.db
            frmMain.g_sbpInfo.Text = "Version Update: Updating FVSIn.db tables ...Stand by";
            string strFVSInFile = ReferenceProjectDirectory.Trim() + "\\fvs\\data\\" + Tables.FIA2FVS.DefaultFvsInputFile;
            if (System.IO.File.Exists(strFVSInFile))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFVSInFile)))
                {
                    conn.Open();

                    if (!oDataMgr.ColumnExist(conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, "SITE_INDEX_BSCALC_YN"))
                    {
                        oDataMgr.AddColumn(conn, Tables.FIA2FVS.DefaultFvsInputStandTableName, "SITE_INDEX_BSCALC_YN", "CHAR", "1");
                    }
                }
            }
            if (oDao != null)
            {
                oDao.m_DaoWorkspace.Close();
                oDao = null;
            }

            // Copy ref_master.db to project if it doesn't already exist
            if (!System.IO.File.Exists(ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultRefMasterDbFile))
            {
                System.IO.File.Copy($@"{frmMain.g_oEnv.strAppDir}\{Tables.Reference.DefaultRefMasterDbFile}", ReferenceProjectDirectory.Trim() + "\\" + Tables.Reference.DefaultRefMasterDbFile, true);
            }

            // update calculated variables for keep, null, and zero options for negative values
            frmMain.g_sbpInfo.Text = "Version Update: Updating Optimizer Calculated Variables ...Stand by";
            string strCalculatedVariablesDb = ReferenceProjectDirectory.Trim() + "\\" + Tables.OptimizerDefinitions.DefaultSqliteDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strCalculatedVariablesDb)))
            {
                conn.Open();

                if (!oDataMgr.ColumnExist(conn, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, "HANDLE_NEGATIVES"))
                {
                    oDataMgr.AddColumn(conn, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, "HANDLE_NEGATIVES", "CHAR", "4");

                    oDataMgr.m_strSQL = "UPDATE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " SET HANDLE_NEGATIVES = CASE WHEN NEGATIVES_YN = 'Y' THEN 'keep' ELSE 'omit' END";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                oDataMgr.m_strSQL = "ALTER TABLE " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " DROP COLUMN NEGATIVES_YN";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
            }
        }

        public void UpdateDatasources_5_12_1()
        {
            DataMgr oDataMgr = new DataMgr();
            // Make sure they have the treesample.db with the treeSampleTvbc table in it
            //@ToDo: Need to test this
            bool bCopyTreeSample = false;
            string strDestFile = frmMain.g_oEnv.strApplicationDataDirectory.Trim() + frmMain.g_strBiosumDataDir + Tables.Reference.DefaultTreeSampleDbFile;
            string strSourceFile = frmMain.g_oEnv.strAppDir + "\\db" + Tables.Reference.DefaultTreeSampleDbFile;
            if (System.IO.File.Exists(strDestFile))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strDestFile)))
                {
                    conn.Open();
                    if (! oDataMgr.TableExist(conn, Tables.Reference.DefaultTreeSampleTvbcTableName))
                    {
                        bCopyTreeSample = true;
                    }
                }
            }
            else
            {
                bCopyTreeSample = true;
            }
            if (bCopyTreeSample)
            {
                // Copy it the database from the app install directory
                System.IO.File.Copy(strSourceFile, strDestFile, true);
            }

        }

            // Method to compare two versions.
            // Returns 1 if v2 is smaller, -1 
            // if v1 is smaller, 0 if equal 
            public int VersionCompare(string v1, string v2)
        {
            // vnum stores each numeric 
            // part of version 
            int vnum1 = 0, vnum2 = 0;

            // loop until both string are 
            // processed 
            for (int i = 0, j = 0; (i < v1.Length
                                    || j < v2.Length);)
            {
                // storing numeric part of 
                // version 1 in vnum1 
                while (i < v1.Length && v1[i] != '.')
                {
                    vnum1 = vnum1 * 10 + (v1[i] - '0');
                    i++;
                }

                // storing numeric part of 
                // version 2 in vnum2 
                while (j < v2.Length && v2[j] != '.')
                {
                    vnum2 = vnum2 * 10 + (v2[j] - '0');
                    j++;
                }

                if (vnum1 > vnum2)
                    return 1;
                if (vnum2 > vnum1)
                    return -1;

                // if equal, reset variables and 
                // go for next numeric part 
                vnum1 = vnum2 = 0;
                i++;
                j++;
            }
            return 0;
        }

        public string ReferenceProjectDirectory
		{
			get {return _strProjDir;}
			set {_strProjDir=value;}
		}
		
	}
}

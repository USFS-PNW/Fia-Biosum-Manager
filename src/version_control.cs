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
                    // Upgraded from 5.12.1 to 5.13.0 (completely rid of Access)
                    if ((Convert.ToInt16(m_strAppVerArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR1]) == 13 &&
                        Convert.ToInt16(m_strAppVerArray[APP_VERSION_MINOR2]) == 0) &&
                        (Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MAJOR]) == 5 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR1]) == 12 &&
                        Convert.ToInt16(m_strProjectVersionArray[APP_VERSION_MINOR2]) == 1))
                    {
                        UpdateDatasources_5_13_0();
                        UpdateProjectVersionFile(strProjVersionFile);
                        bPerformCheck = false;
                    }
                }
            }

            //UpdateDatasources_5_13_1(this.ReferenceProjectDirectory);

            //UpdateDatasources_5_13_1(frmMain.g_oFrmMain.frmProject.uc_project1.m_strProjectDirectory);
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

        public void UpdateDatasources_5_13_0()
        {
            // Remove precipitation column from plot table, remove drybiot, drybiom, drybio_top,
            // drybio_sapling, drybio_wdld_spp from tree table, add drybio_ag to tree table

            DataMgr oDataMgr = new DataMgr();
            frmMain.g_sbpInfo.Text = "Version Update: Updating plot and tree tables ...Stand by";

            Datasource oProjectDs = new Datasource();
            // Find path to existing tables
            oProjectDs.m_strDataSourceDBFile = this.ReferenceProjectDirectory + "\\db\\project.db";
            oProjectDs.m_strDataSourceTableName = "datasource";
            oProjectDs.m_strScenarioId = "";
            oProjectDs.LoadTableColumnNamesAndDataTypes = false;
            oProjectDs.LoadTableRecordCount = false;
            oProjectDs.populate_datasource_array();
            // Assuming the plot and tree tables are in the same db
            int intPlotTable = oProjectDs.getTableNameRow(Datasource.TableTypes.Plot);

            string strDirectoryPath = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.PATH].Trim();
            string strFileName = oProjectDs.m_strDataSource[intPlotTable, FIA_Biosum_Manager.Datasource.DBFILE].Trim();

            string strMasterDb = strDirectoryPath + "\\" + strFileName;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strMasterDb)))
            {
                conn.Open();

                string strTreeTable = frmMain.g_oTables.m_oFIAPlot.DefaultTreeTableName;
                string strPlotTable = frmMain.g_oTables.m_oFIAPlot.DefaultPlotTableName;

                if (oDataMgr.ColumnExist(conn, strPlotTable, "precipitation"))
                {
                    oDataMgr.m_strSQL = "ALTER TABLE " + strPlotTable + " DROP COLUMN precipitation";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                if (oDataMgr.ColumnExist(conn, strTreeTable, "drybiot"))
                {
                    oDataMgr.m_strSQL = "ALTER TABLE " + strTreeTable + " DROP COLUMN drybiot";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                if (oDataMgr.ColumnExist(conn, strTreeTable, "drybiom"))
                {
                    oDataMgr.m_strSQL = "ALTER TABLE " + strTreeTable + " DROP COLUMN drybiom";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                if (oDataMgr.ColumnExist(conn, strTreeTable, "drybio_top"))
                {
                    oDataMgr.m_strSQL = "ALTER TABLE " + strTreeTable + " DROP COLUMN drybio_top";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                if (oDataMgr.ColumnExist(conn, strTreeTable, "drybio_sapling"))
                {
                    oDataMgr.m_strSQL = "ALTER TABLE " + strTreeTable + " DROP COLUMN drybio_sapling";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                if (oDataMgr.ColumnExist(conn, strTreeTable, "drybio_wdld_spp"))
                {
                    oDataMgr.m_strSQL = "ALTER TABLE " + strTreeTable + " DROP COLUMN drybio_wdld_spp";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }

                if (!oDataMgr.ColumnExist(conn, strTreeTable, "drybio_ag"))
                {
                    oDataMgr.AddColumn(conn, strTreeTable, "drybio_ag", "DOUBLE", null);
                }
            }

            string strFvsOutDb = ReferenceProjectDirectory.Trim() + Tables.FVS.DefaultFVSTreeListDbFile;
            if (System.IO.File.Exists(strFvsOutDb))
            {
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strFvsOutDb)))
                {
                    conn.Open();
                    if (oDataMgr.TableExist(conn, Tables.FVS.DefaultFVSCutTreeTvbcTableName))
                    {
                        if (!oDataMgr.FieldExist(conn, $@"select * from { Tables.FVS.DefaultFVSCutTreeTvbcTableName} limit 1", "WOODLAND_YN"))
                        {
                            // The table will be recreated when FVSOut runs for the first time
                            oDataMgr.m_strSQL = $@"DROP TABLE {Tables.FVS.DefaultFVSCutTreeTvbcTableName}";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                        }
                    }
                }
            }

            // delete any existing weighed fvs variable tables
            frmMain.g_sbpInfo.Text = "Version Update: Deleting weighted FVS variable tables ...Stand by";

            string strWeightedFVSVariablesDb = this.ReferenceProjectDirectory + "\\" + Tables.OptimizerScenarioResults.DefaultCalculatedPrePostFVSVariableTableDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strWeightedFVSVariablesDb)))
            {
                conn.Open();

                string[] arrTableNames = oDataMgr.getTableNames(conn);
                foreach (string strTable in arrTableNames)
                {
                    oDataMgr.m_strSQL = "DROP TABLE " + strTable;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }
        }
        public void UpdateDatasources_5_13_1(string strProjectDirectory)
        {
            //@ToDo: Remove below when ready to release
            ReferenceProjectDirectory = strProjectDirectory;
            DataMgr oDataMgr = new DataMgr();
            frmMain.g_sbpInfo.Text = "Version Update: Updating Processor Rule Definitions for Carbon Metrics ...Stand by";

            // Create temporary database to assemble new table
            string strTempDbFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "db");
            string strTempTable = "tmpWoodValue";
            string strProcessorDb = ReferenceProjectDirectory.Trim() + "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultDbFile;
            oDataMgr.CreateDbFile(strTempDbFile);
            long recordCount = 0;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strTempDbFile)))
            {
                conn.Open();
                frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesDollarValuesTable(oDataMgr, conn, strTempTable);
                oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strProcessorDb}' AS SOURCE";
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                oDataMgr.m_strSQL = "INSERT INTO tmpWoodValue(scenario_id,species_group,diam_group,merch_value, chip_value,merch_pct,chip_pct)" +
                    " SELECT scenario_id, species_group, diam_group, merch_value, chip_value," +
                    " CASE WHEN wood_bin = 'M' THEN 100 ELSE 0 END," +
                    " CASE WHEN wood_bin = 'C' THEN 100 ELSE 0 END" +
                    " FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName;
                oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                recordCount = oDataMgr.getRecordCount(conn, "SELECT count(distinct species_group) FROM " + strTempTable, strTempTable);
            }
            if (recordCount > 0)
            {
                // This means we've moved the data to the new table and can drop and re-add the existing one
                using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strProcessorDb)))
                {
                    conn.Open();
                    oDataMgr.SqlNonQuery(conn, $@"DROP TABLE {Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName}");
                    frmMain.g_oTables.m_oProcessorScenarioRuleDefinitions.CreateScenarioTreeSpeciesDollarValuesTable(oDataMgr, conn, 
                        Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName);
                    oDataMgr.m_strSQL = $@"ATTACH DATABASE '{strTempDbFile}' AS SOURCE";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName +
                        "(scenario_id,species_group,diam_group,merch_pct,merch_value,chip_pct,chip_value,wood4_pct,wood4_value,wood5_pct,wood5_value,wood6_pct,wood6_value)" +
                        " SELECT * FROM " + strTempTable;
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    oDataMgr.SqlNonQuery(conn, "DETACH DATABASE 'SOURCE'");
                }
            }

            frmMain.g_sbpInfo.Text = "Version Update: Updating Optimizer Variables for Carbon Metrics ...Stand by";
            string strOptimizerDefinitionsDb = this.ReferenceProjectDirectory + "\\" + Tables.OptimizerDefinitions.DefaultDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strOptimizerDefinitionsDb)))
            {
                conn.Open();

                int intVariables = (int)oDataMgr.getRecordCount(conn, "SELECT COUNT(*) FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName, 
                    Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName);

                oDataMgr.m_strSQL = "SELECT * FROM " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                    " WHERE VARIABLE_NAME = 'wood4_volume_1'";
                if (oDataMgr.getRecordCount(conn, oDataMgr.m_strSQL, Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName) == 0)
                {
                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultCalculatedOptimizerVariablesTableName +
                        " (VARIABLE_NAME, VARIABLE_DESCRIPTION, VARIABLE_TYPE, VARIABLE_SOURCE, HANDLE_NEGATIVES)" +
                        " VALUES ('wood4_volume_1', 'Sum of wood4 volume for 4 cycles. Each cycle is weighted at 1', 'ECON', " +
                        "'ECON_BY_RX_UTILIZED_SUM.wood4_vol_cf', 'omit'), ('wood5_volume_1', 'Sum of wood5 volume for 4 cycles. Each cycle is weighted at 1', " +
                        "'ECON', 'ECON_BY_RX_UTILIZED_SUM.wood5_vol_cf', 'omit'), ('wood6_volume_1', 'Sum of wood6 volume for 4 cycles. Each cycle is weighted at 1', " +
                        "'ECON', 'ECON_BY_RX_UTILIZED_SUM.wood6_vol_cf', 'omit'), ('total_bole_wood_volume_1', " +
                        "'Sum of merch, wood4, wood5, and wood6 volumes for 4 cycles. Each cycle is weighted at 1', 'ECON', 'CALCULATED', 'omit'), " +
                        "('merch_dry_weight_1', 'Sum of merch dry weight for 4 cycles. Each cycle is weighted at 1', 'ECON', 'ECON_BY_RX_UTILIZED_SUM.merch_wt_bdt', " +
                        "'omit'), ('chip_dry_weight_1', 'Sum of chip dry weight for 4 cycles. Each cycle is weighted at 1', 'ECON', " +
                        "'ECON_BY_RX_UTILIZED_SUM.chip_wt_bdt', 'omit'), ('wood4_dry_weight_1', 'Sum of wood4 dry weight for 4 cycles. Each cycle is weighted at 1', " +
                        "'ECON', 'ECON_BY_RX_UTILIZED_SUM.wood4_wt_bdt', 'omit'), ('wood5_dry_weight_1', 'Sum of wood5 dry weight for 4 cycles. Each cycle is weighted at 1', " +
                        "'ECON', 'ECON_BY_RX_UTILIZED_SUM.wood5_wt_bdt', 'omit'), ('wood6_dry_weight_1', 'Sum of wood6 dry weight for 4 cycles. Each cycle is weighted at 1', " +
                        "'ECON', 'ECON_BY_RX_UTILIZED_SUM.wood6_wt_bdt', 'omit'), ('total_bole_wood_dry_weight_1', 'Sum of merch, wood4, wood5, and wood6 dry weights " +
                        "for 4 cycles. Each cycle is weighted at 1', 'ECON', 'CALCULATED', 'omit')";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);

                    oDataMgr.m_strSQL = "INSERT INTO " + Tables.OptimizerDefinitions.DefaultCalculatedEconVariablesTableName +
                        " (calculated_variables_id, rxcycle, weight) VALUES " +
                        "(" + (intVariables + 1) + ", '1', 1), (" + (intVariables + 1) + ", '2', 1), (" + (intVariables + 1) + ", '3', 1), (" + (intVariables + 1) + ", '4', 1), " +
                        "(" + (intVariables + 2) + ", '1', 1), (" + (intVariables + 2) + ", '2', 1), (" + (intVariables + 2) + ", '3', 1), (" + (intVariables + 2) + ", '4', 1), " +
                        "(" + (intVariables + 3) + ", '1', 1), (" + (intVariables + 3) + ", '2', 1), (" + (intVariables + 3) + ", '3', 1), (" + (intVariables + 3) + ", '4', 1), " +
                        "(" + (intVariables + 4) + ", '1', 1), (" + (intVariables + 4) + ", '2', 1), (" + (intVariables + 4) + ", '3', 1), (" + (intVariables + 4) + ", '4', 1), " +
                        "(" + (intVariables + 5) + ", '1', 1), (" + (intVariables + 5) + ", '2', 1), (" + (intVariables + 5) + ", '3', 1), (" + (intVariables + 5) + ", '4', 1), " +
                        "(" + (intVariables + 6) + ", '1', 1), (" + (intVariables + 6) + ", '2', 1), (" + (intVariables + 6) + ", '3', 1), (" + (intVariables + 6) + ", '4', 1), " +
                        "(" + (intVariables + 7) + ", '1', 1), (" + (intVariables + 7) + ", '2', 1), (" + (intVariables + 7) + ", '3', 1), (" + (intVariables + 7) + ", '4', 1), " +
                        "(" + (intVariables + 8) + ", '1', 1), (" + (intVariables + 8) + ", '2', 1), (" + (intVariables + 8) + ", '3', 1), (" + (intVariables + 8) + ", '4', 1), " +
                        "(" + (intVariables + 9) + ", '1', 1), (" + (intVariables + 9) + ", '2', 1), (" + (intVariables + 9) + ", '3', 1), (" + (intVariables + 9) + ", '4', 1), " +
                        "(" + (intVariables + 10) + ", '1', 1), (" + (intVariables + 10) + ", '2', 1), (" + (intVariables + 10) + ", '3', 1), (" + (intVariables + 10) + ", '4', 1)";
                    oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                }
            }


            frmMain.g_sbpInfo.Text = "Version Update: Updating Optimizer Rule Definitions for Carbon Metrics ...Stand by";
            string strOptimizerRuleDefsDb = this.ReferenceProjectDirectory + "\\" + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioTableDbFile;
            using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString(strOptimizerRuleDefsDb)))
            {
                conn.Open();

                if (!oDataMgr.ColumnExist(conn, Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName, "merch_prev_econDG_YN"))
                {
                    string[] arrNewColumns = {"merch_prev_econDG_YN", "wood4_prev_econDG_YN", "wood5_prev_econDG_YN", "wood6_prev_econDG_YN",
                    "chips_prev_econDG_YN", "merch_prev_nofacDG_YN", "wood4_prev_nofacDG_YN", "wood5_prev_nofacDG_YN", "wood6_prev_nofacDG_YN"};
                    
                    foreach (string strColumn in arrNewColumns)
                    {
                        oDataMgr.m_strSQL = "ALTER TABLE " + Tables.OptimizerScenarioRuleDefinitions.DefaultScenarioCostsTableName +
                            " ADD COLUMN " + strColumn + " CHAR(1) DEFAULT 'N'";
                        oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                    }
                }
            }

            // Delete and re-add Processor results tables with new schemas
            string[] arrFolders = System.IO.Directory.GetDirectories(this.ReferenceProjectDirectory + "\\processor");
            for (int i = 0; i < arrFolders.Length; i++)
            {
                string strScenario = System.IO.Path.GetFileName(arrFolders[i]);
                if (!strScenario.Equals("db") && System.IO.File.Exists($@"{arrFolders[i]}\{Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile}"))
                {
                    using (System.Data.SQLite.SQLiteConnection conn = new System.Data.SQLite.SQLiteConnection(oDataMgr.GetConnectionString($@"{arrFolders[i]}\{Tables.ProcessorScenarioRun.DefaultSqliteResultsDbFile}")))
                    {
                        conn.Open();
                        if (!oDataMgr.FieldExist(conn, $@"SELECT * FROM {Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName} limit 1", "wood6_wt_bdt"))
                        {
                            oDataMgr.m_strSQL=$@"DROP TABLE {Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName}";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            oDataMgr.m_strSQL = $@"DROP TABLE {Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName}";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            oDataMgr.m_strSQL = $@"DROP TABLE {Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName}";
                            oDataMgr.SqlNonQuery(conn, oDataMgr.m_strSQL);
                            frmMain.g_oTables.m_oProcessor.CreateHarvestCostsTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName);
                            frmMain.g_oTables.m_oProcessor.CreateTreeVolValSpeciesDiamGroupsTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);
                            frmMain.g_oTables.m_oProcessorScenarioRun.CreateAdditionalKcpCpaTable(oDataMgr,
                                conn, Tables.ProcessorScenarioRun.DefaultAddKcpCpaTableName, false);
                        }
                    }
                }
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

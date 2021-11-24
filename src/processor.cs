using System;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Objects and logic for processing cutlist into BioSum output
    /// </summary>
    public class processor
    {
        private string m_strScenarioId = "";
        private ado_data_access  m_oAdo;
        private SQLite.ADO.DataMgr m_oDataMgr;
        private string m_strOpcostTableName = "OpCost_Input";
        private string m_strTvvTableName = "TreeVolValLowSlope";
        private string m_strDebugFile ="";
        private string m_strOpcostIdealTableName = "OpCost_Ideal_Output";
        private bool m_bUsingSqlite = false;
        //private ado_data_access p_oAdo;
        private System.Collections.Generic.List<tree> m_trees;
        private scenarioHarvestMethod m_scenarioHarvestMethod;
        private scenarioMoveInCost m_scenarioMoveInCost;
        private System.Collections.Generic.IDictionary<string, prescription> m_prescriptions;
        private System.Collections.Generic.IList<harvestMethod> m_harvestMethodList;
        private escalators m_escalators;
        public System.Collections.Generic.List<string> m_standsWithNoYardingDistance;

        public processor(string strDebugFile, string strScenarioId, string strConnectionString, bool bUsingSqlite,
                         string sqliteConnectionString)
        {
            m_strDebugFile = strDebugFile;
            m_strScenarioId = strScenarioId;
            m_bUsingSqlite = bUsingSqlite;
            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(strConnectionString);
            if (m_bUsingSqlite)
            {
                m_oDataMgr = new SQLite.ADO.DataMgr();
                m_oDataMgr.OpenConnection(sqliteConnectionString);
                // Attach to rule definitions database; Seems to be connection-specific
                string strScenarioDB =
                    frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                    "\\processor\\" + Tables.ProcessorScenarioRuleDefinitions.DefaultSqliteDbFile;
                string strSql = "ATTACH DATABASE '" + strScenarioDB + "' AS definitions";
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "Execute SQL: " + strSql + "\r\n");
                m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, strSql);
                // Set up an ODBC DSN for the temp database
                // Check to see if the input SQLite DSN exists and if so, delete so we can add
                ODBCMgr odbcmgr = new ODBCMgr();
                if (odbcmgr.CurrentUserDSNKeyExist(ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName))
                {
                    odbcmgr.RemoveUserDSN(ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName);
                }
                odbcmgr.CreateUserSQLiteDSN(ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName,
                    m_oDataMgr.m_Connection.FileName);
            }
        }
        
        public Queries init()
        {
            Queries p_oQueries = new Queries();
            
            //
            //CREATE LINK IN TEMP MDB TO ALL PROCESSOR SCENARIO TABLES
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: Create Links to the Scenario tables - " + System.DateTime.Now.ToString() + "\r\n");
            dao_data_access oDao = new dao_data_access();
            //
            //SCENARIO MDB
            //
            
            string strScenarioMDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\db\\scenario_processor_rule_definitions.mdb";
            //
            //SCENARIO RESULTS MDB
            //
            string strScenarioResultsMDB =
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() +
                "\\processor\\" + m_strScenarioId + "\\" + Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsDbFile;

            //
            //LOAD PROJECT DATATASOURCES INFO
            //
            p_oQueries.m_oFvs.LoadDatasource = true;
            p_oQueries.m_oReference.LoadDatasource = true;
            p_oQueries.m_oProcessor.LoadDatasource = true;
            p_oQueries.LoadDatasources(true, m_bUsingSqlite, "processor", m_strScenarioId);

            //link to all the scenario rule definition tables
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                "scenario_cost_revenue_escalators",
                strScenarioMDB, "scenario_cost_revenue_escalators", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                "scenario_additional_harvest_costs",
                strScenarioMDB, "scenario_additional_harvest_costs", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
               "scenario_harvest_cost_columns",
               strScenarioMDB, "scenario_harvest_cost_columns", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
              "scenario_harvest_method",
              strScenarioMDB, "scenario_harvest_method", true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName,
                strScenarioMDB, 
                Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName, true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName,
                strScenarioMDB,
                Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName, true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
             "scenario_tree_species_diam_dollar_values",
             strScenarioMDB, "scenario_tree_species_diam_dollar_values", true);
            //link scenario results tables
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName,
                strScenarioResultsMDB,
                Tables.ProcessorScenarioRun.DefaultHarvestCostsTableName, true);
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName,
                strScenarioResultsMDB,
                Tables.ProcessorScenarioRun.DefaultTreeVolValSpeciesDiamGroupsTableName, true);

            // link gis travel_timetable
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                p_oQueries.m_oTravelTime.m_strTravelTimeTable,
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.TravelTime.DefaultTravelTimeTableDbFile,
                p_oQueries.m_oTravelTime.m_strTravelTimeTable, true);

            // link to PRE_FVS_SUMMARY table
            oDao.CreateTableLink(p_oQueries.m_strTempDbFile,
                Tables.FVS.DefaultPreFVSSummaryTableName,
                frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim() + "\\" + Tables.FVS.DefaultPreFVSSummaryDbFile,
                Tables.FVS.DefaultPreFVSSummaryTableName, true);

            oDao.m_DaoDbEngine.Idle(1);
            oDao.m_DaoDbEngine.Idle(8);
            oDao.m_DaoWorkspace.Close();
            oDao.m_DaoDbEngine = null;
            oDao = null;

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "END: Create Links to the Scenario tables - " + System.DateTime.Now.ToString() + "\r\n");

            //
            //CREATE LINK IN TEMP MDB TO ALL VARIANT CUTLIST TABLES
            //
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "START: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");
            RxTools oRxTools = new RxTools();
            oRxTools.CreateTableLinksToFVSOutTreeListTables(p_oQueries, p_oQueries.m_strTempDbFile);
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "END: CreateTableLinksToFVSOutTreeListTables - " + System.DateTime.Now.ToString() + "\r\n");

            m_oAdo = new ado_data_access();
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(p_oQueries.m_strTempDbFile, "", ""));
 
            return p_oQueries;
        }
        
        public int LoadTrees(string p_strVariant, string p_strRxPackage, string p_strCondTableName, string p_strPlotTableName,
                             string p_strRefHarvestMethodTableName, string p_strRxTableName)
        {

            //Load harvest methods; Prescription load depends on harvest methods
            //@ToDo: Table names that are passed in here are still in MS Access
            m_harvestMethodList = LoadHarvestMethods(p_strRefHarvestMethodTableName);
            //If harvest methods didn't load, stop processing
            if (m_harvestMethodList == null)
            {
                return -1;
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.loadTrees BEGIN \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Variant: " + p_strVariant + " Package: " + p_strRxPackage + "\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            //Load prescriptions into reference dictionary
            m_prescriptions = LoadPrescriptions(p_strRxTableName);
            //Load diameter variables into reference object
            m_scenarioHarvestMethod = LoadScenarioHarvestMethod(m_strScenarioId);
            //Load escalators into reference object
            m_escalators = LoadEscalators();
            //Load move-in costs into reference object
            m_scenarioMoveInCost = LoadScenarioMoveInCost(m_strScenarioId);
            m_standsWithNoYardingDistance = new System.Collections.Generic.List<string>();

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Diameter Variables in Use: " + m_scenarioHarvestMethod.ToString() + "\r\n");
            
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";

            if (m_oAdo.m_intError == 0)
            {
                //string strSQL = "SELECT z.biosum_cond_id, c.biosum_plot_id, z.rxCycle, z.rx, z.rxYear, " +
                //                "z.dbh, z.tpa, z.volCfNet, z.drybiot, z.drybiom,z.FvsCreatedTree_YN, " +
                //                "z.fvs_tree_id, z.fvs_species, z.volTsGrs, z.volCfGrs, " +
                //                "c.slope, p.elev, p.gis_yard_dist_ft, min(t.travel_time) as travel_time " +
                //                  "FROM " + strTableName + " z, cond c, plot p, travel_time t " +
                //                  "WHERE z.rxpackage='" + p_strRxPackage + "' AND " +
                //                  "z.biosum_cond_id = c.biosum_cond_id AND " +
                //                  "c.biosum_plot_id = p.biosum_plot_id AND " +
                //                  "c.biosum_plot_id = t.biosum_plot_id AND " +
                //                  "mid(z.fvs_tree_id,1,2)='" + p_strVariant + "' AND " +
                //                  "t.travel_time > 0" +
                //                  " group by z.biosum_cond_id, c.biosum_plot_id, z.rxCycle, z.rx, z.rxYear, " +
                //                  "z.dbh, z.tpa, z.volCfNet, z.drybiot, z.drybiom,z.FvsCreatedTree_YN, " +
                //                  "z.fvs_tree_id, z.fvs_species, z.volTsGrs, z.volCfGrs, " +
                //                  "c.slope, p.elev, p.gis_yard_dist_ft";
                string strSQL = "SELECT z.biosum_cond_id, c.biosum_plot_id, z.rxCycle, z.rx, z.rxYear, z.dbh, z.tpa, " +
                                "z.volCfNet, z.drybiot, z.drybiom,z.FvsCreatedTree_YN, z.fvs_tree_id, " +
                                "z.fvs_species, z.volTsGrs, z.volCfGrs, c.slope, c.elev, c.gis_yard_dist_ft " +
                                "FROM " + strTableName + " z, " +
                                "(SELECT p.biosum_plot_id,p.gis_yard_dist_ft,p.elev,d.biosum_cond_id,d.slope FROM " +
                                p_strPlotTableName + " p INNER JOIN " + p_strCondTableName + " d ON p.biosum_plot_id = d.biosum_plot_id) c " +
                                "WHERE z.rxpackage='" + p_strRxPackage + "' AND " +
                                "z.biosum_cond_id = c.biosum_cond_id";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    m_trees = new System.Collections.Generic.List<tree>();
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        tree newTree = new tree();
                        try
                        {
                            newTree.DebugFile = m_strDebugFile;
                            newTree.CondId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_cond_id"]).Trim();
                            newTree.PlotId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_plot_id"]).Trim();
                            newTree.FvsTreeId = Convert.ToString(m_oAdo.m_OleDbDataReader["fvs_tree_id"]).Trim();
                            newTree.RxCycle = Convert.ToString(m_oAdo.m_OleDbDataReader["rxCycle"]).Trim();
                            newTree.RxPackage = p_strRxPackage;
                            newTree.Rx = Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                            newTree.RxYear = Convert.ToString(m_oAdo.m_OleDbDataReader["rxYear"]).Trim();
                            newTree.Dbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["dbh"]);
                            newTree.Tpa = Convert.ToDouble(m_oAdo.m_OleDbDataReader["tpa"]);
                            // Special processing for saplings where volCfNet may be null
                            if (m_oAdo.m_OleDbDataReader["volCfNet"] == System.DBNull.Value && newTree.IsSapling)
                            {
                                newTree.VolCfNet = 0;
                            }
                            else
                            {
                                newTree.VolCfNet = Convert.ToDouble(m_oAdo.m_OleDbDataReader["volCfNet"]);
                            }
                            newTree.VolTsGrs = Convert.ToDouble(m_oAdo.m_OleDbDataReader["volTsGrs"]);
                            // Special processing for saplings where volCfGrs may be null
                            if (m_oAdo.m_OleDbDataReader["volCfGrs"] == System.DBNull.Value && newTree.IsSapling)
                            {
                                newTree.VolCfGrs = 0;
                            }
                            else
                            {
                                newTree.VolCfGrs = Convert.ToDouble(m_oAdo.m_OleDbDataReader["volCfGrs"]);
                            }
                            newTree.DryBiot = Convert.ToDouble(m_oAdo.m_OleDbDataReader["drybiot"]);
                            // Special processing for saplings where drybiom may be null
                            if (m_oAdo.m_OleDbDataReader["drybiom"] == System.DBNull.Value && newTree.IsSapling)
                            {
                                newTree.DryBiom = 0;
                            }
                            else
                            {
                                newTree.DryBiom = Convert.ToDouble(m_oAdo.m_OleDbDataReader["drybiom"]);
                            }
                        }
                        catch (Exception)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Unable to load fvs_tree_id: " + newTree.FvsTreeId + " due to null values!\r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Process aborted for variant: " + p_strVariant + " rxPackage: " + p_strRxPackage + "!\r\n");
                            string strErrorMessage = "Unable to process selected variant/package due to missing tree weights or volume! " +
                                                     "Please run 'Step 4 - Post-processing audit check' from the 'FVS Output Data' menu" +
                                                     " to identify the trees.";
                            System.Windows.Forms.MessageBox.Show(strErrorMessage, "FIA Biosum");
                            return -1;
                        }

                        newTree.Slope = Convert.ToInt32(m_oAdo.m_OleDbDataReader["slope"]);
                        // find default harvest methods in prescription in case we need them
                        harvestMethod objDefaultHarvestMethodLowSlope = null;
                        harvestMethod objDefaultHarvestMethodSteepSlope = null;
                        prescription currentPrescription = null;
                        m_prescriptions.TryGetValue(newTree.Rx, out currentPrescription);
                        if (currentPrescription != null)
                        {
                            objDefaultHarvestMethodLowSlope = currentPrescription.HarvestMethodLowSlope;
                            objDefaultHarvestMethodSteepSlope = currentPrescription.HarvestMethodSteepSlope;
                        }

                        if (newTree.Slope < m_scenarioHarvestMethod.SteepSlopePct)
                        {
                            // assign low slope harvest method if not rx
                            if ((m_scenarioHarvestMethod.HarvMethodSelection.Equals(HarvestMethodSelection.SELECTED) || 
                                 m_scenarioHarvestMethod.HarvMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST)) &&
                                 m_scenarioHarvestMethod.HarvestMethodLowSlope != null)
                            {
                                newTree.HarvestMethod = m_scenarioHarvestMethod.HarvestMethodLowSlope;
                            }
                            else
                            {
                                newTree.HarvestMethod = objDefaultHarvestMethodLowSlope;
                            }
                        }
                        else
                        {
                            // assign steep slope harvest method if not rx
                            if ((m_scenarioHarvestMethod.HarvMethodSelection.Equals(HarvestMethodSelection.SELECTED) || 
                                 m_scenarioHarvestMethod.HarvMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST)) &&
                                 m_scenarioHarvestMethod.HarvestMethodSteepSlope != null)
                            {
                                newTree.HarvestMethod = m_scenarioHarvestMethod.HarvestMethodSteepSlope;
                            }
                            else
                            {
                                newTree.HarvestMethod = objDefaultHarvestMethodSteepSlope;
                            }
                        }
                        if (Convert.ToString(m_oAdo.m_OleDbDataReader["FvsCreatedTree_YN"]).Trim().ToUpper() == "Y")
                        {
                            newTree.FvsCreatedTree = true;
                            // only use fvs_species from cut list if it is an FVS created tree
                            // convert to int to get rid of leading 0
                            int intSpcd = Convert.ToInt16(m_oAdo.m_OleDbDataReader["fvs_species"]);
                            newTree.SpCd = Convert.ToString(intSpcd);
                        }
                        newTree.Elevation = Convert.ToInt32(m_oAdo.m_OleDbDataReader["elev"]);
                        if (m_oAdo.m_OleDbDataReader["gis_yard_dist_ft"] == System.DBNull.Value)
                            newTree.YardingDistance = -1;
                        else
                            newTree.YardingDistance = Convert.ToDouble(m_oAdo.m_OleDbDataReader["gis_yard_dist_ft"]);

                        // only process the tree if it has a valid yarding distance
                        if (newTree.YardingDistance > 0)
                        { 
                            m_trees.Add(newTree); 
                        }
                        else
                        {
                            // otherwise add it to the list of dropped stands
                            if (!m_standsWithNoYardingDistance.Contains(newTree.CondId))
                                m_standsWithNoYardingDistance.Add(newTree.CondId);
                        }
                        
                    }
                }
            }

            if (m_standsWithNoYardingDistance.Count > 0)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// !! " + Convert.ToString(m_standsWithNoYardingDistance.Count) + " stands had no valid yarding distance and were excluded from processing \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "// Check the gis_yard_dist_ft on the plot table to find these stands !!  \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.loadTrees END \r\n");
                if (m_trees != null)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//Loaded " + m_trees.Count + " trees \r\n");
                }
                else
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//No trees were loaded from the cut list. \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//This query didn't return any rows: \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//SELECT z.biosum_cond_id, c.biosum_plot_id, z.rxCycle, z.rx, z.rxYear, z.dbh, z.tpa, z.volCfNet, z.drybiot, \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//z.drybiom,z.FvsCreatedTree_YN, z.fvs_tree_id, z.fvs_species, z.volTsGrs, z.volCfGrs, c.slope, c.elev, \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//c.gis_yard_dist, t.min_traveltime FROM fvs_tree_IN_BM_P001_TREE_CUTLIST z, (SELECT MIN(ONE_WAY_HOURS) AS  \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//min_traveltime, BIOSUM_PLOT_ID FROM TRAVEL_TIME WHERE ONE_WAY_HOURS > 0 GROUP BY BIOSUM_PLOT_ID) t, (SELECT \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//p.biosum_plot_id,p.gis_yard_dist_ft,p.elev,d.biosum_cond_id,d.slope FROM plot p INNER JOIN cond d ON  \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//p.biosum_plot_id = d.biosum_plot_id) c WHERE z.rxpackage='001' AND z.biosum_cond_id = c.biosum_cond_id AND \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//c.biosum_plot_id = t.biosum_plot_id \r\n");
                }
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            return 0;
        }

        public int UpdateTrees(string p_strVariant, string p_strRxPackage, string p_strTreeTableName, 
            string p_strTravelTimesTableName, string p_strTreeSpeciesTableName, bool blnCreateReconcilationTable)
        {
            if (m_trees == null)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination.\r\nAuxillary tree data cannot be appended",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return -1;
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.updateTrees BEGIN \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            //Load species groups into reference dictionary
            System.Collections.Generic.IDictionary<string, treeSpecies> dictTreeSpecies = LoadTreeSpecies(p_strVariant, p_strTreeSpeciesTableName);
            //Load species diam values into reference dictionary
            System.Collections.Generic.IDictionary<string, speciesDiamValue> dictSpeciesDiamValues = LoadSpeciesDiamValues(m_strScenarioId);
            //Load diameter groups into reference list
            System.Collections.Generic.List<treeDiamGroup> listDiamGroups = LoadTreeDiamGroups();
            System.Collections.Generic.IDictionary<string, double> dictTravelTimes = null;
            if (m_scenarioMoveInCost.MoveInTimeMultiplier > 0)
            {
                //Load travel times
                dictTravelTimes = LoadTravelTimes(p_strTravelTimesTableName);

                //Abort if travel times have not been loaded
                if (dictTravelTimes.Count == 0)
                {
                    System.Windows.MessageBox.Show("Move-in Time Multiplier is set to 0 but there are no valid entries in the travel_time table. " +
                        "This is not a valid configuration. Process halted!",
                        "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return -1;
            }
            }

            if (dictTreeSpecies == null)
            {
                System.Windows.MessageBox.Show("Some reference data is unavailable. Processor cannot continue. Process halted!",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return -1;
            }
            
            //Query TREE table to get original FIA species codes
            if (m_oAdo.m_intError == 0)
            {
            string strTableName = "fvs_tree_IN_" + p_strVariant + "_P" + p_strRxPackage + "_TREE_CUTLIST";
            string strSQL = "SELECT DISTINCT t.fvs_tree_id, t.biosum_cond_id, t.spcd " +
                    "FROM " + p_strTreeTableName + " t, " + strTableName + " z " +
                    "WHERE t.fvs_tree_id = z.fvs_tree_id " +
                    "AND t.biosum_cond_id = z.biosum_cond_id " +
                    "AND z.rxpackage='" + p_strRxPackage + "' " +
                    "GROUP BY t.fvs_tree_id, t.biosum_cond_id, t.spcd";

            m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
            if (m_oAdo.m_OleDbDataReader.HasRows)
            {
                System.Collections.Generic.Dictionary<String, String> dictSpCd = 
                    new System.Collections.Generic.Dictionary<string, string>();
                while (m_oAdo.m_OleDbDataReader.Read())
                {
                    string strTreeId = Convert.ToString(m_oAdo.m_OleDbDataReader["fvs_tree_id"]).Trim();
                    string strCondId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_cond_id"]).Trim();
                    string strSpCd = Convert.ToString(m_oAdo.m_OleDbDataReader["spcd"]).Trim();
                    dictSpCd.Add(strCondId + "_" + strTreeId, strSpCd);
                }

                // Second pass at processing tree properties based on information from the cut list
                System.Collections.Generic.IList<tree> lstRemovetrees = new System.Collections.Generic.List<tree>();
                foreach (tree nextTree in m_trees)
                {
                    string strKey = nextTree.CondId + "_" + nextTree.FvsTreeId;
                    if (!nextTree.FvsCreatedTree)
                    {
                        nextTree.SpCd = dictSpCd[strKey];
                    }
                    // set tree species fields from treeSpecies dictionary
                    if (!dictTreeSpecies.ContainsKey(nextTree.SpCd))
                    {
                        System.Windows.Forms.MessageBox.Show("The tree_species table is missing either an entry or species group for variant " +
                        p_strVariant + " spcd " + nextTree.SpCd + ". Please resolve this issue before running Processor.",
                        "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return -1;
                    }

                    // set tree species fields from treeSpecies dictionary
                    treeSpecies foundSpecies = dictTreeSpecies[nextTree.SpCd];
                    nextTree.SpeciesGroup = foundSpecies.SpeciesGroup;
                    nextTree.OdWgt = foundSpecies.OdWgt;
                    nextTree.DryToGreen = foundSpecies.DryToGreen;
                    nextTree.IsWoodlandSpecies = foundSpecies.IsWoodlandSpecies;

                    // set diameter group from diameter group list
                    foreach (treeDiamGroup nextGroup in listDiamGroups)
                    {
                        if (nextTree.Dbh >= nextGroup.MinDiam &&
                            nextTree.Dbh <= nextGroup.MaxDiam)
                        {
                            nextTree.DiamGroup = nextGroup.DiamGroup;
                            break;
                        }
                    }

                    // set tree properties based on scenario_tree_species_diam_dollar_values
                    string strSpeciesDiamKey = nextTree.DiamGroup + "|" + nextTree.SpeciesGroup;
                    speciesDiamValue treeSpeciesDiam = null;
                    if (dictSpeciesDiamValues.TryGetValue(strSpeciesDiamKey, out treeSpeciesDiam))
                    {
                        nextTree.MerchValue = treeSpeciesDiam.MerchValue;
                        nextTree.ChipValue = treeSpeciesDiam.ChipValue;
                        switch (treeSpeciesDiam.WoodBin)
                        {
                            case "M":
                                nextTree.IsNonCommercial = false;
                                break;
                            case "C":
                                nextTree.IsNonCommercial = true;
                                break;
                        }
                    }
                    else
                    {
                        if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: Missing species diam values for diamGroup|speciesGroup " +
                                strSpeciesDiamKey + " - " + System.DateTime.Now.ToString() + "\r\n");
                    }
                    // saplings are never cull
                    if (nextTree.IsSapling == true)
                    {
                        nextTree.IsCull = false;
                    }
                    // set cull pct based on scenarioHarvestMethod.cullPctThreshold
                    else if (nextTree.VolCfNet < ((100 - m_scenarioHarvestMethod.CullPctThreshold) * nextTree.VolCfGrs / 100))
                    {
                        nextTree.IsCull = true;
                    }
                    else
                    {
                        nextTree.IsCull = false;
                    }
                    
                    //Populate travel_time from database
                    if (dictTravelTimes != null)
                    {
                        double dblTravelTime = 0;
                        if (dictTravelTimes.TryGetValue(nextTree.PlotId, out dblTravelTime))
                        {
                            nextTree.TravelTime = dblTravelTime;
                        }
                    }

                    // Apply move-in hours modifiers
                    if (m_scenarioMoveInCost.MoveInTimeMultiplier >= 0)
                        nextTree.TravelTime = nextTree.TravelTime * m_scenarioMoveInCost.MoveInTimeMultiplier;
                    if (m_scenarioMoveInCost.MoveInHoursAddend > 0)
                        nextTree.TravelTime = nextTree.TravelTime + m_scenarioMoveInCost.MoveInHoursAddend;

                    // Remove trees with no travel time if multiplier > 0
                    if (m_scenarioMoveInCost.MoveInTimeMultiplier > 0 && nextTree.TravelTime == 0)
                        lstRemovetrees.Add(nextTree);
                    
                    //Assign OpCostTreeType
                    nextTree.TreeType = ChooseOpCostTreeType(nextTree);
                    CalculateVolumeAndWeight(nextTree);

                    //Dump OpCostTreeType in .csv format for validation
                    //string strLogEntry = nextTree.Dbh + ", " + nextTree.Slope + ", " + nextTree.IsNonCommercial +
                    //    ", " + nextTree.SpeciesGroup + ", " + nextTree.TreeType.ToString();
                    //if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    //    frmMain.g_oUtils.WriteText(m_strDebugFile, "loadTrees: OpCost tree type, " +
                    //        strLogEntry + "\r\n");
                }
                foreach (tree objTree in lstRemovetrees)
                {
                    m_trees.Remove(objTree);
                }
              }
            }
            
            // Create the reconcilation table, if desired
            if (blnCreateReconcilationTable == true)
            {
                CreateTreeReconcilationTable(p_strVariant, p_strRxPackage, p_strTreeTableName, m_oAdo);
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.updateTrees END \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            return 0;
        }

        private void CreateTreeReconcilationTable(string p_strVariant, string p_strRxPackage, string p_strTreeTableName, 
            ado_data_access p_oAdo)
        {
            if (p_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT fvs_tree_id, biosum_cond_id, tree, cn from " + p_strTreeTableName + " where fvs_tree_id is not null and biosum_cond_id is not null";
  
                p_oAdo.SqlQueryReader(p_oAdo.m_OleDbConnection, strSQL);
                if (p_oAdo.m_OleDbDataReader.HasRows)
                {
                    short idxTree = 0;
                    short idxCn = 1;
                    System.Collections.Generic.IDictionary<string, System.Collections.Generic.List<string>> dictTreeTable = 
                        new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>();
                    while (p_oAdo.m_OleDbDataReader.Read())
                    {
                        System.Collections.Generic.List<string> treeList = 
                            new System.Collections.Generic.List<string>();
                        string strFvsTreeId = Convert.ToString(p_oAdo.m_OleDbDataReader["fvs_tree_id"]).Trim();
                        string strKey = Convert.ToString(p_oAdo.m_OleDbDataReader["biosum_cond_id"]).Trim() + "_" + strFvsTreeId;
                        //This is stored as a number but we convert to string so we can store in list
                        string strTree = Convert.ToString(p_oAdo.m_OleDbDataReader["tree"]);
                        string strCn = Convert.ToString(p_oAdo.m_OleDbDataReader["cn"]).Trim();
                        treeList.Add(strTree);
                        treeList.Add(strCn);
                        dictTreeTable.Add(strKey, treeList);
                    }

                    string strTableName = p_strVariant + "_" + p_strRxPackage + "_reconcile_trees";

                    // drop reconcilation table if it already exists
                    if (p_oAdo.TableExist(p_oAdo.m_OleDbConnection, strTableName) == true)
                        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, "DROP TABLE " + strTableName);


                    frmMain.g_oTables.m_oProcessor.CreateTreeReconcilationTable(p_oAdo, p_oAdo.m_OleDbConnection, strTableName);

                    string strTempCn = "9999";
                    int intTempTree = 9999;
                    foreach (tree nextTree in m_trees)
                    {
                        string strKey = nextTree.CondId + "_" + nextTree.FvsTreeId;
                        if (dictTreeTable.ContainsKey(strKey))
                        {
                            System.Collections.Generic.List<string> treeList = dictTreeTable[strKey];
                            intTempTree = Convert.ToInt16(treeList[idxTree]);
                            strTempCn = treeList[idxCn];
                        }
                        p_oAdo.m_strSQL = "INSERT INTO " + strTableName + " " +
                        "(cn, tree, biosum_cond_id, biosum_plot_id, spcd, merchWtGt, nonMerchWtGt, drybiom, " +
                        "drybiot, volCfNet, volCfGrs, volTsgrs, odWgt, dryToGreen, tpa, dbh, species_group, " +
                        "isSapling, isWoodland, isCull, diam_group, merch_value, opcost_type, biosum_category)" +
                        "VALUES ('" + strTempCn + "', " + intTempTree + ", '" + nextTree.CondId + "', '" + nextTree.PlotId + "', " +
                        nextTree.SpCd + ", " + nextTree.MerchWtGtPa + ", " + nextTree.NonMerchWtGtPa + ", " + nextTree.DryBiom + ", " +
                        nextTree.DryBiot + ", " + nextTree.VolCfNet + ", " + nextTree.VolCfGrs + ", " + nextTree.VolTsGrs + ", " + nextTree.OdWgt +
                        ", " + nextTree.DryToGreen + ", " + nextTree.Tpa + ", " + nextTree.Dbh + ", " + nextTree.SpeciesGroup + ", " +
                        nextTree.IsSapling + ", " + nextTree.IsWoodlandSpecies + ", " + nextTree.IsCull + ", " + nextTree.DiamGroup + 
                        ", " + nextTree.MerchValue + ", '" + nextTree.TreeType + "', " + nextTree.HarvestMethod.BiosumCategory + " )";

                        p_oAdo.SqlNonQuery(p_oAdo.m_OleDbConnection, p_oAdo.m_strSQL);
                    }
                }

            }
        }

        public int CreateOpcostInput(string p_strVariant)
        {
            int intReturnVal = -1;
            int intHwdSpeciesCodeThreshold = 299; // Species codes greater than this are hardwoods
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n The OpCost input file cannot be created",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return intReturnVal;
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createOpcostInput BEGIN \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }
            
            if (m_oAdo.m_intError == 0)
            {

                // create opcost input table
                frmMain.g_oTables.m_oProcessor.CreateNewOpcostInputTable(m_oAdo, m_oAdo.m_OleDbConnection, m_strOpcostTableName);

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Read trees into opcost input - " + System.DateTime.Now.ToString() + "\r\n");

                System.Collections.Generic.IDictionary<string, opcostInput> dictOpcostInput = 
                    new System.Collections.Generic.Dictionary<string, opcostInput>();

                    foreach (tree nextTree in m_trees)
                    {
                        opcostInput nextInput = null;
                        string strStand = nextTree.CondId + nextTree.RxPackage + nextTree.Rx + nextTree.RxCycle;
                        bool blnFound = dictOpcostInput.TryGetValue(strStand, out nextInput);
                        if (!blnFound)
                        { 
                            nextInput = new opcostInput(nextTree.CondId, nextTree.Slope, nextTree.RxCycle, nextTree.RxPackage,
                                                        nextTree.Rx, nextTree.RxYear, nextTree.YardingDistance, nextTree.Elevation,
                                                        nextTree.HarvestMethod, nextTree.TravelTime, m_scenarioMoveInCost);
                            dictOpcostInput.Add(strStand, nextInput);
                        }

                        // All trees add their BaFracCutNumerator to the total, regardless of treeType
                        nextInput.TotalBaFracCutNumerator = nextInput.TotalBaFracCutNumerator + nextTree.BaFracCutNumerator;
                        
                        // Metrics for brush cut trees
                        if (nextTree.TreeType == OpCostTreeType.BC)
                        {
                            nextInput.TotalBcTpa = nextInput.TotalBcTpa + nextTree.Tpa;
                            nextInput.PerAcBcVolCf = nextInput.PerAcBcVolCf + nextTree.BrushCutVolCfPa;
                        }

                        // Metrics for chip trees
                        else if (nextTree.TreeType == OpCostTreeType.CT)
                        {
                            nextInput.TotalChipTpa = nextInput.TotalChipTpa + nextTree.Tpa;
                            nextInput.ChipMerchVolCfPa = nextInput.ChipMerchVolCfPa + nextTree.MerchVolCfPa;
                            nextInput.ChipNonMerchVolCfPa = nextInput.ChipNonMerchVolCfPa + nextTree.NonMerchVolCfPa;
                            nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.TotalVolCfPa;
                            nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.TotalWtGtPa;
                            if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                nextInput.ChipHwdVolCfPa = nextInput.ChipHwdVolCfPa + nextTree.TotalVolCfPa;
                        }

                        // Metrics for small log trees
                        else if (nextTree.TreeType == OpCostTreeType.SL)
                        {
                            nextInput.TotalSmLogTpa = nextInput.TotalSmLogTpa + nextTree.Tpa;
                            nextInput.SmLogMerchVolCfPa = nextInput.SmLogMerchVolCfPa + nextTree.MerchVolCfPa;
                            if (nextTree.IsNonCommercial || nextTree.IsCull)
                            {
                                nextInput.SmLogNonCommMerchVolCfPa = nextInput.SmLogNonCommMerchVolCfPa + nextTree.MerchVolCfPa;
                                nextInput.SmLogNonCommVolCfPa = nextInput.SmLogNonCommVolCfPa + nextTree.TotalVolCfPa;
                            }
                            else
                            {
                                nextInput.SmLogCommNonMerchVolCfPa = nextInput.SmLogCommNonMerchVolCfPa + nextTree.NonMerchVolCfPa;
                            }
                            nextInput.SmLogVolCfPa = nextInput.SmLogVolCfPa + nextTree.TotalVolCfPa;
                            nextInput.SmLogWtGtPa = nextInput.SmLogWtGtPa + nextTree.TotalWtGtPa;
                            if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                nextInput.SmLogHwdVolCfPa = nextInput.SmLogHwdVolCfPa + nextTree.TotalVolCfPa;
                            nextInput.TotalSmLogQmdPa = nextInput.TotalSmLogQmdPa + nextTree.QmdPa;
                        }

                        // Metrics for large log trees
                        else if (nextTree.TreeType == OpCostTreeType.LL)
                        {
                            nextInput.TotalLgLogTpa = nextInput.TotalLgLogTpa + nextTree.Tpa;
                            nextInput.LgLogMerchVolCfPa = nextInput.LgLogMerchVolCfPa + nextTree.MerchVolCfPa;
                            if (nextTree.IsNonCommercial || nextTree.IsCull)
                            {
                                nextInput.LgLogNonCommMerchVolCfPa = nextInput.LgLogNonCommMerchVolCfPa + nextTree.MerchVolCfPa;
                                nextInput.LgLogNonCommVolCfPa = nextInput.LgLogNonCommVolCfPa + nextTree.TotalVolCfPa;
                            }
                            else
                            {
                                nextInput.LgLogCommNonMerchVolCfPa = nextInput.LgLogCommNonMerchVolCfPa + nextTree.NonMerchVolCfPa;
                            }
                            nextInput.LgLogVolCfPa = nextInput.LgLogVolCfPa + nextTree.TotalVolCfPa;
                            nextInput.LgLogWtGtPa = nextInput.LgLogWtGtPa + nextTree.TotalWtGtPa;
                            if (Convert.ToInt32(nextTree.SpCd) > intHwdSpeciesCodeThreshold)
                                nextInput.LgLogHwdVolCfPa = nextInput.LgLogHwdVolCfPa + nextTree.TotalVolCfPa;
                            nextInput.TotalLgLogQmdPa = nextInput.TotalLgLogQmdPa + nextTree.QmdPa;
                        }
                    }
                //System.Windows.MessageBox.Show(dictOpcostInput.Keys.Count + " lines in file");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Finished reading trees - " + System.DateTime.Now.ToString() + "\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Load Basal Area for all conditions from PRE_FVS_SUMMARY - " + System.DateTime.Now.ToString() + "\r\n");

                System.Collections.Generic.IDictionary<string, double> dictFvsPreBasalArea = this.LoadFvsPreBasalArea(p_strVariant);

                if (dictFvsPreBasalArea.Keys.Count == 0)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: NO BASAL AREA (BA) RECORDS FOUND IN PRE_FVS_SUMMARY TABLE!!   \r\n");
                }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createOpcostInput: Begin writing opcost input table - " + System.DateTime.Now.ToString() + "\r\n");
                long lngCount =0;

                foreach (string key in dictOpcostInput.Keys)
                {
                    opcostInput nextStand = dictOpcostInput[key];

                        // Some fields we wait to calculate until we have the totals
                        // *** BRUSH CUT ***
                        double dblBcAvgVolume = 0;

                        if (nextStand.TotalBcTpa > 0)
                        { dblBcAvgVolume = nextStand.PerAcBcVolCf / nextStand.TotalBcTpa; }

                        // *** CHIP TREES ***
                        double dblCtAvgVolume = 0;
                        if (nextStand.TotalChipTpa > 0)
                        { dblCtAvgVolume = nextStand.ChipVolCfPa / nextStand.TotalChipTpa; }
                        double dblCtMerchPctTotal = 0;
                        double dblCtAvgDensity = 0;
                        double dblCtHwdPct = 0;
                        if (nextStand.ChipVolCfPa > 0)
                        {
                            dblCtMerchPctTotal = nextStand.ChipMerchVolCfPa / nextStand.ChipVolCfPa * 100;
                            dblCtAvgDensity = nextStand.ChipWtGtPa * 2000 / nextStand.ChipVolCfPa;
                            dblCtHwdPct = nextStand.ChipHwdVolCfPa / nextStand.ChipVolCfPa * 100;
                        }


                        // *** SMALL LOGS ***
                        double dblSmLogAvgVolume = 0;
                        double dblSmLogAvgVolumeAdj = 0;
                        if (nextStand.TotalSmLogTpa > 0)
                        { dblSmLogAvgVolume = nextStand.SmLogVolCfPa / nextStand.TotalSmLogTpa; }
                        double dblSmLogMerchPctTotal = 0;
                        double dblSmLogChipPct_Cat1_3 = 0;
                        double dblSmLogChipPct_Cat2_4 = 0;
                        double dblSmLogChipPct_Cat5 = 0;
                        double dblSmLogAvgDensity = 0;
                        double dblSmLogHwdPct = 0;
                        if (nextStand.SmLogVolCfPa > 0)
                        {
                            dblSmLogMerchPctTotal = nextStand.SmLogMerchVolCfPa / nextStand.SmLogVolCfPa * 100;
                            dblSmLogChipPct_Cat1_3 = nextStand.SmLogNonCommMerchVolCfPa / nextStand.SmLogVolCfPa * 100;
                            dblSmLogChipPct_Cat2_4 = (nextStand.SmLogNonCommVolCfPa + nextStand.SmLogCommNonMerchVolCfPa) / nextStand.SmLogVolCfPa * 100;
                            dblSmLogChipPct_Cat5 = nextStand.SmLogNonCommVolCfPa / nextStand.SmLogVolCfPa * 100;
                            dblSmLogAvgDensity = nextStand.SmLogWtGtPa * 2000 / nextStand.SmLogVolCfPa;
                            dblSmLogHwdPct = nextStand.SmLogHwdVolCfPa / nextStand.SmLogVolCfPa * 100;
                        }
                        // Apply OpCost value limits
                        if (nextStand.TotalSmLogTpa > 0)
                        {
                            if (nextStand.TotalSmLogTpa < nextStand.HarvestMethod.MinTpa)
                            {
                                nextStand.TotalSmLogTpaUnadj = nextStand.TotalSmLogTpa;
                                nextStand.TotalSmLogTpa = nextStand.HarvestMethod.MinTpa;
                            }
                            if (dblSmLogAvgVolume < nextStand.HarvestMethod.MinAvgTreeVolCf)
                            {
                                dblSmLogAvgVolumeAdj = dblSmLogAvgVolume;
                                dblSmLogAvgVolume = nextStand.HarvestMethod.MinAvgTreeVolCf;
                            }
                        }

                        // *** LARGE LOGS ***
                        double dblLgLogAvgVolume = 0;
                        double dblLgLogAvgVolumeAdj = 0;
                        if (nextStand.TotalLgLogTpa > 0)
                        { dblLgLogAvgVolume = nextStand.LgLogVolCfPa / nextStand.TotalLgLogTpa; }
                        double dblLgLogMerchPctTotal = 0;
                        double dblLgLogChipPct_Cat1_3_4 = 0;
                        double dblLgLogChipPct_Cat2 = 0;
                        double dblLgLogChipPct_Cat5 = 0;
                        double dblLgLogAvgDensity = 0;
                        double dblLgLogHwdPct = 0;
                        if (nextStand.LgLogVolCfPa > 0)
                        {
                            dblLgLogMerchPctTotal = nextStand.LgLogMerchVolCfPa / nextStand.LgLogVolCfPa * 100;
                            dblLgLogChipPct_Cat1_3_4 = nextStand.LgLogNonCommMerchVolCfPa / nextStand.LgLogVolCfPa * 100;
                            dblLgLogChipPct_Cat2 = (nextStand.LgLogNonCommVolCfPa + nextStand.LgLogCommNonMerchVolCfPa) / nextStand.LgLogVolCfPa * 100;
                            dblLgLogChipPct_Cat5 = nextStand.LgLogNonCommVolCfPa / nextStand.LgLogVolCfPa * 100;
                            dblLgLogAvgDensity = nextStand.LgLogWtGtPa * 2000 / nextStand.LgLogVolCfPa;
                            dblLgLogHwdPct = nextStand.LgLogHwdVolCfPa / nextStand.LgLogVolCfPa * 100;
                        }
                        // Apply OpCost value limits
                        if (nextStand.TotalLgLogTpa > 0)
                        {
                            if (dblLgLogAvgVolume < nextStand.HarvestMethod.MinAvgTreeVolCf)
                            {
                                dblLgLogAvgVolumeAdj = dblLgLogAvgVolume;
                                dblLgLogAvgVolume = nextStand.HarvestMethod.MinAvgTreeVolCf;
                            }
                        }

                        // ** BA FRAC CUT (INTENSITY) **
                        if (dictFvsPreBasalArea.ContainsKey(nextStand.OpCostStand))
                        {
                            double dblBaFracCutDenominator = dictFvsPreBasalArea[nextStand.OpCostStand];
                            if (dblBaFracCutDenominator > 0)
                            {
                                double dblTestBaFracCut = nextStand.TotalBaFracCutNumerator / dblBaFracCutDenominator;
                                if (dblTestBaFracCut > 0 &&
                                    dblTestBaFracCut < 1)
                                {
                                    nextStand.BaFracCut = dblTestBaFracCut;
                                }
                                else
                                {
                                    frmMain.g_oUtils.WriteText(m_strDebugFile, "BA_FRAC_CUT --> INVALID VALUE: " + dblTestBaFracCut + " from opcost stand " + nextStand.OpCostStand + "\r\n");
                                }
                            }
                            else
                            {
                                frmMain.g_oUtils.WriteText(m_strDebugFile, "BA_FRAC_CUT --> INVALID VALUE FOR PRE_FVS_SUMMARY.BA: " + dblBaFracCutDenominator + " from opcost stand " + nextStand.OpCostStand + "\r\n");
                            }
                        }
                        else
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "BA_FRAC_CUT --> OPCOST STAND MISSING FROM PRE_FVS_SUMMARY TABLE: " + nextStand.OpCostStand + "\r\n");
                        }

                        m_oAdo.m_strSQL = "INSERT INTO " + m_strOpcostTableName + " " +
                        "(Stand, [Percent Slope], [One-way Yarding Distance], YearCostCalc, " +
                        "[Project Elevation], [Harvesting System], [Chip tree per acre], [Chip trees MerchAsPctOfTotal], " +
                        "[Chip trees average volume(ft3)], [CHIPS Average Density (lbs/ft3)], [CHIPS Hwd Percent], [Small log trees per acre],  " +
                        "[Small log trees MerchAsPctOfTotal], [Small log trees ChipPct_Cat1_3], [Small log trees ChipPct_Cat2_4], " +
                        "[Small log trees ChipPct_Cat5], [Small log trees average volume(ft3)], [Small log trees average density(lbs/ft3)], " +
                        "[Small log trees hardwood percent], [Large log trees per acre], [Large log trees MerchAsPctOfTotal ], " +
                        "[Large log trees ChipPct_Cat1_3_4], [Large log trees ChipPct_Cat2], [Large log trees ChipPct_Cat5], " +
                        "[Large log trees average vol(ft3)], [Large log trees average density(lbs/ft3)], [Large log trees hardwood percent], " +
                        "BrushCutTPA, [BrushCutAvgVol], RxPackage_Rx_RxCycle, biosum_cond_id, RxPackage, Rx, RxCycle, Move_In_Hours, " +
                        "Harvest_Area_Assumed_Acres, [Unadjusted One-way Yarding distance], " +
                        "[Unadjusted Small log trees per acre], [Unadjusted Small log trees average volume (ft3)], " +
                        "[Unadjusted Large log trees per acre], [Unadjusted Large log trees average vol(ft3)], " +
                        "ba_frac_cut, QMD_SL, QMD_LL )" +
                        "VALUES ('" + nextStand.OpCostStand + "', " + nextStand.PercentSlope + ", " + nextStand.YardingDistance + ", '" + nextStand.RxYear + "', " +
                        nextStand.Elev + ", '" + nextStand.HarvestMethod.Method + "', " + nextStand.TotalChipTpa + ", " +
                        dblCtMerchPctTotal + ", " + dblCtAvgVolume + ", " + dblCtAvgDensity + ", " + dblCtHwdPct + ", " +
                        nextStand.TotalSmLogTpa + ", " + dblSmLogMerchPctTotal + ", " + dblSmLogChipPct_Cat1_3 + ", " +
                        dblSmLogChipPct_Cat2_4 + ", " + dblSmLogChipPct_Cat5 + ", " + dblSmLogAvgVolume + ", " + dblSmLogAvgDensity + ", " + dblSmLogHwdPct + ", " +
                        nextStand.TotalLgLogTpa + ", " + dblLgLogMerchPctTotal + ", " + dblLgLogChipPct_Cat1_3_4 + "," +
                        dblLgLogChipPct_Cat2 + "," + dblLgLogChipPct_Cat5 + "," + dblLgLogAvgVolume + ", " +
                        dblLgLogAvgDensity + ", " + dblLgLogHwdPct + "," + nextStand.TotalBcTpa + ", " + dblBcAvgVolume +
                        ",'" + nextStand.RxPackageRxRxCycle + "', '" + nextStand.CondId + "', '" + nextStand.RxPackage + "', '" +
                        nextStand.Rx + "', '" + nextStand.RxCycle + "', " + nextStand.MoveInHours + ", " + 
                        nextStand.HarvestAreaAssumedAc + ", " + nextStand.YardingDistanceUnadj + ", " +
                        nextStand.TotalSmLogTpaUnadj + ", " + dblSmLogAvgVolumeAdj + ", " +
                        nextStand.TotalLgLogTpaUnadj + ", " + dblLgLogAvgVolumeAdj + ", " + nextStand.BaFracCut + ", " +
                        nextStand.QMD_SL + ", " + nextStand.QMD_LL +
                        " )";

                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
                        if (m_oAdo.m_intError != 0)
                        {
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "Processor.createOpcostInput ERROR " + System.DateTime.Now.ToString() + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "SQL: " + m_oAdo.m_strSQL + "\r\n");
                            frmMain.g_oUtils.WriteText(m_strDebugFile, "*********************************************" + "\r\n");
                            break;
                        }

                        lngCount++;
                        intReturnVal = 0;
                    }

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createOpcostInput END \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                }
                return intReturnVal;
            }
            return -1;
        }

        public int CreateTreeVolValWorkTable(string strDateTimeCreated)
        {
            int intReturnVal = -1;
            if (m_trees.Count < 1)
            {
                System.Windows.MessageBox.Show("No cut trees have been loaded for this scenario, variant, package combination. \r\n The tree vol val cannot be created",
                    "FIA Biosum", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                return intReturnVal;
            }

            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createTreeVolValWorkTable BEGIN \r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (m_oAdo.m_intError == 0)
            {
                if (!m_bUsingSqlite)
                {
                    // drop tree vol val work table (TreeVolValLowSlope) if it exists for next variant/package
                    if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, m_strTvvTableName) == true)
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + m_strTvvTableName);

                    // create tree vol val work table (TreeVolValLowSlope); Re-use the sql from tree vol val but don't create the indexes
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, Tables.Processor.CreateTreeVolValSpeciesDiamGroupsTableSQL(m_strTvvTableName));
                }
                else
                {
                    // drop tree vol val work table (TreeVolValLowSlope) if it exists for next variant/package
                    if (m_oDataMgr.TableExist(m_oDataMgr.m_Connection, m_strTvvTableName) == true)
                        m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, "DROP TABLE " + m_strTvvTableName);

                    // create tree vol val work table (TreeVolValLowSlope); Re-use the sql from tree vol val but don't create the indexes
                    m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, Tables.Processor.CreateSqliteTreeVolValSpeciesDiamGroupsTableSQL(m_strTvvTableName, false));

                    // check to see if table link exists; Create it if it doesn't
                    if (! m_oAdo.TableExist(m_oAdo.m_OleDbConnection, m_strTvvTableName))
                    {
                        dao_data_access oDao = new dao_data_access();
                        oDao.CreateSQLiteTableLink(m_oAdo.m_OleDbConnection.DataSource, m_strTvvTableName,
                            m_strTvvTableName, ODBCMgr.DSN_KEYS.ProcessorTemporaryDsnName,
                            m_oDataMgr.m_Connection.FileName);
                        oDao.m_DaoWorkspace.Close();
                        oDao = null;
                    }
                }


                // load opcostIdeal into memory if user asked for low-cost harvest system
                System.Collections.Generic.IDictionary<String, opcostIdeal> dictIdeal = new System.Collections.Generic.Dictionary<String, opcostIdeal>();
                bool blnUseIdeal = false;
                // 2018-30-JUL Disabling lowest-cost harvest method option. blnUseIdeal should always be false and
                // dictIdeal will never be used.
                //if (m_scenarioHarvestMethod.HarvMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST))
                //{
                //    if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                //        frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Load opcost_output_ideal into memory - " + System.DateTime.Now.ToString() + "\r\n");

                //    blnUseIdeal = true;
                //    //key: strCondId + strRxPackage + strRx + strRxCycle;
                //    dictIdeal = loadOpcostIdeal();
                //}
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 2)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Read trees into tree vol val - " + System.DateTime.Now.ToString() + "\r\n");

                string strSeparator = "_";
                System.Collections.Generic.IDictionary<string, treeVolValInput> dictTvvInput = 
                    new System.Collections.Generic.Dictionary<string, treeVolValInput>();
                foreach (tree nextTree in m_trees)
                {
                    if (blnUseIdeal)
                    {
                        string strIdealKey = nextTree.CondId + nextTree.RxPackage + nextTree.Rx + nextTree.RxCycle;
                        opcostIdeal objIdeal = null;
                        bool blnIdealFound = dictIdeal.TryGetValue(strIdealKey, out objIdeal);
                        if (blnIdealFound)
                        {
                            harvestMethod objHarvestMethodLowSlope = null;
                            harvestMethod objHarvestMethodSteepSlope = null;
                            foreach (harvestMethod nextMethod in m_harvestMethodList)
                            {
                                if (nextMethod.Method.Equals(objIdeal.HarvestSystem) && !nextMethod.SteepSlope)
                                {
                                    objHarvestMethodLowSlope = nextMethod;
                                }
                                else if (nextMethod.Method.Equals(objIdeal.HarvestSystem) && nextMethod.SteepSlope)
                                {
                                    objHarvestMethodSteepSlope = nextMethod;
                                }

                                if (nextTree.Slope < m_scenarioHarvestMethod.SteepSlopePct)
                                {
                                    nextTree.LowestCostHarvestMethod = objHarvestMethodLowSlope;
                                }
                                else
                                {
                                    nextTree.LowestCostHarvestMethod = objHarvestMethodSteepSlope;
                                }
                            }
                        }
                    }
                    
                    treeVolValInput nextInput = null;
                    string strKey = nextTree.CondId + strSeparator + nextTree.RxCycle + strSeparator + nextTree.DiamGroup + strSeparator + nextTree.SpeciesGroup;
                    bool blnFound = dictTvvInput.TryGetValue(strKey, out nextInput);
                    if (!blnFound)
                    {
                        //calculate chipMktValPgt; Apply escalators according to rxCycle
                        double chipMktValPgt = nextTree.ChipValue;
                        switch (nextTree.RxCycle)
                        {
                            case "2":
                                chipMktValPgt = chipMktValPgt * m_escalators.EnergyWoodRevCycle2;
                                break;
                            case "3":
                                chipMktValPgt = chipMktValPgt * m_escalators.EnergyWoodRevCycle3;
                                break;
                            case "4":
                                chipMktValPgt = chipMktValPgt * m_escalators.EnergyWoodRevCycle4;
                                break;
                        }
                        nextInput = new treeVolValInput(nextTree.CondId, nextTree.RxCycle, nextTree.RxPackage, nextTree.Rx,
                            nextTree.SpeciesGroup, nextTree.DiamGroup, nextTree.IsNonCommercial, chipMktValPgt, nextTree.HarvestMethod.BiosumCategory);
                        dictTvvInput.Add(strKey, nextInput);
                    }

                    // Metrics for brush cut trees
                    if (nextTree.TreeType == OpCostTreeType.BC)
                    {
                        nextInput.TotalBrushCutWtGtPa = nextInput.TotalBrushCutWtGtPa + nextTree.BrushCutWtGtPa;
                        nextInput.TotalBrushCutVolCfPa = nextInput.TotalBrushCutVolCfPa + nextTree.BrushCutVolCfPa;
                        nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.BrushCutWtGtPa;
                    }

                    //metrics for chip trees
                    else if (nextTree.TreeType == OpCostTreeType.CT)
                    {
                        if (nextTree.BiosumCategory == 1 || nextTree.BiosumCategory == 3)
                        {
                            // Only bole is chipped; nonMerch goes to stand residue
                            nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.MerchVolCfPa;
                            nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.MerchWtGtPa;
                            nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.NonMerchWtGtPa;
                        }
                        else
                        {
                            // Whole tree is chipped
                            nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.TotalVolCfPa;
                            nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.TotalWtGtPa;
                        }
                    }

                    //metrics for small and large trees
                    else if (nextTree.TreeType == OpCostTreeType.SL || nextTree.TreeType == OpCostTreeType.LL)
                    {
                        if (nextTree.BiosumCategory == 1 || nextTree.BiosumCategory == 3)
                        {
                            if (nextTree.IsNonCommercial || nextTree.IsCull)
                            {
                                // Only bole is chipped; nonMerch goes to stand residue
                                nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.MerchVolCfPa;
                                nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.MerchWtGtPa;
                                nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.NonMerchWtGtPa;
                            }
                            else
                            {
                                // Only bole is merch; nonMerch goes to stand residue
                                nextInput.TotalMerchVolCfPa = nextInput.TotalMerchVolCfPa + nextTree.MerchVolCfPa;
                                nextInput.TotalMerchWtGtPa = nextInput.TotalMerchWtGtPa+ nextTree.MerchWtGtPa;
                                nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.NonMerchWtGtPa;
                                nextInput.TotalMerchValDpa = nextInput.TotalMerchValDpa + nextTree.MerchValDpa;
                            }
                        }
                        else if (nextTree.BiosumCategory == 2)
                        {
                            if (nextTree.IsNonCommercial || nextTree.IsCull)
                            {
                                // Entire tree is chipped
                                nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.TotalVolCfPa;
                                nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.TotalWtGtPa;
                            }
                            else
                            {
                                // Bole is merch; nonMerch goes to chips
                                nextInput.TotalMerchVolCfPa = nextInput.TotalMerchVolCfPa + nextTree.MerchVolCfPa;
                                nextInput.TotalMerchWtGtPa = nextInput.TotalMerchWtGtPa + nextTree.MerchWtGtPa;
                                nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.NonMerchVolCfPa;
                                nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.NonMerchWtGtPa;
                                nextInput.TotalMerchValDpa = nextInput.TotalMerchValDpa + nextTree.MerchValDpa;
                            }
                        }
                        else if (nextTree.BiosumCategory == 4)
                        {
                            if (nextTree.TreeType == OpCostTreeType.SL)
                            {
                                if (nextTree.IsNonCommercial || nextTree.IsCull)
                                {
                                    // Entire tree is chipped
                                    nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.TotalVolCfPa;
                                    nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.TotalWtGtPa;
                                }
                                else
                                {
                                    // Bole is merch; nonMerch goes to chips
                                    nextInput.TotalMerchVolCfPa = nextInput.TotalMerchVolCfPa + nextTree.MerchVolCfPa;
                                    nextInput.TotalMerchWtGtPa = nextInput.TotalMerchWtGtPa + nextTree.MerchWtGtPa;
                                    nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.NonMerchVolCfPa;
                                    nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.NonMerchWtGtPa;
                                    nextInput.TotalMerchValDpa = nextInput.TotalMerchValDpa + nextTree.MerchValDpa;
                                }
                            }
                            else if (nextTree.TreeType == OpCostTreeType.LL)
                            {
                                if (nextTree.IsNonCommercial || nextTree.IsCull)
                                {
                                    // Only bole is chipped; nonMerch goes to stand residue
                                    nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.MerchVolCfPa;
                                    nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.MerchWtGtPa;
                                    nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.NonMerchWtGtPa;
                                }
                                else
                                {
                                    // Only bole is merch; nonMerch goes to stand residue
                                    nextInput.TotalMerchVolCfPa = nextInput.TotalMerchVolCfPa + nextTree.MerchVolCfPa;
                                    nextInput.TotalMerchWtGtPa = nextInput.TotalMerchWtGtPa + nextTree.MerchWtGtPa;
                                    nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.NonMerchWtGtPa;
                                    nextInput.TotalMerchValDpa = nextInput.TotalMerchValDpa + nextTree.MerchValDpa;
                                }
                            }
                        }
                        else if (nextTree.BiosumCategory == 5)
                        {
                            if (nextTree.IsNonCommercial || nextTree.IsCull)
                            {
                                // Entire tree is chipped
                                nextInput.ChipVolCfPa = nextInput.ChipVolCfPa + nextTree.TotalVolCfPa;
                                nextInput.ChipWtGtPa = nextInput.ChipWtGtPa + nextTree.TotalWtGtPa;
                            }
                            else
                            {
                                // Only bole is merch; nonMerch goes to stand residue
                                nextInput.TotalMerchVolCfPa = nextInput.TotalMerchVolCfPa + nextTree.MerchVolCfPa;
                                nextInput.TotalMerchWtGtPa = nextInput.TotalMerchWtGtPa + nextTree.MerchWtGtPa;
                                nextInput.StandResidueWtGtPa = nextInput.StandResidueWtGtPa + nextTree.NonMerchWtGtPa;
                                nextInput.TotalMerchValDpa = nextInput.TotalMerchValDpa + nextTree.MerchValDpa;
                            }
                        }
                    }

                }
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Finished reading trees - " + System.DateTime.Now.ToString() + "\r\n");

                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "createTreeVolValWorkTable: Begin writing tree vol val table - " + System.DateTime.Now.ToString() + "\r\n");
                long lngCount =0;

                System.Collections.Generic.IList<string> lstSql = new System.Collections.Generic.List<string>();
                foreach (string key in dictTvvInput.Keys)
                {
                    treeVolValInput nextStand = dictTvvInput[key];
                    string strSQL = "INSERT INTO " + m_strTvvTableName + " " +
                    "(biosum_cond_id, rxpackage, rx, rxcycle, species_group, diam_group, " +
                    "chip_vol_cf, chip_wt_gt, chip_val_dpa, chip_mkt_val_pgt," +
                    "merch_vol_cf, merch_wt_gt, merch_val_dpa, " +
                    "merch_to_chipbin_YN,  " +
                    "bc_vol_cf, bc_wt_gt, stand_residue_wt_gt, " +
                    //"biosum_harvest_method_category, DateTimeCreated, place_holder)" +
                    "DateTimeCreated, place_holder)" +
                    "VALUES ('" + nextStand.CondId + "', '" + nextStand.RxPackage + "', '" + nextStand.Rx + "', '" +
                    nextStand.RxCycle + "', " + nextStand.SpeciesGroup + ", " + nextStand.DiamGroup + ", " +
                    nextStand.ChipVolCfPa + ", " + nextStand.ChipWtGtPa + ", " + (nextStand.ChipWtGtPa * nextStand.ChipMktValPgt) +
                    ", " + nextStand.ChipMktValPgt + ", " + nextStand.TotalMerchVolCfPa + ", " + nextStand.TotalMerchWtGtPa + ", " + nextStand.TotalMerchValDpa +
                    ", '" + nextStand.MerchToChip + "', " + nextStand.TotalBrushCutVolCfPa + "," +
                    nextStand.TotalBrushCutWtGtPa + ", " + nextStand.StandResidueWtGtPa + ", " +
                    //nextStand.HarvestMethodCategory + ", '" + strDateTimeCreated + "', 'N')";
                    "'" + strDateTimeCreated + "', 'N')";
                    lstSql.Add(strSQL);
                }

                if (!m_bUsingSqlite)
                {
                    foreach (var item in lstSql)
                    {
                        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, item);
                        if (m_oAdo.m_intError != 0) break;
                        lngCount++;
                    }
                }
                else
                {
                    //Note: Wrapping this in a transaction made it MUCH faster!!
                    m_oDataMgr.m_Command = m_oDataMgr.m_Connection.CreateCommand();
                    using (m_oDataMgr.m_Transaction = m_oDataMgr.m_Connection.BeginTransaction())
                    {
                        m_oDataMgr.m_Command.Transaction = m_oDataMgr.m_Transaction;
                        foreach (var item in lstSql)
                        {
                            m_oDataMgr.SqlNonQuery(m_oDataMgr.m_Connection, item);
                            if (m_oDataMgr.m_intError != 0) break;
                            lngCount++;
                        }
                        m_oDataMgr.m_Transaction.Commit();
                    }
                }

                //Drop id column because it prevents copying rows into final tree vol val                    
                if (!m_bUsingSqlite)
                {
                    string strSqlAlter = "ALTER TABLE " + m_strTvvTableName + " DROP COLUMN id";
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, strSqlAlter);
                }                    
                
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "END createTreeVolValWorkTable INSERTED " + lngCount + " RECORDS: " + System.DateTime.Now.ToString() + "\r\n");

                intReturnVal = 0;
            }
            if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
            {
                frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                frmMain.g_oUtils.WriteText(m_strDebugFile, "//Processor.createTreeVolValWorkTable \r\n");
                 frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
            }

            if (m_oAdo != null)
            {
                m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
                m_oAdo = null;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//Dispose of ado_data_access object END \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                }
            }
            if (m_oDataMgr != null)
            {
                m_oDataMgr.CloseConnection(m_oDataMgr.m_Connection);
                m_oDataMgr = null;
                if (frmMain.g_bDebug && frmMain.g_intDebugLevel > 1)
                {
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "\r\n//\r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//Dispose of DataMgr object END \r\n");
                    frmMain.g_oUtils.WriteText(m_strDebugFile, "//\r\n");
                }
            }
            return intReturnVal;
        }

        private System.Collections.Generic.List<treeDiamGroup> LoadTreeDiamGroups()
        {
            System.Collections.Generic.List<treeDiamGroup> listDiamGroups = new System.Collections.Generic.List<treeDiamGroup>();
            if (!m_bUsingSqlite)
            {
                if (m_oAdo.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                        " WHERE TRIM(scenario_id)='" + m_strScenarioId.Trim() + "'";
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            int intDiamGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["diam_group"]);
                            double dblMinDiam = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_diam"]);
                            double dblMaxDiam = Convert.ToDouble(m_oAdo.m_OleDbDataReader["max_diam"]);
                            listDiamGroups.Add(new treeDiamGroup(intDiamGroup, dblMinDiam, dblMaxDiam));
                        }
                    }
                }
            }
            else
            {
                if (m_oDataMgr.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM definitions." + Tables.ProcessorScenarioRuleDefinitions.DefaultTreeDiamGroupsTableName +
                        " WHERE TRIM(scenario_id)='" + m_strScenarioId.Trim() + "'";
                    m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, strSQL);
                    if (m_oDataMgr.m_DataReader.HasRows)
                    {
                        while (m_oDataMgr.m_DataReader.Read())
                        {
                            int intDiamGroup = Convert.ToInt32(m_oDataMgr.m_DataReader["diam_group"]);
                            double dblMinDiam = Convert.ToDouble(m_oDataMgr.m_DataReader["min_diam"]);
                            double dblMaxDiam = Convert.ToDouble(m_oDataMgr.m_DataReader["max_diam"]);
                            listDiamGroups.Add(new treeDiamGroup(intDiamGroup, dblMinDiam, dblMaxDiam));
                        }
                        m_oDataMgr.m_DataReader.Close();
                    }
                }
            }

            return listDiamGroups;
        }

        private System.Collections.Generic.IDictionary<String, treeSpecies> LoadTreeSpecies(string p_strVariant, 
            string p_strTreeSpeciesTableName)
        {
            System.Collections.Generic.IDictionary<String, treeSpecies> dictTreeSpecies = 
                new System.Collections.Generic.Dictionary<String, treeSpecies>();
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT DISTINCT s.SPCD, s.SPECIES_GROUP, f.od_wgt, f.Dry_to_Green, f.WOODLAND_YN FROM " +
                         p_strTreeSpeciesTableName + " t, " +
                         Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesGroupsListTableName + " s, " +
                         Tables.ProcessorScenarioRun.DefaultFiaTreeSpeciesRefTableName + " f " +
                         "WHERE t.spcd = s.spcd AND t.spcd = f.spcd and f.spcd = s.spcd " +
                         "AND FVS_VARIANT = '" + p_strVariant + "' " +
                         "AND S.SPCD IS NOT NULL AND S.SPECIES_GROUP IS NOT NULL " +
                         "AND TRIM(UCASE(S.scenario_id)) = '" + m_strScenarioId.Trim().ToUpper() + "' " +
                         "GROUP BY s.SPCD, s.SPECIES_GROUP, f.OD_WGT, f.Dry_to_Green, f.WOODLAND_YN";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strSpCd = Convert.ToString(m_oAdo.m_OleDbDataReader["SPCD"]).Trim();
                        if (dictTreeSpecies.ContainsKey(strSpCd))
                        {
                            System.Windows.Forms.MessageBox.Show("The tree_species table contains duplicate entries for variant " +
                                p_strVariant + " spcd " + strSpCd + ". Please resolve this issue before running Processor.",
                                "FIA Biosum", System.Windows.Forms.MessageBoxButtons.OK, 
                                System.Windows.Forms.MessageBoxIcon.Error);
                            return null;
                        }
                        int intSpcGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["SPECIES_GROUP"]);
                        double dblOdWgt = Convert.ToDouble(m_oAdo.m_OleDbDataReader["OD_WGT"]);
                        double dblDryToGreen = Convert.ToDouble(m_oAdo.m_OleDbDataReader["Dry_to_Green"]);
                        string strIsWoodlandSpecies = Convert.ToString(m_oAdo.m_OleDbDataReader["WOODLAND_YN"]).Trim();
                        bool isWoodlandSpecies = false;
                        if (strIsWoodlandSpecies == "Y")
                        {
                            isWoodlandSpecies = true;
                        }
                        treeSpecies nextTreeSpecies = new treeSpecies(strSpCd, intSpcGroup, dblOdWgt, dblDryToGreen, isWoodlandSpecies);
                        dictTreeSpecies.Add(strSpCd, nextTreeSpecies);
                    }
                }
            }
            return dictTreeSpecies;
        }

        ///<summary>
        /// Loads scenario_tree_species_diam_dollar_values into a reference dictionary
        /// The composite key is intDiamGroup + "|" + intSpcGroup
        /// The value is a speciesDiamValue object
        ///</summary> 
        private System.Collections.Generic.IDictionary<String, speciesDiamValue> LoadSpeciesDiamValues(string p_scenario)
        {
            System.Collections.Generic.IDictionary<String, speciesDiamValue> dictSpeciesDiamValues = 
                new System.Collections.Generic.Dictionary<String, speciesDiamValue>();
            if (! m_bUsingSqlite)
            {
                if (m_oAdo.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " +
                                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName +
                                    " WHERE scenario_id = '" + p_scenario + "'";
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            int intSpcGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["species_group"]);
                            int intDiamGroup = Convert.ToInt32(m_oAdo.m_OleDbDataReader["diam_group"]);
                            string strWoodBin = Convert.ToString(m_oAdo.m_OleDbDataReader["wood_bin"]).Trim();
                            double dblMerchValue = Convert.ToDouble(m_oAdo.m_OleDbDataReader["merch_value"]);
                            double dblChipValue = Convert.ToDouble(m_oAdo.m_OleDbDataReader["chip_value"]);
                            string strKey = intDiamGroup + "|" + intSpcGroup;
                            dictSpeciesDiamValues.Add(strKey, new speciesDiamValue(intDiamGroup, intSpcGroup,
                                strWoodBin, dblMerchValue, dblChipValue));
                        }
                    }
                }
            }
            else
            {
                if (m_oDataMgr.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM definitions." +
                                    Tables.ProcessorScenarioRuleDefinitions.DefaultTreeSpeciesDollarValuesTableName +
                                    " WHERE scenario_id = '" + p_scenario + "'";
                    m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, strSQL);
                    if (m_oDataMgr.m_DataReader.HasRows)
                    {
                        while (m_oDataMgr.m_DataReader.Read())
                        {
                            int intSpcGroup = Convert.ToInt32(m_oDataMgr.m_DataReader["species_group"]);
                            int intDiamGroup = Convert.ToInt32(m_oDataMgr.m_DataReader["diam_group"]);
                            string strWoodBin = Convert.ToString(m_oDataMgr.m_DataReader["wood_bin"]).Trim();
                            double dblMerchValue = Convert.ToDouble(m_oDataMgr.m_DataReader["merch_value"]);
                            double dblChipValue = Convert.ToDouble(m_oDataMgr.m_DataReader["chip_value"]);
                            string strKey = intDiamGroup + "|" + intSpcGroup;
                            dictSpeciesDiamValues.Add(strKey, new speciesDiamValue(intDiamGroup, intSpcGroup,
                                strWoodBin, dblMerchValue, dblChipValue));
                        }
                        m_oDataMgr.m_DataReader.Close();
                    }
                }
            }

            return dictSpeciesDiamValues;
        }

        private System.Collections.Generic.IDictionary<String, prescription> LoadPrescriptions(string p_strRxTable)
        {
            if (m_harvestMethodList == null || m_harvestMethodList.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Harvest methods must be loaded before loading prescriptions", "FIA Biosum", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            
            System.Collections.Generic.IDictionary<String, prescription> dictPrescriptions = 
                new System.Collections.Generic.Dictionary<String, prescription>();
                if (m_oAdo.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " + p_strRxTable;
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            string strRx = Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                            string strHarvestMethodLowSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
                            string strHarvestMethodSteepSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();
                            harvestMethod objHarvestMethodLowSlope = null;
                            harvestMethod objHarvestMethodSteepSlope = null;
                            foreach (harvestMethod nextMethod in m_harvestMethodList)
                            {
                                if (nextMethod.Method.Equals(strHarvestMethodLowSlope) && !nextMethod.SteepSlope)
                                {
                                    objHarvestMethodLowSlope = nextMethod;
                                }
                                else if (nextMethod.Method.Equals(strHarvestMethodSteepSlope) && nextMethod.SteepSlope)
                                {
                                    objHarvestMethodSteepSlope = nextMethod;
                                }
                            }

                            dictPrescriptions.Add(strRx, new prescription(strRx, objHarvestMethodLowSlope, objHarvestMethodSteepSlope));
                        }
                    }
                }
            return dictPrescriptions;
        }

        private scenarioHarvestMethod LoadScenarioHarvestMethod(string p_scenario)
        {
            if (m_harvestMethodList == null || m_harvestMethodList.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Harvest methods must be loaded before loading scenario harvest methods", "FIA Biosum",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            
            scenarioHarvestMethod returnVariables = null;
            if (!m_bUsingSqlite)
            {
                if (m_oAdo.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName +
                                    " WHERE scenario_id = '" + p_scenario + "'";
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        // We should only have one record
                        m_oAdo.m_OleDbDataReader.Read();
                        string strHarvestMethodLowSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodLowSlope"]).Trim();
                        double dblMinChipDbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_chip_dbh"]);
                        double dblMinSmallLogDbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_sm_log_dbh"]);
                        double dblMinLgLogDbh = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_lg_log_dbh"]);
                        int intMinSlopePct = Convert.ToInt32(m_oAdo.m_OleDbDataReader["SteepSlope"]);
                        double dblMinDbhSteepSlope = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_dbh_steep_slope"]);
                        string strHarvestMethodSelection = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSelection"]).Trim();
                        HarvestMethodSelection objHarvestMethodSelection = HarvestMethodSelection.RX;
                        if (strHarvestMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST.Value))
                        {
                            objHarvestMethodSelection = HarvestMethodSelection.LOWEST_COST;
                        }
                        else if (strHarvestMethodSelection.Equals(HarvestMethodSelection.SELECTED.Value))
                        {
                            objHarvestMethodSelection = HarvestMethodSelection.SELECTED;
                        }
                        string strHarvestMethodSteepSlope = Convert.ToString(m_oAdo.m_OleDbDataReader["HarvestMethodSteepSlope"]).Trim();
                        int intSaplingMerchAsPercentOfTotalVol = Convert.ToInt16(m_oAdo.m_OleDbDataReader["SaplingMerchAsPercentOfTotalVol"]);
                        int intWoodlandMerchAsPercentOfTotalVol = Convert.ToInt16(m_oAdo.m_OleDbDataReader["WoodlandMerchAsPercentOfTotalVol"]);
                        int intCullPctThreshold = Convert.ToInt16(m_oAdo.m_OleDbDataReader["CullPctThreshold"]);
                        harvestMethod objHarvestMethodLowSlope = null;
                        harvestMethod objHarvestMethodSteepSlope = null;
                        foreach (harvestMethod nextMethod in m_harvestMethodList)
                        {
                            if (nextMethod.Method.Equals(strHarvestMethodLowSlope) && !nextMethod.SteepSlope)
                            {
                                objHarvestMethodLowSlope = nextMethod;
                            }
                            else if (nextMethod.Method.Equals(strHarvestMethodSteepSlope) && nextMethod.SteepSlope)
                            {
                                objHarvestMethodSteepSlope = nextMethod;
                            }
                        }

                        returnVariables = new scenarioHarvestMethod(dblMinChipDbh, dblMinSmallLogDbh, dblMinLgLogDbh,
                            intMinSlopePct, dblMinDbhSteepSlope,
                            objHarvestMethodLowSlope, objHarvestMethodSteepSlope,
                            intSaplingMerchAsPercentOfTotalVol, intWoodlandMerchAsPercentOfTotalVol, intCullPctThreshold, objHarvestMethodSelection);
                    }
                }
            }
            else
            {
                if (m_oDataMgr.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM definitions." + Tables.ProcessorScenarioRuleDefinitions.DefaultHarvestMethodTableName +
                                    " WHERE scenario_id = '" + p_scenario + "'";
                    m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, strSQL);
                    if (m_oDataMgr.m_DataReader.HasRows)
                    {
                        // We should only have one record
                        m_oDataMgr.m_DataReader.Read();
                        string strHarvestMethodLowSlope = Convert.ToString(m_oDataMgr.m_DataReader["HarvestMethodLowSlope"]).Trim();
                        double dblMinChipDbh = Convert.ToDouble(m_oDataMgr.m_DataReader["min_chip_dbh"]);
                        double dblMinSmallLogDbh = Convert.ToDouble(m_oDataMgr.m_DataReader["min_sm_log_dbh"]);
                        double dblMinLgLogDbh = Convert.ToDouble(m_oDataMgr.m_DataReader["min_lg_log_dbh"]);
                        int intMinSlopePct = Convert.ToInt32(m_oDataMgr.m_DataReader["SteepSlope"]);
                        double dblMinDbhSteepSlope = Convert.ToDouble(m_oDataMgr.m_DataReader["min_dbh_steep_slope"]);
                        string strHarvestMethodSelection = Convert.ToString(m_oDataMgr.m_DataReader["HarvestMethodSelection"]).Trim();
                        HarvestMethodSelection objHarvestMethodSelection = HarvestMethodSelection.RX;
                        if (strHarvestMethodSelection.Equals(HarvestMethodSelection.LOWEST_COST.Value))
                        {
                            objHarvestMethodSelection = HarvestMethodSelection.LOWEST_COST;
                        }
                        else if (strHarvestMethodSelection.Equals(HarvestMethodSelection.SELECTED.Value))
                        {
                            objHarvestMethodSelection = HarvestMethodSelection.SELECTED;
                        }
                        string strHarvestMethodSteepSlope = Convert.ToString(m_oDataMgr.m_DataReader["HarvestMethodSteepSlope"]).Trim();
                        int intSaplingMerchAsPercentOfTotalVol = Convert.ToInt16(m_oDataMgr.m_DataReader["SaplingMerchAsPercentOfTotalVol"]);
                        int intWoodlandMerchAsPercentOfTotalVol = Convert.ToInt16(m_oDataMgr.m_DataReader["WoodlandMerchAsPercentOfTotalVol"]);
                        int intCullPctThreshold = Convert.ToInt16(m_oDataMgr.m_DataReader["CullPctThreshold"]);
                        harvestMethod objHarvestMethodLowSlope = null;
                        harvestMethod objHarvestMethodSteepSlope = null;
                        foreach (harvestMethod nextMethod in m_harvestMethodList)
                        {
                            if (nextMethod.Method.Equals(strHarvestMethodLowSlope) && !nextMethod.SteepSlope)
                            {
                                objHarvestMethodLowSlope = nextMethod;
                            }
                            else if (nextMethod.Method.Equals(strHarvestMethodSteepSlope) && nextMethod.SteepSlope)
                            {
                                objHarvestMethodSteepSlope = nextMethod;
                            }
                        }

                        returnVariables = new scenarioHarvestMethod(dblMinChipDbh, dblMinSmallLogDbh, dblMinLgLogDbh,
                            intMinSlopePct, dblMinDbhSteepSlope,
                            objHarvestMethodLowSlope, objHarvestMethodSteepSlope,
                            intSaplingMerchAsPercentOfTotalVol, intWoodlandMerchAsPercentOfTotalVol, intCullPctThreshold, objHarvestMethodSelection);
                    }
                }
            }

            return returnVariables;
        }

        private scenarioMoveInCost LoadScenarioMoveInCost(string p_scenario)
        {
            scenarioMoveInCost returnVariables = null;
            if (!m_bUsingSqlite)
            {
                if (m_oAdo.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName +
                                    " WHERE scenario_id = '" + p_scenario + "'";
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        // We should only have one record
                        m_oAdo.m_OleDbDataReader.Read();
                        double dblYardDistThreshold = Convert.ToDouble(m_oAdo.m_OleDbDataReader["yard_dist_threshold"]);
                        double dblAssumedHarvestAreaAc = Convert.ToDouble(m_oAdo.m_OleDbDataReader["assumed_harvest_area_ac"]);
                        double dblMoveInTimeMultiplier = Convert.ToDouble(m_oAdo.m_OleDbDataReader["move_in_time_multiplier"]);
                        double dblMoveInHoursAddend = Convert.ToDouble(m_oAdo.m_OleDbDataReader["move_in_hours_addend"]);

                        returnVariables = new scenarioMoveInCost(dblYardDistThreshold, dblAssumedHarvestAreaAc,
                                                                 dblMoveInTimeMultiplier, dblMoveInHoursAddend);
                    }
                }
            }
            else
            {
                if (m_oDataMgr.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " + Tables.ProcessorScenarioRuleDefinitions.DefaultMoveInCostsTableName +
                                    " WHERE scenario_id = '" + p_scenario + "'";
                    m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, strSQL);
                    if (m_oDataMgr.m_DataReader.HasRows)
                    {
                        // We should only have one record
                        m_oDataMgr.m_DataReader.Read();
                        double dblYardDistThreshold = Convert.ToDouble(m_oDataMgr.m_DataReader["yard_dist_threshold"]);
                        double dblAssumedHarvestAreaAc = Convert.ToDouble(m_oDataMgr.m_DataReader["assumed_harvest_area_ac"]);
                        double dblMoveInTimeMultiplier = Convert.ToDouble(m_oDataMgr.m_DataReader["move_in_time_multiplier"]);
                        double dblMoveInHoursAddend = Convert.ToDouble(m_oDataMgr.m_DataReader["move_in_hours_addend"]);
                        returnVariables = new scenarioMoveInCost(dblYardDistThreshold, dblAssumedHarvestAreaAc,
                                                                 dblMoveInTimeMultiplier, dblMoveInHoursAddend);
                        m_oDataMgr.m_DataReader.Close();
                    }
                }
            }

            return returnVariables;
        }

        private escalators LoadEscalators()
        {
            escalators returnEscalators = null;
            if (!m_bUsingSqlite)
            {
                if (m_oAdo.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM " +
                                    Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName +
                                    " WHERE scenario_id = '" + m_strScenarioId + "'";
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        // We should only have one record
                        m_oAdo.m_OleDbDataReader.Read();
                        double dblEnergyWoodRevCycle2 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle2"]);
                        double dblEnergyWoodRevCycle3 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle3"]);
                        double dblEnergyWoodRevCycle4 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorEnergyWoodRevenue_Cycle4"]);
                        double dblMerchWoodRevCycle2 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle2"]);
                        double dblMerchWoodRevCycle3 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle3"]);
                        double dblMerchWoodRevCycle4 = Convert.ToDouble(m_oAdo.m_OleDbDataReader["EscalatorMerchWoodRevenue_Cycle4"]);


                        returnEscalators = new escalators(dblEnergyWoodRevCycle2, dblEnergyWoodRevCycle3, dblEnergyWoodRevCycle4,
                                                          dblMerchWoodRevCycle2, dblMerchWoodRevCycle3, dblMerchWoodRevCycle4);
                    }
                }
            }
            else
            {
                if (m_oDataMgr.m_intError == 0)
                {
                    string strSQL = "SELECT * FROM definitions." +
                                    Tables.ProcessorScenarioRuleDefinitions.DefaultCostRevenueEscalatorsTableName +
                                    " WHERE scenario_id = '" + m_strScenarioId + "'";
                    m_oDataMgr.SqlQueryReader(m_oDataMgr.m_Connection, strSQL);
                    if (m_oDataMgr.m_DataReader.HasRows)
                    {
                        // We should only have one record
                        m_oDataMgr.m_DataReader.Read();
                        double dblEnergyWoodRevCycle2 = Convert.ToDouble(m_oDataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle2"]);
                        double dblEnergyWoodRevCycle3 = Convert.ToDouble(m_oDataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle3"]);
                        double dblEnergyWoodRevCycle4 = Convert.ToDouble(m_oDataMgr.m_DataReader["EscalatorEnergyWoodRevenue_Cycle4"]);
                        double dblMerchWoodRevCycle2 = Convert.ToDouble(m_oDataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle2"]);
                        double dblMerchWoodRevCycle3 = Convert.ToDouble(m_oDataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle3"]);
                        double dblMerchWoodRevCycle4 = Convert.ToDouble(m_oDataMgr.m_DataReader["EscalatorMerchWoodRevenue_Cycle4"]);


                        returnEscalators = new escalators(dblEnergyWoodRevCycle2, dblEnergyWoodRevCycle3, dblEnergyWoodRevCycle4,
                                                          dblMerchWoodRevCycle2, dblMerchWoodRevCycle3, dblMerchWoodRevCycle4);
                    }
                }
            }

            return returnEscalators;
        }

        private System.Collections.Generic.IList<harvestMethod> LoadHarvestMethods(string p_strRefHarvestMethodTableName)
        {
            System.Collections.Generic.IList<harvestMethod> harvestMethodList = null;

                if (m_oAdo.m_intError == 0)
                {
                    // Check to see if the biosum_category column exists in the harvest method table; If not
                    // throw an error and exit the function; Processor won't work without this value
                    if (!m_oAdo.ColumnExist(m_oAdo.m_OleDbConnection, p_strRefHarvestMethodTableName, "min_tpa"))
                    {
                        string strErrMsg = "Your project contains an obsolete version of the " + p_strRefHarvestMethodTableName +
                                           " table that does not contain the 'min_tpa' field. Copy a new version of this table into your project from the" +
                                           " BioSum installation directory before trying to run Processor.";
                        System.Windows.Forms.MessageBox.Show(strErrMsg, "FIA Biosum",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return harvestMethodList;
                    }

                    string strSQL = "SELECT * FROM " + p_strRefHarvestMethodTableName;
                    m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                    if (m_oAdo.m_OleDbDataReader.HasRows)
                    {
                        harvestMethodList = new System.Collections.Generic.List<harvestMethod>();
                        while (m_oAdo.m_OleDbDataReader.Read())
                        {
                            string strSteepYN = Convert.ToString(m_oAdo.m_OleDbDataReader["STEEP_YN"]).Trim();
                            bool blnSteep = false;
                            if (strSteepYN.Equals("Y"))
                            { blnSteep = true; }
                            string strMethod = Convert.ToString(m_oAdo.m_OleDbDataReader["Method"]).Trim();
                            int intBiosumCategory = Convert.ToInt16(m_oAdo.m_OleDbDataReader["biosum_category"]);
                            double dblMinTpa = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_tpa"]);
                            double dblMinYardDistanceFt = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_yard_distance_ft"]);
                            double dblMinAvgTreeVolCf = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_avg_tree_vol_cf"]);
                            harvestMethod newMethod = new harvestMethod(blnSteep, strMethod, intBiosumCategory, dblMinTpa,
                                dblMinYardDistanceFt, dblMinAvgTreeVolCf);
                            harvestMethodList.Add(newMethod);
                        }
                    }
                }
            return harvestMethodList;
        }
        
        private OpCostTreeType ChooseOpCostTreeType(tree p_tree)
        {
            OpCostTreeType returnType = OpCostTreeType.None;

            if (p_tree.Dbh < m_scenarioHarvestMethod.MinChipDbh)
            {
                returnType = OpCostTreeType.BC;
            }
            else if (p_tree.Slope >= m_scenarioHarvestMethod.SteepSlopePct && 
                     p_tree.Dbh < m_scenarioHarvestMethod.MinDbhSteepSlope)
            {
                returnType = OpCostTreeType.BC;
            }
            else if (p_tree.Dbh >= m_scenarioHarvestMethod.MinChipDbh && 
                     p_tree.Dbh < m_scenarioHarvestMethod.MinSmallLogDbh)
            {
                returnType = OpCostTreeType.CT;
            }
            else if (p_tree.Dbh >= m_scenarioHarvestMethod.MinSmallLogDbh &&
                     p_tree.Dbh < m_scenarioHarvestMethod.MinLargeLogDbh)
            {
                returnType = OpCostTreeType.SL;
            }
            else if (p_tree.Dbh >= m_scenarioHarvestMethod.MinLargeLogDbh)
            {
                returnType = OpCostTreeType.LL;
            }

            return returnType;
        }

        private void CalculateVolumeAndWeight(tree p_tree)
        {
            
            //merchVolCfPa
            if (p_tree.IsSapling)
            {
                if (p_tree.VolTsGrs > 0)
                {
                    p_tree.MerchVolCfPa = p_tree.VolTsGrs * ( (double) m_scenarioHarvestMethod.SaplingMerchAsPercentOfTotalVol / 100) * p_tree.Tpa;
                }
                else
                {
                    //convert drybiot to some kind of volume
                    p_tree.MerchVolCfPa = (p_tree.DryBiot / p_tree.OdWgt) * p_tree.Tpa * ((double)m_scenarioHarvestMethod.SaplingMerchAsPercentOfTotalVol / 100);
                }
            }
            else if (p_tree.IsWoodlandSpecies)
            {
                p_tree.MerchVolCfPa = p_tree.VolCfGrs * ( (double) m_scenarioHarvestMethod.WoodlandMerchAsPercentOfTotalVol / 100) * p_tree.Tpa;
            }
            else
            {
                p_tree.MerchVolCfPa = p_tree.VolCfGrs * p_tree.Tpa;
            }

            //merchValDpa
            if (!p_tree.IsNonCommercial && !p_tree.IsCull)
            {
                switch (p_tree.RxCycle)
                {
                    case "1":
                        p_tree.MerchValDpa = p_tree.MerchVolCfPa * p_tree.MerchValue;
                        break;
                    case "2":
                        p_tree.MerchValDpa = p_tree.MerchVolCfPa * p_tree.MerchValue * m_escalators.MerchWoodRevCycle2;
                        break;
                    case "3":
                        p_tree.MerchValDpa = p_tree.MerchVolCfPa * p_tree.MerchValue * m_escalators.MerchWoodRevCycle3;
                        break;
                    case "4":
                        p_tree.MerchValDpa = p_tree.MerchVolCfPa * p_tree.MerchValue * m_escalators.MerchWoodRevCycle4;
                        break;
                }
            }

            //merchWtGtPa; Always calculate merchVolCfPa first because it's used in this equation
            p_tree.MerchWtGtPa = p_tree.MerchVolCfPa * p_tree.OdWgt / p_tree.DryToGreen / 2000;

            //nonMerchVolCfPa
            if (p_tree.IsSapling)
            {
                p_tree.NonMerchVolCfPa = ((p_tree.DryBiot * ((100 - (double) m_scenarioHarvestMethod.SaplingMerchAsPercentOfTotalVol) / 100)) 
                                         * p_tree.Tpa) / p_tree.OdWgt;
            }
            else if (p_tree.IsWoodlandSpecies)
            {
                p_tree.NonMerchVolCfPa = ((p_tree.DryBiot * ((100 - (double) m_scenarioHarvestMethod.WoodlandMerchAsPercentOfTotalVol) / 100)) 
                                         * p_tree.Tpa) / p_tree.OdWgt;
            }
            else
            {
                p_tree.NonMerchVolCfPa = ((p_tree.DryBiot - p_tree.DryBiom) * p_tree.Tpa) / p_tree.OdWgt;
            }

            //nonMerchWtGtPa
            p_tree.NonMerchWtGtPa = p_tree.NonMerchVolCfPa * p_tree.OdWgt / p_tree.DryToGreen / 2000;

            //brushcut
            p_tree.BrushCutVolCfPa = p_tree.DryBiot * p_tree.Tpa / p_tree.OdWgt;
            p_tree.BrushCutWtGtPa = (p_tree.DryBiot * p_tree.Tpa / p_tree.DryToGreen) / 2000;
        }

        enum OpCostTreeType
        {
            None = 0,
            BC = 1,
            CT = 2,
            SL = 3,
            LL = 4
        };
        
        ///<summary>
        ///Represents a tree in the fvs cutlist
        ///</summary>
        private class tree
        {
            string _strCondId = "";
            string _strPlotId = "";
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            string _strRxYear = "";
            double _dblDbh;
            double _dblTpa;
            double _dblVolCfNet;
            double _dblVolTsGrs;
            double _dblVolCfGrs;
            double _dblDryBiot;
            double _dblDryBiom;
            int _intSlope;
            string _strSpcd;
            bool _boolFvsCreatedTree;
            string _strFvsTreeId;
            OpCostTreeType _opCostTreeType;
            int _intSpeciesGroup;
            int _intDiamGroup;
            bool _boolIsNonCommercial;
            double _dblMerchValue;
            double _dblChipValue;
            int _intElev;
            double _dblYardingDistance;
            double _dblOdWgt;
            double _dblDryToGreen;
            double _dblMerchVolCfPa;
            double _dblMerchWtGtPa;
            double _dblBrushCutVolCfPa;
            double _dblBrushCutWtGtPa;
            double _dblNonMerchVolCfPa;
            double _dblNonMerchWtGtPa;
            double _dblMerchValDpa;
            bool _blnIsWoodlandSpecies;
            bool _blnIsCull;
            double _dblTravelTime;
            harvestMethod _objHarvestMethod;
            harvestMethod _objLowestCostHarvestMethod;

            string _strDebugFile = "";

            public tree()
			{
               
			}

            public string CondId
            {
                get { return _strCondId; }
                set { _strCondId = value; }
            }
            public string PlotId
            {
                get { return _strPlotId; }
                set { _strPlotId = value; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
                set { _strRxCycle = value; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
                set { _strRxPackage = value; }
            }            
            public string Rx
            {
                get { return _strRx; }
                set { _strRx = value; }
            }            
            public string RxYear
            {
                get { return _strRxYear; }
                set { _strRxYear = value; }
            }
            public double Dbh
            {
                get { return _dblDbh; }
                set { _dblDbh = value; }
            }
            public double Tpa
            {
                get { return _dblTpa; }
                set { _dblTpa = value; }
            }
            public double VolCfNet
            {
                get { return _dblVolCfNet; }
                set { _dblVolCfNet = value; }
            }
            public double VolTsGrs
            {
                get { return _dblVolTsGrs; }
                set { _dblVolTsGrs = value; }
            }
            public double VolCfGrs
            {
                get { return _dblVolCfGrs; }
                set { _dblVolCfGrs = value; }
            }
            public double DryBiot
            {
                get { return _dblDryBiot; }
                set { _dblDryBiot = value; }
            }
            public double DryBiom
            {
                get { return _dblDryBiom; }
                set { _dblDryBiom = value; }
            }
            public int Slope
            {
                get { return _intSlope; }
                set { _intSlope = value; }
            }
            public string SpCd
            {
                get { return _strSpcd; }
                set { _strSpcd = value; }
            }
            public bool FvsCreatedTree
            {
                get { return _boolFvsCreatedTree; }
                set { _boolFvsCreatedTree = value; }
            }
            public string FvsTreeId
            {
                get { return _strFvsTreeId; }
                set { _strFvsTreeId = value; }
            }
            public OpCostTreeType TreeType
            {
                get { return _opCostTreeType; }
                set { _opCostTreeType = value; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
                set { _intSpeciesGroup = value; }
            }
            public int DiamGroup
            {
                get { return _intDiamGroup; }
                set { _intDiamGroup = value; }
            }
            public bool IsNonCommercial
            {
                get { return _boolIsNonCommercial; }
                set { _boolIsNonCommercial = value; }
            }
            public double MerchValue
            {
                get { return _dblMerchValue; }
                set { _dblMerchValue = value; }
            }
            public double ChipValue
            {
                get { return _dblChipValue; }
                set { _dblChipValue = value; }
            }
            public int Elevation
            {
                get { return _intElev; }
                set { _intElev = value; }
            }
            public double YardingDistance
            {
                get { return _dblYardingDistance; }
                set { _dblYardingDistance = value; }
            }
            public double OdWgt
            {
                get { return _dblOdWgt; }
                set { _dblOdWgt = value; }
            }
            public double DryToGreen
            {
                get { return _dblDryToGreen; }
                set { _dblDryToGreen = value; }
            }
            public double MerchVolCfPa
            {
                get { return _dblMerchVolCfPa; }
                set { _dblMerchVolCfPa = value; }
            }
            public double MerchWtGtPa
            {
                get { return _dblMerchWtGtPa; }
                set { _dblMerchWtGtPa = value; }
            }
            public double BrushCutVolCfPa
            {
                get { return _dblBrushCutVolCfPa; }
                set { _dblBrushCutVolCfPa = value; }
            }
            public double BrushCutWtGtPa
            {
                get { return _dblBrushCutWtGtPa; }
                set { _dblBrushCutWtGtPa = value; }
            }
            public double NonMerchVolCfPa
            {
                get { return _dblNonMerchVolCfPa; }
                set { _dblNonMerchVolCfPa = value; }
            }
            public double NonMerchWtGtPa
            {
                get { return _dblNonMerchWtGtPa; }
                set { _dblNonMerchWtGtPa = value; }
            }
            public harvestMethod HarvestMethod
            {
                get { return _objHarvestMethod; }
                set { _objHarvestMethod = value; }
            }
            public harvestMethod LowestCostHarvestMethod
            {
                get { return _objLowestCostHarvestMethod; }
                set { _objLowestCostHarvestMethod = value; }
            }
            public double MerchValDpa
            {
                get { return _dblMerchValDpa; }
                set { _dblMerchValDpa = value; }
            }
            public double TotalVolCfPa
            {
                get 
                {
                    if (_opCostTreeType != OpCostTreeType.BC)
                    {
                        return _dblMerchVolCfPa + _dblNonMerchVolCfPa;
                    }
                    else
                    {
                        return _dblBrushCutVolCfPa;
                    }
                }
            }
            public double TotalWtGtPa
            {
                get
                {
                    if (_opCostTreeType != OpCostTreeType.BC)
                    {
                        return _dblMerchWtGtPa + _dblNonMerchWtGtPa;
                    }
                    else
                    {
                        return _dblBrushCutWtGtPa;
                    }
                }
            }
            public bool IsWoodlandSpecies
            {
                get { return _blnIsWoodlandSpecies; }
                set { _blnIsWoodlandSpecies = value; }
            }
            public bool IsSapling
            {
                get
                {
                    if (_dblDbh < 5.0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            public bool IsCull
            {
                get { return _blnIsCull; }
                set { _blnIsCull = value; }
            }
            public double TravelTime
            {
                get { return _dblTravelTime; }
                set {_dblTravelTime = value; }
            }

            public int BiosumCategory
            {
                get
                {
                    //if (_objLowestCostHarvestMethod != null)
                    //{
                    //    return _objLowestCostHarvestMethod.BiosumCategory;
                    //}
                    if (_objHarvestMethod != null)
                    {
                        return _objHarvestMethod.BiosumCategory;
                    }
                    else
                    {
                        return -1;
                    }
                }
            }

            public double BaFracCutNumerator
            {
                //tpa * pi * ( dbh/24)^2 
                get { return _dblTpa * Math.PI * Math.Pow((_dblDbh / 24), 2); }
            }

            public double QmdPa
            {
                //tpa * dbh ^ 2
                get { return _dblTpa * Math.Pow(_dblDbh, 2);  }
            }

            public string DebugFile
            {
                set { _strDebugFile = value; }
            }
        }

        private class treeDiamGroup
        {
            int _intDiamGroup;
            double _dblMinDiam;
            double _dblMaxDiam;

            public treeDiamGroup(int diamGroup, double dblMinDiam, double dblMaxDiam)
			{
                _intDiamGroup = diamGroup;
                _dblMinDiam = dblMinDiam;
                _dblMaxDiam = dblMaxDiam;
			}

            public int DiamGroup
            {
                get { return _intDiamGroup; }
            }
            public double MinDiam
            {
                get { return _dblMinDiam; }
            }
            public double MaxDiam
            {
                get { return _dblMaxDiam; }
            }
        }

        private class speciesDiamValue
        {
            int _intSpeciesGroup;
            int _intDiamGroup;
            string _strWoodBin;
            double _dblMerchValue;
            double _dblChipValue;

            public speciesDiamValue(int diamGroup, int speciesGroup, string woodBin, double merchValue, double chipValue)
			{
                _intDiamGroup = diamGroup;
                _intSpeciesGroup = speciesGroup;
                _strWoodBin = woodBin;
                _dblMerchValue = merchValue;
                _dblChipValue = chipValue;
			}

            public int DiamGroup
            {
                get { return _intDiamGroup; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
            }
            public string WoodBin
            {
                get { return _strWoodBin; }
            }
            public double MerchValue
            {
                get { return _dblMerchValue; }
            }
            public double ChipValue
            {
                get { return _dblChipValue; }
            }
        }

        private class scenarioHarvestMethod
        {
            double _dblMinSmallLogDbh;
            double _dblMinLargeLogDbh;
            double _dblMinChipDbh;
            int _intSteepSlopePct;
            double _dblMinDbhSteepSlope;
            harvestMethod _objHarvestMethodLowSlope;
            harvestMethod _objHarvestMethodSteepSlope;
            int _intSaplingMerchAsPercentOfTotalVol;
            int _intWoodlandMerchAsPercentOfTotalVol;
            int _intCullPctThreshold;
            HarvestMethodSelection _objHarvestMethodSelection;

            public scenarioHarvestMethod(double minChipDbh, double minSmallLogDbh, double minLargeLogDbh, int steepSlopePct,
                                         double minDbhSteepSlope,
                                         harvestMethod harvestMethodLowSlope, harvestMethod harvestMethodSteepSlope,
                                         int saplingMerchAsPercentOfTotalVol,
                                         int woodlandMerchAsPercentOfTotalVol, int cullPctThreshold, HarvestMethodSelection harvestMethodSelection)
            {
                _dblMinSmallLogDbh = minSmallLogDbh;
                _dblMinLargeLogDbh = minLargeLogDbh;
                _dblMinChipDbh = minChipDbh;
                _intSteepSlopePct = steepSlopePct;
                _dblMinDbhSteepSlope = minDbhSteepSlope;
                _objHarvestMethodLowSlope = harvestMethodLowSlope;
                _objHarvestMethodSteepSlope = harvestMethodSteepSlope;
                _intSaplingMerchAsPercentOfTotalVol = saplingMerchAsPercentOfTotalVol;
                _intWoodlandMerchAsPercentOfTotalVol = woodlandMerchAsPercentOfTotalVol;
                _intCullPctThreshold = cullPctThreshold;
                _objHarvestMethodSelection = harvestMethodSelection;
            }

            public double MinChipDbh
            {
                get { return _dblMinChipDbh; }
            }
            public double MinSmallLogDbh
            {
                get { return _dblMinSmallLogDbh; }
            }
            public double MinLargeLogDbh
            {
                get { return _dblMinLargeLogDbh; }
            }
            public int SteepSlopePct
            {
                get { return _intSteepSlopePct; }
            }
            public double MinDbhSteepSlope
            {
                get { return _dblMinDbhSteepSlope; }
            }
            public harvestMethod HarvestMethodLowSlope
            {
                get { return _objHarvestMethodLowSlope; }
            }
            public harvestMethod HarvestMethodSteepSlope
            {
                get { return _objHarvestMethodSteepSlope; }
            }
            public int SaplingMerchAsPercentOfTotalVol
            {
                get { return _intSaplingMerchAsPercentOfTotalVol; }
            }
            public int WoodlandMerchAsPercentOfTotalVol
            {
                get { return _intWoodlandMerchAsPercentOfTotalVol; }
            }
            public int CullPctThreshold
            {
                get { return _intCullPctThreshold; }
            }
            public HarvestMethodSelection HarvMethodSelection
            {
                get { return _objHarvestMethodSelection; }
            }
            // Overriding the ToString method for debugging purposes
            public override string ToString()
            {
                string strScenarioHarvestMethodLowSlope = "";
                string strScenarioHarvestMethodSteepSlope = "";
                if (_objHarvestMethodLowSlope != null)
                    strScenarioHarvestMethodLowSlope = _objHarvestMethodLowSlope.Method;
                if (_objHarvestMethodSteepSlope != null)
                    strScenarioHarvestMethodSteepSlope = _objHarvestMethodSteepSlope.Method;

                return string.Format("MinChipDbh: {0}, MinSmallLogDbh: {1}, MinLargeLogDbh: {2}, SteepSlopePct: {3}, MinDbhSteepSlope: {4}, " +
                    "HarvestMethodLowSlope: {5}, HarvestMethodSteepSlope: {6}, " +
                    "SaplingMerchAsPercentOfTotalVol: {7} WoodlandMerchAsPercentOfTotalVol: {8} CullPctThreshold: {9} " +
                    "HarvestMethodSelection: {10} ]",
                    _dblMinChipDbh, _dblMinSmallLogDbh, _dblMinLargeLogDbh, _intSteepSlopePct, _dblMinDbhSteepSlope,
                    strScenarioHarvestMethodLowSlope, strScenarioHarvestMethodSteepSlope,
                    _intSaplingMerchAsPercentOfTotalVol, _intWoodlandMerchAsPercentOfTotalVol, _intCullPctThreshold,
                    _objHarvestMethodSelection.Value);
            }
        }

        private class scenarioMoveInCost
        {
            double _dblYardDistThreshold;
            double _dblAssumedHarvestAreaAc;
            double _dblMoveInTimeMultiplier;
            double _dblMoveInHoursAddend;

            public scenarioMoveInCost(double yardDistThreshold, double assumedHarvestAreaAc, 
                                      double moveInTimeMultiplier, double moveInHoursAddend)
            {
                _dblYardDistThreshold = yardDistThreshold;
                _dblAssumedHarvestAreaAc = assumedHarvestAreaAc;
                _dblMoveInTimeMultiplier = moveInTimeMultiplier;
                _dblMoveInHoursAddend = moveInHoursAddend;
            }

            public double YardDistThreshold
            {
                get { return _dblYardDistThreshold; }
            }
            public double AssumedHarvestAreaAc
            {
                get { return _dblAssumedHarvestAreaAc; }
            }
            public double MoveInTimeMultiplier
            {
                get { return _dblMoveInTimeMultiplier; }
            }
            public double MoveInHoursAddend
            {
                get { return _dblMoveInHoursAddend; }
            }
        }

        private class prescription
        {
            string _strRx = "";
            harvestMethod _objHarvestMethodLowSlope;
            harvestMethod _objHarvestMethodSteepSlope;

            public prescription(string rx, harvestMethod harvestMethodLowSlope, harvestMethod harvestMethodSteepSlope)
            {
                _strRx = rx;
                _objHarvestMethodLowSlope = harvestMethodLowSlope;
                _objHarvestMethodSteepSlope = harvestMethodSteepSlope;
            }

            public string Rx
            {
                get { return _strRx; }
            }
            public harvestMethod HarvestMethodLowSlope
            {
                get { return _objHarvestMethodLowSlope; }
            }
            public harvestMethod HarvestMethodSteepSlope
            {
                get { return _objHarvestMethodSteepSlope; }
            }
        }

        /// <summary>
        /// An opcostInput object represents a line in the opcostInput file
        /// The metrics are aggregated by stand with is a unique concatenation of
        /// conditionId, rxPackage, rx, and rxCycle
        /// </summary>
        private class opcostInput
        {
            string _strCondId = "";
            int _intPercentSlope;
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            string _strRxYear = "";
            double _dblYardingDistance;
            int _intElev;
            harvestMethod _objHarvestMethod;
            scenarioMoveInCost _objScenarioMoveInCost;
            double _dblTotalBcTpa;
            double _dblPerAcBcVolCf;
            double _dblTotalChipTpa;
            double _dblChipNonMerchVolCfPa;
            double _dblChipMerchVolCfPa;
            double _dblChipVolCfPa;
            double _dblChipWtGtPa;
            double _dblChipHwdVolCfPa;
            double _dblTotalSmLogTpa;
            double _dblSmLogNonCommVolCfPa;
            double _dblSmLogMerchVolCfPa;
            double _dblSmLogNonCommMerchVolCfPa;
            double _dblSmLogCommNonMerchVolCfPa;
            double _dblSmLogVolCfPa;
            double _dblSmLogWtGtPa;
            double _dblSmLogHwdVolCfPa;
            double _dblTotalLgLogTpa;
            double _dblLgLogMerchVolCfPa;
            double _dblLgLogNonCommMerchVolCfPa;
            double _dblLgLogNonCommVolCfPa;
            double _dblLgLogCommNonMerchVolCfPa;
            double _dblLgLogVolCfPa;
            double _dblLgLogWtGtPa;
            double _dblLgLogHwdVolCfPa;
            double _dblMoveInHours;
            double _dblHarvestAreaAssumedAc;
            double _dblYardingDistanceUnadj;
            double _dblTotalSmLogTpaUnadj;
            double _dblTotalLgLogTpaUnadj;
            double _dblTotalBaFracCutNumerator;
            double _dblBaFracCut = -1;
            double _dblSmLogQmdPa;
            double _dblLgLogQmdPa;

            public opcostInput(string condId, int percentSlope, string rxCycle, string rxPackage, string rx,
                               string rxYear, double yardingDistance, int elev, harvestMethod harvestMethod, double moveInHours,
                               scenarioMoveInCost scenarioMoveInCost)
            {
                _strCondId = condId;
                _intPercentSlope = percentSlope;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _strRxYear = rxYear;
                _dblYardingDistance = yardingDistance;
                _intElev = elev;
                _objHarvestMethod = harvestMethod;
                _dblMoveInHours = moveInHours;
                _dblHarvestAreaAssumedAc = scenarioMoveInCost.AssumedHarvestAreaAc;
                _objScenarioMoveInCost = scenarioMoveInCost;

                // Apply move-in costs yarding threshold; Note that this implementation doesn't record the
                // original yarding distance when adjusting to move-in yarding threshold
                if (_dblYardingDistance < _objScenarioMoveInCost.YardDistThreshold)
                    _dblYardingDistance = _objScenarioMoveInCost.YardDistThreshold;

                
                //Apply yarding distance minimum
                if (_dblYardingDistance < _objHarvestMethod.MinYardDistanceFt)
                {
                    _dblYardingDistanceUnadj = _dblYardingDistance;
                    _dblYardingDistance = _objHarvestMethod.MinYardDistanceFt;
                }
            }

            public string OpCostStand    
            {
                get { return _strCondId + _strRxPackage + _strRx + _strRxCycle; }
            }
            public int PercentSlope
            {
                get { return _intPercentSlope; }
            }
            public double YardingDistance
            {
                get { return _dblYardingDistance; }
            }
            public string RxYear
            {
                get { return _strRxYear; }
            }
            public int Elev
            {
                get { return _intElev; }
            }
            public string RxPackageRxRxCycle
            {
                get { return _strRxPackage + _strRx + _strRxCycle; }
            }
            public harvestMethod HarvestMethod
            {
                get { return _objHarvestMethod; }
            }
            public double TotalBcTpa
            {
                set { _dblTotalBcTpa = value; }
                get { return _dblTotalBcTpa; }
            }
            public double PerAcBcVolCf
            {
                set { _dblPerAcBcVolCf = value; }
                get { return _dblPerAcBcVolCf; }
            }
            public double TotalChipTpa
            {
                set { _dblTotalChipTpa = value; }
                get { return _dblTotalChipTpa; }
            }
            public double ChipMerchVolCfPa
            {
                set { _dblChipMerchVolCfPa = value; }
                get { return _dblChipMerchVolCfPa; }
            }
            public double ChipNonMerchVolCfPa
            {
                set { _dblChipNonMerchVolCfPa = value; }
                get { return _dblChipNonMerchVolCfPa; }
            }
            public double ChipVolCfPa
            {
                set { _dblChipVolCfPa = value; }
                get { return _dblChipVolCfPa; }
            }
            public double ChipWtGtPa
            {
                set { _dblChipWtGtPa = value; }
                get { return _dblChipWtGtPa; }
            }
            public double ChipHwdVolCfPa
            {
                set { _dblChipHwdVolCfPa = value; }
                get { return _dblChipHwdVolCfPa; }
            }
            public double TotalSmLogTpa
            {
                set { _dblTotalSmLogTpa = value; }
                get { return _dblTotalSmLogTpa; }
            }
            public double SmLogMerchVolCfPa
            {
                set { _dblSmLogMerchVolCfPa = value; }
                get { return _dblSmLogMerchVolCfPa; }
            }
            public double SmLogNonCommVolCfPa
            {
                set { _dblSmLogNonCommVolCfPa = value; }
                get { return _dblSmLogNonCommVolCfPa; }
            }
            public double SmLogNonCommMerchVolCfPa
            {
                set { _dblSmLogNonCommMerchVolCfPa = value; }
                get { return _dblSmLogNonCommMerchVolCfPa; }
            }
            public double SmLogCommNonMerchVolCfPa
            {
                set { _dblSmLogCommNonMerchVolCfPa = value; }
                get { return _dblSmLogCommNonMerchVolCfPa; }
            }
            public double SmLogVolCfPa
            {
                set { _dblSmLogVolCfPa = value; }
                get { return _dblSmLogVolCfPa; }
            }
            public double SmLogWtGtPa
            {
                set { _dblSmLogWtGtPa = value; }
                get { return _dblSmLogWtGtPa; }
            }
            public double SmLogHwdVolCfPa
            {
                set { _dblSmLogHwdVolCfPa = value; }
                get { return _dblSmLogHwdVolCfPa; }
            }
            public double TotalLgLogTpa
            {
                set { _dblTotalLgLogTpa = value; }
                get { return _dblTotalLgLogTpa; }
            }
            public double LgLogNonCommMerchVolCfPa
            {
                set { _dblLgLogNonCommMerchVolCfPa = value; }
                get { return _dblLgLogNonCommMerchVolCfPa; }
            }
            public double LgLogMerchVolCfPa
            {
                set { _dblLgLogMerchVolCfPa = value; }
                get { return _dblLgLogMerchVolCfPa; }
            }
            public double LgLogNonCommVolCfPa
            {
                set { _dblLgLogNonCommVolCfPa = value; }
                get { return _dblLgLogNonCommVolCfPa; }
            }
            public double LgLogCommNonMerchVolCfPa
            {
                set { _dblLgLogCommNonMerchVolCfPa = value; }
                get { return _dblLgLogCommNonMerchVolCfPa; }
            }
            public double LgLogVolCfPa
            {
                set { _dblLgLogVolCfPa = value; }
                get { return _dblLgLogVolCfPa; }
            }
            public double LgLogWtGtPa
            {
                set { _dblLgLogWtGtPa = value; }
                get { return _dblLgLogWtGtPa; }
            }
            public double LgLogHwdVolCfPa
            {
                set { _dblLgLogHwdVolCfPa = value; }
                get { return _dblLgLogHwdVolCfPa; }
            }
            public string CondId
            {
                set { _strCondId = value; }
                get { return _strCondId; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
            }
            public string Rx
            {
                get { return _strRx; }
            }
            public double MoveInHours
            {
                get { return _dblMoveInHours; }
            }
            public double HarvestAreaAssumedAc
            {
                get { return _dblHarvestAreaAssumedAc; }
            }
            public double YardingDistanceUnadj
            {
                get { return _dblYardingDistanceUnadj; }
            }
            public double TotalSmLogTpaUnadj
            {
                set { _dblTotalSmLogTpaUnadj = value; }
                get { return _dblTotalSmLogTpaUnadj; }
            }
            public double TotalLgLogTpaUnadj
            {
                set { _dblTotalLgLogTpaUnadj = value; }
                get { return _dblTotalLgLogTpaUnadj; }
            }
            public double TotalBaFracCutNumerator
            {
                set { _dblTotalBaFracCutNumerator = value; }
                get { return _dblTotalBaFracCutNumerator; }
            }
            public double BaFracCut
            {
                set { _dblBaFracCut = value; }
                get { return _dblBaFracCut; }
            }
            public double TotalSmLogQmdPa
            {
                set { _dblSmLogQmdPa = value; }
                get { return _dblSmLogQmdPa; }
            }
            public double TotalLgLogQmdPa
            {
                set { _dblLgLogQmdPa = value; }
                get { return _dblLgLogQmdPa; }
            }
            public double QMD_SL
            {
                get 
                {
                    // We may have overwritten the actual tpa value with the adjusted tpa value so we have to
                    // check before using it
                    double dblTotalTpa = _dblTotalSmLogTpa;
                    if (_dblTotalSmLogTpaUnadj > 0)
                    {
                        dblTotalTpa = _dblTotalSmLogTpaUnadj;
                    }
                    if (_dblSmLogQmdPa == 0 || dblTotalTpa == 0)
                    {
                        return 0;
                    }
                    else
                    {
                        return Math.Sqrt(_dblSmLogQmdPa / dblTotalTpa);
                    }
                }
            }
            public double QMD_LL
            {
                get
                {
                    double dblTotalTpa = _dblTotalLgLogTpa;
                    if (_dblTotalLgLogTpaUnadj > 0)
                    {
                        dblTotalTpa = _dblTotalLgLogTpaUnadj;
                    }
                    if (_dblLgLogQmdPa == 0 || dblTotalTpa == 0)
                    {
                        return 0;
                    }
                    else
                    {
                        return Math.Sqrt(_dblLgLogQmdPa / dblTotalTpa);
                    }
                }
            }

        }

        /// <summary>
        /// An opcost_ideal object represents a line in the opcost_ideal_output table
        /// This table is linked back to a tree using condId, rxCycle, rxPackage, and rx
        /// </summary>
        private class opcostIdeal
        {
            string _strCondId = "";
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            string _strHarvestSystem = "";

            public opcostIdeal(string condId, string rxCycle, string rxPackage, string rx,
                               string harvestSystem)
            {
                _strCondId = condId;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _strHarvestSystem = harvestSystem;
            }

            public string CondId
            {
                get { return _strCondId; }
            }
            public string Rx
            {
                get { return _strRx; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
            }
            public string HarvestSystem
            {
                get { return _strHarvestSystem; }
            }
        }

        /// <summary>
        /// An treeVolValInput object represents a line in the tree vol val file
        /// The metrics are aggregated by conditionId, rxCycle, speciesGroup, and diameterGroup
        /// </summary>
        private class treeVolValInput
        {
            string _strCondId = "";
            string _strRxCycle = "";
            string _strRxPackage = "";
            string _strRx = "";
            int _intSpeciesGroup;
            int _intDiamGroup;
            int _intHarvestMethodCategory;
            string _strMerchToChip;
            double _dblChipMktValPgt;
            double _dblTotalrushCutVolCfPa;
            double _dblTotalBrushCutWtGtPa;
            double _dblStandResidueWtGtPa;
            double _dblChipVolCfPa;
            double _dblChipWtGtPa;
            double _dblTotalMerchVolCfPa;
            double _dblTotalMerchWtGtPa;
            double _dblTotalMerchValDpa;


            public treeVolValInput(string condId, string rxCycle, string rxPackage, string rx,
                                    int speciesGroup, int diamGroup, bool isNonCommercial,
                                    double chipMktValPgt, int harvestMethodCategory)
            {
                _strCondId = condId;
                _strRxCycle = rxCycle;
                _strRxPackage = rxPackage;
                _strRx = rx;
                _intSpeciesGroup = speciesGroup;
                _intDiamGroup = diamGroup;
                _intHarvestMethodCategory = harvestMethodCategory;
                if (isNonCommercial)
                {
                    _strMerchToChip = "Y";
                }
                else
                {
                    _strMerchToChip = "N";
                }
                _dblChipMktValPgt = chipMktValPgt;
            }
            
            public string CondId
            {
                get { return _strCondId; }
            }
            public string Rx
            {
                get { return _strRx; }
            }
            public string RxCycle
            {
                get { return _strRxCycle; }
            }
            public string RxPackage
            {
                get { return _strRxPackage; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
                set { _intSpeciesGroup = value; }
            }
            public int DiamGroup
            {
                get { return _intDiamGroup; }
                set { _intDiamGroup = value; }
            }
            public string MerchToChip
            {
                get { return _strMerchToChip; }
            }
            public double ChipMktValPgt
            {
                get { return _dblChipMktValPgt; }
            }
            public int HarvestMethodCategory
            {
                get { return _intHarvestMethodCategory; }
            }
            public double TotalBrushCutVolCfPa
            {
                get { return _dblTotalrushCutVolCfPa; }
                set { _dblTotalrushCutVolCfPa = value; }
            }
            public double TotalBrushCutWtGtPa
            {
                get { return _dblTotalBrushCutWtGtPa; }
                set { _dblTotalBrushCutWtGtPa = value; }
            }
            public double StandResidueWtGtPa
            {
                get { return _dblStandResidueWtGtPa; }
                set { _dblStandResidueWtGtPa = value; }
            }
            public double ChipVolCfPa
            {
                get { return _dblChipVolCfPa; }
                set { _dblChipVolCfPa = value; }
            }
            public double ChipWtGtPa
            {
                get { return _dblChipWtGtPa; }
                set { _dblChipWtGtPa = value; }
            }
            public double TotalMerchVolCfPa
            {
                get { return _dblTotalMerchVolCfPa; }
                set { _dblTotalMerchVolCfPa = value; }
            }
            public double TotalMerchWtGtPa
            {
                get { return _dblTotalMerchWtGtPa; }
                set { _dblTotalMerchWtGtPa = value; }
            }
            public double TotalMerchValDpa
            {
                get { return _dblTotalMerchValDpa; }
                set { _dblTotalMerchValDpa = value; }
            }

        }

        private class treeSpecies
        {
            string _strSpcd = "";
            int _intSpeciesGroup;
            double _dblOdWgt;
            double _dblDryToGreen;
            bool _blnIsWoodlandSpecies;

            public treeSpecies(string spCd, int speciesGroup, double odWgt, double dryToGreen, bool isWoodlandSpecies)
            {
                _strSpcd = spCd;
                _intSpeciesGroup = speciesGroup;
                _dblOdWgt = odWgt;
                _dblDryToGreen = dryToGreen;
                _blnIsWoodlandSpecies = isWoodlandSpecies;
            }

            public string Spcd
            {
                get { return _strSpcd; }
            }
            public int SpeciesGroup
            {
                get { return _intSpeciesGroup; }
            }
            public double OdWgt
            {
                get { return _dblOdWgt; }
            }
            public double DryToGreen
            {
                get { return _dblDryToGreen; }
            }
            public bool IsWoodlandSpecies
            {
                get { return _blnIsWoodlandSpecies; }
            }
        }

        private class escalators
        {
            double _dblEnergyWoodRevCycle2;
            double _dblEnergyWoodRevCycle3;
            double _dblEnergyWoodRevCycle4;
            double _dblMerchWoodRevCycle2;
            double _dblMerchWoodRevCycle3;
            double _dblMerchWoodRevCycle4;


            public escalators(double energyWoodRevCycle2, double energyWoodRevCycle3, double energyWoodRevCycle4,
                              double merchWoodRevCycle2, double merchWoodRevCycle3, double merchWoodRevCycle4)
            {
                _dblEnergyWoodRevCycle2 = energyWoodRevCycle2;
                _dblEnergyWoodRevCycle3 = energyWoodRevCycle3;
                _dblEnergyWoodRevCycle4 = energyWoodRevCycle4;
                _dblMerchWoodRevCycle2 = merchWoodRevCycle2;
                _dblMerchWoodRevCycle3 = merchWoodRevCycle3;
                _dblMerchWoodRevCycle4 = merchWoodRevCycle4;
            }
            
            public double EnergyWoodRevCycle2
            {
                get { return _dblEnergyWoodRevCycle2; }
            }
            public double EnergyWoodRevCycle3
            {
                get { return _dblEnergyWoodRevCycle3; }
            }
            public double EnergyWoodRevCycle4
            {
                get { return _dblEnergyWoodRevCycle4; }
            }
            public double MerchWoodRevCycle2
            {
                get { return _dblMerchWoodRevCycle2; }
            }
            public double MerchWoodRevCycle3
            {
                get { return _dblMerchWoodRevCycle3; }
            }
            public double MerchWoodRevCycle4
            {
                get { return _dblMerchWoodRevCycle4; }
            }
        }

        private class harvestMethod
        {
            bool _blnSteepSlope;
            string _strMethod;
            int _intBiosumCategory;
            double _dblMinTpa;
            double _dblMinYardDistanceFt;
            double _dblMinAvgTreeVolCf;

            public harvestMethod(bool steepSlope, string method, int biosumCategory, double minTpa, double minYardDistanceFt,
                                 double minAvgTreeVolCf)
            {
                _blnSteepSlope = steepSlope;
                _strMethod = method;
                _intBiosumCategory = biosumCategory;
                _dblMinTpa = minTpa;
                _dblMinYardDistanceFt = minYardDistanceFt;
                _dblMinAvgTreeVolCf = minAvgTreeVolCf;
            }

            public bool SteepSlope
            {
                get { return _blnSteepSlope; }
            }
            public string Method
            {
                get { return _strMethod; }
            }
            public int BiosumCategory
            {
                get { return _intBiosumCategory; }
            }
            public double MinTpa
            {
                get { return _dblMinTpa; }
            }
            public double MinYardDistanceFt
            {
                get { return _dblMinYardDistanceFt; }
            }
            public double MinAvgTreeVolCf
            {
                get { return _dblMinAvgTreeVolCf; }
            }
        }

        private System.Collections.Generic.IDictionary<String, opcostIdeal> LoadOpcostIdeal()
        {
            System.Collections.Generic.IDictionary<String, opcostIdeal> dictOpcostIdeal =
                new System.Collections.Generic.Dictionary<String, opcostIdeal>();
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT * FROM " +
                                m_strOpcostIdealTableName;
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strCondId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_cond_id"]).Trim();
                        string strRxPackage = Convert.ToString(m_oAdo.m_OleDbDataReader["rxPackage"]).Trim();
                        string strRx = Convert.ToString(m_oAdo.m_OleDbDataReader["rx"]).Trim();
                        string strRxCycle = Convert.ToString(m_oAdo.m_OleDbDataReader["rxCycle"]).Trim();
                        string strHarvestSystem = Convert.ToString(m_oAdo.m_OleDbDataReader["harvest_system"]).Trim();

                        string key = strCondId + strRxPackage + strRx + strRxCycle;
                        opcostIdeal newIdeal = new opcostIdeal(strCondId, strRxCycle, strRxPackage, strRx, strHarvestSystem);
                        dictOpcostIdeal.Add(key, newIdeal);
                    }
                }
            }
            return dictOpcostIdeal;
        }

        private System.Collections.Generic.IDictionary<String, double> LoadTravelTimes(string p_strTravelTimesTableName)
        {
            System.Collections.Generic.IDictionary<String, double> dictTravelTimes =
                new System.Collections.Generic.Dictionary<String, double>();
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT MIN(ONE_WAY_HOURS) AS min_one_way_hours, BIOSUM_PLOT_ID " +
                                "FROM " + p_strTravelTimesTableName +
                                " WHERE ONE_WAY_HOURS > 0 " +
                                "GROUP BY BIOSUM_PLOT_ID";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strPlotId = Convert.ToString(m_oAdo.m_OleDbDataReader["biosum_plot_id"]).Trim();
                        double dblTravelTime = Convert.ToDouble(m_oAdo.m_OleDbDataReader["min_one_way_hours"]);
                        dictTravelTimes.Add(strPlotId, dblTravelTime);
                    }
                }
            }
            return dictTravelTimes;
        }

        private System.Collections.Generic.IDictionary<String, double> LoadFvsPreBasalArea(string p_strVariant)
        {
            System.Collections.Generic.IDictionary<String, double> dictPreBasalArea =
                new System.Collections.Generic.Dictionary<String, double>();
            if (m_oAdo.m_intError == 0)
            {
                string strSQL = "SELECT TRIM(biosum_cond_id) + TRIM(rxpackage)  + TRIM(rx) + TRIM(rxcycle) as [OpCostStandId], BA" +
                                " FROM " + Tables.FVS.DefaultPreFVSSummaryTableName + 
                                " WHERE fvs_variant = '" + p_strVariant + "'";
                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, strSQL);
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        string strOpCostStandId = Convert.ToString(m_oAdo.m_OleDbDataReader["OpCostStandId"]).Trim();
                        double dblBasalArea = Convert.ToDouble(m_oAdo.m_OleDbDataReader["BA"]);
                        dictPreBasalArea.Add(strOpCostStandId, dblBasalArea);
                    }
                }
            }
            return dictPreBasalArea;
        }

    }
}

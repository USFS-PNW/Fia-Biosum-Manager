using System;
using System.Collections.Generic;

namespace FIA_Biosum_Manager
{
    /// <summary>
    /// Summary description for Tables.
    /// </summary>
    public class Tables
    {
        public Project m_oProject = new Project();
        public OptimizerScenarioResults m_oOptimizerScenarioResults = new OptimizerScenarioResults();
        public OptimizerDefinitions m_oOptimizerDef = new OptimizerDefinitions();
        public OptimizerScenarioRuleDefinitions m_oOptimizerScenarioRuleDef = new OptimizerScenarioRuleDefinitions();
        public FIAPlot m_oFIAPlot = new FIAPlot();
        public FVS m_oFvs = new FVS();
        public TravelTime m_oTravelTime = new TravelTime();
        public Processor m_oProcessor = new Processor();
        public Scenario m_oScenario = new Scenario();
        public Audit m_oAudit = new Audit();
        public Reference m_oReference = new Reference();
        public ProcessorScenarioRun m_oProcessorScenarioRun = new ProcessorScenarioRun();
        public ProcessorScenarioRuleDefinitions m_oProcessorScenarioRuleDefinitions = new ProcessorScenarioRuleDefinitions();
        public VolumeAndBiomass m_oVolumeAndBiomass = new VolumeAndBiomass();


        public Tables()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public class Project
        {
            static public string DefaultProjectTableName { get { return "project"; } }
            static public string DefaultProjectDatasourceTableName { get { return "datasource"; } }
            public Project()
            {
            }
            public void CreateDatasourceTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateDatasourceTableSQL(p_strTableName));
            }
            public string CreateDatasourceTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "table_type CHAR(60)," +
                    "path CHAR(254)," +
                    "file CHAR(50)," +
                    "table_name CHAR(50))";
            }

            public void CreateProjectTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateProjectTableSQL(p_strTableName));
            }
            public string CreateProjectTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "proj_id CHAR(100) PRIMARY KEY," +
                    "created_by CHAR(30)," +
                    "created_date CHAR(22)," +
                    "organization CHAR(100)," +
                    "description CHAR(1200)," +
                    "notes CHAR(255)," +
                    "project_directory CHAR(254)," +
                    "application_version CHAR(11))";     
            }


           
        }
        public class OptimizerScenarioResults
        {

            static public string DefaultScenarioResultsEffectiveTableSuffix { get { return "_effective"; } }
            static public string DefaultScenarioResultsBestRxSummaryTableSuffix { get { return "_best_rx_summary"; } }
            static public string DefaultScenarioResultsBestRxSummaryBeforeTiebreaksTableSuffix { get { return "_best_rx_summary_before_tiebreaks"; } }
            static public string DefaultScenarioResultsOptimizationTableSuffix { get { return "_optimization"; } }
            static public string DefaultScenarioResultsTieBreakerTableName { get { return "tiebreaker"; } }

            //other
            static public string DefaultScenarioResultsValidCombosFVSPrePostTableName { get { return "validcombos_fvsprepost"; } }
            static public string DefaultScenarioResultsValidCombosFVSPreTableName { get { return "validcombos_fvspre"; } }
            static public string DefaultScenarioResultsValidCombosFVSPostTableName { get { return "validcombos_fvspost"; } }
            static public string DefaultScenarioResultsValidCombosTableName { get { return "validcombos"; } }
            static public string DefaultScenarioResultsTreeVolValSumTableName { get { return "tree_vol_val_sum_by_rx_cycle_work"; } }
            static public string DefaultCalculatedPrePostFVSVariableTableDbFile { get { return @"optimizer\db\prepost_fvs_weighted.db"; } }
            static public string DefaultScenarioResultsPostEconomicWeightedTableName { get { return @"post_economic_weighted"; } }
            static public string DefaultScenarioResultsDbFile { get { return @"db\optimizer_results.db"; } }
            static public string DefaultScenarioResultsEconByRxCycleTableName { get { return @"econ_by_rx_cycle"; } }
            static public string DefaultScenarioResultsEconByRxUtilSumTableName { get { return @"econ_by_rx_utilized_sum"; } }
            static public string DefaultScenarioResultsPSiteAccessibleWorkTableName { get { return @"psite_accessible_work_table"; } }
            static public string DefaultScenarioResultsHaulCostsTableName { get { return @"haul_costs"; } }
            static public string DefaultScenarioResultsCondPsiteTableName { get { return @"cond_psite"; } }
            static public string DefaultScenarioResultsContextDbFile { get { return @"db\context.db"; } }
            static public string DefaultScenarioResultsHarvestMethodRefTableName { get { return @"harvest_method_ref_C"; } }
            static public string DefaultScenarioResultsRxPackageRefTableName { get { return @"rxpackage_ref_C"; } }
            static public string DefaultScenarioResultsDiameterSpeciesGroupRefTableName { get { return @"diameter_spp_grp_ref_C"; } }
            static public string DefaultScenarioResultsSpeciesGroupRefTableName { get { return @"spp_grp_ref_C"; } }
            static public string DefaultScenarioResultsFvsWeightedVariablesRefTableName { get { return @"fvs_weighted_variables_ref_C"; } }
            static public string DefaultScenarioResultsEconWeightedVariablesRefTableName { get { return @"econ_weighted_variables_ref_C"; } }
            static public string DefaultScenarioResultsVersionTableName { get { return @"version"; } }

            public OptimizerScenarioResults()
            {
            }
            //
            //EFFECTIVE TABLE
            //
            public void CreateEffectiveTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn,
                                                   string p_strTable, string p_strFilterColumnName)
            {
                string strTableName = p_strTable;
                p_oDataMgr.SqlNonQuery(p_oConn, CreateEffectiveTableSQL(strTableName, p_strFilterColumnName));
            }
            static public string CreateEffectiveTableSQL(string p_strTableName, string p_strFilterColumnName)
            {
                string strSql = "CREATE TABLE " + p_strTableName + " (" +
                         "biosum_cond_id CHAR(25)," +
                         "rxpackage CHAR(3)," +
                         "rx CHAR(3)," +
                         "rxcycle CHAR(1),";
                         //"nr_dpa DOUBLE,";
                if (!String.IsNullOrEmpty(p_strFilterColumnName))
                {
                    strSql += p_strFilterColumnName + " DOUBLE,";
                }
                strSql += "pre_variable1_name CHAR(100)," +
                         "post_variable1_name CHAR(100)," +
                         "pre_variable1_value DOUBLE," +
                         "post_variable1_value DOUBLE," +
                         "variable1_change DOUBLE," +
                         "variable1_better_yn CHAR(1)," +
                         "variable1_worse_yn CHAR(1)," +
                         "variable1_effective_yn CHAR(1)," +
                         "pre_variable2_name CHAR(100)," +
                         "post_variable2_name CHAR(100)," +
                         "pre_variable2_value DOUBLE," +
                         "post_variable2_value DOUBLE," +
                         "variable2_change DOUBLE," +
                         "variable2_better_yn CHAR(1)," +
                         "variable2_worse_yn CHAR(1)," +
                         "variable2_effective_yn CHAR(1)," +
                         "pre_variable3_name CHAR(100)," +
                         "post_variable3_name CHAR(100)," +
                         "pre_variable3_value DOUBLE," +
                         "post_variable3_value DOUBLE," +
                         "variable3_change DOUBLE," +
                         "variable3_better_yn CHAR(1)," +
                         "variable3_worse_yn CHAR(1)," +
                         "variable3_effective_yn CHAR(1)," +
                         "pre_variable4_name CHAR(100)," +
                         "post_variable4_name CHAR(100)," +
                         "pre_variable4_value DOUBLE," +
                         "post_variable4_value DOUBLE," +
                         "variable4_change DOUBLE," +
                         "variable4_better_yn CHAR(1)," +
                         "variable4_worse_yn CHAR(1)," +
                         "variable4_effective_yn CHAR(1)," +
                         "overall_effective_yn CHAR(1)," +
                         "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";

                return strSql;
            }
            
            //
            //TIE BREAKER TABLE
            //
            public void CreateTieBreakerTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTieBreakerTableSQL(p_strTableName));
            }
            static public string CreateTieBreakerTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "last_tiebreak_rank INTEGER," +
                    "pre_variable1_name CHAR(100)," +
                    "post_variable1_name CHAR(100)," +
                    "pre_variable1_value DOUBLE," +
                    "post_variable1_value DOUBLE," +
                    "variable1_change DOUBLE," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }
            //
            //VALID COMBO TABLE
            //
            public void CreateValidComboTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateValidComboTableSQL(p_strTableName));
            }
            static public string CreateValidComboTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }
          
            public void CreateValidComboFVSPostTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateValidComboFVSPostTableSQL(p_strTableName));
            }
            static public string CreateValidComboFVSPostTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "variable1_yn CHAR(1)," +
                    "variable2_yn CHAR(1)," +
                    "variable3_yn CHAR(1)," +
                    "variable4_yn CHAR(1)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }
            
            public void CreateValidComboFVSPreTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateValidComboFVSPreTableSQL(p_strTableName));
            }
            static public string CreateValidComboFVSPreTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "variable1_yn CHAR(1)," +
                    "variable2_yn CHAR(1)," +
                    "variable3_yn CHAR(1)," +
                    "variable4_yn CHAR(1)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }
            public void CreateValidComboFVSPrePostTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateValidComboFVSPrePostTableSQL(p_strTableName));
            }
            static public string CreateValidComboFVSPrePostTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "fvs_variant CHAR(2)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }
            //
            //BEST TREATMENT TABLE
            //
            public void CreateBestRxSummaryTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateBestRxSummaryTableSQL(p_strTableName));
            }
            static public string CreateBestRxSummaryTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "acres double," +
                    "owngrpcd INTEGER," +
                    "optimization_value DOUBLE," +
                    "tiebreaker_value DOUBLE," +
                    "last_tiebreak_rank INTEGER," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rx))";
            }
            
            static public string CreateBestRxSummaryCycle1TieBreakerTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "acres DOUBLE," +
                    "owngrpcd INTEGER," +
                    "optimization_value DOUBLE," +
                    "tiebreaker_value DOUBLE," +
                    "last_tiebreak_rank INTEGER)";
            }


            //
            //OPTIMIZATION VARIABLE
            //
            public void CreateOptimizationTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn,
                                    string p_strTableName, string p_strFilterColumnName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateOptimizationTableSQL(p_strTableName, p_strFilterColumnName));
            }
            static public string CreateOptimizationTableSQL(string p_strTableName, string p_strFilterColumnName)
            {
                string strSQL = "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "pre_variable_name CHAR(100)," +
                    "post_variable_name CHAR(100)," +
                    "pre_variable_value DOUBLE," +
                    "post_variable_value DOUBLE," +
                    "change_value DOUBLE," +
                    "affordable_YN CHAR(1),";
                if (!String.IsNullOrEmpty(p_strFilterColumnName))
                {
                    strSQL += p_strFilterColumnName + " DOUBLE,";
                }
                strSQL += "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";

                return strSQL;
            }
            //
            //INTENSITY WORK TABLE
            //
            public void CreateIntensityWorkTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateIntensityWorkTableSQL(p_strTableName));
            }
            static public string CreateIntensityWorkTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "number_value DOUBLE," +
                    "number_value2 DOUBLE," +
                    "min_intensity INTEGER," +
                    "PRIMARY KEY (biosum_cond_id,rxpackage,rx,rxcycle))";
            }
            //
            //HAUL COST TABLE
            //
            public void CreateHaulCostTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateHaulCostTableSQL(p_strTableName));
                CreateHaulCostTableIndexes(p_oDataMgr, p_oConn, p_strTableName);


            }
            public void CreateHaulCostTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "psite_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "railhead_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "biosum_plot_id");

            }

            static public string CreateHaulCostTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "haul_cost_id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "biosum_plot_id CHAR(24)," +
                    "railhead_id INTEGER," +
                    "psite_id INTEGER," +
                    "transfer_cost_dpgt DOUBLE DEFAULT 0," +
                    "road_cost_dpgt DOUBLE DEFAULT 0," +
                    "rail_cost_dpgt DOUBLE DEFAULT 0," +
                    "complete_haul_cost_dpgt DOUBLE DEFAULT 0," +
                    "materialcd CHAR(2))";
            }
            public void CreateHaulCostRailroadTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateHaulCostRailroadTableSQL(p_strTableName));
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "psite_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "railhead_id");
            }
            static public string CreateHaulCostRailroadTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "railhead_id INTEGER," +
                    "psite_id INTEGER," +
                    "transfer_cost_dpgt DOUBLE DEFAULT 0," +
                    "road_cost_dpgt DOUBLE DEFAULT 0," +
                    "rail_cost_dpgt DOUBLE DEFAULT 0," +
                    "complete_haul_cost_dpgt DOUBLE DEFAULT 0," +
                    "materialcd CHAR(2)," +
                    "PRIMARY KEY(psite_id, railhead_id))";
            }
            //
            //TREE VOLUME AND VALUE SUM BY RX TABLE
            //
            public void CreateTreeVolValSumTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTreeVolValSumTableSQL(p_strTableName));
            }
            static public string CreateTreeVolValSumTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "chip_vol_cf DOUBLE," +
                    "chip_wt_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "merch_vol_cf DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "place_holder CHAR(1) DEFAULT 'N'," +
                    "PRIMARY KEY (biosum_cond_id,rxpackage,rx,rxcycle))";
            }

            //
            //PRODUCT YIELDS NET REVENUE/COSTS SUMMARY TABLE
            //
            public void CreateProductYieldsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateProductYieldsTableSQL(p_strTableName));
            }
            static public string CreateProductYieldsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "chip_vol_cf DOUBLE," +
                    "merch_vol_cf DOUBLE," +
                    "chip_wt_gt DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "harvest_onsite_cost_dpa DOUBLE," +
                    "chip_haul_cost_dpa DOUBLE," +
                    "merch_haul_cost_dpa DOUBLE," +
                    "merch_chip_nr_dpa DOUBLE," +
                    "merch_nr_dpa DOUBLE," +
                    "usebiomass_yn CHAR(1)," +
                    "max_nr_dpa DOUBLE," +
                    "acres DOUBLE," +
                    "owngrpcd INTEGER," +
                    "haul_costs_dpa CHAR(255)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }

            //
            //PRODUCT YIELDS NET REVENUE/COSTS SUMMARY BY PACKAGE TABLE
            //
            public void CreateEconByRxUtilSumTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateEconByRxUtilTableSQL(p_strTableName));
            }
            static public string CreateEconByRxUtilTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "chip_vol_cf_utilized DOUBLE," +
                    "merch_vol_cf DOUBLE," +
                    "chip_wt_gt_utilized DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "chip_val_dpa_utilized DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "harvest_onsite_cost_dpa DOUBLE," +
                    "chip_haul_cost_dpa_utilized DOUBLE," +
                    "merch_haul_cost_dpa DOUBLE," +
                    "merch_chip_nr_dpa DOUBLE," +
                    "merch_nr_dpa DOUBLE," +
                    "max_nr_dpa DOUBLE," +
                    "acres DOUBLE," +
                    "treated_acres DOUBLE," +
                    "owngrpcd INTEGER," +
                    "haul_costs_dpa CHAR(255)," +
                    "hvst_type_by_cycle CHAR(4)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage))";
            }
            //
            //POST ECONOMIC WEIGHTED TABLE
            //
            public void CreatePostEconomicWeightedTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePostEconomicWeightedTableSQL(p_strTableName));
            }
            static public string CreatePostEconomicWeightedTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "biosum_cond_id CHAR(25), " +
                       "rxpackage CHAR(3) )";
            }
            //
            //PSITES WORKTABLE
            //
            static public string CreatePSitesWorktableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_plot_id CHAR(24)," +
                    "biosum_cond_id CHAR(25)," +
                    "merch_haul_cost_id INTEGER," +
                    "merch_haul_psite INTEGER," +
                    "merch_haul_psite_name CHAR(255)," +
                    "merch_haul_cost_dpgt DOUBLE DEFAULT 0," +
                    "chip_haul_cost_id INTEGER," +
                    "chip_haul_psite INTEGER," +
                    "chip_haul_psite_name CHAR(255)," +
                    "chip_haul_cost_dpgt DOUBLE DEFAULT 0," +
                    "cond_too_far_steep_yn CHAR(1) DEFAULT 'N'," +
                    "cond_accessible_yn CHAR(1) DEFAULT 'Y'," +
                    "PRIMARY KEY (biosum_plot_id, biosum_cond_id))";
            }
            public void CreatePSitesWorktable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePSitesWorktableSQL(p_strTableName));
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_plot_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_cond_id");
            }

            //
            //COND_PSITE TABLE
            //
            public void CreateCondPsiteTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateCondPsiteTableSQL(p_strTableName));
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
            }
            static public string CreateCondPsiteTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "BIOSUM_COND_ID CHAR(25)," +
                    "MERCH_PSITE_NUM INTEGER," +
                    "MERCH_PSITE_NAME CHAR(255)," +
                    "CHIP_PSITE_NUM INTEGER," +
                    "CHIP_PSITE_NAME CHAR(255)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (BIOSUM_COND_ID))";
            }

            //
            //VERSION TABLE
            //
            static public string CreateVersionTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "APPLICATION_VERSION CHAR(25))";
            }
            public void CreateVersionTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateVersionTableSQL(p_strTableName));
            }

            //
            //HARVEST_METHOD_REF TABLE
            //
            public void CreateHarvestMethodRefTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateHarvestMethodRefTableSQL(p_strTableName));
            }

            static public string CreateHarvestMethodRefTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "RX CHAR(3)," +
                    "RX_HARVEST_METHOD_LOW CHAR(50)," +
                    "RX_HARVEST_METHOD_LOW_ID INTEGER," +
                    "RX_HARVEST_METHOD_LOW_CATEGORY INTEGER," +
                    "RX_HARVEST_METHOD_LOW_CATEGORY_DESCR CHAR(100)," +
                    "RX_HARVEST_METHOD_STEEP CHAR(50)," +
                    "RX_HARVEST_METHOD_STEEP_ID INTEGER," +
                    "RX_HARVEST_METHOD_STEEP_CATEGORY INTEGER," +
                    "RX_HARVEST_METHOD_STEEP_CATEGORY_DESCR CHAR(100)," +
                    "USE_RX_HARVEST_METHOD_YN CHAR(1)," +
                    "STEEP_SLOPE_PCT INTEGER," +
                    "SCENARIO_HARVEST_METHOD_LOW CHAR(50)," +
                    "SCENARIO_HARVEST_METHOD_LOW_ID INTEGER," +
                    "SCENARIO_HARVEST_METHOD_LOW_CATEGORY INTEGER," +
                    "SCENARIO_HARVEST_METHOD_LOW_CATEGORY_DESCR CHAR(100)," +
                    "SCENARIO_HARVEST_METHOD_STEEP CHAR(50)," +
                    "SCENARIO_HARVEST_METHOD_STEEP_ID INTEGER," +
                    "SCENARIO_HARVEST_METHOD_STEEP_CATEGORY INTEGER," +
                    "SCENARIO_HARVEST_METHOD_STEEP_CATEGORY_DESCR CHAR(100)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (RX))";
            }

            //
            //RXPACKAGE_REF TABLE
            //
            public void CreateRxPackageRefTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateRxPackageRefTableSQL(p_strTableName));
            }
            static public string CreateRxPackageRefTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "RXPACKAGE CHAR(3)," +
                    "DESCRIPTION CHAR(255)," +
                    "SIMYEAR1_RX CHAR(3)," +
                    "SIMYEAR1_RX_DESCRIPTION CHAR(255)," +
                    "SIMYEAR2_RX CHAR(3)," +
                    "SIMYEAR2_RX_DESCRIPTION CHAR(255)," +
                    "SIMYEAR3_RX CHAR(3)," +
                    "SIMYEAR3_RX_DESCRIPTION CHAR(255)," +
                    "SIMYEAR4_RX CHAR(3)," +
                    "SIMYEAR4_RX_DESCRIPTION CHAR(255)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (RXPACKAGE))";
            }

            //
            //DIAMETER_SPP_GRP_REF_C TABLE
            //
            public void CreateDiameterSpeciesGroupRefTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateDiameterSpeciesGroupRefTableSQL(p_strTableName));
            }

            static public string CreateDiameterSpeciesGroupRefTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "DBH_CLASS_NUM INTEGER," +
                    "DBH_RANGE_INCHES CHAR(15)," +
                    "SPP_GRP_CODE INTEGER," +
                    "SPP_GRP CHAR(50)," +
                    "TO_CHIPS CHAR(1)," +
                    "MERCH_VAL_DpCF DOUBLE," +
                    "VALUE_IF_CHIPPED_DpGT DOUBLE," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (DBH_CLASS_NUM, SPP_GRP_CODE))";
            }

            //
            //SPP_GRP_REF_C TABLE
            //
            public void CreateSpeciesGroupRefTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateSpeciesGroupRefTableSQL(p_strTableName));
            }
            static public string CreateSpeciesGroupRefTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "SPP_GRP_CD INTEGER," +
                    "COMMON_NAME CHAR(50)," +
                    "FIA_SPCD INTEGER )";
            }

            //
            //FVS_WEIGHTED_VARIABLES_REF TABLE
            //
            public void CreateFvsWeightedVariableRefTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFvsWeightedVariableRefTableSQL(p_strTableName));
            }

            static public string CreateFvsWeightedVariableRefTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "VARIABLE_NAME CHAR(40)," +
                    "VARIABLE_DESCRIPTION CHAR(255)," +
                    "BASELINE_RXPACKAGE CHAR(3)," +
                    "VARIABLE_SOURCE CHAR(100)," +
                    "weight_1_pre DOUBLE," +
                    "weight_1_post DOUBLE," +
                    "weight_2_pre DOUBLE," +
                    "weight_2_post DOUBLE," +
                    "weight_3_pre DOUBLE," +
                    "weight_3_post DOUBLE," +
                    "weight_4_pre DOUBLE," +
                    "weight_4_post DOUBLE," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (VARIABLE_NAME))";
            }
            //
            //ECON_WEIGHTED_VARIABLES_REF TABLE
            //
            public void CreateEconWeightedVariableRefTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateEconWeightedVariableRefTableSQL(p_strTableName));
            }
            static public string CreateEconWeightedVariableRefTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "VARIABLE_NAME CHAR(40)," +
                    "VARIABLE_DESCRIPTION CHAR(255)," +
                    "VARIABLE_SOURCE CHAR(100)," +
                    "CYCLE_1_WEIGHT DOUBLE," +
                    "CYCLE_2_WEIGHT DOUBLE," +
                    "CYCLE_3_WEIGHT DOUBLE," +
                    "CYCLE_4_WEIGHT DOUBLE," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (VARIABLE_NAME))";
            }
        }
        public class OptimizerScenarioRuleDefinitions
        {
            static public string DefaultScenarioFvsVariablesTieBreakerTableName { get { return "scenario_fvs_variables_tiebreaker"; } }
            static public string DefaultScenarioFvsVariablesOptimizationTableName { get { return "scenario_fvs_variables_optimization"; } }
            static public string DefaultScenarioFvsVariablesOverallEffectiveTableName { get { return "scenario_fvs_variables_overall_effective"; } }
            static public string DefaultScenarioFvsVariablesTableName { get { return "scenario_fvs_variables"; } }
            static public string DefaultScenarioLastTieBreakRankTableName { get { return "scenario_last_tiebreak_rank"; } }
            static public string DefaultScenarioPSitesTableName { get { return "scenario_psites"; } }
            static public string DefaultScenarioPlotFilterTableName { get { return "scenario_plot_filter"; } }
            static public string DefaultScenarioLandOwnerGroupsTableName { get { return "scenario_land_owner_groups"; } }
            static public string DefaultScenarioHarvestCostColumnsTableName { get { return "scenario_harvest_cost_columns"; } }
            static public string DefaultScenarioDatasourceTableName { get { return @"scenario_datasource"; } }
            static public string DefaultScenarioCostsTableName { get { return "scenario_costs"; } }
            static public string DefaultScenarioTableDbFile { get { return @"optimizer\db\scenario_optimizer_rule_definitions.db"; } }
            static public string DefaultScenarioTableName { get { return "scenario"; } }
            static public string DefaultScenarioCondFilterMiscTableName { get { return "scenario_cond_filter_misc"; } }
            static public string DefaultScenarioCondFilterTableName { get { return "scenario_cond_filter"; } }
            static public string DefaultScenarioProcessorScenarioSelectTableName { get { return "scenario_processor_scenario_select"; } }



            public OptimizerScenarioRuleDefinitions()
            {
            }
            public void CreateScenarioCostsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioCostsTable(p_strTableName));
            }
            static public string CreateScenarioCostsTable(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20) PRIMARY KEY, " +
                    "chip_mkt_val_pgt DOUBLE DEFAULT 0, " +
                    "road_haul_cost_pgt_per_hour DOUBLE DEFAULT 0, " +
                    "rail_haul_cost_pgt_per_mile DOUBLE DEFAULT 0, " +
                    "rail_chip_transfer_pgt DOUBLE DEFAULT 0, " +
                    "rail_merch_transfer_pgt DOUBLE DEFAULT 0)";
            }
            public void CreateScenarioProcessorScenarioSelectTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioProcessorScenarioSelectTableSQL(p_strTableName));
            }
            static public string CreateScenarioProcessorScenarioSelectTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20) PRIMARY KEY, " +
                    "processor_scenario_id CHAR(20), " +
                    "FullDetailsYN CHAR(1))";
            }
            static public string CreateScenarioHarvestCostColumnsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "ColumnName CHAR(50)," +
                    "Description CHAR(255))";
            }
            public void CreateScenarioHarvestCostColumnsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioHarvestCostColumnsTableSQL(p_strTableName));
                CreateScenarioHarvestCostColumnsTableIndex(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioHarvestCostColumnsTableIndex(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }
            static public string CreateScenarioLandOwnerGroupsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "owngrpcd INTEGER)";
            }
            public void CreateScenarioLandOwnerGroupsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioLandOwnerGroupsTableSQL(p_strTableName));
                CreateScenarioLandOwnerGroupsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioLandOwnerGroupsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }
            static public string CreateScenarioPlotFilterTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "sql_command VARCHAR(4000)," +
                    "current_yn CHAR(1)," +
                    "table_list CHAR(200))";
            }
            public void CreateScenarioPlotFilterTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioPlotFilterTableSQL(p_strTableName));
                CreateScenarioPlotFilterTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioPlotFilterTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }
            static public string CreateScenarioCondFilterTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "sql_command VARCHAR(4000)," +
                    "current_yn CHAR(1)," +
                    "table_list CHAR(200))";
            }
            public void CreateScenarioCondFilterTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioCondFilterTableSQL(p_strTableName));
                CreateScenarioCondFilterTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioCondFilterTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }
            static public string CreateScenarioCondFilterMiscTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + "  (" +
                    "scenario_id CHAR(20)," +
                    "yard_dist INTEGER," +
                    "yard_dist2 INTEGER)";
            }
            public void CreateScenarioCondFilterMiscTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioCondFilterMiscTableSQL(p_strTableName));
                CreateScenarioCondFilterMiscTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioCondFilterMiscTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }

            public void CreateScenarioPSitesTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioPSitesTableSQL(p_strTableName));
                CreateScenarioPSitesTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioPSitesTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }
            static public string CreateScenarioPSitesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20), " +
                    "psite_id INTEGER, " +
                    "name CHAR(100), " +
                    "trancd CHAR(1), " +
                    "biocd CHAR(1), " +
                    "selected_yn CHAR(1), " +
                    "PRIMARY KEY (scenario_id, psite_id))";
            }
            public void CreateScenarioLastTieBreakRankTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioLastTieBreakTableSQL(p_strTableName));
                CreateScenarioLastTieBreakRankTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioLastTieBreakRankTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "scenario_id");
            }
            static public string CreateScenarioLastTieBreakTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "rxpackage CHAR(3)," +
                    "last_tiebreak_rank INTEGER, " +
                    "PRIMARY KEY (scenario_id, rxpackage))";
            }
            static public string CreateScenarioFVSVariablesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "rxcycle CHAR(1)," +
                    "variable_number INTEGER," +
                    "fvs_variables_list CHAR(255)," +
                    "pre_fvs_variable CHAR(100)," +
                    "post_fvs_variable CHAR(100)," +
                    "better_expression VARCHAR(4000)," +
                    "worse_expression VARCHAR(4000)," +
                    "effective_expression VARCHAR(4000)," +
                    "current_yn CHAR(1))";
            }
            public void CreateScenarioFVSVariablesTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioFVSVariablesTableSQL(p_strTableName));
            }
            static public string CreateScenarioFVSVariablesOverallEffectiveTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "rxcycle CHAR(1)," +
                    "fvs_variables_list CHAR(255)," +
                    "overall_effective_expression VARCHAR(4000)," +
                    "nr_dpa_filter_enabled_yn CHAR(1)," +
                    "nr_dpa_filter_operator CHAR(2)," +
                    "nr_dpa_filter_value DOUBLE DEFAULT 0," +
                    "current_yn CHAR(1))";
            }
            public void CreateScenarioFVSVariableOverallEffectiveTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioFVSVariablesOverallEffectiveTableSQL(p_strTableName));
            }
            //
            //scenario rule definitions fvs variables optimization selection
            //
            static public string CreateScenarioFVSVariablesOptimizationTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "rxcycle CHAR(1)," +
                    "optimization_variable CHAR(100)," +
                    "fvs_variable_name CHAR(100)," +
                    "value_source CHAR(20)," +
                    "max_yn CHAR(1)," +
                    "min_yn CHAR(1)," +
                    "filter_enabled_yn CHAR(1)," +
                    "filter_operator CHAR(2)," +
                    "filter_value DOUBLE," +
                    "checked_yn CHAR(1)," +
                    "current_yn CHAR(1)," +
                    "revenue_attribute CHAR(100))";
            }
            public void CreateScenarioFVSVariablesOptimizationTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioFVSVariablesOptimizationTableSQL(p_strTableName));
            }
            static public string CreateScenarioFVSVariablesTieBreakerTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "rxcycle CHAR(1)," +
                    "tiebreaker_method CHAR(100)," +
                    "fvs_variable_name CHAR(100)," +
                    "value_source CHAR(20)," +
                    "max_yn CHAR(1)," +
                    "min_yn CHAR(1)," +
                    "checked_yn CHAR(1))";
            }
            public void CreateScenarioFVSVariablesTieBreakerTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioFVSVariablesTieBreakerTableSQL(p_strTableName));
            }

            public void CreateScenarioFvsVariableWeightsReferenceTable(SQLite.ADO.DataMgr oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                oDataMgr.SqlNonQuery(p_oConn, CreateScenarioFvsVariableWeightsReferenceTableSQL(p_strTableName));
            }
            static public string CreateScenarioFvsVariableWeightsReferenceTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxcycle CHAR(3)," +
                    "pre_or_post CHAR(4)," +
                    "rxyear INTEGER," +
                    "weight DOUBLE," +
                    "PRIMARY KEY (rxcycle, pre_or_post, rxyear))";
            }

        }

        public class OptimizerDefinitions
        {
            static public string DefaultCalculatedOptimizerVariablesTableName { get { return "calculated_optimizer_variables"; } }
            static public string DefaultCalculatedEconVariablesTableName { get { return "calculated_econ_variables_definition"; } }
            static public string DefaultCalculatedFVSVariablesTableName { get { return "calculated_fvs_variables_definition"; } }
            static public string DefaultOptimizerProjectConfigTableName { get { return "project_config"; } }
            static public string DefaultDbFile { get { return @"optimizer\db\optimizer_definitions.db"; } }

        }


        public class VolumeAndBiomass
        {
            static public string BiosumVolumesInputTable { get { return "biosum_volumes_input"; } }
            static public string FcsBiosumVolumesInputTable { get { return "fcs_biosum_volumes_input"; } }
            static public string BiosumCalcOutputTable { get { return "biosum_calc_output"; } }
            static public string DefaultSqliteWorkDatabase { get { return "fcs_tree.db"; } }
            static public string DefaultSqliteConfigDatabase { get { return "BiosumSpeciesConfig.db"; } }
            static public string SqliteWorkTable { get{ return "volume_and_biomass_work_table"; } }
            static public string BiosumVolumeCalcTable { get { return "BIOSUM_CALC"; } }
            static public string BioSumCompsJar { get { return "BioSumComps.jar"; } }
            static public string FcsTreeCalcBat { get { return "fcs_tree_calc.bat"; } }
            static public string FiaTreeVBCJar { get { return "FIA_TreeVBC.jar"; } }
            static public string TvbcTreeCalcBat { get { return "tvbc_tree_calc.bat"; } }
            static public string DefaultTvbcWorkDatabase { get { return "tvbc_tree_data.db"; } }
            static public string TvbcTreeDataTable { get { return "tvbc_tree_data"; } }           
            static public string TvbcTreeDataCalcTable { get { return "tvbc_tree_data_calc"; } }
            static public string ExportBiosumVolumesDatabase { get { return "TreeTroubleshooter.db"; } }
            static public string DefaultEcoSubCdRefTable { get { return "ECOSUBCD_REF"; } }

            // These are the columns on the biosum_calc table
            static public List<Tuple<string, utils.DataType>> ColumnsAndDataTypes
            {
                get
                {
                    return new List<Tuple<string, utils.DataType>>
                    {
                        Tuple.Create("STATECD", utils.DataType.INTEGER),
                        Tuple.Create("COUNTYCD", utils.DataType.INTEGER),
                        Tuple.Create("PLOT", utils.DataType.INTEGER),
                        Tuple.Create("INVYR", utils.DataType.INTEGER),
                        Tuple.Create("SUBP", utils.DataType.INTEGER),
                        Tuple.Create("TREE", utils.DataType.INTEGER),
                        Tuple.Create("VOL_LOC_GRP", utils.DataType.STRING),
                        Tuple.Create("SPCD", utils.DataType.INTEGER),
                        Tuple.Create("PRECIPITATION", utils.DataType.DOUBLE),
                        Tuple.Create("BALIVE", utils.DataType.DOUBLE),
                        Tuple.Create("SITREE", utils.DataType.INTEGER),
                        Tuple.Create("WDLDSTEM", utils.DataType.INTEGER),
                        Tuple.Create("DIAHTCD", utils.DataType.INTEGER),
                        Tuple.Create("DIA", utils.DataType.DOUBLE),
                        Tuple.Create("HT", utils.DataType.INTEGER),
                        Tuple.Create("ACTUALHT", utils.DataType.INTEGER),
                        Tuple.Create("UPPER_DIA", utils.DataType.DOUBLE),
                        Tuple.Create("UPPER_DIA_HT", utils.DataType.DOUBLE),
                        Tuple.Create("CENTROID_DIA", utils.DataType.DOUBLE),
                        Tuple.Create("CENTROID_DIA_HT_ACTUAL", utils.DataType.DOUBLE),
                        Tuple.Create("SAWHT", utils.DataType.INTEGER),
                        Tuple.Create("HTDMP", utils.DataType.DOUBLE),
                        Tuple.Create("BOLEHT", utils.DataType.INTEGER),
                        Tuple.Create("FORMCL", utils.DataType.INTEGER),
                        Tuple.Create("CR", utils.DataType.INTEGER),
                        Tuple.Create("STATUSCD", utils.DataType.INTEGER),
                        Tuple.Create("STANDING_DEAD_CD", utils.DataType.INTEGER),
                        Tuple.Create("TREECLCD", utils.DataType.INTEGER),
                        Tuple.Create("ROUGHCULL", utils.DataType.INTEGER),
                        Tuple.Create("CULL", utils.DataType.INTEGER),
                        Tuple.Create("CULLBF", utils.DataType.INTEGER),
                        Tuple.Create("CULLCF", utils.DataType.INTEGER),
                        Tuple.Create("CULL_FLD", utils.DataType.INTEGER),
                        Tuple.Create("CULLDEAD", utils.DataType.INTEGER),
                        Tuple.Create("CULLFORM", utils.DataType.INTEGER),
                        Tuple.Create("CULLMSTOP", utils.DataType.INTEGER),
                        Tuple.Create("CFSND", utils.DataType.INTEGER),
                        Tuple.Create("BFSND", utils.DataType.INTEGER),
                        Tuple.Create("DECAYCD", utils.DataType.INTEGER),
                        Tuple.Create("TOTAGE", utils.DataType.INTEGER),
                        Tuple.Create("ECODIV", utils.DataType.STRING),
                        Tuple.Create("STDORGCD", utils.DataType.INTEGER),
                        Tuple.Create("PLT_CN", utils.DataType.STRING),
                        Tuple.Create("CND_CN", utils.DataType.STRING),
                        Tuple.Create("TRE_CN", utils.DataType.STRING),
                        Tuple.Create("VOLCFGRS_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("VOLCFNET_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("VOLCFSND_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("VOLCSGRS_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("VOLTSGRS_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("DRYBIOM_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("DRYBIOT_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("DRYBIO_BOLE_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("DRYBIO_TOP_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("DRYBIO_SAPLING_CALC", utils.DataType.DOUBLE),
                        Tuple.Create("DRYBIO_WDLD_SPP_CALC", utils.DataType.DOUBLE)
                    };
                }
            }
            public void CreateBiosumVolumesOutputTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateBiosumVolumesOutputTableSQL(p_strTableName));
                CreateBiosumVolumesOutputTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateBiosumVolumesOutputTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "tre_cn");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_cond_id");
                //p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "id");
            }
            //These columns come from the TreeSample table
            public string CreateBiosumVolumesOutputTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER PRIMARY KEY," +
                    "tre_cn CHAR(34)," +
                    "biosum_cond_id CHAR(25) NOT NULL," +
                    "statecd INTEGER," +
                    "countycd INTEGER," +
                    "plot INTEGER," +
                    "subp INTEGER," +
                    "tree INTEGER," +
                    "InvYr CHAR(4)," +
                    "spcd INTEGER," +
                    "dia    DOUBLE," +
                    "ht DOUBLE," +
                    "vol_loc_grp    CHAR(10)," +
                    "actualht   DOUBLE," +
                    "statuscd   CHAR(1)," +
                    "treeclcd CHAR(1)," +
                    "cr DOUBLE," +
                    "cull   DOUBLE," +
                    "roughcull DOUBLE," +
                    "decaycd    INTEGER," +
                    "balive DOUBLE, " +
                    "centroid_dia DOUBLE, " +
                    "centroid_dia_ht_actual DOUBLE, " +
                    "cull_fld INTEGER," +
                    "cullform INTEGER," +
                    "cullmstop  INTEGER," +
                    "diahtcd INTEGER," +
                    "htdmp  DOUBLE," +
                    "sitree INTEGER," +
                    "standing_dead_cd   INTEGER," +
                    "upper_dia DOUBLE," +
                    "upper_dia_ht DOUBLE," +
                    "wdldstem INTEGER," +
                    "stdorgcd INTEGER," +
                    "ecosubcd CHAR(7)," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcfsnd DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "drybio_bole DOUBLE," +
                    "voltsgrs_bark DOUBLE," +
                    "volcfgrs_bark DOUBLE," +
                    "volcfsnd_bark DOUBLE," +
                    "volcfnet_bark DOUBLE," +
                    "volcsgrs_bark DOUBLE," +
                    "volbsnet DOUBLE," +
                    "drybio_stem DOUBLE," +
                    "drybio_stem_bark DOUBLE," +
                    "drybio_stump DOUBLE," +
                    "drybio_stump_bark DOUBLE," +
                    "drybio_bole_bark  DOUBLE," +
                    "drybio_branch DOUBLE," +
                    "drybio_foliage DOUBLE," +
                    "drybio_ag DOUBLE," +
                    "drybio_bg DOUBLE," +
                    "carbon_ag DOUBLE," +
                    "carbon_bg DOUBLE " +
                    ")";
            }
            // These are the columns on the biosum_calc table
            static public string[] TvbcVolAndBio
            {
                get
                {
                    return new string []
                    {
                        "wood_sg_adj",
                        "bark_sg_adj",
                        "volcfgrs",
                        "volcfnet",
                        "volcfsnd",
                        "volcsgrs",
                        "voltsgrs",
                        "drybio_bole",
                        "voltsgrs_bark",
                        "volcfgrs_bark",
                        "volcfsnd_bark",
                        "volcfnet_bark",
                        "volcsgrs_bark",
                        "volbsnet",
                        "drybio_stem",
                        "drybio_stem_bark",
                        "drybio_stump",
                        "drybio_stump_bark",
                        "drybio_bole_bark",
                        "drybio_branch",
                        "drybio_foliage",
                        "drybio_ag",
                        "drybio_bg",
                        "carbon_ag",
                        "carbon_bg"
                    };
                }
            }
        }

        public class FVS
        {
            public static string[] g_strFVSOutTablesArray =  {"FVS_CASES",
                                                              "FVS_SUMMARY",
                                                              "FVS_SUMMARY_EAST",
                                                              "FVS_COMPUTE",
                                                              "FVS_TREELIST",
                                                              "FVS_ATRTLIST",
                                                              "FVS_CUTLIST",
                                                              "FVS_STRCLASS",
                                                              "FVS_POTFIRE",
                                                              "FVS_POTFIRE_EAST",
                                                              "FVS_FUELS",
                                                              "FVS_BURNREPORT",
                                                              "FVS_SNAGSUM",
                                                              "FVS_CARBON",
                                                              "FVS_HRV_CARBON",
                                                              "FVS_DOWN_WOOD_COV",
                                                              "FVS_DOWN_WOOD_VOL",
                                                              "FVS_DM_STND_SUM",
                                                              "FVS_CANPROFILE",
                                                              "FVS_CUSTOM",
                                                              "FVS_ECONSUMMARY",
                                                              "FVS_CONSUMPTION"};

            static public string DefaultRxTableName { get { return "rx"; } }
            static public string DefaultRxHarvestCostColumnsTableName { get { return "rx_harvest_cost_columns"; } }
            static public string DefaultRxPackageTableName { get { return "rxpackage"; } }
            static public string DefaultRxPackageDbFile { get { return @"db\master.db"; } }
            static public string DefaultFVSTreeTableName { get { return "FVS_Tree"; } }
            static public string DefaultFVSCutTreeTableName { get { return "FVS_CutTree"; } }
            static public string DefaultFVSInForestTreeTableName { get { return "FVS_InForestTree"; } }
            static public string DefaultFVSCutTreeTvbcTableName { get { return "FVS_CutTreeTvbc"; } }
            static public string DefaultFVSTreeListDbFile { get { return @"\fvs\data\FVSOUT_TREE_LIST.db"; } }
            static public string DefaultFVSOutDbFile { get { return @"\fvs\data\FVSOut.db"; } }
            static public string DefaultFVSOutBiosumDbFile { get { return @"\fvs\data\FVSOut_BioSum.db"; } }
            static public string DefaultFVSOutPrePostDbFile { get { return @"\fvs\db\PREPOST_FVSOUT.db"; } }
            static public string DefaultFVSAuditsDbFile { get { return @"\fvs\db\FVS_AUDITS.db"; } }
            static public string DefaultFVSPrePostSeqNumTable { get { return "fvs_output_prepost_seqnum"; } }
            static public string DefaultFVSPrePostSeqNumTableDbFile { get { return @"db\master.db"; } }
            static public string DefaultFVSPrePostSeqNumRxPackageAssgnTable { get { return "fvs_output_prepost_seqnum_rxpackage_assignment"; } }
            static public string DefaultPreFVSSummaryTableName { get { return "PRE_FVS_SUMMARY"; } }

            static public string DefaultPreFVSComputeTableName { get { return "PRE_FVS_COMPUTE"; } }
            static public string DefaultPostFVSComputeTableName { get { return "POST_FVS_COMPUTE"; } }
            static public string DefaultFVSCasesTableName { get { return "FVS_CASES"; } }
            static public string DefaultFVSCasesTempTableName { get { return "FVS_CASES_TEMP"; } }
            static public string DefaultFVSSummaryTableName { get { return "FVS_SUMMARY"; } }
            static public string DefaultFVSCutListTableName { get { return "FVS_CUTLIST"; } }
            static public string DefaultFVSPotFireTableName { get { return "FVS_POTFIRE"; } }
            static public string DefaultFVSPotFireEastTableName { get { return "FVS_POTFIRE_EAST"; } }
            static public string DefaultFVSTreeListTableName { get { return "FVS_TREELIST"; } }
            static public string DefaultFVSAtrtListTableName { get { return "FVS_ATRTLIST"; } }

            public FVS()
            {
            }
            public void CreateFVSOutTreeTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSOutTreeTableSQL(p_strTableName));
                CreateFVSOutTreeTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateFVSOutTreeTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
            }
            // Note: When fields are added here, they also need to be added in CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL,
            // CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL
            // and the code that decides whether to use that SQL to rebuild the table also needs to be updated.
            public string CreateFVSOutTreeTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "estht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "statuscd INTEGER," +
                    "decaycd INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "drybio_bole double," +
                    "drybio_sapling double," +
                    "drybio_top double," +
                    "drybio_wdld_spp double," +
                    "volcfsnd double," +
                    "drybiom DOUBLE," +
                    "drybiot DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATE)";
            }
            public void CreateFVSInForestTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSInForestTableSQL(p_strTableName));
                CreateFVSInForestTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateFVSInForestTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
            }
            public string CreateFVSInForestTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "simYear INTEGER,"  +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "drybio_bole double," +
                    "drybio_sapling double," +
                    "drybio_top double," +
                    "drybio_wdld_spp double," +
                    "volcfsnd double," +
                    "drybiom DOUBLE," +
                    "drybiot DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATE)";
            }

            public void CreateFVSOutTreeTvbcTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSOutTreeTvbcTableSQL(p_strTableName));
                CreateFVSOutTreeTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }

            // Note: When fields are added here, they also need to be added in CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL,
            // CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL
            // and the code that decides whether to use that SQL to rebuild the table also needs to be updated.
            public string CreateFVSOutTreeTvbcTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "estht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "statuscd INTEGER," +
                    "decaycd INTEGER," +
                    "wood_sg_adj INTEGER," +
                    "bark_sg_adj INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "drybio_bole double," +
                    "volcfsnd double," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "voltsgrs_bark DOUBLE," +
                    "volcfgrs_bark DOUBLE," +
                    "volcfsnd_bark DOUBLE," +
                    "volcfnet_bark DOUBLE," +
                    "volcsgrs_bark DOUBLE," +
                    "volbsnet DOUBLE," +
                    "drybio_stem DOUBLE," +
                    "drybio_stem_bark DOUBLE," +
                    "drybio_stump DOUBLE," +
                    "drybio_stump_bark DOUBLE," +                    
                    "drybio_bole_bark  DOUBLE," +
                    "drybio_branch DOUBLE," +
                    "drybio_foliage DOUBLE," +
                    "drybio_ag  DOUBLE," +
                    "drybio_bg DOUBLE," +
                    "carbon_ag DOUBLE," +
                    "carbon_bg DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATE)";
            }
            public void CreateTreeListWorkTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTreeListWorkTableSQL(p_strTableName));
            }
            public string CreateTreeListWorkTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rowid INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "caseid CHAR(255)," +
                    "standid CHAR(255)," +
                    "year INTEGER," +
                    "treeid CHAR(255)," +
                    "treeindex INTEGER)";
            }

            //
            //FVS_tree table audit
            //
            public void CreateFVSTreeIdAudit(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSTreeIdAuditTableSQL(p_strTableName));
                CreateFVSTreeIdAuditTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateFVSTreeIdAuditTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
            }
            public string CreateFVSTreeIdAuditTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER PRIMARY KEY," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_tree_id CHAR(10)," +
                    "Found_FvsTreeId_YN CHAR(1) DEFAULT 'N')";
            }
            public void CreateInputBiosumVolumesTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateInputBiosumVolumesTableSQL(p_strTableName));
                CreateInputBiosumVolumesTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateInputBiosumVolumesTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "fvs_tree_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_cond_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "id");
            }
            public string CreateInputBiosumVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER," +
                    "biosum_cond_id CHAR(25)," +
                    "invyr INTEGER," +
                    "fvs_variant CHAR(2)," +
                    "spcd INTEGER," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "actualht DOUBLE," +
                    "vol_loc_grp CHAR(10)," +
                    "statuscd INTEGER," +
                    "treeclcd INTEGER," +
                    "cr DOUBLE," +
                    "cull DOUBLE," +
                    "roughcull DOUBLE," +
                    "decaycd INTEGER," +
                    "totage DOUBLE," +
                    "SUBP INTEGER," +
                    "FORMCL INTEGER," +
                    "CULLBF DOUBLE," +
                    "sitree INTEGER, " +
                    "wdldstem INTEGER," +
                    "upper_dia DOUBLE," +
                    "upper_dia_ht DOUBLE," +
                    "centroid_dia DOUBLE," +
                    "centroid_dia_ht_actual DOUBLE," +
                    "sawht DOUBLE," +
                    "htdmp DOUBLE," +
                    "boleht DOUBLE," +
                    "cullcf DOUBLE," +
                    "cull_fld DOUBLE," +
                    "culldead DOUBLE," +
                    "cullform DOUBLE," +
                    "cullmstop DOUBLE," +
                    "cfsnd DOUBLE," +
                    "bfsnd DOUBLE," +
                    "precipitation DOUBLE," +
                    "balive DOUBLE," +
                    "diahtcd INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "volcfsnd_calc DOUBLE," +
                    "drybio_bole_calc DOUBLE," +
                    "drybio_top_calc DOUBLE," +
                    "drybio_sapling_calc DOUBLE," +
                    "drybio_wdld_spp_calc DOUBLE," +
                    "ecosubcd CHAR(7)," +
                    "stdorgcd INTEGER," +
                    "volcfnet DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "drybiom DOUBLE," +
                    "drybiot DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(id))";
            }
            public void CreateTvbcInputBiosumVolumesTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTvbcInputBiosumVolumesTableSQL(p_strTableName));
                CreateInputBiosumVolumesTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }

            public string CreateTvbcInputBiosumVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "id INTEGER," +
                    "biosum_cond_id CHAR(25)," +
                    "statecd INTEGER," +
                    "countycd INTEGER," +
                    "plot INTEGER," +
                    "invyr INTEGER," +
                    //"fvs_variant CHAR(2)," +
                    "spcd INTEGER," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "actualht DOUBLE," +
                    "vol_loc_grp CHAR(10)," +
                    "statuscd INTEGER," +
                    "treeclcd INTEGER," +
                    "cr DOUBLE," +
                    "cull DOUBLE," +
                    "roughcull DOUBLE," +
                    "decaycd INTEGER," +
                    //"totage DOUBLE," +
                    //"SUBP INTEGER," +
                    //"FORMCL INTEGER," +
                    //"CULLBF DOUBLE," +
                    "sitree INTEGER, " +
                    "wdldstem INTEGER," +
                    "upper_dia DOUBLE," +
                    "upper_dia_ht DOUBLE," +
                    "centroid_dia DOUBLE," +
                    "centroid_dia_ht_actual DOUBLE," +
                    //"sawht DOUBLE," +
                    "htdmp DOUBLE," +
                    //"boleht DOUBLE," +
                    //"cullcf DOUBLE," +
                    "cull_fld DOUBLE," +
                    //"culldead DOUBLE," +
                    "cullform DOUBLE," +
                    "cullmstop DOUBLE," +
                    //"cfsnd DOUBLE," +
                    //"bfsnd DOUBLE," +
                    //"precipitation DOUBLE," +
                    "balive DOUBLE," +
                    "diahtcd INTEGER," +
                    "standing_dead_cd INTEGER," +
                    //"volcfsnd_calc DOUBLE," +
                    //"drybio_bole_calc DOUBLE," +
                    //"drybio_top_calc DOUBLE," +
                    //"drybio_sapling_calc DOUBLE," +
                    //"drybio_wdld_spp_calc DOUBLE," +
                    "ecosubcd CHAR(7)," +
                    "stdorgcd INTEGER," +
                    //"volcfnet DOUBLE," +
                    //"volcfgrs DOUBLE," +
                    //"volcsgrs DOUBLE," +
                    //"drybiom DOUBLE," +
                    //"drybiot DOUBLE," +
                    //"voltsgrs DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(id))";
            }
            public void CreateInputFCSBiosumVolumesTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateInputFCSBiosumVolumesTableSQL(p_strTableName));
            }
            public string CreateInputFCSBiosumVolumesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "STATECD INTEGER, " +
                    "COUNTYCD INTEGER, " +
                    "PLOT INTEGER, " +
                    "INVYR INTEGER, " +
                    "TREE INTEGER, " +
                    "VOL_LOC_GRP CHAR(10), " +
                    "SPCD INTEGER, " +
                    "DIA DOUBLE, " +
                    "HT DOUBLE, " +
                    "ACTUALHT DOUBLE, " +
                    "CR DOUBLE, " +
                    "STATUSCD INTEGER, " +
                    "TREECLCD INTEGER, " +
                    "ROUGHCULL DOUBLE, " +
                    "CULL DOUBLE, " +
                    "DECAYCD INTEGER, " +
                    "TOTAGE DOUBLE, " +
                    "SUBP INTEGER, " +
                    "FORMCL INTEGER, " +
                    "CULLBF DOUBLE, " +
                    "SITREE INTEGER, " +
                    "WDLDSTEM INTEGER, " +
                    "UPPER_DIA DOUBLE, " +
                    "UPPER_DIA_HT DOUBLE, " +
                    "CENTROID_DIA DOUBLE, " +
                    "CENTROID_DIA_HT_ACTUAL DOUBLE, " +
                    "SAWHT INTEGER, " +
                    "HTDMP DOUBLE, " +
                    "BOLEHT INTEGER, " +
                    "CULLCF INTEGER, " +
                    "CULL_FLD INTEGER, " +
                    "CULLDEAD INTEGER, " +
                    "CULLFORM INTEGER, " +
                    "CULLMSTOP INTEGER, " +
                    "CFSND INTEGER, " +
                    "BFSND INTEGER, " +
                    "PRECIPITATION DOUBLE, " +
                    "BALIVE DOUBLE, " +
                    "DIAHTCD INTEGER, " +
                    "STANDING_DEAD_CD INTEGER, " +
                    "VOLCFSND_CALC DOUBLE, " +
                    "DRYBIO_BOLE_CALC DOUBLE, " +
                    "DRYBIO_TOP_CALC DOUBLE, " +
                    "DRYBIO_SAPLING_CALC DOUBLE, " +
                    "DRYBIO_WDLD_SPP_CALC DOUBLE, " +
                    "STDORGCD INTEGER, " +
                    "ECODIV CHAR(7), " +
                    "TRE_CN CHAR(34) PRIMARY KEY, " +
                    "CND_CN CHAR(34), " +
                    "PLT_CN CHAR(34), " +
                    "VOLCFGRS_CALC DOUBLE, " +
                    "VOLCSGRS_CALC DOUBLE, " +
                    "VOLCFNET_CALC DOUBLE, " +
                    "DRYBIOM_CALC DOUBLE, " +
                    "DRYBIOT_CALC DOUBLE, " +
                    "VOLTSGRS_CALC DOUBLE)";
            }
            public void CreateInputFCSBiosumVolumesWorkTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateInputFCSBiosumVolumesWorkTableSQL(p_strTableName));
            }

            public string CreateInputFCSBiosumVolumesWorkTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "ID INTEGER PRIMARY KEY AUTOINCREMENT," +
                    "STATECD INTEGER," +
                    "COUNTYCD INTEGER," +
                    "PLOT INTEGER," +
                    "INVYR INTEGER," +
                    "TREE INTEGER," +
                    "VOL_LOC_GRP CHAR(10)," +
                    "SPCD INTEGER," +
                    "DIA DOUBLE," +
                    "HT DOUBLE," +
                    "ACTUALHT DOUBLE," +
                    "CR DOUBLE," +
                    "STATUSCD INTEGER," +
                    "TREECLCD INTEGER," +
                    "ROUGHCULL DOUBLE," +
                    "CULL DOUBLE," +
                    "DECAYCD INTEGER," +
                    "TOTAGE DOUBLE," +
                    "SUBP INTEGER," +
                    "FORMCL INTEGER," +
                    "CULLBF DOUBLE," +
                    "SITREE INTEGER, " +
                    "WDLDSTEM INTEGER," +
                    "UPPER_DIA DOUBLE," +
                    "UPPER_DIA_HT DOUBLE," +
                    "CENTROID_DIA DOUBLE," +
                    "CENTROID_DIA_HT_ACTUAL DOUBLE," +
                    "SAWHT INTEGER," +
                    "HTDMP DOUBLE," +
                    "BOLEHT INTEGER," +
                    "CULLCF INTEGER," +
                    "CULL_FLD INTEGER," +
                    "CULLDEAD INTEGER," +
                    "CULLFORM INTEGER," +
                    "CULLMSTOP INTEGER," +
                    "CFSND INTEGER," +
                    "BFSND INTEGER," +
                    "PRECIPITATION DOUBLE," +
                    "BALIVE DOUBLE," +
                    "DIAHTCD INTEGER," +
                    "STANDING_DEAD_CD INTEGER," +
                    "VOLCFSND_CALC DOUBLE," +
                    "DRYBIO_BOLE_CALC DOUBLE," +
                    "DRYBIO_TOP_CALC DOUBLE," +
                    "DRYBIO_SAPLING_CALC DOUBLE," +
                    "DRYBIO_WDLD_SPP_CALC DOUBLE," +
                    "ECODIV CHAR(7)," +
                    "STDORGCD INTEGER," +
                    "TRE_CN CHAR(34)," +
                    "CND_CN CHAR(34)," +
                    "PLT_CN CHAR(34)," +
                    "VOLCFGRS_CALC DOUBLE," +
                    "VOLCSGRS_CALC DOUBLE," +
                    "VOLCFNET_CALC DOUBLE," +
                    "DRYBIOM_CALC DOUBLE," +
                    "DRYBIOT_CALC DOUBLE," +
                    "VOLTSGRS_CALC DOUBLE)";
            }

            //
            //RX table
            //
            /// <summary>
            /// Create the treatment table
            /// </summary>
            /// <param name="p_oAdo"></param>
            /// <param name="p_oConn"></param>
            /// <param name="p_strTableName"></param>
            public void CreateRxTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateRxTableSQL(p_strTableName));
            }
            public string CreateRxTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rx CHAR(3)," +
                    "description CHAR(2000)," +
                    "HarvestMethodLowSlope CHAR(50)," +
                    "HarvestMethodSteepSlope CHAR(50)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(rx))";
            }

            public void CreateRxHarvestCostColumnTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateRxHarvestCostColumnTableSQL(p_strTableName));
            }
            static public string CreateRxHarvestCostColumnTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rx CHAR(3)," +
                    "ColumnName CHAR(50)," +
                    "Description CHAR(255)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (RX, COLUMNNAME))";
            }
           public void CreateRxPackageTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateRxPackageTableSQL(p_strTableName));
            }
            public string CreateRxPackageTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "rxpackage CHAR(3)," +
                    "description CHAR(2000)," +
                    "rxcycle_length CHAR(2)," +
                    "simyear1_rx CHAR(3)," +
                    "simyear1_fvscycle CHAR(2)," +
                    "simyear2_rx CHAR(3)," +
                    "simyear2_fvscycle CHAR(2)," +
                    "simyear3_rx CHAR(3)," +
                    "simyear3_fvscycle CHAR(2)," +
                    "simyear4_rx CHAR(3)," +
                    "simyear4_fvscycle CHAR(2)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(rxpackage))";
            }

            //
            //FVS Output PRE-POST Sequence Number Definitions
            //
            /// <summary>
            /// Create the table that defines the FVS Output PRE-POST Sequence Number Definitions.
            /// </summary>
            /// <param name="p_oAdo"></param>
            /// <param name="p_oConn"></param>
            /// <param name="p_strTableName"></param>
            public void CreateFVSOutputPrePostSeqNumTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumTableSQL(p_strTableName));
                CreateFVSOutputPrePostSeqNumTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateFVSOutputPrePostSeqNumTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "tablename");
            }
            public string CreateFVSOutputPrePostSeqNumTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "PREPOST_SEQNUM_ID INTEGER," +
                    "TABLENAME CHAR(75)," +
                    "TYPE CHAR(1)," +
                    "RXCYCLE1_PRE_SEQNUM INTEGER," +
                    "RXCYCLE1_POST_SEQNUM INTEGER," +
                    "RXCYCLE2_PRE_SEQNUM INTEGER," +
                    "RXCYCLE2_POST_SEQNUM INTEGER," +
                    "RXCYCLE3_PRE_SEQNUM INTEGER," +
                    "RXCYCLE3_POST_SEQNUM INTEGER," +
                    "RXCYCLE4_PRE_SEQNUM INTEGER," +
                    "RXCYCLE4_POST_SEQNUM INTEGER," +
                    "RXCYCLE1_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE2_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE3_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE4_PRE_BASEYR_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE1_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE1_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "RXCYCLE2_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE2_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "RXCYCLE3_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE3_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "RXCYCLE4_PRE_BEFORECUT_YN CHAR(1) DEFAULT 'N'," +
                    "RXCYCLE4_POST_BEFORECUT_YN CHAR(1) DEFAULT 'Y'," +
                    "USE_SUMMARY_TABLE_SEQNUM_YN CHAR(1) DEFAULT 'N'," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(PREPOST_SEQNUM_ID))";

            }
            //
            //FVS Output PRE-POST Sequence Number RX Package Assignments
            //

            public void CreateFVSOutputPrePostSeqNumRxPackageAssgnTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumRxPackageAssgnTableSQL(p_strTableName));
            }
            public string CreateFVSOutputPrePostSeqNumRxPackageAssgnTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                  "RXPACKAGE CHAR(3)," +
                  "PREPOST_SEQNUM_ID INTEGER," +
                  "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(RXPACKAGE,PREPOST_SEQNUM_ID))";
            }
            //
            //FVS Output PRE-POST Sequence Number Audit
            //
            public void CreateFVSOutputPrePostSeqNumAuditGenericTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumAuditGenericTableSQL(p_strTableName));
                CreateFVSOutputPrePostSeqNumAuditGenericTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateFVSOutputPrePostSeqNumAuditGenericTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "STANDID,[YEAR]");
            }
            public string CreateFVSOutputPrePostSeqNumAuditGenericTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                  "SEQNUM INTEGER," +
                  "STANDID CHAR(255)," +
                  "YEAR INTEGER," +
                  "CYCLE1_PRE_YN CHAR(1)," +
                  "CYCLE1_POST_YN CHAR(1)," +
                  "CYCLE2_PRE_YN CHAR(1)," +
                  "CYCLE2_POST_YN CHAR(1)," +
                  "CYCLE3_PRE_YN CHAR(1)," +
                  "CYCLE3_POST_YN CHAR(1)," +
                  "CYCLE4_PRE_YN CHAR(1)," +
                  "CYCLE4_POST_YN CHAR(1)," +
                  "RXPACKAGE CHAR(3), " +
                  "FVS_VARIANT CHAR(2), " +
                  "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY(SEQNUM, STANDID, RXPACKAGE, FVS_VARIANT))";
            }

            //
            //FVS Output StrClass PRE-POST Sequence Number Audit
            //
            public void CreateFVSOutputPrePostSeqNumAuditStrClassTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateFVSOutputPrePostSeqNumAuditStrClassTableSQL(p_strTableName));
            }
            public string CreateFVSOutputPrePostSeqNumAuditStrClassTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                  "SEQNUM INTEGER," +
                  "STANDID CHAR(255)," +
                  "[YEAR] INTEGER," +
                  "REMOVAL_CODE INTEGER," +
                  "CYCLE1_PRE_YN CHAR(1)," +
                  "CYCLE1_POST_YN CHAR(1)," +
                  "CYCLE2_PRE_YN CHAR(1)," +
                  "CYCLE2_POST_YN CHAR(1)," +
                  "CYCLE3_PRE_YN CHAR(1)," +
                  "CYCLE3_POST_YN CHAR(1)," +
                  "CYCLE4_PRE_YN CHAR(1)," +
                  "CYCLE4_POST_YN CHAR(1), " +
                  "RXPACKAGE CHAR(3), " +
                  "FVS_VARIANT CHAR(2), " +
                  "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (SEQNUM,STANDID,REMOVAL_CODE,RXPACKAGE,FVS_VARIANT))";
            }

            public static class Audit
            {
                public static class Post
                {
                    // Note: Be sure that this schema is reconciled with CreateFVSOutTreeTableSQL and CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL
                    public static string CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "COLUMN_NAME CHAR(30)," +
                            "ERROR_DESC CHAR(60)," +
                            "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "estht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "statuscd INTEGER," +
                    "decaycd INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "drybio_bole double," +
                    "drybio_sapling double," +
                    "drybio_top double," +
                    "drybio_wdld_spp double," +
                    "volcfsnd double," +
                    "drybiom DOUBLE," +
                    "drybiot DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATETIME)";

                    }
                    public static string CreateFVSPostAuditTvbcCutlistERROR_OUTPUTtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "COLUMN_NAME CHAR(30)," +
                            "ERROR_DESC CHAR(60)," +
                            "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "estht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "statuscd INTEGER," +
                    "decaycd INTEGER," +
                    "wood_sg_adj INTEGER," +
                    "bark_sg_adj INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "drybio_bole double," +
                    "volcfsnd double," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "voltsgrs_bark DOUBLE," +
                    "volcfgrs_bark DOUBLE," +
                    "volcfsnd_bark DOUBLE," +
                    "volcfnet_bark DOUBLE," +
                    "volcsgrs_bark DOUBLE," +
                    "volbsnet DOUBLE," +
                    "drybio_stem DOUBLE," +
                    "drybio_stem_bark DOUBLE," +
                    "drybio_stump DOUBLE," +
                    "drybio_stump_bark DOUBLE," +
                    "drybio_bole_bark  DOUBLE," +
                    "drybio_branch DOUBLE," +
                    "drybio_foliage DOUBLE," +
                    "drybio_ag  DOUBLE," +
                    "drybio_bg DOUBLE," +
                    "carbon_ag DOUBLE," +
                    "carbon_bg DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATE)";

                    }
                    /// <summary>
                    /// Creates the POST-FVS audit table that DETAILS items in the BIOSUM FVS_TREE table
                    /// that are not found in other tables. For example, a BIOSUM_COND_ID in the 
                    /// BIOSUM FVS_TREE table should also exist in the BIOSUM CONDITION table. If it does not
                    /// then this table will be used to document the error.
                    /// Note: Be sure that this schema is reconciled with CreateFVSOutTreeTableSQL and CreateFVSPostAuditCutlistERROR_OUTPUTtableSQL

                    /// </summary>
                    /// <param name="p_strTableName"></param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditCutlistNOTFOUND_ERRORtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "COLUMN_NAME CHAR(30)," +
                            "NOTFOUND_VALUE CHAR(50)," +
                            "ERROR_DESC CHAR(60)," +
                            "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "estht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "statuscd INTEGER," +
                    "decaycd INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "drybio_bole double," +
                    "drybio_sapling double," +
                    "drybio_top double," +
                    "drybio_wdld_spp double," +
                    "volcfsnd double," +
                    "drybiom DOUBLE," +
                    "drybiot DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATETIME)";

                    }

                    /// </summary>
                    /// <param name="p_strTableName"></param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditTvbcCutlistNOTFOUND_ERRORtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "COLUMN_NAME CHAR(30)," +
                            "NOTFOUND_VALUE CHAR(50)," +
                            "ERROR_DESC CHAR(60)," +
                            "id LONG," +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "rxyear CHAR(4)," +
                    "fvs_variant CHAR(2)," +
                    "fvs_species CHAR(6)," +
                    "tpa DOUBLE," +
                    "dbh DOUBLE," +
                    "ht DOUBLE," +
                    "estht DOUBLE," +
                    "pctcr DOUBLE," +
                    "treeval INTEGER," +
                    "mortpa DOUBLE," +
                    "mdefect INTEGER," +
                    "bapctile DOUBLE," +
                    "htg DOUBLE," +
                    "dg DOUBLE," +
                    "statuscd INTEGER," +
                    "decaycd INTEGER," +
                    "wood_sg_adj INTEGER," +
                    "bark_sg_adj INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "drybio_bole double," +
                    "volcfsnd double," +
                    "volcfgrs DOUBLE," +
                    "volcfnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "voltsgrs_bark DOUBLE," +
                    "volcfgrs_bark DOUBLE," +
                    "volcfsnd_bark DOUBLE," +
                    "volcfnet_bark DOUBLE," +
                    "volcsgrs_bark DOUBLE," +
                    "volbsnet DOUBLE," +
                    "drybio_stem DOUBLE," +
                    "drybio_stem_bark DOUBLE," +
                    "drybio_stump DOUBLE," +
                    "drybio_stump_bark DOUBLE," +
                    "drybio_bole_bark  DOUBLE," +
                    "drybio_branch DOUBLE," +
                    "drybio_foliage DOUBLE," +
                    "drybio_ag  DOUBLE," +
                    "drybio_bg DOUBLE," +
                    "carbon_ag DOUBLE," +
                    "carbon_bg DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "FvsCreatedTree_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated DATE)";

                    }
                    /// <summary>
                    ///Creates the POST-FVS audit table that SUMMARIZES the validation of the BIOSUM FVS_TREE table data 
                    /// </summary>
                    /// <param name="p_strTableName"></param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditCutlistSUMMARYtableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                          "IDX CHAR(3)," +      //Rename index to idx in case sqlite doesn't like it
                          "FVS_VARIANT CHAR(2)," +
                          "rxpackage CHAR(3)," +
                          "COLUMN_NAME CHAR(30)," +
                          "NOVALUE_ERROR CHAR(10)," +
                          "NF_IN_COND_TABLE_ERROR CHAR(10)," +
                          "NF_IN_PLOT_TABLE_ERROR CHAR(10)," +
                          "VALUE_ERROR CHAR(20)," +
                          "NF_IN_RX_TABLE_ERROR CHAR(10)," +
                          "NF_RXPACKAGE_RXCYCLE_RX_ERROR CHAR(10)," +
                          "NF_IN_RXPACKAGE_TABLE_ERROR CHAR(10)," +
                          "NF_IN_TREE_TABLE_ERROR CHAR(10)," +
                          "TREE_SPECIES_CHANGE_WARNING CHAR(10)," +
                          "DATETIMECREATED DATETIME," +
                          "PRIMARY KEY (IDX, FVS_VARIANT, RXPACKAGE, COLUMN_NAME))";
                    }
                    /// <summary>
                    ///Create the audit table used to check the tree data after appending FVS CUTLIST table data to the BIOSUM FVS_TREE table.
                    ///The purpose of the table is to contain matching FVS trees to FIA trees (by FVS_TREE_ID) to determine these items: 
                    ///1. Check if treatment cycle 1 FVS tree column data match FIA tree column data (ERROR item);
                    ///2. Check if the FVS tree species is different than the FIA tree species (WARNING item) 
                    /// </summary>
                    /// <param name="p_strTableName">Table name to create</param>
                    /// <param name="p_strDescriptionColumnName">Name of the column that will hold the warning or error description</param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditCutlistFVSFIA_TREEMATCHINGtableSQL(string p_strTableName, string p_strDescriptionColumnName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "COLUMN_NAME CHAR(30)," +
                             p_strDescriptionColumnName + " CHAR(100)," +
                            "ID LONG," +
                            "BIOSUM_COND_ID CHAR(25)," +
                            "FVS_VARIANT CHAR(2)," +
                            "RXPACKAGE CHAR(3)," +
                            "RXCYCLE CHAR(1)," +
                            "FVS_TREE_FVS_TREE_ID CHAR(10)," +
                            "FIA_TREE_FVS_TREE_ID CHAR(10)," +
                            "FVS_TREE_SPCD INTEGER," +
                            "FIA_TREE_SPCD INTEGER," +
                            "FVS_TREE_DIA  SINGLE," +
                            "FIA_TREE_DIA SINGLE," +
                            "FVS_TREE_ESTHT DOUBLE," +
                            "FIA_TREE_ESTHT DOUBLE," +
                            "FVS_TREE_ACTUALHT DOUBLE," +
                            "FIA_TREE_ACTUALHT DOUBLE," +
                            "FVS_TREE_CR DOUBLE," +
                            "FIA_TREE_CR DOUBLE," +
                            "FVS_TREE_VOLCSGRS DOUBLE," +
                            "FIA_TREE_VOLCSGRS DOUBLE," +
                            "FVS_TREE_VOLCFGRS DOUBLE," +
                            "FIA_TREE_VOLCFGRS DOUBLE," +
                            "FVS_TREE_VOLCFNET DOUBLE," +
                            "FIA_TREE_VOLCFNET DOUBLE," +
                            "FVS_TREE_VOLTSGRS DOUBLE," +
                            "FIA_TREE_VOLTSGRS DOUBLE," +
                            "FVS_TREE_DRYBIOT DOUBLE," +
                            "FIA_TREE_DRYBIOT DOUBLE," +
                            "FVS_TREE_DRYBIOM DOUBLE," +
                            "FIA_TREE_DRYBIOM DOUBLE," +
                            "FIA_TREE_STATUSCD BYTE," +
                            "FIA_TREE_TREECLCD BYTE," +
                            "FIA_TREE_CULL  DOUBLE," +
                            "FIA_TREE_ROUGHCULL DOUBLE," +
                            "FVSCREATEDTREE_YN CHAR(1) DEFAULT 'N')";


                    }

                    /// <summary>
                    ///Create the audit table used to check the tree data after appending FVS CUTLIST table data to the BIOSUM FVS_TREE table.
                    ///The purpose of the table is to contain matching FVS trees to FIA trees (by FVS_TREE_ID) to determine these items: 
                    ///1. Check if treatment cycle 1 FVS tree column data match FIA tree column data (ERROR item);
                    ///2. Check if the FVS tree species is different than the FIA tree species (WARNING item) 
                    /// </summary>
                    /// <param name="p_strTableName">Table name to create</param>
                    /// <param name="p_strDescriptionColumnName">Name of the column that will hold the warning or error description</param>
                    /// <returns></returns>
                    public static string CreateFVSPostAuditTvbcCutlistFVSFIA_TREEMATCHINGtableSQL(string p_strTableName, string p_strDescriptionColumnName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                            "COLUMN_NAME CHAR(30)," +
                             p_strDescriptionColumnName + " CHAR(100)," +
                            "ID LONG," +
                            "BIOSUM_COND_ID CHAR(25)," +
                            "FVS_VARIANT CHAR(2)," +
                            "RXPACKAGE CHAR(3)," +
                            "RXCYCLE CHAR(1)," +
                            "FVS_TREE_FVS_TREE_ID CHAR(10)," +
                            "FIA_TREE_FVS_TREE_ID CHAR(10)," +
                            "FVS_TREE_SPCD INTEGER," +
                            "FIA_TREE_SPCD INTEGER," +
                            "FVS_TREE_DIA  SINGLE," +
                            "FIA_TREE_DIA SINGLE," +
                            "FVS_TREE_ESTHT DOUBLE," +
                            "FIA_TREE_ESTHT DOUBLE," +
                            "FVS_TREE_ACTUALHT DOUBLE," +
                            "FIA_TREE_ACTUALHT DOUBLE," +
                            "FVS_TREE_CR DOUBLE," +
                            "FIA_TREE_CR DOUBLE," +
                            "FVS_TREE_VOLCSGRS DOUBLE," +
                            "FIA_TREE_VOLCSGRS DOUBLE," +
                            "FVS_TREE_VOLCFGRS DOUBLE," +
                            "FIA_TREE_VOLCFGRS DOUBLE," +
                            "FVS_TREE_VOLCFNET DOUBLE," +
                            "FIA_TREE_VOLCFNET DOUBLE," +
                            "FVS_TREE_VOLTSGRS DOUBLE," +
                            "FIA_TREE_VOLTSGRS DOUBLE," +
                            "FVS_TREE_DRYBIO_AG DOUBLE," +
                            "FIA_TREE_DRYBIOT DOUBLE," +
                            "FVS_TREE_DRYBIO_BOLE DOUBLE," +
                            "FIA_TREE_DRYBIOM DOUBLE," +
                            "FIA_TREE_STATUSCD BYTE," +
                            "FIA_TREE_TREECLCD BYTE," +
                            "FIA_TREE_CULL  DOUBLE," +
                            "FIA_TREE_ROUGHCULL DOUBLE," +
                            "FVSCREATEDTREE_YN CHAR(1) DEFAULT 'N')";


                    }
                }
                public static class Pre
                {
                    public static string CreateFVSPreYearCountsTableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                               "STANDID VARCHAR(255)," +
                                "totalrows INTEGER," +
                                "fvs_variant CHAR(2), " +
                                "rxPackage CHAR(3)," +
                                "pre_cycle1rows INTEGER," +
                                "post_cycle1rows INTEGER, " +
                                "pre_cycle2rows INTEGER," +
                                "post_cycle2rows INTEGER," +
                                "pre_cycle3rows INTEGER," +
                                "post_cycle3rows INTEGER," +
                                "pre_cycle4rows INTEGER," +
                                "post_cycle4rows INTEGER)";                    }

                    public static string CreateFVSPreAuditCountsTableSQL(string p_strTableName)
                    {
                        return "CREATE TABLE " + p_strTableName + " (" +
                                "SeqNum INTEGER, " +
                                "StandID VARCHAR(255), " +
                                "Year INTEGER, " +
                                "fvs_variant CHAR(2), " +
                                "rxPackage CHAR(3), " +
                                "PRIMARY KEY (StandID, Year, fvs_variant, rxPackage))";
                    }
                }
            }
        }
        public class TravelTime
        {
            public static string DefaultTravelTimeTableName { get { return "travel_time"; } }
            public static string DefaultProcessingSiteTableName { get { return "processing_site"; } }
            public static string DefaultMasterTravelTimeDbFile { get { return "gis_travel_times_master.db"; } }
            public static string DefaultTravelTimeDbFile { get { return "gis_travel_times.db"; } }
            public static string DefaultTravelTimePathAndDbFile { get { return @"gis\db\" + DefaultTravelTimeDbFile; } }
            public static string DefaultPlotGisTableName { get { return "plot_gis"; } }
            public static string DefaultGisAuditDbFile { get { return "gis_audit.db"; } }
            public static string DefaultGisAuditPathAndDbFile { get { return @"gis\db\" + DefaultGisAuditDbFile; } }
            public static string DefaultGisPlotDistanceAuditTable { get { return "plot_distance_audit"; } }

            public TravelTime()
            {
            }
            public void CreateProcessingSiteTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateProcessingSiteTableSQL(p_strTableName));
            }
            public string CreateProcessingSiteTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "PSITE_ID INTEGER," +
                    "PSITE_CN CHAR(12)," +
                    "NAME CHAR(100)," +
                    "TRANCD INTEGER," +
                    "TRANCD_DEF CHAR(20)," +
                    "BIOCD INTEGER," +
                    "BIOCD_DEF CHAR(15)," +
                    "EXISTS_YN CHAR(1) DEFAULT 'N'," +
                    "LAT DOUBLE," +
                    "LON DOUBLE," +
                    "STATE CHAR(2)," +
                    "CITY CHAR(40)," +
                    "MILL_TYPE CHAR(40)," +
                    "COUNTY CHAR(40)," +
                    "STATUS CHAR(40)," +
                    "NOTES CHAR(50)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (PSITE_ID))";
            }
            
            public void CreateTravelTimeTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTravelTimeTableSQL(p_strTableName));
                CreateTravelTimeTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateTravelTimeTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "PSITE_ID");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "COLLECTOR_ID");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "RAILHEAD_ID");
            }

            public string CreateTravelTimeTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "TRAVELTIME_ID INTEGER PRIMARY KEY AUTOINCREMENT," + 
                    "PSITE_ID INTEGER," +
                    "BIOSUM_PLOT_ID CHAR(24)," +
                    "COLLECTOR_ID INTEGER," +
                    "RAILHEAD_ID INTEGER," +
                    "TRAVEL_MODE CHAR(1)," +
                    "ONE_WAY_HOURS DOUBLE," +
                    "PLOT INTEGER," +
                    "STATECD INTEGER)";
            }

            public void CreatePlotDistanceAuditTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePlotDistanceAuditTableSQL(p_strTableName));
            }
            public string CreatePlotDistanceAuditTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "BIOSUM_PLOT_ID CHAR(24)," +
                    "PLOT INTEGER," +
                    "STATECD INTEGER," +
                    "GIS_YARD_DIST_FT DOUBLE," +
                    "MoveDist_ft_REPLACEMENT DOUBLE)";
            }
        }
        public class Audit
        {
            public static string DefaultCondAuditTableName { get { return "cond_audit"; } }
            public static string DefaultCondRxAuditTableName { get { return "cond_rx_audit"; } }
            public static string DefaultCondAuditTableDbFile { get { return @"audit.db"; } }

            public Audit()
            {
            }
            public void CreateCondAuditTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateCondAuditTableSQL(p_strTableName));
            }
            public string CreateCondAuditTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25) PRIMARY KEY," +
                    "fvs_prepost_variables_yn CHAR(1)," +
                    "gis_travel_times_yn CHAR(1)," +
                    "processor_tree_vol_val_yn CHAR(1), " +
                    "harvest_costs_yn CHAR(1), " +
                    "cond_too_far_steep_yn CHAR(1), " +
                    "psite_merch_yn CHAR(1), " +
                    "psite_chip_yn CHAR(1)) ";
            }
            public void CreatePlotCondRxAuditTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePlotCondRxAuditTableSQL(p_strTableName));
            }
            public string CreatePlotCondRxAuditTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "fvs_prepost_variables_yn CHAR(1)," +
                    "processor_tree_vol_val_yn CHAR(1)," +
                    "harvest_costs_yn CHAR(1)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id, rxpackage, rx, rxcycle))";
            }

        }
        public class FIAPlot
        {
            public FIAPlot()
            {
            }
            public string DefaultPlotTableSqliteDbFile { get { return @"db\master.db"; } }
            public string DefaultPlotTableName { get { return "plot"; } }
            public string DefaultConditionTableName { get { return "cond"; } }
            public string DefaultTreeTableName { get { return "tree"; } }
            public string DefaultPopEstnUnitTableName { get { return "pop_estn_unit"; } }
            public string DefaultPopEvalTableName { get { return "pop_eval"; } }
            public string DefaultPopPlotStratumAssgnTableName { get { return "pop_plot_stratum_assgn"; } }
            public string DefaultPopStratumTableName { get { return "pop_stratum"; } }
            public string DefaultBiosumPopStratumAdjustmentFactorsTableName { get { return "biosum_pop_stratum_adjustment_factors"; } }
            public string DefaultSiteTreeTableName { get { return "sitetree"; } }
            public string DefaultPopTableDbFile { get { return @"db\master.db"; } }
            public string DefaultSeedlingTableName { get { return "fiadb_seedling_input"; } }

            public string DefaultDWMCoarseWoodyDebrisName { get { return "DWM_COARSE_WOODY_DEBRIS"; } }
            public string DefaultDWMDuffLitterFuelName { get { return "DWM_DUFF_LITTER_FUEL"; } }
            public string DefaultDWMFineWoodyDebrisName { get { return "DWM_FINE_WOODY_DEBRIS"; } }
            public string DefaultDWMTransectSegmentName { get { return "DWM_TRANSECT_SEGMENT"; } }

            
            public void CreatePlotTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePlotTableSQL(p_strTableName));
                CreatePlotTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreatePlotTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "num_cond");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "biosum_status_cd");
            }
            public string CreatePlotTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_plot_id CHAR(24)," +
                    "statecd INTEGER," +
                    "invyr INTEGER," +
                    "unitcd INTEGER," +
                    "countycd INTEGER," +
                    p_strTableName + " INTEGER," +
                    "measyear INTEGER," +
                    "elev INTEGER," +
                    "fvs_variant CHAR(2)," +
                    "fvs_loc_cd INTEGER," +
                    "half_state CHAR(10)," +
                    "gis_yard_dist_ft DOUBLE," +
                    "num_cond INTEGER," +
                    "one_cond_yn CHAR(1)," +
                    "lat DOUBLE," +
                    "lon DOUBLE," +
                    "macro_breakpoint_dia INTEGER," +
                    "ecosubcd CHAR(7)," +
                    "biosum_status_cd INTEGER," +
                    "cn CHAR(34)," +
                    "RDDISTCD INTEGER," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_plot_id))";
            }

            public void CreateConditionTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateConditionTableSQL(p_strTableName));
                CreateConditionTableIndexes(p_oDataMgr, p_oConn, p_strTableName);

            }
            public void CreateConditionTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_plot_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "condid");
            }
            public string CreateConditionTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "biosum_plot_id CHAR(24)," +
                    "invyr INTEGER," +
                    "condid INTEGER," +
                    "condprop DOUBLE," +
                    "cond_status_cd INTEGER," +
                    "fortypcd INTEGER," +
                    "owncd INTEGER," +
                    "owngrpcd INTEGER," +
                    "reservcd INTEGER," +
                    "siteclcd INTEGER," +
                    "sibase INTEGER," +
                    "sicond INTEGER," +
                    "sisp INTEGER," +
                    "slope INTEGER," +
                    "aspect INTEGER," +
                    "stdage INTEGER," +
                    "stdszcd INTEGER," +
                    "habtypcd1 CHAR(10)," +
                    "adforcd INTEGER," +
                    "qmd_all_inch DOUBLE," +
                    "qmd_hwd_inch DOUBLE," +
                    "qmd_swd_inch DOUBLE," +
                    "acres DOUBLE," +
                    "unitcd INTEGER," +
                    "vol_loc_grp CHAR(10)," +
                    "tpacurr DOUBLE," +
                    "hwd_tpacurr DOUBLE," +
                    "swd_tpacurr DOUBLE," +
                    "ba_ft2_ac DOUBLE," +
                    "hwd_ba_ft2_ac DOUBLE," +
                    "swd_ba_ft2_ac DOUBLE," +
                    "vol_ac_grs_stem_ttl_ft3 DOUBLE," +
                    "hwd_vol_ac_grs_stem_ttl_ft3 DOUBLE," +
                    "swd_vol_ac_grs_stem_ttl_ft3 DOUBLE," +
                    "vol_ac_grs_ft3 DOUBLE," +
                    "hwd_vol_ac_grs_ft3 DOUBLE," +
                    "swd_vol_ac_grs_ft3 DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "hwd_volcsgrs DOUBLE," +
                    "swd_volcsgrs DOUBLE," +
                    "gsstkcd DOUBLE," +
                    "alstkcd DOUBLE," +
                    "condprop_unadj DOUBLE," +
                    "micrprop_unadj DOUBLE," +
                    "subpprop_unadj DOUBLE," +
                    "macrprop_unadj DOUBLE," +
                    "cn CHAR(34)," +
                    "biosum_status_cd INTEGER, " +
                    "dwm_fuelbed_typcd CHAR(3)," +
                    "balive DOUBLE, " +
                    "stdorgcd INTEGER," + 
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id))";
            }
            
            public void CreateTreeTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTreeTableSQL(p_strTableName));
                CreateTreeTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateTreeTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "subp");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "tree");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "condid");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx5", "fvs_tree_id");
            }
            public string CreateTreeTableSQL (string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + "(" +
                    "biosum_cond_id CHAR(25)," +
                    "invyr INTEGER," +
                    "statecd INTEGER," +
                    "unitcd INTEGER," +
                    "countycd INTEGER," +
                    "subp INTEGER," +
                    "tree INTEGER," +
                    "condid INTEGER," +
                    "statuscd INTEGER," +
                    "spcd INTEGER," +
                    "spgrpcd INTEGER," +
                    "dia DOUBLE," +
                    "diahtcd INTEGER," +
                    "ht DOUBLE," +
                    "htcd INTEGER," +
                    "actualht DOUBLE," +
                    "formcl INTEGER," +
                    "treeclcd INTEGER," +
                    "cr DOUBLE," +
                    "cclcd INTEGER DEFAULT 0," +
                    "cull DOUBLE," +
                    "roughcull DOUBLE," +
                    "decaycd INTEGER," +
                    "stocking DOUBLE," +
                    "tpacurr DOUBLE," +
                    "wdldstem INTEGER," +
                    "volcfnet DOUBLE," +
                    "volcfgrs DOUBLE," +
                    "volcsnet DOUBLE," +
                    "volcsgrs DOUBLE," +
                    "volbfnet DOUBLE," +
                    "volbfgrs DOUBLE," +
                    "voltsgrs DOUBLE," +
                    "bhage INTEGER," +
                    "cullbf DOUBLE," +
                    "cullcf DOUBLE," +
                    "totage DOUBLE," +
                    "mist_cl_cd INTEGER," +
                    "agentcd INTEGER," +
                    "damtyp1 INTEGER," +
                    "damsev1 INTEGER," +
                    "damtyp2 INTEGER," +
                    "damsev2 INTEGER," +
                    "tpa_unadj DOUBLE," +
                    "condprop_specific DOUBLE," +
                    "sitree INTEGER," +
                    "upper_dia DOUBLE," +
                    "upper_dia_ht DOUBLE," +
                    "centroid_dia DOUBLE," +
                    "centroid_dia_ht_actual DOUBLE," +
                    "sawht INTEGER," +
                    "htdmp DOUBLE," +
                    "boleht INTEGER," +
                    "cull_fld INTEGER," +
                    "culldead INTEGER," +
                    "cullform INTEGER," +
                    "cullmstop INTEGER," +
                    "cfsnd INTEGER," +
                    "bfsnd INTEGER," +
                    "standing_dead_cd INTEGER," +
                    "volcfsnd DOUBLE," +
                    "drybio_ag DOUBLE," +
                    "drybio_bole DOUBLE," +
                    "fvs_tree_id CHAR(10)," +
                    "cn CHAR(34)," +
                    "biosum_status_cd INTEGER)";
            }
            
            public void CreateSiteTreeTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateSiteTreeTableSQL(p_strTableName));
                CreateSiteTreeTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateSiteTreeTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "condid");
            }
            public string CreateSiteTreeTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + "(" +
                    "biosum_plot_id CHAR(24)," +
                    "invyr INTEGER," +
                    "condid INTEGER," +
                    "tree INTEGER," +
                    "spcd INTEGER," +
                    "dia DOUBLE," +
                    "ht DOUBLE," +
                    "agedia INTEGER," +
                    "spgrpcd INTEGER," +
                    "sitree INTEGER," +
                    "sibase INTEGER," +
                    "subp INTEGER," +
                    "method INTEGER," +
                    "validcd INTEGER," +
                    "condlist INTEGER," +
                    "biosum_status_cd INTEGER)";
            }
           
            public void CreatePopEstnUnitTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePopEstnUnitTableSQL(p_strTableName));
                CreatePopEstnUnitTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreatePopEstnUnitTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "evalid");
            }
            public string CreatePopEstnUnitTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "cn VARCHAR (34)," +
                    "eval_cn VARCHAR (34)," +
                    "rscd INTEGER," +
                    "evalid INTEGER," +
                    "estn_unit INTEGER," +
                    "estn_unit_descr VARCHAR (255)," +
                    "statecd INTEGER," +
                    "arealand_eu FLOAT," +
                    "areatot_eu FLOAT," +
                    "area_used FLOAT," +
                    "area_source VARCHAR (50)," +
                    "p1pntcnt_eu INTEGER," +
                    "p1source VARCHAR (50)," +
                    "biosum_status_cd VARCHAR (1), " +
                    "modified_date DATE," +
                    "PRIMARY KEY(RSCD, EVALID, ESTN_UNIT) )";
            }
            
            public void CreatePopEvalTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePopEvalTableSQL(p_strTableName));
                CreatePopEvalTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreatePopEvalTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "evalid");
            }
            // These data types match FIADB from which the data is loaded
            public string CreatePopEvalTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "cn VARCHAR (34)," +
                    "rscd INTEGER," +
                    "evalid INTEGER," +
                    "eval_descr VARCHAR (255)," +
                    "statecd INTEGER," +
                    "location_nm VARCHAR (255)," +
                    "report_year_nm VARCHAR (255)," +
                    "notes VARCHAR (2000)," +
                    "start_invyr INTEGER," +
                    "end_invyr INTEGER," +
                    "growth_acct VARCHAR (1)," +
                    "land_only VARCHAR (1)," +
                    "biosum_status_cd VARCHAR (1)," +
                    "MODIFIED_DATE DATE," +
                    "PRIMARY KEY(evalid, rscd))";
            }
            
            public void CreatePopPlotStratumAssgnTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePopPlotStratumAssgnTableSQL(p_strTableName));
                CreatePopPlotStratumAssgnTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreatePopPlotStratumAssgnTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "evalid");
            }
            public string CreatePopPlotStratumAssgnTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "cn VARCHAR (34)," +
                    "stratum_cn VARCHAR (34)," +
                    "plt_cn VARCHAR (34)," +
                    "statecd INTEGER," +
                    "invyr INTEGER," +
                    "unitcd INTEGER," +
                    "countycd INTEGER," +
                    "plot INTEGER," +
                    "rscd INTEGER," +
                    "evalid INTEGER," +
                    "estn_unit INTEGER," +
                    "stratumcd INTEGER," +
                    "biosum_status_cd VARCHAR (1)," +
                    "modified_date DATE, " +
                    "PRIMARY KEY(PLT_CN, STRATUM_CN, RSCD))";
            }
            
            public void CreatePopStratumTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreatePopStratumTableSQL(p_strTableName));
                CreatePopStratumTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreatePopStratumTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "evalid");
            }
            public string CreatePopStratumTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "cn VARCHAR (34)," +
                    "estn_unit_cn VARCHAR (34)," +
                    "rscd INTEGER," +
                    "evalid INTEGER," +
                    "estn_unit INTEGER," +
                    "stratumcd INTEGER," +
                    "stratum_descr VARCHAR (255)," +
                    "statecd INTEGER," +
                    "p1pointcnt INTEGER," +
                    "p2pointcnt INTEGER," +
                    "expns FLOAT," +
                    "adj_factor_macr FLOAT," +
                    "adj_factor_subp FLOAT," +
                    "adj_factor_micr FLOAT," +
                    "biosum_status_cd VARCHAR (1)," +
                    "modified_date DATE, " +
                    "PRIMARY KEY(RSCD, EVALID, ESTN_UNIT, STRATUMCD))";
            }
           
            public void CreateBiosumPopStratumAdjustmentFactorsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateBiosumPopStratumAdjustmentFactorsTableSQL(p_strTableName));
                CreateBiosumPopStratumAdjustmentFactorsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateBiosumPopStratumAdjustmentFactorsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "rscd,evalid");
            }
            public string CreateBiosumPopStratumAdjustmentFactorsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "stratum_cn CHAR(34)," +
                    "rscd DOUBLE," +
                    "evalid DOUBLE," +
                    "eval_descr CHAR(255)," +
                    "estn_unit INTEGER," +
                    "estn_unit_descr CHAR(255)," +
                    "stratumcd INTEGER," +
                    "p2pointcnt_man DOUBLE," +
                    "double_sampling INTEGER," +
                    "stratum_area DOUBLE," +
                    "expns DOUBLE," +
                    "pmh_macr DOUBLE," +
                    "pmh_sub DOUBLE," +
                    "pmh_micr DOUBLE," +
                    "pmh_cond DOUBLE," +
                    "adj_factor_macr DOUBLE," +
                    "adj_factor_subp DOUBLE," +
                    "adj_factor_micr DOUBLE," +
                    "biosum_status_cd CHAR(1))";
            }

           
            public void CreateDWMCoarseWoodyDebrisTable(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateDWMCoarseWoodyDebrisTableSQL(p_strTableName));
                CreateDWMCoarseWoodyDebrisTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }

            public void CreateDWMCoarseWoodyDebrisTableIndexes(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "plt_cn");
            }

            public string CreateDWMCoarseWoodyDebrisTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "biosum_cond_id CHAR(25)" +
                       ",biosum_status_cd INTEGER" +
                       ",CN CHAR(34) PRIMARY KEY" +
                       ",PLT_CN CHAR(34)" +
                       ",INVYR INTEGER" +
                       ",STATECD INTEGER" +
                       ",COUNTYCD INTEGER" +
                       ",PLOT INTEGER" +
                       ",SUBP INTEGER" +
                       ",TRANSECT INTEGER" +
                       ",CWDID DOUBLE" +
                       ",MEASYEAR INTEGER" +
                       ",CONDID INTEGER" +
                       ",SPCD INTEGER" +
                       ",DECAYCD INTEGER" +
                       ",TRANSDIA INTEGER" +
                       ",LENGTH INTEGER" +
                       ",CWD_SAMPLE_METHOD CHAR(6)" +
                       ",INCLINATION INTEGER" +
                       ")";
            }

           
            public void CreateDWMFineWoodyDebrisTable(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateDWMFineWoodyDebrisTableSQL(p_strTableName));
                CreateDWMFineWoodyDebrisTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateDWMFineWoodyDebrisTableIndexes(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "plt_cn");
            }

            public string CreateDWMFineWoodyDebrisTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "biosum_cond_id CHAR(25)" +
                       ",biosum_status_cd INTEGER" +
                       ",CN CHAR(34) PRIMARY KEY" +
                       ",PLT_CN CHAR(34)" +
                       ",INVYR INTEGER" +
                       ",STATECD INTEGER" +
                       ",COUNTYCD INTEGER" +
                       ",PLOT INTEGER" +
                       ",TRANSECT INTEGER" +
                       ",SUBP INTEGER" +
                       ",CONDID INTEGER" +
                       ",MEASYEAR INTEGER" +
                       ",SMALLCT INTEGER" +
                       ",MEDIUMCT INTEGER" +
                       ",LARGECT INTEGER" +
                       ",RSNCTCD INTEGER" +
                       ",SMALL_TL_COND DOUBLE" +
                       ",MEDIUM_TL_COND DOUBLE" +
                       ",LARGE_TL_COND DOUBLE" +
                       ",FWD_NONSAMPLE_REASN_CD INTEGER" +
                       ",FWD_SAMPLE_METHOD CHAR(6)" +
                       ")";
            }

           
            public void CreateDWMDuffLitterFuelTable(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateDWMDuffLitterFuelTableSQL(p_strTableName));
                CreateDWMDuffLitterFuelTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }

            public void CreateDWMDuffLitterFuelTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn,
                string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "plt_cn");
            }

            public string CreateDWMDuffLitterFuelTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "biosum_cond_id CHAR(25)" +
                       ",biosum_status_cd INTEGER" +
                       ",CN CHAR(34) PRIMARY KEY" +
                       ",PLT_CN CHAR(34)" +
                       ",INVYR INTEGER" +
                       ",STATECD INTEGER" +
                       ",COUNTYCD INTEGER" +
                       ",PLOT INTEGER" +
                       ",TRANSECT INTEGER" +
                       ",SUBP INTEGER" +
                       ",MEASYEAR INTEGER" +
                       ",CONDID INTEGER" +
                       ",DUFFDEP DOUBLE" +
                       ",LITTDEP DOUBLE" +
                       ",FUELDEP DOUBLE" +
                       ",DUFF_METHOD INTEGER" +
                       ",DUFF_NONSAMPLE_REASN_CD INTEGER" +
                       ",LITTER_METHOD INTEGER" +
                       ",LITTER_NONSAMPLE_REASN_CD INTEGER" +
                       ")";
            }

           
            public void CreateDWMTransectSegmentTable(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateDWMTransectSegmentTableSQL(p_strTableName));
                CreateDWMTransectSegmentTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }

            public void CreateDWMTransectSegmentTableIndexes(SQLite.ADO.DataMgr p_oDataMgr,
                System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "plt_cn");
            }

            public string CreateDWMTransectSegmentTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                       "biosum_cond_id CHAR(25)" +
                       ",biosum_status_cd INTEGER" +
                       ",CN CHAR(34) PRIMARY KEY" +
                       ",PLT_CN CHAR(34)" +
                       ",INVYR INTEGER" +
                       ",STATECD INTEGER" +
                       ",COUNTYCD INTEGER" +
                       ",PLOT INTEGER" +
                       ",SUBP INTEGER" +
                       ",TRANSECT INTEGER" +
                       ",SEGMNT INTEGER" +
                       ",MEASYEAR INTEGER" +
                       ",CONDID INTEGER" +
                       ",SLOPE_BEGNDIST DOUBLE" +
                       ",SLOPE_ENDDIST DOUBLE" +
                       ",SLOPE INTEGER" +
                       ",HORIZ_LENGTH DOUBLE" +
                       ",HORIZ_BEGNDIST DOUBLE" +
                       ",HORIZ_ENDDIST DOUBLE" +
                       ")";
            }
        }
        public class Processor
        {
            public Processor()
            {
            }

            public void CreateHarvestCostsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.Processor.CreateHarvestCostsTableSQL(p_strTableName));
            }
            static public string CreateHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "additional_cpa DOUBLE DEFAULT 0," +
                    "harvest_cpa DOUBLE ," +
                    "complete_cpa DOUBLE DEFAULT 0," +
                    "chip_cpa DOUBLE ," +
                    "assumed_movein_cpa DOUBLE ," +
                    "place_holder CHAR(1) DEFAULT 'N'," +
                    "override_YN CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated CHAR(22)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (biosum_cond_id,rxpackage,rx,rxcycle))";
            }
            static public string CreateAdditionalHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "biosum_cond_id TEXT," +
                    "rx TEXT," +
                    "PRIMARY KEY(biosum_cond_id,rx))";
            }

            public void CreateTreeVolValSpeciesDiamGroupsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, 
                string p_strTableName, bool p_bCreateIdColumn)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.Processor.CreateTreeVolValSpeciesDiamGroupsTableSQL(p_strTableName, p_bCreateIdColumn));
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "biosum_cond_id,rxpackage,rx,rxcycle");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx2", "rx");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx3", "species_group");
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx4", "diam_group");
            }
            static public string CreateTreeVolValSpeciesDiamGroupsTableSQL(string p_strTableName, bool p_bCreateIdColumn)
            {
                string strSQL = "CREATE TABLE " + p_strTableName + " (";
                if (p_bCreateIdColumn)
                {
                    strSQL = strSQL + "ID INTEGER PRIMARY KEY AUTOINCREMENT,";  //SQLite requires this be integer rather than long
                }
                strSQL = strSQL + "biosum_cond_id CHAR(25)," +
                    "rxpackage CHAR(3)," +
                    "rx CHAR(3)," +
                    "rxcycle CHAR(1)," +
                    "species_group INTEGER," +
                    "diam_group INTEGER DEFAULT 0," +
                    "biosum_harvest_method_category INTEGER DEFAULT 0," +
                    "chip_vol_cf DOUBLE," +
                    "chip_wt_gt DOUBLE," +
                    "chip_val_dpa DOUBLE," +
                    "chip_mkt_val_pgt DOUBLE DEFAULT 0," +
                    "merch_vol_cf DOUBLE," +
                    "merch_wt_gt DOUBLE," +
                    "merch_val_dpa DOUBLE," +
                    "merch_to_chipbin_YN CHAR(1) DEFAULT 'N'," +
                    "bc_vol_cf DOUBLE," +
                    "bc_wt_gt DOUBLE," +
                    "stand_residue_wt_gt DOUBLE," +
                    "place_holder CHAR(1) DEFAULT 'N'," +
                    "DateTimeCreated CHAR(22))";
                return strSQL;
            }

            public void CreateOpcostInputTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.Processor.CreateOpcostInputTableSQL(p_strTableName));
                // No indexes currently on OpCost input table
            }

            static public string CreateOpcostInputTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                             "(Stand text (255)," +
                             " [Percent Slope] short," +
                             " [One-way Yarding Distance] DOUBLE," +
                             " YearCostCalc INTEGER," +
                             " [Project Elevation] short," +
                             " [Harvesting System] text (50)," +
                             " [Chip tree per acre] single," +
                             " [Chip trees MerchAsPctOfTotal] single," +
                             " [Chip trees average volume(ft3)] single," +
                            " [CHIPS Average Density (lbs/ft3)] single," +
                            " [CHIPS Hwd Percent] single," +
                            " [Small log trees per acre] single," +
                            " [Small log trees MerchAsPctOfTotal] single," +
                            " [Small log trees ChipPct_Cat1_3] single," +
                            " [Small log trees ChipPct_Cat2_4] single," +
                            " [Small log trees ChipPct_Cat5] single," +
                            " [Small log trees average volume(ft3)] single," +
                            " [Small log trees average density(lbs/ft3)] single," +
                            " [Small log trees hardwood percent] single," +
                            " [Large log trees per acre] single," +
                            " [Large log trees MerchAsPctOfTotal] single," +
                            " [Large log trees ChipPct_Cat1_3_4] single," +
                            " [Large log trees ChipPct_Cat2] single," +
                            " [Large log trees ChipPct_Cat5] single," +
                            " [Large log trees average vol(ft3)] single," +
                            " [Large log trees average density(lbs/ft3)] single," +
                            " [Large log trees hardwood percent] single," +
                            " BrushCutTPA single," +
                            " BrushCutAvgVol single," +
                            " RxPackage_Rx_RxCycle text (255)," +
                            " RxCycle text (255), " +
                            " biosum_cond_id text (255)," +
                            " RxPackage text (255)," +
                            " Rx text (255)," +
                            " Move_In_Hours single," +
                            " Harvest_area_assumed_acres single," +
                            " [Unadjusted One-way Yarding distance] double," +
                            " [Unadjusted Small log trees per acre] single," +
                            " [Unadjusted Small log trees average volume (ft3)] single," +
                            " [Unadjusted Large log trees per acre] single," +
                            " [Unadjusted Large log trees average vol(ft3)] single," +
                            " ba_frac_cut single," +
                            " QMD_SL single," +
                            " QMD_LL single" +
                            " )";
            }

            public void CreateTreeReconcilationTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.Processor.CreateTreeReconcilationTableSQL(p_strTableName));
                // No indexes currently on OpCost input table
            }

            static public string CreateTreeReconcilationTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName +
                       " (cn CHAR (34)," +
                       " fvs_tree_id CHAR(10)," +
                       " biosum_cond_id CHAR (25)," +
                       " biosum_plot_id CHAR (24)," +
                       " spcd INTEGER," +
                       " merchWtGt double," +
                       " nonMerchWtGt double," +
                       " drybiom double," +
                       " drybiot double," +
                       " volCfNet double," +
                       " volCfGrs double," +
                       " volTsGrs double," +
                       " odWgt double," +
                       " dryToGreen double, " +
                       " tpa double, " +
                       " dbh double, " +
                       " isSapling CHAR(1), " +
                       " isWoodland CHAR(1), " +
                       " isCull CHAR(1), " +
                       " species_group integer, " +
                       " diam_group integer, " +
                       " merch_value double, " +
                       " opcost_type CHAR (5), " +
                       " biosum_category integer)";
            }
        }

        public class Scenario
        {
            private string strSQL = "";
            public static string DefaultScenarioTableName { get { return "scenario"; } }
            public static string DefaultScenarioDatasourceTableName { get { return "scenario_datasource"; } }

            public Scenario()
            {
            }
            
            public void CreateScenarioTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioTableSQL(p_strTableName));
            }
            public static string CreateScenarioTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20) NOT NULL PRIMARY KEY," +
                    "description VARCHAR(4000)," +
                    "path CHAR(254)," +
                    "file CHAR(50)," +
                    "notes VARCHAR(4000))";
            }

            public void CreateScenarioDatasourceTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioDatasourceTableSQL(p_strTableName));
                CreateScenarioDatasourceTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioDatasourceTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx", "scenario_id");
            }
            public static string CreateScenarioDatasourceTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "table_type CHAR(60)," +
                    "path CHAR(254)," +
                    "file CHAR(50)," +
                    "table_name CHAR(50))";
            }

        }
        public class ProcessorScenarioRun
        {
            static public string DefaultHarvestCostsTableName { get { return "harvest_costs"; } }
            static public string DefaultTreeVolValSpeciesDiamGroupsTableName { get { return "tree_vol_val_by_species_diam_groups"; } }
            static public string DefaultFiaTreeSpeciesRefTableName { get { return "FIA_TREE_SPECIES_REF"; } }
            static public string DefaultOpcostErrorsTableName { get { return @"opcost_errors"; } }
            static public string DefaultSqliteResultsDbFile { get { return @"db\scenario_results.db"; } }
            static public string DefaultAddKcpCpaTableName { get { return @"additional_kcp_cpa"; } }
            static public string DefaultScenarioResultsTableDbFile { get { return @"db\scenario_results.db"; } }

            public ProcessorScenarioRun()
            {
            }

            public void CreateTotalAdditionalHarvestCostsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn,
                string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateTotalAdditionalHarvestCostsTableSQL(p_strTableName));           
            }
            static public string CreateTotalAdditionalHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " " +
                          "(biosum_cond_id char(25)," +
                          "RX char(3)," +
                          "complete_additional_cpa double," +
                          "PRIMARY KEY(biosum_cond_id, rx))";
            }

            public void CreateAdditionalKcpCpaTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn,
                string p_strTableName, bool bIsWorktable)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateAdditionalKcpCpaTableSQL(p_strTableName, bIsWorktable));
            }
            static public string CreateAdditionalKcpCpaTableSQL(string p_strTableName, bool bIsWorktable)
            {
                string lastField = "DateTimeCreated CHAR(22),";
                if (bIsWorktable)
                {
                    lastField = "additional_cpa DOUBLE,";
                }
                string strSql = $@"CREATE TABLE {p_strTableName}
                          (biosum_cond_id CHAR(25),
                           rxPackage CHAR(3),
                           rx CHAR(3),
                           rxCycle CHAR(1),
                           {lastField} PRIMARY KEY (biosum_cond_id,rxPackage,rx,rxCycle))";
                return strSql;
            }


        }

        public class ProcessorScenarioRuleDefinitions
        {

            static public string DefaultTreeSpeciesDollarValuesTableName { get { return "scenario_tree_species_diam_dollar_values"; } }
            static public string DefaultHarvestMethodTableName { get { return "scenario_harvest_method"; } }
            static public string DefaultMoveInCostsTableName { get { return "scenario_move_in_costs"; } }
            static public string DefaultHarvestCostColumnsTableName { get { return "scenario_harvest_cost_columns"; } }
            static public string DefaultCostRevenueEscalatorsTableName { get { return "scenario_cost_revenue_escalators"; } }
            static public string DefaultAdditionalHarvestCostsDbFile { get { return @"db\scenario_processor_rule_definitions.db"; } }
            static public string DefaultAdditionalHarvestCostsTableName { get { return "scenario_additional_harvest_costs"; } }
            static public string DefaultTreeDiamGroupsTableName { get { return "scenario_tree_diam_groups"; } }
            static public string DefaultTreeSpeciesGroupsTableName { get { return "scenario_tree_species_groups"; } }
            static public string DefaultTreeSpeciesGroupsListTableName { get { return "scenario_tree_species_groups_list"; } }
            static public string DefaultDbFile { get { return @"db\scenario_processor_rule_definitions.db"; } }



            public ProcessorScenarioRuleDefinitions()
            {
            }

            public void CreateScenarioTreeSpeciesDollarValuesTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioTreeSpeciesDollarValuesTableSQL(p_strTableName));
                // This table didn't have any indexes in the MS AccessVersion
            }
            static public string CreateScenarioTreeSpeciesDollarValuesTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "species_group INTEGER," +
                    "diam_group INTEGER," +
                    "wood_bin CHAR(1) DEFAULT 'M'," +
                    "merch_value DOUBLE DEFAULT 0," +
                    "chip_value DOUBLE DEFAULT 0)";
            }
            public void CreateScenarioHarvestMethodTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioHarvestMethodTableSQL(p_strTableName));
                CreateScenarioHarvestMethodTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioHarvestMethodTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");
            }
            static public string CreateScenarioHarvestMethodTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "HarvestMethodLowSlope CHAR(50)," +
                    "HarvestMethodSteepSlope CHAR(50)," +
                    "MaxCableYardingDistance DOUBLE," +
                    "MaxHelicopterCableYardingDistance DOUBLE," +
                    "min_chip_dbh DOUBLE," +
                    "min_sm_log_dbh DOUBLE," +
                    "min_lg_log_dbh DOUBLE," +
                    "SteepSlope INTEGER," +
                    "min_dbh_steep_slope DOUBLE," +
                    "ProcessLowSlopeYN CHAR(1) DEFAULT 'Y'," +
                    "ProcessSteepSlopeYN CHAR(1) DEFAULT 'Y', " +
                    "WoodlandMerchAsPercentOfTotalVol INTEGER, " +
                    "SaplingMerchAsPercentOfTotalVol INTEGER, " +
                    "CullPctThreshold INTEGER, " +
                    "HarvestMethodSelection CHAR(15))";
            }

            public void CreateScenarioMoveInCostsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioMoveInCostsTableSQL(p_strTableName));
                CreateScenarioMoveInCostsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioMoveInCostsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");
            }
            static public string CreateScenarioMoveInCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "yard_dist_threshold DOUBLE," +
                    "assumed_harvest_area_ac DOUBLE," +
                    "move_in_time_multiplier DOUBLE," +
                    "move_in_hours_addend DOUBLE)";
            }

            public void CreateScenarioCostRevenueEscalatorsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioCostRevenueEscalatorsTableSQL(p_strTableName));
                CreateScenarioCostRevenueEscalatorsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioCostRevenueEscalatorsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
               p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");
            }
            static public string CreateScenarioCostRevenueEscalatorsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "EscalatorOperatingCosts_Cycle2 DOUBLE DEFAULT 1," +
                    "EscalatorOperatingCosts_Cycle3 DOUBLE DEFAULT 1," +
                    "EscalatorOperatingCosts_Cycle4 DOUBLE DEFAULT 1," +
                    "EscalatorMerchWoodRevenue_Cycle2 DOUBLE DEFAULT 1," +
                    "EscalatorMerchWoodRevenue_Cycle3 DOUBLE DEFAULT 1," +
                    "EscalatorMerchWoodRevenue_Cycle4 DOUBLE DEFAULT 1," +
                    "EscalatorEnergyWoodRevenue_Cycle2 DOUBLE DEFAULT 1," +
                    "EscalatorEnergyWoodRevenue_Cycle3 DOUBLE DEFAULT 1," +
                    "EscalatorEnergyWoodRevenue_Cycle4 DOUBLE DEFAULT 1)";
            }
            public void CreateScenarioHarvestCostColumnsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioHarvestCostColumnsTableSQL(p_strTableName));
                CreateScenarioHarvestCostColumnsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioHarvestCostColumnsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");
            }
            static public string CreateScenarioHarvestCostColumnsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                     "scenario_id CHAR(20)," +
                     "ColumnName CHAR(50)," +
                     "rx CHAR(3)," +
                     "Default_CPA DOUBLE," +
                     "Description CHAR(255))";
            }

            public void CreateScenarioAdditionalHarvestCostsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, Tables.ProcessorScenarioRuleDefinitions.CreateScenarioAdditionalHarvestCostsTableSQL(p_strTableName));
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_plotrx", "biosum_cond_id,rx");
            }
            static public string CreateScenarioAdditionalHarvestCostsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "biosum_cond_id CHAR(25)," +
                    "rx CHAR(3)," +
                    "CONSTRAINT " + p_strTableName + "_pk PRIMARY KEY (scenario_id,biosum_cond_id,rx))";
            }
            public void CreateScenarioTreeDiamGroupsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioTreeDiamGroupsTableSQL(p_strTableName));
                CreateScenarioTreeDiamGroupsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);

            }
            public void CreateScenarioTreeDiamGroupsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");

            }
            public string CreateScenarioTreeDiamGroupsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "scenario_id CHAR(20)," +
                    "diam_group INTEGER DEFAULT 0," +
                    "diam_class CHAR(15)," +
                    "min_diam DOUBLE DEFAULT 0," +
                    "max_diam DOUBLE DEFAULT 0, " +
                    "PRIMARY KEY(diam_group,scenario_id))";
            }
            public void CreateScenarioTreeSpeciesGroupsTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioTreeSpeciesGroupsTableSQL(p_strTableName));
                CreateScenarioTreeSpeciesGroupsTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioTreeSpeciesGroupsTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_ScenarioId", "scenario_id");

            }
            public string CreateScenarioTreeSpeciesGroupsTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "species_group INTEGER," +
                    "species_label CHAR(50)," +
                    "scenario_id CHAR(20), " +
                    "PRIMARY KEY(species_group,scenario_id))";
            }

            public string CreateScenarioTreeSpeciesGroupsListTableSQL(string p_strTableName)
            {
                return "CREATE TABLE " + p_strTableName + " (" +
                    "species_group INTEGER," +
                    "common_name CHAR(50), " +
                    "spcd INTEGER," +
                    "scenario_id CHAR(20))";
            }
            public void CreateScenarioTreeSpeciesGroupsListTable(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.SqlNonQuery(p_oConn, CreateScenarioTreeSpeciesGroupsListTableSQL(p_strTableName));
                CreateScenarioTreeSpeciesGroupsListTableIndexes(p_oDataMgr, p_oConn, p_strTableName);
            }
            public void CreateScenarioTreeSpeciesGroupsListTableIndexes(SQLite.ADO.DataMgr p_oDataMgr, System.Data.SQLite.SQLiteConnection p_oConn, string p_strTableName)
            {
                p_oDataMgr.AddIndex(p_oConn, p_strTableName, p_strTableName + "_idx1", "species_group");
            }
        }
        public class Reference
        {
            public Reference()
            {
            }
            static public string DefaultFiadbFVSVariantTableName { get { return "fiadb_fvs_variant"; } }
            static public string DefaultHarvestMethodsTableName { get { return "harvest_methods"; } }
            static public string DefaultTreeMacroPlotBreakPointDiaTableDbFile { get { return @"biosum_ref.db"; } }
            static public string DefaultTreeMacroPlotBreakPointDiaTableName { get { return "TreeMacroPlotBreakPointDia"; } }
            static public string DefaultBiosumReferenceVersionTableName { get { return "REF_VERSION"; } }
            static public string DefaultOpCostReferenceDbFile { get { return @"db\opcost_ref.db"; } }
            static public string DefaultBiosumReferenceFile { get { return "biosum_ref.db"; } }
            static public string DefaultSiteIndexEquationsTable { get { return "site_index_equations"; } }
            static public string DefaultFIATreeSpeciesTableName { get { return "FIA_TREE_SPECIES_REF"; } }
            static public string DefaultRefMasterDbFile { get { return @"db\ref_master.db"; } }
            static public string DefaultTreeSampleDbFile { get { return @"\treesample.db"; } }
            static public string DefaultTreeSampleTvbcTableName { get { return "TreeSampleTvbc"; } }
        }
        public class FIA2FVS
        {
            static public string DefaultFvsInputFile { get { return "FVSIn.db"; } }
            static public string DefaultFvsInputStandTableName { get { return "FVS_STANDINIT_COND"; } }
            static public string DefaultFvsInputTreeTableName { get { return "FVS_TREEINIT_COND"; } }
            static public string DefaultFvsInputKeywordsTableName { get { return "FVS_GROUPADDFILESANDKEYWORDS"; } }
            static public string DefaultFvsInputFolderName { get { return @"\fvs\data"; } }
            static public string KcpFileBiosumKeywords { get { return "BioSum_Keywords.kcp"; } }
            static public string KcpFileExtension { get { return ".template"; } }
        }
    }
}
